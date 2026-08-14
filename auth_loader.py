# -*- coding: utf-8 -*-
"""多店铺鉴权加载器（2026-08-13 上线）

职责（用户决策 2026-08-13）：
    1. 从 config/{shop_id}/{jm,jzt,sz}_cookie.json 读 Cookie（浏览器 DevTools 导出格式）
    2. 从 config/{shop_id}/h5st.json 读 h5st（含 captured_at 时间戳）
    3. 自动检查 Cookie 过期（expires 字段）+ h5st 过期（30 分钟）
    4. 转换 Cookie JSON 格式 → "name=val; name=val" 字符串（向后兼容 requests 库）
    5. 为影刀 RPA 预留事件驱动接口：Cookie/h5st 过期时自动调 yingdao.exe 重抓

设计原则：
    - 不影响现有 .txt 文件路径（向后兼容，main.py 仍可读 config/{sz,jzt,jm}_cookie.txt）
    - 优先读 JSON（如果 config/{shop_id}/*.json 存在），否则 fallback 到 .txt
    - 环境变量 SHOP_ID 选店铺（默认 "FYA箱包旗舰店"）
    - 环境变量 AUTH_RPA_CLI 可关闭 RPA 调用（设 "0" 跳过自动重抓）
    - 所有过期检查抛标准异常，main.py 现有 CookieExpiredError 不用改

使用示例（main.py 改造后）：
    from auth_loader import AuthLoader
    auth = AuthLoader()
    cookie_str = auth.get_cookie_str(biz_type="sz")    # → "pin=FYA8888; thor=...; "
    h5st = auth.get_h5st()                             # → "20260813..."
    if auth.is_h5st_expired():
        raise H5stExpiredError("h5st 过期，请 RPA 重抓")

异常：
    CookieExpiredError - Cookie 过期（继承 main.py 的同名异常，可直接 raise）
    H5stExpiredError   - h5st 过期（30 分钟）
    AuthFileNotFound   - 文件不存在
"""
import os
import json
import time
import logging
import subprocess
from datetime import datetime
from typing import Optional, Dict, List

# 配置目录
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_DIR = os.path.join(BASE_DIR, "config")

# h5st 有效期（30 分钟，配合抓包时限，参考 .trae/skills/jd-api-analyze/SKILL.md）
H5ST_EXPIRE_SECONDS = 30 * 60

# h5st 子类型文件映射（2026-08-14 实测：不同业务页面的 h5st 不能跨业务复用）
H5ST_KEY_MAP = {
    "jm_order": "h5st_jm_order.json",          # 京麦订单明细（项目14）
    "jm_after_sale": "h5st_jm_after_sale.json", # 京麦售后明细（项目16）
    "jzt": "h5st_jzt.json",                    # 京准通（项目1）
}

# 业务域 → 业务类型 / Cookie 文件名映射
BIZ_TYPE_MAP = {
    "sz": "sz_cookie",  # 商智（sz.jd.com）
    "jzt": "jzt_cookie",  # 京准通（jzt.jd.com）
    "jm": "jm_cookie",  # 京麦（shop.jd.com）
}

# RPA CLI 默认占位命令（用户需替换为影刀实际可执行文件路径）
# 例：C:/Program Files/Yingdao/yingdao.exe
DEFAULT_RPA_CLI = "yingdao.exe"
DEFAULT_RPA_TASK = "jd_refresh"  # 影刀里"重抓鉴权"任务名

# ============================ 异常定义 ============================

class AuthError(Exception):
    """鉴权相关异常基类"""
    pass


class AuthFileNotFound(AuthError):
    """鉴权文件不存在"""
    pass


class CookieExpiredError(AuthError):
    """Cookie 过期（按 expires 字段判断）"""
    pass


class H5stExpiredError(AuthError):
    """h5st 过期（30 分钟）"""
    pass


# ============================ 主类 ============================

class AuthLoader:
    """多店铺鉴权加载器（单例模式）

    配置文件查找顺序（按优先级）：
        1. config/{shop_id}/{biz}_cookie.json      ← RPA 推荐（带 expires）
        2. config/{shop_id}/{biz}_cookie.txt       ← 兼容旧版
        3. config/{biz}_cookie.json                ← 兼容演示目录
        4. config/{biz}_cookie.txt                 ← 兼容根目录（最旧）
    """

    _instance = None

    def __new__(cls, *args, **kwargs):
        """单例模式（同一进程只读一次文件，避免 IO 抖动）"""
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self, shop_id: Optional[str] = None, config_dir: Optional[str] = None):
        """初始化

        参数:
            shop_id    - 店铺 ID，默认从环境变量 SHOP_ID 读，否则用 "FYA箱包旗舰店"
            config_dir - 配置根目录，默认项目根下的 config/
        """
        # 单例 + 允许 reload：检测参数是否变了，变了就清缓存
        if hasattr(self, "_initialized") and self._initialized:
            if shop_id == self.shop_id and config_dir == self.config_dir:
                return
        self.shop_id = shop_id or os.getenv("SHOP_ID", "FYA箱包旗舰店")
        self.config_dir = config_dir or CONFIG_DIR
        self._initialized = True
        # 缓存（避免重复读文件，5 秒内复用）
        self._cache: Dict[str, dict] = {}
        self._cache_ts: Dict[str, float] = {}
        self._cache_ttl = 5  # 秒
        self._setup_logger()

    def _setup_logger(self):
        """初始化 logger"""
        self.logger = logging.getLogger(f"AuthLoader[{self.shop_id}]")
        if not self.logger.handlers:
            handler = logging.StreamHandler()
            handler.setFormatter(logging.Formatter(
                "[%(asctime)s] %(name)s %(levelname)s: %(message)s",
                datefmt="%H:%M:%S"
            ))
            self.logger.addHandler(handler)
            self.logger.setLevel(logging.INFO)

    # ====================== 文件查找 ======================

    def _find_auth_file(self, biz_type: str, file_ext: str = "json") -> str:
        """按优先级查找鉴权文件

        参数:
            biz_type - 业务类型: sz/jzt/jm
            file_ext - 文件扩展名: json/txt

        返回:
            str - 第一个存在的文件路径

        异常:
            AuthFileNotFound - 所有路径都不存在
        """
        if biz_type not in BIZ_TYPE_MAP:
            raise ValueError(f"未知 biz_type={biz_type}，合法值: {list(BIZ_TYPE_MAP.keys())}")
        basename = BIZ_TYPE_MAP[biz_type]

        # 4 级候选路径
        candidates = [
            os.path.join(self.config_dir, self.shop_id, f"{basename}.{file_ext}"),
            os.path.join(self.config_dir, self.shop_id, f"{basename}.txt" if file_ext == "json" else f"{basename}.{file_ext}"),
            os.path.join(self.config_dir, f"{basename}.{file_ext}"),
            os.path.join(self.config_dir, f"{basename}.txt" if file_ext == "json" else f"{basename}.{file_ext}"),
        ]
        for path in candidates:
            if os.path.isfile(path):
                return path
        raise AuthFileNotFound(
            f"鉴权文件不存在：biz_type={biz_type}, file_ext={file_ext}\n"
            f"已尝试路径：\n  " + "\n  ".join(candidates) +
            f"\n请 RPA 抓取后写入 {candidates[0]}"
        )

    def _find_h5st_file(self, h5st_key: str = "jm_order") -> str:
        """查找 h5st 文件（json 优先，txt fallback）

        参数:
            h5st_key - h5st 子类型（默认 jm_order）

        返回:
            str - h5st 文件路径

        异常:
            AuthFileNotFound - 文件不存在
        """
        # ⚠️ 2026-08-14 改造：3 个独立 h5st 文件（按 h5st_key 区分）
        # 兼容旧版 h5st.json / h5st.txt（无 h5st_key 时 fallback 到旧文件）
        if h5st_key in H5ST_KEY_MAP:
            primary_filename = H5ST_KEY_MAP[h5st_key]
        else:
            primary_filename = "h5st.json"

        candidates = [
            # 优先级 1：新版独立 h5st 文件
            os.path.join(self.config_dir, self.shop_id, primary_filename),
            # 优先级 2：旧版兼容 h5st.json
            os.path.join(self.config_dir, self.shop_id, "h5st.json"),
            # 优先级 3：根目录 h5st.json（向后兼容）
            os.path.join(self.config_dir, "h5st.json"),
            os.path.join(self.config_dir, "h5st.txt"),
        ]
        for path in candidates:
            if os.path.isfile(path):
                return path
        raise AuthFileNotFound(
            f"h5st 文件不存在\n已尝试路径：\n  " + "\n  ".join(candidates) +
            f"\n请 RPA 抓取后写入 {candidates[0]}"
        )

    # ====================== Cookie 读取 ======================

    def get_cookie_str(self, biz_type: str, check_expire: bool = True) -> str:
        """获取 Cookie 字符串（HTTP 请求头格式）

        参数:
            biz_type    - 业务类型: sz/jzt/jm
            check_expire - 是否检查过期（默认 True）

        返回:
            str - "name1=val1; name2=val2; " 格式

        异常:
            AuthFileNotFound - 文件不存在
            CookieExpiredError - Cookie 过期（如果 check_expire=True）
        """
        # 优先读 JSON（新格式）
        cache_key = f"cookie_str:{biz_type}:{check_expire}"
        if self._is_cache_valid(cache_key):
            return self._cache[cache_key]

        # 1. 尝试读 JSON
        try:
            json_path = self._find_auth_file(biz_type, "json")
            with open(json_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            cookies = data.get("cookies", [])
            if not cookies:
                raise ValueError(f"JSON 文件 {json_path} 不含 cookies 字段")
            cookie_str = "; ".join(
                f"{c['name']}={c['value']}" for c in cookies if c.get("name")
            ) + "; "
            # 检查 expires
            if check_expire:
                self._check_cookie_expires(cookies, source=json_path)
            self._set_cache(cache_key, cookie_str)
            self.logger.info(f"✅ 从 JSON 读取 Cookie：{json_path}（{len(cookies)} 个字段）")
            return cookie_str
        except AuthFileNotFound:
            pass  # JSON 不存在，尝试 txt

        # 2. Fallback 读 txt（旧格式）
        try:
            txt_path = self._find_auth_file(biz_type, "txt")
            with open(txt_path, "r", encoding="utf-8") as f:
                cookie_str = f.read().strip()
            if not cookie_str:
                raise ValueError(f"TXT 文件 {txt_path} 为空")
            # TXT 没有 expires 字段，无法精确判断过期
            if check_expire:
                self.logger.warning(
                    f"⚠️ {txt_path} 是旧 txt 格式，无 expires 字段，无法精确判断过期。"
                    f"建议 RPA 重抓后存为 json 格式。"
                )
            self._set_cache(cache_key, cookie_str)
            self.logger.info(f"✅ 从 TXT 读取 Cookie：{txt_path}（{len(cookie_str)} 字节）")
            return cookie_str
        except AuthFileNotFound:
            raise AuthFileNotFound(
                f"找不到 {biz_type} 类型的 Cookie 文件（json 和 txt 都没有）\n"
                f"店铺：{self.shop_id}\n"
                f"配置目录：{self.config_dir}"
            )

    def _check_cookie_expires(self, cookies: List[dict], source: str):
        """检查 Cookie 列表中是否有过期项

        参数:
            cookies - 浏览器导出的 cookies[] 列表
            source  - 数据来源路径（用于错误信息）

        异常:
            CookieExpiredError - 至少一个关键 Cookie 已过期
        """
        now = time.time()
        expired = []
        for c in cookies:
            # sessionCookie=true 表示会话级 Cookie（关闭浏览器即失效）
            if c.get("sessionCookie"):
                continue
            exp = c.get("expires")
            if exp and exp > 0 and exp < now:
                expired.append((c.get("name"), exp))
        if expired:
            names = ", ".join(n for n, _ in expired)
            self._try_rpa_refresh(reason=f"Cookie 过期: {names}")
            raise CookieExpiredError(
                f"❌ {source} 中 {len(expired)} 个 Cookie 已过期：{names}\n"
                f"   → 请 RPA 重新抓取（影刀任务：{DEFAULT_RPA_TASK}）"
            )

    def is_cookie_expired(self, biz_type: str) -> bool:
        """判断 Cookie 是否过期（不抛异常版本）

        返回:
            bool - True=过期，False=有效
        """
        try:
            self.get_cookie_str(biz_type, check_expire=True)
            return False
        except CookieExpiredError:
            return True

    # ====================== h5st 读取 ======================

    def get_h5st(self, check_expire: bool = True, h5st_key: str = "jm_order") -> str:
        """获取 h5st 字符串

        参数:
            check_expire - 是否检查 30 分钟过期（默认 True）
            h5st_key - h5st 子类型（默认 jm_order）：
                "jm_order"      → 京麦订单明细（项目14）
                "jm_after_sale" → 京麦售后明细（项目16）
                "jzt"           → 京准通（项目1）

        返回:
            str - h5st 字符串

        异常:
            AuthFileNotFound - 文件不存在
            H5stExpiredError - h5st 过期

        关键（2026-08-14 实测）：
            不同业务页面的 h5st 不能跨业务复用！必须传正确的 h5st_key。
            错误使用售后页 h5st 跑订单明细 → 服务端返回 code=1001 未登录
        """
        if h5st_key not in H5ST_KEY_MAP:
            raise ValueError(f"未知 h5st_key={h5st_key!r}，合法值: {list(H5ST_KEY_MAP.keys())}")
        cache_key = f"h5st:{check_expire}:{h5st_key}"  # ⚠️ 加 h5st_key 避免缓存串
        if self._is_cache_valid(cache_key):
            return self._cache[cache_key]

        # 1. 尝试读 JSON（带 captured_at 字段）
        try:
            json_path = self._find_h5st_file(h5st_key)
            if json_path.endswith(".json"):
                with open(json_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                h5st_value = data.get("h5st", "").strip()
                if not h5st_value:
                    raise ValueError(f"JSON 文件 {json_path} 不含 h5st 字段")
                if check_expire:
                    captured_at = data.get("captured_at", 0)
                    if captured_at > 0:
                        age = time.time() - captured_at
                        if age > H5ST_EXPIRE_SECONDS:
                            self._try_rpa_refresh(reason=f"h5st 过期 {age:.0f}秒 > {H5ST_EXPIRE_SECONDS}秒")
                            raise H5stExpiredError(
                                f"❌ {json_path} 的 h5st 已过期 {age/60:.1f} 分钟 > 30 分钟\n"
                                f"   captured_at: {datetime.fromtimestamp(captured_at/1000).isoformat() if captured_at > 1e12 else datetime.fromtimestamp(captured_at).isoformat()}\n"
                                f"   → 请 RPA 重新抓取（影刀任务：{DEFAULT_RPA_TASK}）"
                            )
                self._set_cache(cache_key, h5st_value)
                self.logger.info(f"✅ 从 JSON 读取 h5st：{json_path}")
                return h5st_value
            # 2. TXT 文件（旧格式，无 captured_at）
            else:
                with open(json_path, "r", encoding="utf-8") as f:
                    h5st_value = f.read().strip()
                if not h5st_value:
                    raise ValueError(f"TXT 文件 {json_path} 为空")
                if check_expire:
                    # TXT 不知道抓取时间，只能用文件 mtime 推算
                    mtime = os.path.getmtime(json_path)
                    age = time.time() - mtime
                    if age > H5ST_EXPIRE_SECONDS:
                        self._try_rpa_refresh(reason=f"h5st.txt mtime {age:.0f}秒 > {H5ST_EXPIRE_SECONDS}秒")
                        raise H5stExpiredError(
                            f"❌ {json_path} 已 {age/60:.1f} 分钟未更新（mtime），疑似 h5st 过期\n"
                            f"   → 请 RPA 重新抓取，或升级为 JSON 格式带 captured_at 字段"
                        )
                self._set_cache(cache_key, h5st_value)
                self.logger.warning(
                    f"⚠️ 从 TXT 读取 h5st（无 captured_at，靠文件 mtime 推算）：{json_path}"
                )
                return h5st_value
        except AuthFileNotFound:
            raise

    def is_h5st_expired(self, h5st_key: str = "jm_order") -> bool:
        """判断 h5st 是否过期（不抛异常版本）"""
        try:
            self.get_h5st(check_expire=True, h5st_key=h5st_key)
            return False
        except H5stExpiredError:
            return True

    def get_h5st_age_seconds(self, h5st_key: str = "jm_order") -> Optional[float]:
        """获取 h5st 已捕获的秒数（用于日志/UI）

        参数:
            h5st_key - h5st 子类型（默认 jm_order）

        返回:
            float - 距捕获的秒数；文件不存在返回 None
        """
        try:
            json_path = self._find_h5st_file(h5st_key)
            if json_path.endswith(".json"):
                with open(json_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                captured_at = data.get("captured_at", 0)
                if captured_at > 1e12:  # 毫秒时间戳
                    return (time.time() * 1000 - captured_at) / 1000
                elif captured_at > 0:
                    return time.time() - captured_at
            else:
                mtime = os.path.getmtime(json_path)
                return time.time() - mtime
        except AuthFileNotFound:
            return None

    # ====================== 一次性获取 ======================

    def get_credentials(self, biz_type: str) -> Dict[str, str]:
        """一次性获取 Cookie + h5st（电商业务最常用）

        参数:
            biz_type - 业务类型: sz/jzt/jm

        返回:
            dict - {
                "shop_id": "FYA箱包旗舰店",
                "biz_type": "sz",
                "cookie_str": "pin=FYA8888; thor=...; ",
                "h5st": "20260813..."  # 可能为 ""（cookie 域名下不需要 h5st）
            }
        """
        return {
            "shop_id": self.shop_id,
            "biz_type": biz_type,
            "cookie_str": self.get_cookie_str(biz_type),
            "h5st": self._try_get_h5st_safe(),
        }

    def _try_get_h5st_safe(self, h5st_key: str = "jm_order") -> str:
        """尝试获取 h5st，失败返回空字符串（不抛异常）"""
        try:
            return self.get_h5st(check_expire=True, h5st_key=h5st_key)
        except (H5stExpiredError, AuthFileNotFound):
            return ""

    # ====================== RPA 接口 ======================

    def _try_rpa_refresh(self, reason: str = ""):
        """尝试调 RPA 重抓鉴权（事件驱动，2026-08-13 用户决策）

        参数:
            reason - 触发原因（用于日志）

        行为:
            - 检查环境变量 AUTH_RPA_CLI：设 "0" 跳过；未设用 DEFAULT_RPA_CLI
            - 调用 subprocess.Popen（非阻塞），失败仅 warning 不抛异常
        """
        # 用户决策：可关闭自动 RPA 调用
        rpa_cli = os.getenv("AUTH_RPA_CLI", DEFAULT_RPA_CLI)
        if rpa_cli == "0" or rpa_cli.lower() == "false":
            self.logger.info(f"[RPA] 已禁用（环境变量 AUTH_RPA_CLI={rpa_cli}）：{reason}")
            return
        rpa_task = os.getenv("AUTH_RPA_TASK", DEFAULT_RPA_TASK)
        self.logger.warning(
            f"🔄 [RPA] 触发影刀重抓任务：{rpa_cli} {rpa_task}（{reason}）"
        )
        try:
            subprocess.Popen(
                [rpa_cli, rpa_task],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=0x00000008 if os.name == "nt" else 0,  # DETACHED_PROCESS
            )
            self.logger.info(f"✅ [RPA] 已派发任务：{rpa_task}（异步执行）")
        except FileNotFoundError:
            self.logger.warning(
                f"⚠️ [RPA] 未找到 {rpa_cli}，跳过自动重抓。\n"
                f"   请手动执行：{rpa_cli} {rpa_task}\n"
                f"   或设置环境变量 AUTH_RPA_CLI=0 禁用自动触发"
            )
        except Exception as e:
            self.logger.warning(f"⚠️ [RPA] 派发失败：{e}（不影响主流程）")

    # ====================== 缓存 ======================

    def _is_cache_valid(self, key: str) -> bool:
        if key not in self._cache:
            return False
        if time.time() - self._cache_ts.get(key, 0) > self._cache_ttl:
            return False
        return True

    def _set_cache(self, key: str, value):
        self._cache[key] = value
        self._cache_ts[key] = time.time()

    def clear_cache(self):
        """清空缓存（调试用，或 RPA 写完文件后强制重读）"""
        self._cache.clear()
        self._cache_ts.clear()
        self.logger.debug("缓存已清空")


# ============================ CLI 入口 ============================

def _cli():
    """CLI 调试入口：python auth_loader.py [shop_id] [biz_type]"""
    import sys
    shop_id = sys.argv[1] if len(sys.argv) > 1 else "FYA箱包旗舰店"
    biz_type = sys.argv[2] if len(sys.argv) > 2 else "sz"
    auth = AuthLoader(shop_id=shop_id)
    print(f"\n=== AuthLoader 调试（shop_id={shop_id}, biz_type={biz_type}）===")
    try:
        cookie = auth.get_cookie_str(biz_type)
        print(f"\n✅ Cookie 字符串（前 80 字符）：{cookie[:80]}...")
        print(f"   长度：{len(cookie)} 字节")
    except (CookieExpiredError, AuthFileNotFound) as e:
        print(f"\n❌ {e}")
    try:
        h5st = auth.get_h5st()
        age = auth.get_h5st_age_seconds()
        print(f"\n✅ h5st（前 30 字符）：{h5st[:30]}...")
        print(f"   距捕获：{age:.0f} 秒（{age/60:.1f} 分钟）")
    except (H5stExpiredError, AuthFileNotFound) as e:
        print(f"\n❌ {e}")


if __name__ == "__main__":
    _cli()
