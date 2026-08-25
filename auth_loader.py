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
import sys
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
# ⚠️ 2026-08-15 现场命名：用户保留拼音文件名 (jm_dingdan_h5st.json / jm_shouhou_h5st.json)
#    与英文语义 key（jm_order / jm_after_sale）通过映射解耦
# ⚠️ 2026-08-15 用户决策：京准通 add/list 接口实测不需要 h5st（HTTP 200 不被拦截），
#    所以 H5ST_KEY_MAP 不再包含 jzt 项。如未来京准通新增业务又需要 h5st，再补回。
H5ST_KEY_MAP = {
    "jm_order": "jm_dingdan_h5st.json",          # 京麦订单明细（项目14）
    "jm_after_sale": "jm_shouhou_h5st.json",     # 京麦售后明细（项目16）
}

# 业务域 → Cookie 文件名候选列表（按优先级查找）
# ⚠️ 2026-08-15 现场命名：京麦域 cookie 拆成拼音命名 (jm_dingdan_cookie.json / jm_shouhou_cookie.json)
#    旧命名 jm_cookie.json 仍作为兜底
BIZ_TYPE_MAP = {
    "sz":  ["sz_cookie"],                                    # 商智（sz.jd.com）
    "jzt": ["jzt_cookie"],                                   # 京准通（jzt.jd.com）
    "jm":  ["jm_dingdan_cookie", "jm_shouhou_cookie", "jm_cookie"],  # 京麦（shop.jd.com）
}

# ⚠️ 2026-08-20 Phase 2.5 改造：店名前缀映射改从 config.xlsx「店铺清单」sheet 读取
# 替代原硬编码 SHOP_ID_TO_PREFIX / PREFIX_TO_SHOP_ID 字典（AGENTS.md 第3条禁止硬编码店铺列表）
# 数据源：biz_config_loader.list_shops() / get_shop_prefix()
# 双花括号 {{xxx}} 是影刀 RPA 模板语法未替换的副产品，字面保留读取
from biz_config_loader import get_shop_prefix as _cfg_get_shop_prefix
from biz_config_loader import list_shop_ids as _cfg_list_shop_ids


def resolve_file_prefix(shop_id: str) -> str:
    """shop_id → RPA 实际文件命名前缀；查不到就用 shop_id 自己包花括号兜底

    示例:
        resolve_file_prefix("FYA箱包旗舰店") → "{{FYA}}"
        resolve_file_prefix("未知店") → "{{未知店}}"（兜底不抛异常）

    数据源：config.xlsx「店铺清单」sheet（由 biz_config_loader 读取）
    兜底：sheet 缺失时 biz_config_loader 自动回退到代码兜底常量并打印警告
    """
    return _cfg_get_shop_prefix(shop_id)


def list_known_shops() -> list:
    """列出所有已知店铺（含停用的；如需只看启用店，调 biz_config_loader.list_shop_ids(True)）

    数据源：config.xlsx「店铺清单」sheet
    """
    # 默认 enabled_only=False，与旧 SHOP_ID_TO_PREFIX.keys() 语义一致（含 OTA 等停用店）
    return _cfg_list_shop_ids(enabled_only=False)


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
            shop_id    - 店铺 ID；不传则从环境变量 SHOP_ID 读，仍未设置则 SystemExit(3)
            config_dir - 配置根目录，默认项目根下的 config/
        """
        # 单例 + 允许 reload：检测参数是否变了，变了就清缓存
        if hasattr(self, "_initialized") and self._initialized:
            if shop_id == self.shop_id and config_dir == self.config_dir:
                return
        # H-13 修复（2026-08-24 审计）：禁止静默回落到 "FYA箱包旗舰店"
        # 背景：MIYO/OTA 跑业务时若 SHOP_ID 未传，会读到 FYA 鉴权 → 数据污染
        # 修复策略：仿照 H-06 风格，未设置时 SystemExit(3)，让上层（CLI/RPA/影刀）显式 set
        if shop_id:
            self.shop_id = shop_id
        else:
            env_shop_id = os.getenv("SHOP_ID", "").strip()
            if env_shop_id:
                self.shop_id = env_shop_id
            else:
                print(
                    "[FATAL] AuthLoader 初始化失败：未提供 shop_id 且环境变量 SHOP_ID 未设置\n"
                    "   → 调用方必须显式传 shop_id 或 set SHOP_ID=FYA箱包旗舰店/MIYO箱包旗舰店/OTA箱包旗舰店\n"
                    "   → 与 H-06 修复风格一致，禁止静默回落",
                    file=sys.stderr,
                )
                sys.exit(3)
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

    def _get_prefix_variants(self) -> list:
        """获取文件命名前缀的所有变体（2026-08-20 修复前缀不匹配）

        resolve_file_prefix 返回 config.xlsx 中的值（如 {{FYA}}），
        但实际 RPA 输出文件可能用 FYA（无花括号）。
        本方法返回两种变体，确保都能找到文件。

        返回:
            list - 前缀字符串列表，如 ["{{FYA}}_", "FYA_"]
        """
        file_prefix = resolve_file_prefix(self.shop_id)
        variants = [f"{file_prefix}_"]
        # 去掉双花括号后的变体（如 {{FYA}} → FYA）
        stripped = file_prefix.replace("{{", "").replace("}}", "")
        stripped_prefix = f"{stripped}_"
        if stripped_prefix not in variants:
            variants.append(stripped_prefix)
        return variants

    def _find_auth_file(self, biz_type: str, file_ext: str = "json") -> str:
        """按优先级查找鉴权文件（2026-08-20 修复：兼容 {{FYA}}_ 和 FYA_ 两种前缀）

        参数:
            biz_type - 业务类型: sz/jzt/jm
            file_ext - 文件扩展名: json/txt

        返回:
            str - 第一个存在的文件路径

        异常:
            AuthFileNotFound - 所有路径都不存在

        布局约定（用户决策 2026-08-17）：
            影刀 RPA 输出固定为 config/{短前缀}_<basename>.<ext> 平铺格式。
            2026-08-20 实测：RPA 实际输出文件名为 FYA_xxx（无花括号），
            但 config.xlsx 中 file_prefix 配置为 {{FYA}}（带花括号）。
            本方法同时尝试两种前缀，确保都能找到文件。
        """
        if biz_type not in BIZ_TYPE_MAP:
            raise ValueError(f"未知 biz_type={biz_type}，合法值: {list(BIZ_TYPE_MAP.keys())}")
        basenames = BIZ_TYPE_MAP[biz_type]
        if isinstance(basenames, str):  # 向后兼容旧 str 单值
            basenames = [basenames]

        # 2026-08-20：获取所有前缀变体（{{FYA}}_ 和 FYA_）
        prefix_variants = self._get_prefix_variants()

        # 对每个 basename × 每个前缀变体生成候选路径
        candidates = []
        for basename in basenames:
            for shop_prefix in prefix_variants:
                candidates.extend([
                    # 优先级：{前缀}{basename}.json
                    os.path.join(self.config_dir, f"{shop_prefix}{basename}.{file_ext}"),
                    # 优先级：{前缀}{basename}.txt（json→txt 兜底）
                    os.path.join(self.config_dir, f"{shop_prefix}{basename}.txt"),
                ])
            # 兜底：旧单店根目录布局（无前缀）
            candidates.extend([
                os.path.join(self.config_dir, f"{basename}.{file_ext}"),
                os.path.join(self.config_dir, f"{basename}.txt"),
            ])
        for path in candidates:
            if os.path.isfile(path):
                return path
        raise AuthFileNotFound(
            f"鉴权文件不存在：biz_type={biz_type}, file_ext={file_ext}\n"
            f"已尝试路径：\n  " + "\n  ".join(candidates) +
            f"\n请 RPA 抓取后写入 {candidates[0]}"
        )

    def _find_h5st_file(self, h5st_key: str = "jm_order") -> str:
        """查找 h5st 文件（2026-08-20 修复：兼容 {{FYA}}_ 和 FYA_ 两种前缀）

        参数:
            h5st_key - h5st 子类型（默认 jm_order）

        返回:
            str - h5st 文件路径

        异常:
            AuthFileNotFound - 文件不存在

        布局（与 _find_auth_file 一致）：
            RPA 输出为 config/{短前缀}_<h5st_filename> 平铺
            2026-08-20：同时尝试 {{FYA}}_ 和 FYA_ 两种前缀
        """
        # ⚠️ 2026-08-14 改造：3 个独立 h5st 文件（按 h5st_key 区分）
        if h5st_key in H5ST_KEY_MAP:
            primary_filename = H5ST_KEY_MAP[h5st_key]
        else:
            primary_filename = "h5st.json"

        # 2026-08-20：获取所有前缀变体（{{FYA}}_ 和 FYA_）
        prefix_variants = self._get_prefix_variants()

        candidates = []
        for shop_prefix in prefix_variants:
            candidates.extend([
                # {前缀}{primary_filename}（影刀 RPA 当前输出格式）
                os.path.join(self.config_dir, f"{shop_prefix}{primary_filename}"),
                # {前缀}h5st.json（兼容旧版同名 RPA 输出）
                os.path.join(self.config_dir, f"{shop_prefix}h5st.json"),
            ])
        # 兜底：根目录 h5st 文件（向后兼容）
        candidates.extend([
            os.path.join(self.config_dir, "h5st.json"),
            os.path.join(self.config_dir, "h5st.txt"),
        ])
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
        # H-10 修复（2026-08-24 审计）：cache_key 加 shop_id 前缀，避免多店串库
        # 背景：单例 + 5 秒缓存，原 cache_key 不含 shop_id，影刀切店后前 5 秒仍命中上店 Cookie
        cache_key = f"{self.shop_id}:cookie_str:{biz_type}:{check_expire}"
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

        绕过（2026-08-15 用户决策）：
            AGENTS.md 第 2 节说「Cookie 是否失效由接口返回码判定」。
            设置环境变量 AUTH_SKIP_COOKIE_EXPIRE_CHECK=1 可跳过本检查。
            适用场景：调试/紧急跑业务时，_gia_d 等高频轮换字段过期但 pin/light_key 等核心字段还有效。
        """
        # 跳过检查开关（环境变量 AUTH_SKIP_COOKIE_EXPIRE_CHECK=1）
        if os.getenv("AUTH_SKIP_COOKIE_EXPIRE_CHECK", "0") == "1":
            return

        # ⚠️ 2026-08-22 修复：以下字段是京东后台埋点/分析字段，业务请求不需要
        # 即使 expires 过期也不阻塞整体 Cookie 校验（AGENTS.md 第2节：接口返回码判定才作数）
        _SKIP_EXPIRE_CHECK_NAMES = {"_gia_d", "sdtoken", "__jdb", "pinId", "_jpCls"}

        now = time.time()
        expired = []
        for c in cookies:
            # sessionCookie=true 表示会话级 Cookie（关闭浏览器即失效）
            if c.get("sessionCookie"):
                continue
            # 白名单字段（埋点/分析）即使 expires 过期也跳过
            if c.get("name") in _SKIP_EXPIRE_CHECK_NAMES:
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

    def _extract_h5st_from_cdp_log(self, json_path: str) -> Dict:
        """从 CDP 抓包日志 JSON 数组中抽取 h5st

        场景（2026-08-15 实测）：
            影刀 RPA 用 [信息] 日志格式输出 CDP 抓包，每行一条 JSON 字符串。
            实际保存到 .json 文件时是 [log_line, log_line, ...] JSON 数组。
            每个 log_line 形如：
                [{'type': 'XHR', 'url': '...exportCenterService.createdExportTask',
                  'requestHeaders': {'h5st': '20260815172442388;ijn5jin75aebjn54;...;...'},
                  'headers': {'x-rp-sdtoken': 'set;1800;...'}}, ...]

        参数:
            json_path - CDP 日志 JSON 文件路径

        返回:
            dict - {
                "h5st": "20260815172442388;...;...;...;...",
                "captured_at": 1786785877997,    # 毫秒时间戳
                "source": "CDP日志 last 200-request h5st, ...ExportList",
                "api": "dsm.order.export.exportCenterService.createdExportTask"
            }

        异常:
            ValueError - CDP 日志里找不到 h5st
        """
        # ⚠️ 2026-08-17 大改：RPA 抓包文件 dict 嵌套太深 + 行被截断，
        #    ast.literal_eval 经常解析失败。改成正则扫原始文本，直接抠 h5st + url。
        with open(json_path, "r", encoding="utf-8", errors="replace") as f:
            text = f.read()

        # h5st 格式固定：13位毫秒戳;xxx;...;...（多段，分号分隔）
        # 抠出所有 (h5st_value, url) 对，按 url 内含 sff.jd.com + status=200 的优先
        import re
        # 正则：抓 'h5st': 'xxx' 字段
        h5st_pattern = re.compile(r"'h5st':\s*'([^']{40,})'", re.S)
        # 正则：抓 'url': 'xxx' 字段（注意 url 可能在截断处提前结束）
        url_pattern = re.compile(r"'url':\s*'([^']+)'", re.S)
        # 正则：抓 'status': 200 字段
        status_pattern = re.compile(r"'status':\s*200")

        candidates = []
        # 简化策略：逐个找 h5st 值，向前后各看 5000 字符找最近的 url + status
        for m in h5st_pattern.finditer(text):
            h5st_value = m.group(1)
            # ⚠️ 2026-08-17 修复：向前 + 向后各看 5000 字符（之前只向前看，h5st 在文件前面会找不到）
            ctx_start = max(0, m.start() - 5000)
            ctx_end = min(len(text), m.end() + 5000)
            ctx = text[ctx_start:ctx_end]
            url_match = None
            for um in url_pattern.finditer(ctx):
                url_match = um  # 取最后一个
            if not url_match:
                continue
            url = url_match.group(1)
            status_ok = bool(status_pattern.search(ctx))
            if not status_ok:
                continue
            # ⚠️ 三级匹配（按优先级）
            if "sff.jd.com" in url and "exportCenterService" in url:
                candidates.append({"h5st": h5st_value, "url": url, "priority": 3, "pos": m.start()})
            elif "sff.jd.com" in url:
                candidates.append({"h5st": h5st_value, "url": url, "priority": 2, "pos": m.start()})
            else:
                candidates.append({"h5st": h5st_value, "url": url, "priority": 1, "pos": m.start()})

        if not candidates:
            raise ValueError(
                f"CDP 日志里没找到 status=200 且带 h5st 的请求: {json_path}"
            )

        # 2. 优先选最后一个（最新），并按业务类型筛
        #    exportCenterService.createdExportTask = 创建导出任务（最严）
        #    queryNeedRollBackAndPermission / queryExportTaskInfo = 查询（较宽）
        #    取业务对应：订单 → "createdExportTask" / 售后 → "createExportTask"
        best = candidates[-1]
        # ⚠️ 2026-08-17 按优先级选最佳：priority 越大越好（最严匹配）
        candidates_sorted = sorted(candidates, key=lambda c: (c.get("priority", 0), c.get("pos", 0)))
        best = candidates_sorted[-1]
        # 如果有 createdExportTask/createExportTask，优先选它（即使不是 priority 最高）
        for c in reversed(candidates):
            if "createdExportTask" in c["url"] or "createExportTask" in c["url"]:
                best = c
                break

        # 3. 抽取 x-rp-sdtoken（解析 `;1800;XXX` 第二段 → 30 分钟 ttl 参考）
        # ⚠️ 2026-08-17 大改：从原始文本正则抠 x-rp-sdtoken（不一定有，缺失用默认 1800s）
        import re as _re_local
        # 在 best["pos"] 附近往后查 'x-rp-sdtoken': 'set;1800;xxx'
        sdtoken = ""
        sdtoken_pattern = _re_local.compile(r"'x-rp-sdtoken':\s*'(set;\d+;[^']*)'", _re_local.S)
        sdtoken_m = sdtoken_pattern.search(text, best.get("pos", 0))
        if sdtoken_m:
            sdtoken = sdtoken_m.group(1)
        if sdtoken.startswith("set;"):
            parts = sdtoken.split(";")
            if len(parts) >= 2:
                try:
                    ttl = int(parts[1])
                except ValueError:
                    ttl = 1800
            else:
                ttl = 1800
        else:
            ttl = 1800

        # 4. captured_at 用文件 mtime（CDP 日志通常在抓取时 mtime）
        file_mtime_ms = int(os.path.getmtime(json_path) * 1000)

        # 5. 从 h5st 字符串里取时间戳（第 7 段，如 "20260815172442388"）
        #    h5st 格式：yyyyMMddHHmmssSSS;uuid;...;ttl;...;timestamp
        try:
            h5st_parts = best["h5st"].split(";")
            # 真实抓包实测：第 7 段是 13 位毫秒时间戳（如 1786785877997）
            for part in h5st_parts:
                if part.isdigit() and len(part) == 13:
                    captured_at = int(part)
                    break
            else:
                captured_at = file_mtime_ms
        except Exception:
            captured_at = file_mtime_ms

        # 6. API 名称（用于日志）
        api_name = "unknown"
        if "createdExportTask" in best["url"]:
            api_name = "dsm.order.export.exportCenterService.createdExportTask"
        elif "createExportTask" in best["url"]:
            api_name = "dsm.seller.afs.bff.ExportDsmService.createExportTask"
        elif "queryExportTaskInfo" in best["url"]:
            api_name = "dsm.order.export.exportCenterService.queryExportTaskInfo"

        return {
            "h5st": best["h5st"],
            "captured_at": captured_at,
            "source": f"CDP日志 last 200-request h5st, api={api_name}",
            "api": api_name,
            "ttl_seconds": ttl,
            "file_mtime_ms": file_mtime_ms,
        }

    def _extract_h5st_from_pinyin_cdp(self, json_path: str) -> Dict:
        """拼音命名 (jm_dingdan / jm_shouhou) CDP 日志专用入口

        ⚠️ 2026-08-15 真实文件格式（影刀 RPA 输出）：
            每行结构: "[信息] [2026-08-15 17:24:42.212] [{'type': 'XHR', 'url': '...', ...}]"
            其中 [...] 是 Python repr 格式（单引号），不是合法 JSON。

        处理：
            1. 用 ast.literal_eval 安全解析单引号 dict 列表
            2. 累加所有合法数组，去重
            3. 复用 _extract_h5st_from_cdp_log 抽 h5st
        """
        import ast
        # 处理 BOM
        with open(json_path, "rb") as f:
            raw = f.read()
        if raw.startswith(b"\xef\xbb\xbf"):
            raw = raw[3:]
        text = raw.decode("utf-8", errors="replace")

        # ⚠️ 2026-08-17 修复：影刀 RPA 写入文件时每行被截断到 ~10 万字符，
        #    不能按行扫描。改成扫描整文本里的 `[{...}]` 完整 dict 配对。
        #    使用 ast.parse 把整个文件当一个表达式解析，逐个提取 dict 字面量。
        data = []

        # 思路：定位所有 [{ 和 }] 配对，逐段提取 ast.literal_eval
        # 因为单引号 dict 可能跨"行内换行"被截断，要从 [{ 开始到下一个完整 }] 结束
        i = 0
        n = len(text)
        parse_attempts = 0
        while i < n - 2:
            # 找下一个 [{
            if text[i] == '[' and i + 1 < n and text[i + 1] == '{':
                # 找匹配的 ]（从后往前找最近的 }]）
                # 启发式：dict 闭合后必是 ], 或 ]) 或 ] 之类
                end = text.find('}]', i + 2)
                if end < 0:
                    # 没找到完整配对，可能文件被严重截断
                    i += 2
                    continue
                # 尝试解析 [i, end+2) 这个片段
                expr = text[i:end + 2]
                parse_attempts += 1
                try:
                    obj = ast.literal_eval(expr)
                    if isinstance(obj, list):
                        data.extend(obj)
                    elif isinstance(obj, dict):
                        data.append(obj)
                    i = end + 2  # 跳到 ] 之后继续找下一个 [{
                    continue
                except (ValueError, SyntaxError):
                    # 这段可能被截断，向后挪1字符继续
                    i += 1
                    continue
            else:
                i += 1

        # ⚠️ 兜底：如果上面完全没匹配到（极少见），尝试按"行"扫描（向后兼容）
        if not data:
            for line in text.splitlines():
                line = line.strip()
                if not line:
                    continue
                start = line.find('[{')
                if start < 0:
                    continue
                if not line.endswith(']'):
                    continue
                expr = line[start:]
                try:
                    obj = ast.literal_eval(expr)
                    if isinstance(obj, list):
                        data.extend(obj)
                    elif isinstance(obj, dict):
                        data.append(obj)
                except (ValueError, SyntaxError):
                    continue

        if not data:
            raise ValueError(
                f"无法解析 CDP 日志（{json_path}）："
                f"未找到任何 [{...}] 数组行。文件前 200 字符：\n{text[:200]!r}"
            )

        # ⚠️ 2026-08-17 大改：ast 解析经常失败（dict 嵌套太深），
        #    改成直接调 _extract_h5st_from_cdp_log 扫原始 text（该函数已改成正则扫文本）
        try:
            result = self._extract_h5st_from_cdp_log(json_path)
            result["source"] = result["source"] + f" | file={os.path.basename(json_path)}"
            return result
        except ValueError:
            # 正则也没找到 → 抛出更详细错误（提示 ast 解析失败也无济于事）
            raise
        # 旧的 tmp_path 复用方式已废弃（ast 解析不能保证数据完整，会丢 h5st）

    # ====================== h5st 读取 ======================

    def get_h5st(self, check_expire: bool = True, h5st_key: str = "jm_order") -> str:
        """获取 h5st 字符串

        参数:
            check_expire - 是否检查 30 分钟过期（默认 True）
            h5st_key - h5st 子类型（默认 jm_order）：
                "jm_order"      → 京麦订单明细（项目14）
                "jm_after_sale" → 京麦售后明细（项目16）

        返回:
            str - h5st 字符串

        异常:
            AuthFileNotFound - 文件不存在
            H5stExpiredError - h5st 过期

        关键（2026-08-14 实测）：
            不同业务页面的 h5st 不能跨业务复用！必须传正确的 h5st_key。
            错误使用售后页 h5st 跑订单明细 → 服务端返回 code=1001 未登录

        ⚠️ 2026-08-15 修订：京准通业务不在此 API 范围内。
        京准通 add/list 接口实测不需要 h5st（HTTP 200 不被拦截），
        项目 7 章节的"h5st 必需"结论已修正。详见 SKILL.md 京准通分区。

        ⚠️ 2026-08-20：环境变量 AUTH_SKIP_H5ST_EXPIRE_CHECK=1 可跳过过期检查
        （与 AUTH_SKIP_COOKIE_EXPIRE_CHECK 同理：由接口返回码判定是否失效）
        """
        if h5st_key not in H5ST_KEY_MAP:
            raise ValueError(f"未知 h5st_key={h5st_key!r}，合法值: {list(H5ST_KEY_MAP.keys())}")
        # 2026-08-20：跳过 h5st 过期检查（与 Cookie 同理，由接口返回码判定）
        if check_expire and os.getenv("AUTH_SKIP_H5ST_EXPIRE_CHECK", "0") == "1":
            check_expire = False
        # H-10 修复（2026-08-24 审计）：cache_key 加 shop_id 前缀，避免多店串库
        cache_key = f"{self.shop_id}:h5st:{check_expire}:{h5st_key}"  # ⚠️ 加 h5st_key 避免缓存串
        if self._is_cache_valid(cache_key):
            return self._cache[cache_key]

        # 1. 尝试读 JSON（带 captured_at 字段）
        try:
            json_path = self._find_h5st_file(h5st_key)
            if json_path.endswith(".json"):
                try:
                    with open(json_path, "r", encoding="utf-8") as f:
                        data = json.load(f)
                    h5st_value = data.get("h5st", "").strip() if isinstance(data, dict) else ""
                except json.JSONDecodeError:
                    # ⚠️ 2026-08-15 实测：拼音命名文件存的是 CDP 抓包日志
                    #    真实文件格式：每行 "[信息] [时间] [{...}, ...]"（Python repr 单引号）
                    #    json.load() 会失败，需要走 _extract_h5st_from_pinyin_cdp
                    self.logger.info(
                        f"🔍 {json_path} 不是干净 JSON，尝试按 CDP 抓包日志格式解析"
                    )
                    cdp_result = self._extract_h5st_from_pinyin_cdp(json_path)
                    h5st_value = cdp_result["h5st"]
                    captured_at_cdp = cdp_result["captured_at"]
                    self.logger.info(
                        f"✅ 从 CDP 日志抽取 h5st：{json_path} | {cdp_result['api']} | ttl={cdp_result['ttl_seconds']}s"
                    )
                    # 过期检查（统一毫秒）
                    if check_expire and captured_at_cdp > 0:
                        age_seconds = (time.time() * 1000 - captured_at_cdp) / 1000
                        if age_seconds > H5ST_EXPIRE_SECONDS:
                            self._try_rpa_refresh(reason=f"h5st 过期 {age_seconds:.0f}秒 > {H5ST_EXPIRE_SECONDS}秒")
                            raise H5stExpiredError(
                                f"❌ {json_path} 抽出的 h5st 已过期 {age_seconds/60:.1f} 分钟 > 30 分钟\n"
                                f"   captured_at: {datetime.fromtimestamp(captured_at_cdp/1000).isoformat()}\n"
                                f"   → 请 RPA 重新抓取（影刀任务：{DEFAULT_RPA_TASK}）"
                            )
                    self._set_cache(cache_key, h5st_value)
                    return h5st_value
                if not h5st_value:
                    # data 是合法 dict 但没 h5st 字段，且不是 dict 类型 → 走 CDP 分支
                    if isinstance(data, list):
                        # 顶层是 list（可能是 CDP 抓包数组）
                        tmp_path = json_path + ".tmp"
                        with open(tmp_path, "w", encoding="utf-8") as f:
                            json.dump(data, f, ensure_ascii=False)
                        try:
                            cdp_result = self._extract_h5st_from_cdp_log(tmp_path)
                        finally:
                            try:
                                os.remove(tmp_path)
                            except OSError as e:
                                # M-22 修复（2026-08-24 审计）：临时文件清理失败加 DEBUG 日志
                                # 背景：Windows 上文件被占用/权限不够时 os.remove 经常失败，
                                #       静默吞掉导致孤儿 .tmp 文件累积，磁盘满/磁盘IO受影响
                                self.logger.debug(
                                    f"⚠️ 临时文件清理失败（可能文件被占用）：{tmp_path} - "
                                    f"{type(e).__name__}: {e}"
                                )
                        h5st_value = cdp_result["h5st"]
                        captured_at_cdp = cdp_result["captured_at"]
                        self.logger.info(
                            f"✅ 从 CDP 日志（list）抽取 h5st：{json_path} | {cdp_result['api']} | ttl={cdp_result['ttl_seconds']}s"
                        )
                        if check_expire and captured_at_cdp > 0:
                            age_seconds = (time.time() * 1000 - captured_at_cdp) / 1000
                            if age_seconds > H5ST_EXPIRE_SECONDS:
                                self._try_rpa_refresh(reason=f"h5st 过期 {age_seconds:.0f}秒 > {H5ST_EXPIRE_SECONDS}秒")
                                raise H5stExpiredError(
                                    f"❌ {json_path} 抽出的 h5st 已过期 {age_seconds/60:.1f} 分钟 > 30 分钟\n"
                                    f"   captured_at: {datetime.fromtimestamp(captured_at_cdp/1000).isoformat()}\n"
                                    f"   → 请 RPA 重新抓取（影刀任务：{DEFAULT_RPA_TASK}）"
                                )
                        self._set_cache(cache_key, h5st_value)
                        return h5st_value
                    raise ValueError(f"JSON 文件 {json_path} 不含 h5st 字段（且不是 CDP 日志格式）")
                if check_expire:
                    captured_at = data.get("captured_at", 0)
                    if captured_at > 0:
                        # H-25 修复（2026-08-24 测试发现）：captured_at 可能是毫秒(13位)或秒(10位)
                        # auth_writer.py 写入 13 位毫秒时间戳；旧代码按秒算 → age 恒负 → 过期检查永不触发
                        captured_at_s = captured_at / 1000 if captured_at > 1e12 else captured_at
                        age = time.time() - captured_at_s
                        if age > H5ST_EXPIRE_SECONDS:
                            self._try_rpa_refresh(reason=f"h5st 过期 {age:.0f}秒 > {H5ST_EXPIRE_SECONDS}秒")
                            raise H5stExpiredError(
                                f"❌ {json_path} 的 h5st 已过期 {age/60:.1f} 分钟 > 30 分钟\n"
                                f"   captured_at: {datetime.fromtimestamp(captured_at_s).isoformat()}\n"
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
        # H-21 修复（2026-08-24 审计）：Cookie 脱敏打印，仅 DEBUG_AUTH=1 时显示前 80 字符
        print(f"\n✅ Cookie 长度：{len(cookie)} 字节（已脱敏）")
        print(f"   字段数：{len(cookie.split(';'))} 个")
        if os.environ.get("DEBUG_AUTH") == "1":
            print(f"   [DEBUG] cookie[:80]={cookie[:80]!r}")
    except (CookieExpiredError, AuthFileNotFound) as e:
        print(f"\n❌ {e}")
    try:
        h5st = auth.get_h5st()
        age = auth.get_h5st_age_seconds()
        # H-22 修复（2026-08-24 审计）：h5st 脱敏打印（首 6 + 尾 6），避免 stdout/日志泄漏
        h5st_masked = f"{h5st[:6]}***{h5st[-6:]}" if len(h5st) > 12 else "***"
        print(f"\n✅ h5st（首尾脱敏）：{h5st_masked}（共 {len(h5st)} 字符）")
        print(f"   距捕获：{age:.0f} 秒（{age/60:.1f} 分钟）")
    except (H5stExpiredError, AuthFileNotFound) as e:
        print(f"\n❌ {e}")


if __name__ == "__main__":
    _cli()
