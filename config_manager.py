# -*- coding: utf-8 -*-
"""多店铺配置管理器（2026-08-12 启动）

职责：
    1. SQLite 存店铺配置（shop_id ↔ Cookie 路径映射）
    2. 自动扫描 config/ 目录加载所有店铺
    3. 提供 h5st/cookie 智能读取（带过期检查）
    4. 为后续 RPA 写入预留接口

⚠️ 用户决策 2026-08-12：使用 SQLite（轻量、零部署）
⚠️ 集成方式：作为可选模块，不影响现有项目1-16 的运行

目录结构（推荐）：
    config/
    ├── _global/
    │   └── h5st_shared.txt         # 跨店铺共享的h5st（可选）
    ├── FYA旗舰店/
    │   ├── sz_cookie.txt
    │   ├── jzt_cookie.txt
    │   ├── jm_cookie.txt
    │   └── h5st.txt
    ├── FYA箱包旗舰店/
    │   └── ...
    └── shops.db                    # SQLite 数据库
"""
import os
import sqlite3
import time
from datetime import datetime, timedelta
from typing import Optional, Dict, List

# 配置目录（与 main.py 同级）
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_DIR = os.path.join(BASE_DIR, "config")
DB_PATH = os.path.join(CONFIG_DIR, "shops.db")

# h5st 有效期（30分钟，配合抓包时限）
H5ST_EXPIRE_MINUTES = 30


class ConfigManager:
    """多店铺配置管理器（单例）"""

    _instance = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._init_db()
        return cls._instance

    # -------------------- 数据库初始化 --------------------
    def _init_db(self):
        """初始化 SQLite + 建表（幂等）"""
        os.makedirs(CONFIG_DIR, exist_ok=True)
        self.conn = sqlite3.connect(DB_PATH, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        cur = self.conn.cursor()

        # 店铺表
        cur.execute("""
            CREATE TABLE IF NOT EXISTS shops (
                shop_id          TEXT PRIMARY KEY,
                pin              TEXT NOT NULL,
                shop_name        TEXT,
                sz_cookie_path   TEXT,
                jzt_cookie_path  TEXT,
                jm_cookie_path   TEXT,
                h5st_path        TEXT,
                created_at       TEXT DEFAULT (datetime('now', 'localtime')),
                updated_at       TEXT DEFAULT (datetime('now', 'localtime'))
            )
        """)

        # h5st 抓取记录
        cur.execute("""
            CREATE TABLE IF NOT EXISTS h5st_log (
                shop_id          TEXT PRIMARY KEY,
                h5st_value       TEXT NOT NULL,
                captured_at      TEXT DEFAULT (datetime('now', 'localtime')),
                expires_at       TEXT
            )
        """)

        # 导出历史（可选，未来用）
        cur.execute("""
            CREATE TABLE IF NOT EXISTS export_history (
                id               INTEGER PRIMARY KEY AUTOINCREMENT,
                shop_id          TEXT NOT NULL,
                biz_key          TEXT NOT NULL,
                task_id          TEXT,
                status           TEXT,
                file_path        TEXT,
                row_count        INTEGER,
                created_at       TEXT DEFAULT (datetime('now', 'localtime'))
            )
        """)

        self.conn.commit()

    # -------------------- 店铺管理 --------------------
    def register_shop(self, shop_id: str, pin: str, shop_name: str = "") -> bool:
        """注册店铺（如果不存在则插入，存在则跳过）"""
        # 自动推断文件路径
        shop_dir = os.path.join(CONFIG_DIR, shop_id)
        os.makedirs(shop_dir, exist_ok=True)

        sz_cookie = os.path.join(shop_dir, "sz_cookie.txt") if os.path.isfile(os.path.join(shop_dir, "sz_cookie.txt")) else None
        jzt_cookie = os.path.join(shop_dir, "jzt_cookie.txt") if os.path.isfile(os.path.join(shop_dir, "jzt_cookie.txt")) else None
        jm_cookie = os.path.join(shop_dir, "jm_cookie.txt") if os.path.isfile(os.path.join(shop_dir, "jm_cookie.txt")) else None
        h5st_path = os.path.join(shop_dir, "h5st.txt") if os.path.isfile(os.path.join(shop_dir, "h5st.txt")) else None

        cur = self.conn.cursor()
        cur.execute("""
            INSERT OR IGNORE INTO shops
                (shop_id, pin, shop_name, sz_cookie_path, jzt_cookie_path, jm_cookie_path, h5st_path)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (shop_id, pin, shop_name or shop_id, sz_cookie, jzt_cookie, jm_cookie, h5st_path))
        self.conn.commit()
        return cur.rowcount > 0

    def list_shops(self) -> List[Dict]:
        """列出所有店铺"""
        cur = self.conn.cursor()
        cur.execute("SELECT * FROM shops ORDER BY shop_id")
        return [dict(row) for row in cur.fetchall()]

    def get_shop(self, shop_id: str) -> Optional[Dict]:
        """按 shop_id 查询"""
        cur = self.conn.cursor()
        cur.execute("SELECT * FROM shops WHERE shop_id = ?", (shop_id,))
        row = cur.fetchone()
        return dict(row) if row else None

    # -------------------- 自动扫描 config/ 目录 --------------------
    def scan_config_dir(self) -> List[str]:
        """扫描 config/ 目录，自动发现店铺子目录

        规则：
            - config/FYA箱包旗舰店/  → shop_id="FYA箱包旗舰店"
            - 目录内必须有 sz_cookie.txt / jzt_cookie.txt / jm_cookie.txt 至少一个
        """
        if not os.path.isdir(CONFIG_DIR):
            return []

        discovered = []
        for entry in os.listdir(CONFIG_DIR):
            full_path = os.path.join(CONFIG_DIR, entry)
            # 跳过文件和非店铺目录
            if not os.path.isdir(full_path) or entry.startswith("_"):
                continue
            # 至少要有 1 个 cookie 文件
            has_cookie = any(
                os.path.isfile(os.path.join(full_path, f"{suffix}_cookie.txt"))
                for suffix in ("sz", "jzt", "jm")
            )
            if has_cookie:
                discovered.append(entry)
        return discovered

    def auto_register_from_disk(self) -> List[str]:
        """从 config/ 目录自动注册所有店铺（提取 pin）"""
        new_shops = []
        for shop_id in self.scan_config_dir():
            # 从 sz_cookie.txt 提取 pin（优先）
            pin = None
            for suffix in ("sz", "jzt", "jm"):
                cookie_path = os.path.join(CONFIG_DIR, shop_id, f"{suffix}_cookie.txt")
                if os.path.isfile(cookie_path):
                    pin = self._extract_pin(cookie_path)
                    if pin:
                        break
            if pin:
                if self.register_shop(shop_id, pin):
                    new_shops.append(shop_id)
        return new_shops

    @staticmethod
    def _extract_pin(cookie_path: str) -> Optional[str]:
        """从 cookie 文本里提取 pin=XXXXX 字段"""
        try:
            with open(cookie_path, "r", encoding="utf-8") as f:
                content = f.read()
            for part in content.split(";"):
                part = part.strip()
                if part.startswith("pin="):
                    return part[4:].strip()
        except Exception:
            pass
        return None

    # -------------------- h5st 读写 --------------------
    def save_h5st(self, shop_id: str, h5st_value: str) -> str:
        """保存 h5st（DB + 文件双写）

        Returns:
            文件路径（便于 RPA 调用后确认）
        """
        expires_at = (datetime.now() + timedelta(minutes=H5ST_EXPIRE_MINUTES)).strftime("%Y-%m-%d %H:%M:%S")

        # 1. 写 SQLite
        cur = self.conn.cursor()
        cur.execute("""
            INSERT OR REPLACE INTO h5st_log (shop_id, h5st_value, captured_at, expires_at)
            VALUES (?, ?, datetime('now', 'localtime'), ?)
        """, (shop_id, h5st_value, expires_at))
        self.conn.commit()

        # 2. 写文件（供 main.py 直接读取）
        shop_dir = os.path.join(CONFIG_DIR, shop_id)
        os.makedirs(shop_dir, exist_ok=True)
        h5st_path = os.path.join(shop_dir, "h5st.txt")
        with open(h5st_path, "w", encoding="utf-8") as f:
            f.write(h5st_value)

        # 3. 更新 shops 表的 h5st_path
        cur.execute("UPDATE shops SET h5st_path = ?, updated_at = datetime('now', 'localtime') WHERE shop_id = ?", (h5st_path, shop_id))
        self.conn.commit()

        return h5st_path

    def get_h5st(self, shop_id: str) -> Optional[str]:
        """读取 h5st（优先 DB，再读文件）

        Returns:
            h5st 字符串（None 表示未抓到）
        """
        cur = self.conn.cursor()
        cur.execute("SELECT h5st_value, expires_at FROM h5st_log WHERE shop_id = ?", (shop_id,))
        row = cur.fetchone()
        if row:
            h5st_value, expires_at = row["h5st_value"], row["expires_at"]
            # 检查是否过期
            try:
                exp = datetime.strptime(expires_at, "%Y-%m-%d %H:%M:%S")
                if datetime.now() < exp:
                    return h5st_value
                else:
                    print(f"⚠️ 店铺 {shop_id} 的 h5st 已过期（{expires_at}），请重新抓取")
            except ValueError:
                pass

        # DB 没有或过期 → 尝试读文件
        shop = self.get_shop(shop_id)
        if shop and shop.get("h5st_path") and os.path.isfile(shop["h5st_path"]):
            with open(shop["h5st_path"], "r", encoding="utf-8") as f:
                return f.read().strip()
        return None

    def is_h5st_expired(self, shop_id: str) -> bool:
        """检查 h5st 是否过期"""
        cur = self.conn.cursor()
        cur.execute("SELECT expires_at FROM h5st_log WHERE shop_id = ?", (shop_id,))
        row = cur.fetchone()
        if not row:
            return True
        try:
            exp = datetime.strptime(row["expires_at"], "%Y-%m-%d %H:%M:%S")
            return datetime.now() >= exp
        except ValueError:
            return True

    # -------------------- Cookie 读写 --------------------
    def save_cookie(self, shop_id: str, biz_type: str, cookie_content: str) -> str:
        """保存 Cookie（按业务类型 sz/jzt/jm 分文件）

        Args:
            shop_id: 店铺标识
            biz_type: sz / jzt / jm
            cookie_content: 完整 Cookie 字符串
        """
        if biz_type not in ("sz", "jzt", "jm"):
            raise ValueError(f"biz_type 必须是 sz/jzt/jm 之一，得到 {biz_type}")

        shop_dir = os.path.join(CONFIG_DIR, shop_id)
        os.makedirs(shop_dir, exist_ok=True)
        cookie_path = os.path.join(shop_dir, f"{biz_type}_cookie.txt")
        with open(cookie_path, "w", encoding="utf-8") as f:
            f.write(cookie_content)

        # 更新 shops 表
        cur = self.conn.cursor()
        cur.execute(f"UPDATE shops SET {biz_type}_cookie_path = ?, updated_at = datetime('now', 'localtime') WHERE shop_id = ?", (cookie_path, shop_id))
        self.conn.commit()

        return cookie_path

    def get_cookie_path(self, shop_id: str, biz_type: str) -> Optional[str]:
        """获取指定业务类型的 Cookie 文件路径"""
        shop = self.get_shop(shop_id)
        if not shop:
            return None
        return shop.get(f"{biz_type}_cookie_path")

    # -------------------- 导出历史 --------------------
    def log_export(self, shop_id: str, biz_key: str, status: str, task_id: str = "", file_path: str = "", row_count: int = 0):
        """记录导出历史（便于后续统计/重跑）"""
        cur = self.conn.cursor()
        cur.execute("""
            INSERT INTO export_history (shop_id, biz_key, task_id, status, file_path, row_count)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (shop_id, biz_key, task_id, status, file_path, row_count))
        self.conn.commit()

    def get_recent_exports(self, shop_id: str, limit: int = 10) -> List[Dict]:
        """查最近 N 条导出记录"""
        cur = self.conn.cursor()
        cur.execute("""
            SELECT * FROM export_history
            WHERE shop_id = ?
            ORDER BY created_at DESC
            LIMIT ?
        """, (shop_id, limit))
        return [dict(row) for row in cur.fetchall()]


# -------------------- 全局便捷函数 --------------------
_default_manager: Optional[ConfigManager] = None


def get_config_manager() -> ConfigManager:
    """获取全局 ConfigManager 单例"""
    global _default_manager
    if _default_manager is None:
        _default_manager = ConfigManager()
    return _default_manager


# -------------------- CLI 入口 --------------------
if __name__ == "__main__":
    import sys

    cm = get_config_manager()

    if len(sys.argv) < 2:
        print("用法：")
        print("  python config_manager.py scan            # 扫描config/目录自动注册")
        print("  python config_manager.py list            # 列出所有店铺")
        print("  python config_manager.py h5st <shop_id>  # 看h5st是否过期")
        sys.exit(1)

    cmd = sys.argv[1]
    if cmd == "scan":
        new = cm.auto_register_from_disk()
        if new:
            print(f"✅ 新注册 {len(new)} 个店铺：{new}")
        else:
            print("无新店铺")
        print(f"\n当前所有店铺：")
        for shop in cm.list_shops():
            print(f"  - {shop['shop_id']} (pin={shop['pin']})")

    elif cmd == "list":
        for shop in cm.list_shops():
            print(f"  - {shop['shop_id']} (pin={shop['pin']})")
            for k in ("sz_cookie_path", "jzt_cookie_path", "jm_cookie_path", "h5st_path"):
                v = shop.get(k)
                print(f"      {k}: {v if v else '(未配置)'}")

    elif cmd == "h5st":
        if len(sys.argv) < 3:
            print("用法：python config_manager.py h5st <shop_id>")
            sys.exit(1)
        shop_id = sys.argv[2]
        expired = cm.is_h5st_expired(shop_id)
        h5st = cm.get_h5st(shop_id)
        print(f"店铺: {shop_id}")
        print(f"h5st 是否过期: {'❌ 过期' if expired else '✅ 有效'}")
        print(f"h5st 值（截断）: {h5st[:60] + '...' if h5st else 'None'}")

    else:
        print(f"未知命令：{cmd}")