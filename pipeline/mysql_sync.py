"""SQLite 到 MySQL Standard 表的安全增量同步。

禁止调用旧的 build_mysql_standard_tables.py：该脚本会 DROP 全部 Standard 表。
本模块只覆盖“当前店铺 + 当前业务 + 当前任务日期”这一最小范围，保留其他历史日期。
"""

from __future__ import annotations

import configparser
import re
import sqlite3
from pathlib import Path


_IDENTIFIER = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")


def _qid(name: str) -> str:
    """严格校验并引用表/列标识符，避免动态 SQL 注入。"""
    if not _IDENTIFIER.fullmatch(name):
        raise ValueError(f"非法数据库标识符：{name!r}")
    return f"`{name}`"


class MySQLIncrementalSync:
    """仅同步本次任务涉及的 SQLite 业务数据。"""

    def __init__(self, project_root: str | Path):
        self.root = Path(project_root)
        self.sqlite_path = self.root / "data" / "jd_report.db"
        self.config_path = self.root / "config" / "mysql.local.ini"

    def _connect_mysql(self):
        """延迟导入 MySQL 驱动，dry-run 和纯计划不依赖该包。"""
        import mysql.connector

        parser = configparser.ConfigParser()
        if not parser.read(self.config_path, encoding="utf-8") or "mysql" not in parser:
            raise RuntimeError("MySQL 配置缺失：config/mysql.local.ini")
        cfg = parser["mysql"]
        return mysql.connector.connect(
            host=cfg.get("host"),
            port=cfg.getint("port"),
            user=cfg.get("user"),
            password=cfg.get("password"),
            database=cfg.get("database"),
            charset=cfg.get("charset", "utf8mb4"),
        )

    def _ensure_target_table(self, cur, table: str, columns: list[str]) -> None:
        """只创建缺失的 Standard 表或缺失列，绝不 DROP/TRUNCATE。"""
        definitions = ["`id` BIGINT NOT NULL AUTO_INCREMENT PRIMARY KEY"]
        definitions.extend(f"{_qid(column)} LONGTEXT NULL" for column in columns)
        definitions.extend([
            "KEY `ix_shop_pin` (`shop_pin`(128))",
            "KEY `ix_stat_date` (`stat_date`(32))",
        ])
        cur.execute(
            f"CREATE TABLE IF NOT EXISTS {_qid(table)} ({', '.join(definitions)}) "
            "CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci"
        )
        cur.execute(
            "SELECT COLUMN_NAME FROM information_schema.columns "
            "WHERE table_schema=DATABASE() AND table_name=%s",
            (table,),
        )
        existing = {row[0] for row in cur.fetchall()}
        for column in columns:
            if column not in existing:
                cur.execute(f"ALTER TABLE {_qid(table)} ADD COLUMN {_qid(column)} LONGTEXT NULL")

    def sync_task(self, *, source_table: str, shop_pin: str, start_date: str, end_date: str) -> int:
        """同步一个任务范围，返回写入 MySQL 的行数。

        同一店铺同一业务日期先删除再插入，语义与 SQLite 的按日期覆盖一致；
        范围以外数据不会触碰，重复运行也不会累加。
        """
        if not self.sqlite_path.exists():
            raise RuntimeError(f"SQLite 数据库不存在：{self.sqlite_path}")
        if not _IDENTIFIER.fullmatch(source_table):
            raise ValueError(f"非法源表名：{source_table!r}")
        target_table = f"std_{source_table}"

        sqlite_conn = sqlite3.connect(self.sqlite_path)
        try:
            info = sqlite_conn.execute(f"PRAGMA table_info({_qid(source_table)})").fetchall()
            if not info:
                raise RuntimeError(f"SQLite 源表不存在：{source_table}")
            columns = [row[1] for row in info if row[1].lower() != "id"]
            required = {"shop_pin", "stat_date"}
            if not required.issubset(columns):
                raise RuntimeError(f"SQLite 源表缺少公共字段：{source_table}")
            rows = sqlite_conn.execute(
                f"SELECT {', '.join(_qid(column) for column in columns)} FROM {_qid(source_table)} "
                "WHERE shop_pin=? AND stat_date BETWEEN ? AND ?",
                (shop_pin, start_date, end_date),
            ).fetchall()
        finally:
            sqlite_conn.close()

        mysql_conn = self._connect_mysql()
        try:
            cur = mysql_conn.cursor()
            self._ensure_target_table(cur, target_table, columns)
            # 仅刷新当前任务范围，避免影响更老历史数据。
            cur.execute(
                f"DELETE FROM {_qid(target_table)} WHERE shop_pin=%s AND stat_date BETWEEN %s AND %s",
                (shop_pin, start_date, end_date),
            )
            if rows:
                placeholders = ", ".join(["%s"] * len(columns))
                cur.executemany(
                    f"INSERT INTO {_qid(target_table)} ({', '.join(_qid(column) for column in columns)}) "
                    f"VALUES ({placeholders})",
                    [tuple(None if value is None else str(value) for value in row) for row in rows],
                )
            mysql_conn.commit()
            return len(rows)
        except Exception:
            mysql_conn.rollback()
            raise
        finally:
            try:
                cur.close()
            except UnboundLocalError:
                pass
            mysql_conn.close()

    def derived_health_check(self) -> None:
        """验证 Derived View 可读取；View 本身不在每次 Run 重建。"""
        conn = self._connect_mysql()
        try:
            cur = conn.cursor()
            cur.execute("SELECT 1 FROM vw_derived_jm_order_full LIMIT 1")
            cur.fetchone()
            cur.close()
        finally:
            conn.close()
