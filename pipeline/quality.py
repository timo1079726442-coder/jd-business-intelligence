"""流水线最小数据质量检查。

检查范围严格限制在本次成功任务的店铺、业务、日期范围，避免每次全库扫描。
"""

from __future__ import annotations

import sqlite3
from pathlib import Path


class DataQualityChecker:
    """检查 SQLite 最新数据的基本完整性、隔离性和精确重复候选。"""

    # 超过阈值时不做全行哈希，以免大报表让每次 Run 变成全库扫描。
    EXACT_DUPLICATE_ROW_LIMIT = 50_000

    def __init__(self, sqlite_path: str | Path):
        self.sqlite_path = Path(sqlite_path)

    def check_task(self, *, source_table: str, shop_pin: str, start_date: str, end_date: str) -> dict:
        """检查一个已同步任务的当前 SQLite 状态。"""
        result = {
            "source_table": source_table,
            "shop_pin": shop_pin,
            "start_date": start_date,
            "end_date": end_date,
            "row_count": 0,
            "latest_data_date": None,
            "empty_shop_pin_rows": 0,
            "exact_duplicate_rows": None,
            "duplicate_check": "NOT_RUN",
            "hard_failures": [],
        }
        if not self.sqlite_path.exists():
            result["hard_failures"].append("SQLite 数据库不存在")
            return result
        conn = sqlite3.connect(self.sqlite_path)
        try:
            info = conn.execute(f'PRAGMA table_info("{source_table}")').fetchall()
            if not info:
                result["hard_failures"].append("本次业务源表不存在")
                return result
            columns = [row[1] for row in info]
            if "shop_pin" not in columns or "stat_date" not in columns:
                result["hard_failures"].append("源表缺少 shop_pin/stat_date 公共字段")
                return result
            row = conn.execute(
                f'SELECT COUNT(*), MAX(stat_date) FROM "{source_table}" '
                "WHERE shop_pin=? AND stat_date BETWEEN ? AND ?",
                (shop_pin, start_date, end_date),
            ).fetchone()
            result["row_count"], result["latest_data_date"] = int(row[0] or 0), row[1]
            result["empty_shop_pin_rows"] = int(conn.execute(
                f'SELECT COUNT(*) FROM "{source_table}" WHERE (shop_pin IS NULL OR shop_pin=\'\') '
                "AND stat_date BETWEEN ? AND ?",
                (start_date, end_date),
            ).fetchone()[0] or 0)
            if result["empty_shop_pin_rows"]:
                result["hard_failures"].append("发现无店铺标识的数据，存在跨店污染风险")

            # 无唯一键业务使用按日期覆盖语义；这里仅检查“业务字段完全相同”的重复行。
            if result["row_count"] <= self.EXACT_DUPLICATE_ROW_LIMIT:
                business_cols = [col for col in columns if col.lower() not in {"id", "etl_time"}]
                select_cols = ", ".join(f'"{col}"' for col in business_cols)
                rows = conn.execute(
                    f'SELECT {select_cols} FROM "{source_table}" WHERE shop_pin=? AND stat_date BETWEEN ? AND ?',
                    (shop_pin, start_date, end_date),
                ).fetchall()
                seen, duplicates = set(), 0
                for values in rows:
                    marker = tuple("" if value is None else str(value) for value in values)
                    if marker in seen:
                        duplicates += 1
                    else:
                        seen.add(marker)
                result["exact_duplicate_rows"] = duplicates
                result["duplicate_check"] = "EXACT_ROW_HASH"
                if duplicates:
                    result["hard_failures"].append(f"发现 {duplicates} 条完全重复候选行")
            else:
                result["duplicate_check"] = "SKIPPED_LARGE_TASK"
        finally:
            conn.close()
        return result
