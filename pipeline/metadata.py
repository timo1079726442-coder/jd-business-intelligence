"""ETL Run/Task 元数据追踪。

历史业务表不增加 batch_id；新流水线只在独立元数据表中记录未来批次。
这避免伪造历史批次，也避免 ALTER 所有 Raw/Standard 表。
"""

from __future__ import annotations

import json
import platform
import sqlite3
import sys
from datetime import datetime
from pathlib import Path
from typing import Any


def now_iso() -> str:
    """返回带时区的当前时间，便于跨运行追踪。"""
    return datetime.now().astimezone().isoformat(timespec="seconds")


class MetadataStore:
    """SQLite 元数据存储；仅在非 dry-run 时写入。"""

    def __init__(self, db_path: str | Path):
        self.db_path = Path(db_path)

    def _connect(self) -> sqlite3.Connection:
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        return sqlite3.connect(self.db_path, timeout=30)

    def ensure_schema(self) -> None:
        """创建独立追踪表，不触碰任何 biz_ 历史业务表。"""
        conn = self._connect()
        try:
            conn.executescript(
                """
                CREATE TABLE IF NOT EXISTS etl_run (
                    run_id TEXT PRIMARY KEY,
                    started_at TEXT NOT NULL,
                    finished_at TEXT,
                    trigger_source TEXT NOT NULL,
                    status TEXT NOT NULL,
                    host TEXT,
                    python_version TEXT,
                    app_version TEXT,
                    config_version TEXT,
                    shops_requested TEXT,
                    error_summary TEXT
                );
                CREATE TABLE IF NOT EXISTS etl_task_run (
                    task_id TEXT PRIMARY KEY,
                    run_id TEXT NOT NULL,
                    shop_key TEXT NOT NULL,
                    shop_pin TEXT NOT NULL,
                    biz_key TEXT NOT NULL,
                    requested_start_date TEXT NOT NULL,
                    requested_end_date TEXT NOT NULL,
                    granularity TEXT,
                    task_type TEXT NOT NULL,
                    status TEXT NOT NULL,
                    record_count INTEGER,
                    started_at TEXT,
                    finished_at TEXT,
                    error_type TEXT,
                    error_message TEXT,
                    source_request_id TEXT,
                    FOREIGN KEY(run_id) REFERENCES etl_run(run_id)
                );
                CREATE INDEX IF NOT EXISTS ix_etl_task_run_lookup
                    ON etl_task_run(run_id, shop_key, biz_key, requested_start_date);
                """
            )
            conn.commit()
        finally:
            conn.close()

    def start_run(self, run_id: str, trigger_source: str, shops_requested: list[str], app_version: str) -> None:
        """写入一次流水线的开始记录。"""
        conn = self._connect()
        try:
            conn.execute(
                """INSERT INTO etl_run(
                    run_id, started_at, trigger_source, status, host, python_version,
                    app_version, config_version, shops_requested
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (
                    run_id,
                    now_iso(),
                    trigger_source,
                    "RUNNING",
                    platform.node(),
                    sys.version.split()[0],
                    app_version,
                    "config.xlsx",
                    json.dumps(shops_requested, ensure_ascii=False),
                ),
            )
            conn.commit()
        finally:
            conn.close()

    def upsert_task(self, task: Any) -> None:
        """写入或更新任务记录；错误消息由调用方先脱敏。"""
        conn = self._connect()
        try:
            conn.execute(
                """INSERT INTO etl_task_run(
                    task_id, run_id, shop_key, shop_pin, biz_key,
                    requested_start_date, requested_end_date, granularity, task_type,
                    status, record_count, started_at, finished_at,
                    error_type, error_message, source_request_id
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(task_id) DO UPDATE SET
                    status=excluded.status,
                    record_count=excluded.record_count,
                    started_at=excluded.started_at,
                    finished_at=excluded.finished_at,
                    error_type=excluded.error_type,
                    error_message=excluded.error_message,
                    source_request_id=excluded.source_request_id
                """,
                (
                    task.task_id,
                    task.run_id,
                    task.shop_key,
                    task.shop_pin,
                    task.biz_key,
                    task.requested_start_date,
                    task.requested_end_date,
                    task.granularity,
                    task.task_type,
                    task.status,
                    task.record_count,
                    task.started_at,
                    task.finished_at,
                    task.error_type,
                    task.error_message,
                    task.source_request_id,
                ),
            )
            conn.commit()
        finally:
            conn.close()

    def finish_run(self, run_id: str, status: str, error_summary: str | None = None) -> None:
        """写入最终状态，确保中断后也能追溯本次 Run。"""
        conn = self._connect()
        try:
            conn.execute(
                "UPDATE etl_run SET finished_at=?, status=?, error_summary=? WHERE run_id=?",
                (now_iso(), status, error_summary, run_id),
            )
            conn.commit()
        finally:
            conn.close()
