"""流水线计划与执行结果的数据结构。

使用 dataclass（数据类）让任务计划可以同时写入 SQLite 元数据表和 JSON Summary。
"""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any


@dataclass
class PlannedTask:
    """一个“店铺 + 业务 + 日期范围”的实际采集任务。"""

    task_id: str
    run_id: str
    shop_id: str
    shop_key: str
    shop_pin: str
    biz_key: str
    module: str
    requested_start_date: str
    requested_end_date: str
    granularity: str
    task_type: str
    supports_range: bool
    auth_required: list[str] = field(default_factory=list)
    auth_status: str = "NOT_CHECKED"
    status: str = "PLANNED"
    record_count: int | None = None
    error_type: str | None = None
    error_message: str | None = None
    source_request_id: str | None = None
    started_at: str | None = None
    finished_at: str | None = None

    def to_dict(self) -> dict[str, Any]:
        """转换为可安全写入 JSON 的普通字典。"""
        return asdict(self)
