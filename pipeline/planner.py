"""按配置和已有日期生成数据计划（Data Plan）。"""

from __future__ import annotations

import hashlib
from datetime import date, datetime, timedelta
from typing import Iterable

from pipeline.models import PlannedTask


def shop_key_from_id(shop_id: str) -> str:
    """店铺全名转稳定展示 Key，例如 MIYO箱包旗舰店 -> MIYO。"""
    return shop_id.replace("箱包旗舰店", "").strip().upper()


def date_ranges(dates: Iterable[str]) -> list[tuple[str, str]]:
    """将连续的日粒度缺口压缩为最少区间。"""
    parsed = sorted({datetime.strptime(value, "%Y-%m-%d").date() for value in dates})
    if not parsed:
        return []
    result: list[tuple[str, str]] = []
    start = previous = parsed[0]
    for current in parsed[1:]:
        if current != previous + timedelta(days=1):
            result.append((start.isoformat(), previous.isoformat()))
            start = current
        previous = current
    result.append((start.isoformat(), previous.isoformat()))
    return result


def month_ranges(dates: Iterable[str]) -> list[tuple[str, str]]:
    """将缺失日期按自然月合并；月度业务绝不退化为逐日请求。"""
    months = sorted({value[:7] for value in dates})
    result: list[tuple[str, str]] = []
    for month in months:
        start = date.fromisoformat(f"{month}-01")
        next_month = (start.replace(day=28) + timedelta(days=4)).replace(day=1)
        result.append((start.isoformat(), (next_month - timedelta(days=1)).isoformat()))
    return result


def _task_id(run_id: str, shop_key: str, biz_key: str, start: str, end: str) -> str:
    """生成可追溯且不泄露鉴权信息的任务 ID。"""
    digest = hashlib.sha1(f"{run_id}|{shop_key}|{biz_key}|{start}|{end}".encode("utf-8")).hexdigest()[:10]
    return f"task-{digest}"


def auth_requirements(module: str, biz_key: str) -> list[str]:
    """声明任务所需鉴权材料；仅返回类别，绝不返回凭据内容。"""
    if module == "京麦":
        h5st = "jm_order_h5st" if "订单" in biz_key else "jm_after_sale_h5st"
        return ["jm_cookie", h5st]
    if module == "京准通":
        return ["jzt_cookie"]
    return ["sz_cookie"]


def build_tasks(
    *,
    run_id: str,
    shop_id: str,
    shop_pin: str,
    configs: list,
    missing_dates_provider,
    effective_strategy,
    range_supported_provider,
    start_override: str | None = None,
    end_override: str | None = None,
    today: date | None = None,
) -> list[PlannedTask]:
    """由既有配置生成任务，不调用接口、不写数据库。

    ``missing_dates_provider`` 复用现有 SQLite 缺口判断；因此不重新定义旧业务规则。
    """
    del today  # 预留参数，便于测试时固定日期；日期计算仍交给既有配置解析器。
    tasks: list[PlannedTask] = []
    shop_key = shop_key_from_id(shop_id)
    for cfg in configs:
        strategy = effective_strategy(cfg)
        supports_range = bool(range_supported_provider(cfg.biz_key))
        if start_override and end_override:
            ranges = [(start_override, end_override)]
            task_type = "BACKFILL"
            granularity = "month" if strategy == "月度" else "day"
        elif strategy == "近30天":
            ranges = [cfg.resolve_dates()]
            task_type, granularity = "ROLLING_REFRESH", "day"
        else:
            missing = list(missing_dates_provider(cfg.biz_key, cfg, shop_pin))
            if strategy == "月度":
                ranges = month_ranges(missing)
                task_type, granularity = "BACKFILL", "month"
            elif strategy == "区间" or supports_range:
                ranges = date_ranges(missing)
                task_type, granularity = "BACKFILL", "day"
            else:
                ranges = [(value, value) for value in missing]
                task_type, granularity = "BACKFILL", "day"

        for start, end in ranges:
            tasks.append(
                PlannedTask(
                    task_id=_task_id(run_id, shop_key, cfg.biz_key, start, end),
                    run_id=run_id,
                    shop_id=shop_id,
                    shop_key=shop_key,
                    shop_pin=shop_pin,
                    biz_key=cfg.biz_key,
                    module=cfg.module,
                    requested_start_date=start,
                    requested_end_date=end,
                    granularity=granularity,
                    task_type=task_type,
                    supports_range=supports_range,
                    auth_required=auth_requirements(cfg.module, cfg.biz_key),
                )
            )
    return tasks
