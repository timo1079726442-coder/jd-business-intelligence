"""统一数据流水线编排器。

职责：读取配置、生成计划、检查鉴权、调用既有采集器、增量同步 MySQL、验证并输出 Summary。
职责边界：不重新实现京东 API，不修改影刀，也不调用危险的全表重建脚本。
"""

from __future__ import annotations

import json
import logging
import os
import re
import sqlite3
import uuid
from datetime import datetime
from pathlib import Path
from typing import Iterable

from pipeline.metadata import MetadataStore, now_iso
from pipeline.models import PlannedTask
from pipeline.mysql_sync import MySQLIncrementalSync
from pipeline.planner import build_tasks, shop_key_from_id
from pipeline.quality import DataQualityChecker
from pipeline.status import (
    FAILED_AUTH,
    FAILED_DATABASE,
    FAILED_PARSE,
    FAILED_REQUEST,
    FAILED_VALIDATION,
    DRY_RUN_READY,
    SUCCESS_EMPTY,
    SUCCESS_WITH_DATA,
    choose_exit_code,
)

APP_VERSION = "pipeline-1.0.0"
_SECRET_PATTERN = re.compile(r"(?i)(cookie|h5st|password|token)(\s*[=:]\s*)([^\s,;]+)")


def mask_sensitive(text: object) -> str:
    """脱敏异常文本，防止 Cookie/h5st/密码写入日志或 Summary。"""
    return _SECRET_PATTERN.sub(r"\1\2***", str(text))


class _ContextFilter(logging.Filter):
    """将 run/task/shop/biz 上下文注入每行日志。"""

    def __init__(self, run_id: str):
        super().__init__()
        self.run_id = run_id
        self.task: PlannedTask | None = None

    def filter(self, record: logging.LogRecord) -> bool:
        task = self.task
        record.run_id = self.run_id
        record.task_id = task.task_id if task else "-"
        record.shop_key = task.shop_key if task else "-"
        record.biz_key = task.biz_key if task else "-"
        return True


class PipelineOrchestrator:
    """统一入口的业务编排对象。"""

    def __init__(self, project_root: str | Path, *, trigger_source: str, dry_run: bool = False):
        self.root = Path(project_root)
        self.trigger_source = trigger_source
        self.dry_run = dry_run
        self.run_id = f"run-{datetime.now():%Y%m%d-%H%M%S}-{uuid.uuid4().hex[:8]}"
        self.started_at = now_iso()
        self.metadata = MetadataStore(self.root / "data" / "jd_report.db")
        self.mysql_sync = MySQLIncrementalSync(self.root)
        self.logger, self.log_filter = self._build_logger()

    def _build_logger(self):
        """创建 UTF-8 文件日志，且不重复添加 handler。"""
        logger = logging.getLogger(f"pipeline.{self.run_id}")
        logger.setLevel(logging.INFO)
        logger.propagate = False
        context_filter = _ContextFilter(self.run_id)
        log_dir = self.root / "logs"
        log_dir.mkdir(parents=True, exist_ok=True)
        handler = logging.FileHandler(log_dir / f"jd_pipeline_{datetime.now():%Y%m%d}.log", encoding="utf-8")
        handler.setFormatter(logging.Formatter(
            "%(asctime)s [%(levelname)s] [run=%(run_id)s][task=%(task_id)s]"
            "[shop=%(shop_key)s][biz=%(biz_key)s] %(message)s"
        ))
        handler.addFilter(context_filter)
        logger.handlers.clear()
        logger.addHandler(handler)
        return logger, context_filter

    def resolve_shops(self, requested: list[str] | None) -> list[tuple[str, str]]:
        """把 CLI 的 MIYO/FYA/OTA 或店铺全名解析为店铺全名与 pin。"""
        from biz_config_loader import list_shops
        from runtime_config import get_shop_pin

        available = list_shops(enabled_only=True)
        selected = []
        wanted = {value.strip().upper() for value in requested or [] if value.strip()}
        for cfg in available:
            short = shop_key_from_id(cfg.shop_id)
            if wanted and cfg.shop_id.upper() not in wanted and short not in wanted:
                continue
            old_shop, old_pin = os.environ.get("SHOP_ID"), os.environ.get("SHOP_PIN")
            try:
                os.environ["SHOP_ID"] = cfg.shop_id
                os.environ.pop("SHOP_PIN", None)
                selected.append((cfg.shop_id, get_shop_pin()))
            finally:
                if old_shop is None:
                    os.environ.pop("SHOP_ID", None)
                else:
                    os.environ["SHOP_ID"] = old_shop
                if old_pin is None:
                    os.environ.pop("SHOP_PIN", None)
                else:
                    os.environ["SHOP_PIN"] = old_pin
        if wanted and not selected:
            raise ValueError(f"未找到启用店铺：{', '.join(sorted(wanted))}")
        return selected

    @staticmethod
    def _is_auth_error(exc: Exception) -> bool:
        return type(exc).__name__ in {"CookieExpiredError", "H5stExpiredError", "AuthFileNotFound"}

    def _check_auth(self, task: PlannedTask) -> str:
        """最低成本检查本任务所需鉴权文件；不输出鉴权内容。"""
        from auth_loader import AuthLoader

        loader = AuthLoader(shop_id=task.shop_id)
        if task.module == "京麦":
            loader.get_cookie_str("jm", check_expire=True)
            loader.get_h5st(check_expire=True, h5st_key="jm_order" if "订单" in task.biz_key else "jm_after_sale")
        elif task.module == "京准通":
            loader.get_cookie_str("jzt", check_expire=True)
        else:
            loader.get_cookie_str("sz", check_expire=True)
        return "READY"

    def _record_count(self, task: PlannedTask) -> int:
        """统计本任务范围 SQLite 当前行数，用于 Summary 与 MySQL 核对。"""
        from db_utils import biz_key_to_table_name, get_db_path

        db_path = get_db_path()
        if not os.path.exists(db_path):
            return 0
        table = biz_key_to_table_name(task.biz_key)
        conn = sqlite3.connect(db_path)
        try:
            row = conn.execute(
                f'SELECT COUNT(*) FROM "{table}" WHERE shop_pin=? AND stat_date BETWEEN ? AND ?',
                (task.shop_pin, task.requested_start_date, task.requested_end_date),
            ).fetchone()
            return int(row[0] or 0)
        finally:
            conn.close()

    def _execute_collector(self, task: PlannedTask):
        """调用既有 handler 和既有 xlsx→SQLite 归档逻辑。"""
        import main
        from run_daily import _extract_and_backfill, _get_jm_h5st

        handler = main.get_business_handler(task.biz_key)
        kwargs = {}
        if task.module == "京麦":
            kwargs["h5st"] = _get_jm_h5st(task.biz_key)
            if not kwargs["h5st"]:
                raise RuntimeError("京麦 h5st 读取失败")
        if task.requested_start_date == task.requested_end_date and task.task_type != "ROLLING_REFRESH":
            kwargs["date"] = task.requested_start_date
        else:
            kwargs["start_date"] = task.requested_start_date
            kwargs["end_date"] = task.requested_end_date
        if task.granularity == "month":
            # 月度业务若 handler 支持 granularity，会使用 month；不影响不读取该参数的旧业务。
            kwargs["granularity"] = "month"
        result = handler(**kwargs)
        _extract_and_backfill(
            type("Config", (), {"biz_key": task.biz_key})(),
            result,
            task.shop_id,
            fallback_date=task.requested_end_date,
        )
        return result

    def _flush_excel(self, shop_id: str, shop_pin: str) -> None:
        """一次 Run 内每店只增量刷一次总表。"""
        import excel_master

        excel_master.flush_shop(shop_id, shop_pin)

    def _metric_smoke(self, shop_key: str, start: str, end: str) -> None:
        """确认 Metric Service 能读取最新 Derived 数据，不依赖浏览器页面。"""
        from metrics.service import MetricService

        MetricService().get_shop_overview(shop_key, start, end)

    def _write_summary(self, summary: dict) -> Path:
        """写运行结果文件；内容不含 Cookie、h5st 或数据库密码。"""
        target = self.root / "runtime" / "runs" / f"{self.run_id}.json"
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
        return target

    def close(self) -> None:
        """显式关闭文件日志，避免 Windows 测试或连续 Run 锁住日志文件。"""
        for handler in list(self.logger.handlers):
            handler.close()
            self.logger.removeHandler(handler)

    def run(
        self,
        *,
        shop_filters: list[str] | None = None,
        biz_filters: list[str] | None = None,
        start_override: str | None = None,
        end_override: str | None = None,
    ) -> dict:
        """执行完整流水线；dry-run 只读配置/SQLite/鉴权文件，不采集或写库。"""
        if bool(start_override) != bool(end_override):
            raise ValueError("--start 与 --end 必须同时传入")
        from config_download_reader import load_download_configs
        import run_daily

        configs = load_download_configs()
        wanted_biz = {value.strip() for value in (biz_filters or []) if value.strip()}
        if wanted_biz:
            configs = [cfg for cfg in configs if cfg.biz_key in wanted_biz]
            unknown = wanted_biz - {cfg.biz_key for cfg in configs}
            if unknown:
                raise ValueError(f"未找到启用业务：{', '.join(sorted(unknown))}")
        shops = self.resolve_shops(shop_filters)
        all_tasks: list[PlannedTask] = []
        for shop_id, shop_pin in shops:
            all_tasks.extend(build_tasks(
                run_id=self.run_id,
                shop_id=shop_id,
                shop_pin=shop_pin,
                configs=configs,
                missing_dates_provider=run_daily._scan_missing_dates,
                effective_strategy=run_daily._effective_strategy,
                range_supported_provider=lambda key: bool(__import__("biz_config_loader").get_biz_feature(key).get("supports_range")),
                start_override=start_override,
                end_override=end_override,
            ))

        if not self.dry_run:
            self.metadata.ensure_schema()
            self.metadata.start_run(self.run_id, self.trigger_source, [name for name, _ in shops], APP_VERSION)

        for task in all_tasks:
            self.log_filter.task = task
            try:
                task.auth_status = self._check_auth(task)
            except Exception as exc:
                task.status = FAILED_AUTH
                task.error_type = type(exc).__name__
                task.error_message = mask_sensitive(exc)
            if not self.dry_run and task.status == "PLANNED":
                task.started_at = now_iso()
                try:
                    old_shop, old_pin = os.environ.get("SHOP_ID"), os.environ.get("SHOP_PIN")
                    os.environ["SHOP_ID"], os.environ["SHOP_PIN"] = task.shop_id, task.shop_pin
                    try:
                        collector_result = self._execute_collector(task)
                    finally:
                        if old_shop is None:
                            os.environ.pop("SHOP_ID", None)
                        else:
                            os.environ["SHOP_ID"] = old_shop
                        if old_pin is None:
                            os.environ.pop("SHOP_PIN", None)
                        else:
                            os.environ["SHOP_PIN"] = old_pin
                    if isinstance(collector_result, dict):
                        # 各接口字段名不同，按已知安全键提取；拿不到时保留 NULL，不伪造 request id。
                        task.source_request_id = next(
                            (str(collector_result[key]) for key in ("source_request_id", "request_id", "task_id")
                             if collector_result.get(key) is not None),
                            None,
                        )
                    task.record_count = self._record_count(task)
                    task.status = SUCCESS_WITH_DATA if task.record_count else SUCCESS_EMPTY
                    from db_utils import biz_key_to_table_name
                    self.mysql_sync.sync_task(
                        source_table=biz_key_to_table_name(task.biz_key),
                        shop_pin=task.shop_pin,
                        start_date=task.requested_start_date,
                        end_date=task.requested_end_date,
                    )
                except Exception as exc:
                    task.error_type = type(exc).__name__
                    task.error_message = mask_sensitive(exc)
                    if self._is_auth_error(exc):
                        task.status = FAILED_AUTH
                    elif "mysql" in type(exc).__name__.lower() or "数据库" in str(exc):
                        task.status = FAILED_DATABASE
                    elif isinstance(exc, (ValueError, KeyError, UnicodeError)):
                        task.status = FAILED_PARSE
                    else:
                        task.status = FAILED_REQUEST
                finally:
                    task.finished_at = now_iso()
            if self.dry_run and task.status == "PLANNED":
                task.status = DRY_RUN_READY
            if not self.dry_run:
                self.metadata.upsert_task(task)

        if not self.dry_run:
            # 对至少有一项采集成功的店铺刷一次 Excel、检查 Derived、验证 Metric。
            for shop_id, shop_pin in shops:
                shop_tasks = [task for task in all_tasks if task.shop_id == shop_id]
                if any(task.status in {SUCCESS_WITH_DATA, SUCCESS_EMPTY} for task in shop_tasks):
                    try:
                        self._flush_excel(shop_id, shop_pin)
                        self.mysql_sync.derived_health_check()
                        successful = next(task for task in shop_tasks if task.status in {SUCCESS_WITH_DATA, SUCCESS_EMPTY})
                        self._metric_smoke(successful.shop_key, successful.requested_start_date, successful.requested_end_date)
                    except Exception as exc:
                        for task in shop_tasks:
                            if task.status in {SUCCESS_WITH_DATA, SUCCESS_EMPTY}:
                                task.status = FAILED_VALIDATION
                                task.error_type = type(exc).__name__
                                task.error_message = mask_sensitive(exc)
                                self.metadata.upsert_task(task)

        quality_checks = []
        if not self.dry_run:
            from db_utils import biz_key_to_table_name, get_db_path

            checker = DataQualityChecker(get_db_path())
            for task in all_tasks:
                if task.status not in {SUCCESS_WITH_DATA, SUCCESS_EMPTY}:
                    continue
                check = checker.check_task(
                    source_table=biz_key_to_table_name(task.biz_key),
                    shop_pin=task.shop_pin,
                    start_date=task.requested_start_date,
                    end_date=task.requested_end_date,
                )
                quality_checks.append(check)
                if check["hard_failures"]:
                    task.status = FAILED_VALIDATION
                    task.error_type = "DataQualityError"
                    task.error_message = "; ".join(check["hard_failures"])
                    self.metadata.upsert_task(task)

        statuses = [task.status for task in all_tasks]
        exit_code = choose_exit_code(statuses)
        run_status = "DRY_RUN" if self.dry_run else ("SUCCESS" if exit_code == 0 else "FAILED")
        summary = {
            "run_id": self.run_id,
            "started_at": self.started_at,
            "finished_at": now_iso(),
            "trigger_source": self.trigger_source,
            "dry_run": self.dry_run,
            "status": run_status,
            "shops": {shop_key_from_id(shop): {"shop_id": shop, "shop_pin": pin} for shop, pin in shops},
            "tasks": [task.to_dict() for task in all_tasks],
            "data_quality": {
                "checks": quality_checks,
                "failed_task_count": sum(task.status == FAILED_VALIDATION for task in all_tasks),
            },
            "exit_code": exit_code,
            "meta": {"app_version": APP_VERSION, "config_version": "config.xlsx"},
        }
        summary_path = self._write_summary(summary)
        summary["summary_path"] = str(summary_path)
        if not self.dry_run:
            self.metadata.finish_run(self.run_id, run_status)
        self.close()
        return summary
