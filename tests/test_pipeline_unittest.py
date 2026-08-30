"""阶段 1B 流水线的纯单元测试。

这些测试不读取真实 Cookie、不请求京东，也不连接 MySQL，适合每次改计划逻辑后快速运行。
"""

import sqlite3
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from pipeline.metadata import MetadataStore
from pipeline.orchestrator import PipelineOrchestrator, mask_sensitive
from pipeline.planner import build_tasks, date_ranges, month_ranges
from pipeline.status import (
    EXIT_AUTH_FAILURE,
    EXIT_DATABASE_FAILURE,
    EXIT_SUCCESS,
    FAILED_AUTH,
    FAILED_DATABASE,
    SUCCESS_EMPTY,
    choose_exit_code,
)


class PipelinePlannerTests(unittest.TestCase):
    """验证 rolling、区间、单日、月度四种计划不会互相退化。"""

    @staticmethod
    def _config(module="商智", biz_key="商品流量来源_搜索"):
        return SimpleNamespace(
            module=module,
            report_name=biz_key,
            biz_key=biz_key,
            resolve_dates=lambda: ("2026-08-01", "2026-08-30"),
        )

    @staticmethod
    def _missing(_biz_key, _cfg, _pin):
        return ["2026-08-24", "2026-08-25", "2026-08-26"]

    def test_date_ranges_merge_only_contiguous_days(self):
        self.assertEqual(
            date_ranges(["2026-08-24", "2026-08-25", "2026-08-27"]),
            [("2026-08-24", "2026-08-25"), ("2026-08-27", "2026-08-27")],
        )

    def test_month_ranges_never_become_single_day_tasks(self):
        self.assertEqual(
            month_ranges(["2026-08-02", "2026-08-28", "2026-09-01"]),
            [("2026-08-01", "2026-08-31"), ("2026-09-01", "2026-09-30")],
        )

    def test_range_business_merges_three_missing_days(self):
        tasks = build_tasks(
            run_id="run-test",
            shop_id="MIYO箱包旗舰店",
            shop_pin="miyo-周",
            configs=[self._config(module="京准通", biz_key="测试区间业务")],
            missing_dates_provider=self._missing,
            effective_strategy=lambda _cfg: "区间",
            range_supported_provider=lambda _key: True,
        )
        self.assertEqual(len(tasks), 1)
        self.assertEqual((tasks[0].requested_start_date, tasks[0].requested_end_date), ("2026-08-24", "2026-08-26"))
        self.assertEqual(tasks[0].task_type, "BACKFILL")

    def test_day_business_keeps_each_missing_day_independent(self):
        tasks = build_tasks(
            run_id="run-test",
            shop_id="MIYO箱包旗舰店",
            shop_pin="miyo-周",
            configs=[self._config()],
            missing_dates_provider=self._missing,
            effective_strategy=lambda _cfg: "单日",
            range_supported_provider=lambda _key: False,
        )
        self.assertEqual(len(tasks), 3)
        self.assertTrue(all(task.requested_start_date == task.requested_end_date for task in tasks))

    def test_only_jingmai_rolling_task_is_marked_rolling_refresh(self):
        tasks = build_tasks(
            run_id="run-test",
            shop_id="MIYO箱包旗舰店",
            shop_pin="miyo-周",
            configs=[self._config(module="京麦", biz_key="京麦订单明细_完整一键导出")],
            missing_dates_provider=self._missing,
            effective_strategy=lambda _cfg: "近30天",
            range_supported_provider=lambda _key: True,
        )
        self.assertEqual(len(tasks), 1)
        self.assertEqual(tasks[0].task_type, "ROLLING_REFRESH")
        self.assertEqual(tasks[0].auth_required, ["jm_cookie", "jm_order_h5st"])


class PipelineMetadataTests(unittest.TestCase):
    """验证批次表独立创建，不依赖或修改业务数据表。"""

    def test_metadata_schema_and_task_write(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            db_path = Path(temp_dir) / "pipeline.db"
            store = MetadataStore(db_path)
            store.ensure_schema()
            store.start_run("run-test", "manual", ["MIYO箱包旗舰店"], "test")
            task = build_tasks(
                run_id="run-test",
                shop_id="MIYO箱包旗舰店",
                shop_pin="miyo-周",
                configs=[PipelinePlannerTests._config()],
                missing_dates_provider=lambda *_: ["2026-08-24"],
                effective_strategy=lambda _cfg: "单日",
                range_supported_provider=lambda _key: False,
            )[0]
            task.status, task.record_count = SUCCESS_EMPTY, 0
            store.upsert_task(task)
            # 同一 task_id 重写状态不增加第二行，保证元数据本身可重复执行。
            store.upsert_task(task)
            store.finish_run("run-test", "SUCCESS")
            conn = sqlite3.connect(db_path)
            try:
                self.assertEqual(conn.execute("SELECT status FROM etl_run").fetchone()[0], "SUCCESS")
                self.assertEqual(conn.execute("SELECT status, record_count FROM etl_task_run").fetchone(), (SUCCESS_EMPTY, 0))
                self.assertEqual(conn.execute("SELECT COUNT(*) FROM etl_task_run").fetchone()[0], 1)
            finally:
                # Windows 会锁定未关闭的 SQLite 文件；显式关闭，避免测试临时目录无法清理。
                conn.close()

    def test_status_exit_code_and_secret_masking(self):
        self.assertEqual(choose_exit_code([SUCCESS_EMPTY]), EXIT_SUCCESS)
        self.assertEqual(choose_exit_code([FAILED_AUTH]), EXIT_AUTH_FAILURE)
        self.assertEqual(choose_exit_code([FAILED_DATABASE]), EXIT_DATABASE_FAILURE)
        self.assertNotIn("abc123", mask_sensitive("Cookie=abc123 h5st:xyz987"))


class PipelineDryRunTests(unittest.TestCase):
    """验证 dry-run 只产出计划，绝不初始化元数据或触发采集。"""

    def test_dry_run_does_not_write_database_or_call_collector(self):
        config = PipelinePlannerTests._config()
        with tempfile.TemporaryDirectory() as temp_dir, \
             patch("config_download_reader.load_download_configs", return_value=[config]), \
             patch("run_daily._scan_missing_dates", return_value=["2026-08-24"]), \
             patch.object(PipelineOrchestrator, "resolve_shops", return_value=[("MIYO箱包旗舰店", "miyo-周")]), \
             patch.object(PipelineOrchestrator, "_check_auth", return_value="READY"), \
             patch.object(PipelineOrchestrator, "_write_summary", return_value=Path(temp_dir) / "result.json"), \
             patch("pipeline.orchestrator.MetadataStore.ensure_schema", side_effect=AssertionError("dry-run 不应写元数据")), \
             patch.object(PipelineOrchestrator, "_execute_collector", side_effect=AssertionError("dry-run 不应采集")):
            summary = PipelineOrchestrator(temp_dir, trigger_source="manual", dry_run=True).run(shop_filters=["MIYO"])
        self.assertEqual(summary["status"], "DRY_RUN")
        self.assertEqual(summary["exit_code"], EXIT_SUCCESS)
        self.assertEqual(summary["tasks"][0]["status"], "DRY_RUN_READY")

    def test_dry_run_auth_failure_has_stable_exit_code(self):
        config = PipelinePlannerTests._config()
        with tempfile.TemporaryDirectory() as temp_dir, \
             patch("config_download_reader.load_download_configs", return_value=[config]), \
             patch("run_daily._scan_missing_dates", return_value=["2026-08-24"]), \
             patch.object(PipelineOrchestrator, "resolve_shops", return_value=[("MIYO箱包旗舰店", "miyo-周")]), \
             patch.object(PipelineOrchestrator, "_check_auth", side_effect=RuntimeError("鉴权文件不存在")), \
             patch.object(PipelineOrchestrator, "_write_summary", return_value=Path(temp_dir) / "result.json"):
            summary = PipelineOrchestrator(temp_dir, trigger_source="manual", dry_run=True).run(shop_filters=["MIYO"])
        self.assertEqual(summary["exit_code"], EXIT_AUTH_FAILURE)
        self.assertEqual(summary["tasks"][0]["status"], FAILED_AUTH)


if __name__ == "__main__":
    unittest.main()
