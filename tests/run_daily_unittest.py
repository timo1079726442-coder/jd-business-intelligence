"""run_daily 调度边界回归测试。"""

import unittest
import os
from types import SimpleNamespace
from unittest.mock import patch

import run_daily


class RunDailyBoundaryTests(unittest.TestCase):
    def test_only_jingmai_order_and_after_sale_use_30_days(self):
        jm = SimpleNamespace(module="京麦", biz_key="京麦订单明细_完整一键导出")
        jm_after = SimpleNamespace(module="京麦", biz_key="京麦售后明细_完整一键导出")
        jzt = SimpleNamespace(module="京准通", biz_key="京准通全站推广-计划")
        sz = SimpleNamespace(module="商智", biz_key="商品流量来源_搜索")
        self.assertEqual(run_daily._effective_strategy(jm), "近30天")
        self.assertEqual(run_daily._effective_strategy(jm_after), "近30天")
        self.assertEqual(run_daily._effective_strategy(jzt), "单日")
        self.assertEqual(run_daily._effective_strategy(sz), "单日")

    def test_dry_run_does_not_read_auth_or_database(self):
        cfg = SimpleNamespace(
            module="京麦",
            report_name="订单明细",
            biz_key="京麦订单明细_完整一键导出",
            strategy="近30天",
            start_date="2026-08-01",
            end_date="2026-08-26",
            resolve_dates=lambda: ("2026-07-28", "2026-08-26"),
        )
        with patch.object(run_daily, "_get_jm_h5st", side_effect=AssertionError("dry-run 不应读取 h5st")), patch(
            "main.get_business_handler", side_effect=AssertionError("dry-run 不应初始化业务 handler")
        ):
            self.assertTrue(run_daily._run_single_config(cfg, "店铺", "pin", dry_run=True))

    def test_first_fill_window_is_capped_at_yesterday(self):
        cfg = SimpleNamespace(start_date="2026-08-25", end_date="2099-01-01")
        with patch.object(run_daily, "_get_existing_dates", return_value=set()), patch.object(
            run_daily, "datetime"
        ) as dt:
            dt.now.return_value.date.return_value = run_daily.date(2026, 8, 27)
            dt.strptime = __import__("datetime").datetime.strptime
            result = run_daily._scan_missing_dates("业务", cfg, "pin")
        self.assertEqual(result, ["2026-08-25", "2026-08-26"])

    def test_auth_exit_code_is_not_swallowed(self):
        with patch.dict(os.environ, {"SHOP_PIN": "pin"}), patch.object(
            run_daily, "_run_single_config", side_effect=SystemExit(run_daily.EXIT_AUTH_EXPIRED)
        ):
            with self.assertRaises(SystemExit) as ctx:
                run_daily.run_shop("店铺", [object()], dry_run=False)
        self.assertEqual(ctx.exception.code, run_daily.EXIT_AUTH_EXPIRED)


if __name__ == "__main__":
    unittest.main()
