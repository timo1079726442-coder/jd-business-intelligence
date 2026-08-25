# -*- coding: utf-8 -*-
"""商智搜索/推荐/购物车 区间逐日拆分 mock 单元测试（2026-08-10，P1-3）。

测试目标（按用户决策 2026-08-10）：
    1. split_date_range() 纯函数：
       - 正常区间 → 逐天日期列表（含首尾边界）
       - 单日（start==end）→ 只返回1天
       - start晚于end → ValueError
       - 日期格式非法 → ValueError
       - 超过最大天数（默认31天）→ ValueError；恰好31天通过
    2. download_sku() 区间模式（P0 核心）：
       - 检测到 start_date != end_date 时进入逐日循环，不再走单日路径
       - 循环次数 = 区间天数；每天接口入参 date=startDate=endDate=当天
       - 每天插入当天日期列，concat 合并后行数 = 每天行数之和
       - 合并文件名标注区间（如 搜索流量_2026-08-03_2026-08-05.xlsx）
       - 输出子目录用最后一天 end_date
       - 区间循环不复用 config 里的旧 date（修复原 known issue）
    3. download_sku() 单日模式（兼容性）：
       - start_date == end_date 时不进入循环，保持原有执行路径（request 只调1次）
    4. 边界保护：
       - 超过31天的区间直接抛 ValueError，不发起任何 request

mock 思路（不依赖真实 Cookie/接口）：
    - object.__new__(ProductFlowAPI) 跳过 __init__（避免读 Cookie 文件）
    - 注入 self.config / self.logger / self.output_dir / self.request（MagicMock）
    - self.request 根据入参 data["date"] 返回当天模拟 xlsx 字节流（无日期列，含SKU/访客数）
    - 用临时目录隔离输出，跑完清理

运行方式：
    python tests/test_flow_sku_range_unittest.py
或：
    python -m unittest tests.test_flow_sku_range_unittest -v
"""
import io
import os
import sys
import tempfile
import unittest
import shutil
from unittest.mock import MagicMock, patch

# 让 tests/ 目录能找到 main.py
THIS_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(THIS_DIR)
sys.path.insert(0, PROJECT_ROOT)

import main as M  # noqa: E402
from main import ProductFlowAPI, split_date_range  # noqa: E402


# ====================== 模拟响应工厂 ======================

class MockResponse:
    """模拟 requests.Response：content 为接口返回的单天 xlsx 字节流。"""

    def __init__(self, content=b""):
        self.status_code = 200
        self.content = content
        self.headers = {}


def make_day_xlsx(rows):
    """生成「无日期列」的单天 xlsx 字节流（模拟商智 downSkuTable.ajax 返回）。

    列名刻意不含"日期/时间"，触发 prepare_date_columns 的"插入日期列"分支。
    每行含 SKU（数字字符串，验证不转科学计数）与访客数。
    """
    import pandas as pd
    df = pd.DataFrame(rows, columns=["SKU", "访客数"])
    buf = io.BytesIO()
    df.to_excel(buf, index=False, engine="openpyxl")
    buf.seek(0)
    return buf.read()


def make_flow_api(request_side_effect, config_date="2026-07-29"):
    """构造一个跳过 __init__ 的 ProductFlowAPI 实例（注入 mock 依赖）。

    入参:
        request_side_effect - self.request 的 side_effect 函数（接收 data，返回 MockResponse）
        config_date         - 模拟 config.xlsx 里的 date（默认旧日期，验证区间不复用它）
    出参:
        (api, temp_dir)
            api      - 构造好的实例
            temp_dir - 输出临时目录（测试结束需清理）
    """
    api = object.__new__(ProductFlowAPI)
    api.config = {
        "date": config_date,                 # 模拟 config 旧日期
        "interval": "DAY",
        "dateType": "day",
        "limit": "5000",
    }
    api.logger = MagicMock()
    api.output_dir = tempfile.mkdtemp(prefix="flow_range_test_")
    api.request = MagicMock(side_effect=request_side_effect)
    return api, api.output_dir


def request_side_effect_factory(rows_by_day):
    """按 data["date"] 返回对应天数的模拟 xlsx 字节流。

    入参:
        rows_by_day - dict，如 {"2026-08-03": [("SKU1", "100"), ...], ...}
    出参:
        side_effect 函数（供 api.request 使用）
    """
    def _side_effect(url, data, uuid_prefix=None):
        day = data.get("date")
        rows = rows_by_day.get(day, [("默认SKU", "1")])
        return MockResponse(make_day_xlsx(rows))
    return _side_effect


# ====================== split_date_range 纯函数测试 ======================

class TestSplitDateRange(unittest.TestCase):
    """P0 边界保护工具函数测试。"""

    def test_normal_range(self):
        """正常3天区间 → 逐天列表含首尾。"""
        result = split_date_range("2026-08-03", "2026-08-05", max_days=31)
        self.assertEqual(
            result,
            ["2026-08-03", "2026-08-04", "2026-08-05"],
        )

    def test_single_day(self):
        """start==end → 只返回1天（不触发循环的前提条件）。"""
        result = split_date_range("2026-08-09", "2026-08-09", max_days=31)
        self.assertEqual(result, ["2026-08-09"])

    def test_start_after_end(self):
        """start 晚于 end → ValueError。"""
        with self.assertRaises(ValueError) as ctx:
            split_date_range("2026-08-09", "2026-08-03", max_days=31)
        self.assertIn("不能晚于", str(ctx.exception))

    def test_invalid_format(self):
        """日期格式非法 → ValueError。"""
        with self.assertRaises(ValueError):
            split_date_range("20260803", "2026-08-05", max_days=31)

    def test_exceed_max_days(self):
        """超过31天 → ValueError（防大批量压接口触发风控）。"""
        with self.assertRaises(ValueError) as ctx:
            split_date_range("2026-08-01", "2026-09-01", max_days=31)  # 32天
        self.assertIn("超过最大限制", str(ctx.exception))

    def test_exactly_max_days_ok(self):
        """恰好31天 → 通过（边界值兼容）。"""
        result = split_date_range("2026-08-01", "2026-08-31", max_days=31)
        self.assertEqual(len(result), 31)
        self.assertEqual(result[0], "2026-08-01")
        self.assertEqual(result[-1], "2026-08-31")


# ====================== download_sku 区间模式测试（P0 核心） ======================

class TestDownloadSkuRangeMode(unittest.TestCase):
    """区间输入：自动拆成逐天循环，合并输出。"""

    DAYS = ["2026-08-03", "2026-08-04", "2026-08-05"]

    def setUp(self):
        """构造 mock API：每天返回2行数据，行内容按天区分。"""
        rows_by_day = {
            day: [(f"SKU{day[-2:]}_1", "100"), (f"SKU{day[-2:]}_2", "200")]
            for day in self.DAYS
        }
        self.api, self.tmp_dir = make_flow_api(
            request_side_effect_factory(rows_by_day),
            config_date="2026-07-29",  # config 旧日期，验证区间不复用它
        )

    def tearDown(self):
        shutil.rmtree(self.tmp_dir, ignore_errors=True)

    def test_loop_count_and_date_params(self):
        """循环次数=天数，且每天入参 date=startDate=endDate=当天。"""
        self.api.download_sku(
            biz_key="商品流量来源_搜索",
            start_date="2026-08-03",
            end_date="2026-08-05",
        )
        # ① 循环次数 = 区间天数
        self.assertEqual(self.api.request.call_count, 3)
        # ② 每天入参 date=startDate=endDate=当天（且不复用 config 旧日期 07-29）
        for i, day in enumerate(self.DAYS):
            call_args = self.api.request.call_args_list[i].args
            data = call_args[1]
            self.assertEqual(data["date"], day)
            self.assertEqual(data["startDate"], day)
            self.assertEqual(data["endDate"], day)

    def test_merged_file_row_count_and_dates(self):
        """合并后行数=每天行数之和，日期列每天正确插入当天。"""
        import pandas as pd
        file_path = self.api.download_sku(
            biz_key="商品流量来源_搜索",
            start_date="2026-08-03",
            end_date="2026-08-05",
        )
        # ① 文件存在，文件名标注区间
        self.assertTrue(os.path.exists(file_path))
        self.assertIn("搜索流量_2026-08-03_2026-08-05.xlsx", file_path)
        # ② 输出路径含 end_date（2026-08-10 改为平面文件名，不再分子目录）
        self.assertIn("2026-08-05", file_path)
        # ③ 输出目录应在 tmp_dir 下（不再分子目录）
        self.assertTrue(file_path.startswith(self.tmp_dir))
        # ③ 合并行数 = 3天×2行 = 6行
        df = pd.read_excel(file_path, dtype=str, na_filter=False)
        self.assertEqual(len(df), 6)
        # ④ 首列是【日期】列，3个不同日期各2行
        #   ⚠️ 日期列在Excel中以datetime存储（单元格格式yyyy/m/d，打开显示2026/8/3），
        #      dtype=str 读出为 '2026-08-03 00:00:00'，取前10位日期部分断言
        self.assertEqual(df.columns[0], "日期")
        date_counts = df["日期"].str[:10].value_counts().to_dict()
        self.assertEqual(date_counts, {"2026-08-03": 2, "2026-08-04": 2, "2026-08-05": 2})

    def test_merged_file_column_structure(self):
        """合并文件的列结构与单天一致（日期列+原业务列）。"""
        import pandas as pd
        file_path = self.api.download_sku(
            biz_key="商品流量来源_搜索",
            start_date="2026-08-03",
            end_date="2026-08-05",
        )
        df = pd.read_excel(file_path, dtype=str, na_filter=False)
        self.assertEqual(list(df.columns), ["日期", "SKU", "访客数"])

    def test_warn_log_printed(self):
        """区间模式打印 CLI 警告（P1-2：不支持区间→自动拆分逐日循环）。"""
        with patch("builtins.print") as mock_print:
            self.api.download_sku(
                biz_key="商品流量来源_搜索",
                start_date="2026-08-03",
                end_date="2026-08-05",
            )
        all_text = "\n".join(str(call.args[0]) for call in mock_print.call_args_list)
        self.assertIn("不支持多日区间导出", all_text)
        self.assertIn("自动拆分为逐日循环", all_text)

    def test_over_max_days_raises_no_request(self):
        """超过31天区间 → ValueError 且不发任何 request。"""
        with self.assertRaises(ValueError):
            self.api.download_sku(
                biz_key="商品流量来源_搜索",
                start_date="2026-08-01",
                end_date="2026-09-01",  # 32天
            )
        self.api.request.assert_not_called()


# ====================== download_sku 单日模式测试（兼容性） ======================

class TestDownloadSkuSingleDayMode(unittest.TestCase):
    """单日期不进入循环，保持原有执行路径（P0 兼容要求）。"""

    def setUp(self):
        rows_by_day = {"2026-08-09": [("SKU9_1", "100")]}
        self.api, self.tmp_dir = make_flow_api(
            request_side_effect_factory(rows_by_day),
            config_date="2026-08-09",
        )

    def tearDown(self):
        shutil.rmtree(self.tmp_dir, ignore_errors=True)

    def test_single_day_no_loop(self):
        """单日（date 传入，start/end 未传）→ request 只调1次，入参=当天。"""
        file_path = self.api.download_sku(
            biz_key="商品流量来源_搜索",
            date="2026-08-09",
        )
        self.api.request.assert_called_once()
        data = self.api.request.call_args.args[1]
        self.assertEqual(data["date"], "2026-08-09")
        self.assertEqual(data["startDate"], "2026-08-09")
        self.assertEqual(data["endDate"], "2026-08-09")
        # 单日文件名格式不变
        self.assertIn("搜索流量_2026-08-09.xlsx", file_path)
        # 2026-08-10 改造：不再分子目录
        self.assertTrue(file_path.startswith(self.tmp_dir))

    def test_single_day_same_start_end(self):
        """显式传 start_date==end_date（单日区间）→ 同样不进入循环。"""
        self.api.download_sku(
            biz_key="商品流量来源_搜索",
            start_date="2026-08-09",
            end_date="2026-08-09",
        )
        self.api.request.assert_called_once()
        data = self.api.request.call_args.args[1]
        self.assertEqual(data["date"], "2026-08-09")


# ====================== 入口 ======================

if __name__ == "__main__":
    unittest.main(verbosity=2)
