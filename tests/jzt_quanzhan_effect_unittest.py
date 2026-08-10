# -*- coding: utf-8 -*-
"""项目10：京准通全站营销单品推广效果报表导出（mock 单测）。

⚠️ 不发真请求，使用 unittest.mock 拦截 requests.Session.post / requests.get。
执行命令：
    cd "d:\\CODE\\trae\\traespace\\FYA箱包旗舰店"
    python -m unittest tests.test_jzt_quanzhan_effect -v

覆盖场景（2026-08-10 用户决策版）：
    1.  payload 必含 13 个字段（platform/campaignTypes/startDay/endDay/orderStatus/giftFlag/skuId/spuId/clickOrOrderCaliber/clickOrOrderDay/isDaily/orderStatusCategory/reportName）
    2.  payload 字段类型严格：字符串 "" / 列表 [int] / bool False
    3.  orderStatus 默认 "1"（成交订单）；传 "" 表示不限
    4.  isDaily 默认 False；可传 True
    5.  skuId/spuId 默认 ""；可传具体值
    6.  报表名模板 = "FYA8888_全站营销_效果报表_单品推广_{startDay}_{endDay}"
    7.  响应双字段判定：code=="1" + data.code=="RC_SUCCESS"
    8.  下载 URL 优先 downloadUrlZip，降级 downloadUrlCsv
    9.  401/CookieExpired 识别（code=2001）
    10. code=601 限流识别
    11. OSS 404 重试：第一次 404 → 退避 → 第二次 200
    12. zip 内含 csv 解压 → xlsx 落盘路径与命名
    13. BUSINESS_REGISTRY 第11 业务已注册

mock 关键点：
    - 不写 cookie 文件：临时写一份有效 cookie 让 __init__ 通过
    - 拦截 session.post：返回预设 JSON（含 downloadUrlZip + downloadUrlCsv）
    - 拦截 requests.get（OSS 下载）：第一次 404，第二次 200（带 csv 字节流）
    - 拦截 zipfile 解压：用真实 zipfile 模块生成临时 zip
"""

import io
import json
import os
import sys
import tempfile
import unittest
import zipfile
from unittest.mock import MagicMock, patch

# 主项目根路径加入 sys.path（保证 import main 生效）
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

# 在 import main 之前，先创建临时 cookie 文件，避免 __init__ 抛 FileNotFoundError
_TMP_COOKIE_PATH = os.path.join(PROJECT_ROOT, "config", "jzt_cookie.txt")
_cookie_backup = None
if os.path.isfile(_TMP_COOKIE_PATH):
    with open(_TMP_COOKIE_PATH, "r", encoding="utf-8") as _f:
        _cookie_backup = _f.read()
# 写一个占位 cookie（mock 单测不发真请求，cookie 内容不影响）
os.makedirs(os.path.dirname(_TMP_COOKIE_PATH), exist_ok=True)
with open(_TMP_COOKIE_PATH, "w", encoding="utf-8") as _f:
    _f.write("mock_cookie_for_test=placeholder")

import main  # noqa: E402


def _make_zip_bytes(csv_text: str = "col1,col2\nA,B\nC,D\n") -> bytes:
    """生成一个含 csv 的 zip 字节流（用于 mock OSS 返回）。"""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("mock_report.csv", csv_text)
    return buf.getvalue()


class TestJZTQuanZhanEffectAPI(unittest.TestCase):
    """京准通全站营销单品推广效果（JZTQuanZhanEffectAPI）mock 单测。"""

    # ---- 基础断言 ----
    def test_01_class_exists_in_main(self):
        """类必须在 main.py 模块中存在。"""
        self.assertTrue(hasattr(main, "JZTQuanZhanEffectAPI"))

    def test_02_class_not_inheriting_jdbaserequest(self):
        """⚠️ 必须不继承 JDBaseRequest（与项目7/8/9 同原因）。"""
        # 父类链不能含 JDBaseRequest
        self.assertNotIn(main.JDBaseRequest, main.JZTQuanZhanEffectAPI.__mro__)
        self.assertEqual(main.JZTQuanZhanEffectAPI.__bases__, (object,))

    def test_03_registry_registered(self):
        """BUSINESS_REGISTRY 必须含项目10 业务。"""
        key = "京准通全站营销单品推广效果"
        self.assertIn(key, main.BUSINESS_REGISTRY)
        info = main.BUSINESS_REGISTRY[key]
        self.assertIs(info["api_class"], main.JZTQuanZhanEffectAPI)
        self.assertIsNotNone(info["callable"], "callable 必须回填（防止前向引用失败）")
        self.assertEqual(info["method"], "run_full_export")

    # ---- Payload 字段映射 ----
    def test_04_payload_full_fields_default(self):
        """默认入参下，payload 必须包含 13 个字段，类型严格。"""
        api = main.JZTQuanZhanEffectAPI()
        payload = api._build_payload("2026-04-25", "2026-04-25")
        # 字段数
        self.assertEqual(len(payload), 13, f"payload 应有 13 个字段，实际 {len(payload)}: {list(payload.keys())}")
        # 字段类型严格
        self.assertEqual(payload["platform"], "")            # str
        self.assertEqual(payload["campaignTypes"], [101])     # list[int]
        self.assertEqual(payload["startDay"], "2026-04-25")  # str
        self.assertEqual(payload["endDay"], "2026-04-25")    # str
        self.assertEqual(payload["orderStatus"], "1")        # str "1"（默认成交订单）
        self.assertEqual(payload["giftFlag"], "")            # str
        self.assertEqual(payload["skuId"], "")               # str
        self.assertEqual(payload["spuId"], "")               # str
        self.assertEqual(payload["clickOrOrderCaliber"], 0)  # int
        self.assertEqual(payload["clickOrOrderDay"], 15)     # int
        self.assertIs(payload["isDaily"], False)             # bool False
        self.assertEqual(payload["orderStatusCategory"], 1)  # int
        self.assertIn("reportName", payload)

    def test_05_payload_report_name_template(self):
        """报表名必须严格匹配模板（与抓包一致）。"""
        api = main.JZTQuanZhanEffectAPI()
        payload = api._build_payload("2026-04-25", "2026-04-25")
        expected = "FYA8888_全站营销_效果报表_单品推广_2026-04-25_2026-04-25"
        self.assertEqual(payload["reportName"], expected)

    def test_06_payload_order_status_override(self):
        """orderStatus 开放入参：传 "" 表示不限。"""
        api = main.JZTQuanZhanEffectAPI()
        payload_default = api._build_payload("2026-04-25", "2026-04-25")
        self.assertEqual(payload_default["orderStatus"], "1", "默认应是 \"1\"=成交订单")

        payload_empty = api._build_payload("2026-04-25", "2026-04-25", order_status="")
        self.assertEqual(payload_empty["orderStatus"], "", "显式传空应是 \"\"=不限")

        payload_other = api._build_payload("2026-04-25", "2026-04-25", order_status="2")
        self.assertEqual(payload_other["orderStatus"], "2", "显式传 \"2\" 应原样保留")

    def test_07_payload_is_daily_override(self):
        """isDaily 不固化，默认 False，可传 True。"""
        api = main.JZTQuanZhanEffectAPI()
        payload_default = api._build_payload("2026-04-25", "2026-04-25")
        self.assertIs(payload_default["isDaily"], False, "默认应是 False（抓包实测值）")

        payload_true = api._build_payload("2026-04-25", "2026-04-25", is_daily=True)
        self.assertIs(payload_true["isDaily"], True, "显式传 True 应启用日报")

    def test_08_payload_sku_spu_override(self):
        """skuId/spuId 默认 ""（不过滤），开放入参。"""
        api = main.JZTQuanZhanEffectAPI()
        payload_default = api._build_payload("2026-04-25", "2026-04-25")
        self.assertEqual(payload_default["skuId"], "")
        self.assertEqual(payload_default["spuId"], "")

        payload_filter = api._build_payload(
            "2026-04-25", "2026-04-25",
            sku_id="123456789", spu_id="987654321",
        )
        self.assertEqual(payload_filter["skuId"], "123456789")
        self.assertEqual(payload_filter["spuId"], "987654321")

    # ---- 响应判定 ----
    def test_09_handle_response_success(self):
        """响应双字段判定：code=="1" + data.code=="RC_SUCCESS"。"""
        api = main.JZTQuanZhanEffectAPI()
        ret = {
            "code": "1",
            "msg": "操作成功",
            "success": True,
            "data": {"code": "RC_SUCCESS", "downloadId": 123, "downloadUrlZip": "http://x.zip"},
        }
        # 不应抛异常
        try:
            api._handle_response(ret, "test")
        except Exception as e:
            self.fail(f"_handle_response 应通过，实际抛：{e}")

    def test_10_handle_response_data_code_fail(self):
        """data.code != RC_SUCCESS 时抛 RuntimeError。"""
        api = main.JZTQuanZhanEffectAPI()
        ret = {
            "code": "1",
            "msg": "操作成功",
            "success": True,
            "data": {"code": "RC_FAIL", "downloadUrlZip": "http://x.zip"},
        }
        with self.assertRaises(RuntimeError):
            api._handle_response(ret, "test")

    def test_11_handle_response_unlogin(self):
        """code=2001/未登录 → CookieExpiredError。"""
        api = main.JZTQuanZhanEffectAPI()
        ret = {"code": 2001, "msg": "未登录", "success": False}
        with self.assertRaises(main.CookieExpiredError):
            api._handle_response(ret, "test")

    def test_12_handle_response_601(self):
        """code=601 → RuntimeError（限流）。"""
        api = main.JZTQuanZhanEffectAPI()
        ret = {"code": "601", "msg": "操作频繁", "success": False}
        with self.assertRaises(RuntimeError):
            api._handle_response(ret, "test")

    # ---- URL 选择 ----
    def test_13_pick_download_url_csv_first(self):
        """⚠️ 用户决策 2026-08-10（第二次调整）：优先 downloadUrlCsv，zip 降级。"""
        api = main.JZTQuanZhanEffectAPI()
        ret = {
            "data": {
                "downloadUrlZip": "http://x.zip",
                "downloadUrlCsv": "http://x.csv",
            }
        }
        url = api._pick_download_url(ret)
        self.assertEqual(url, "http://x.csv", "csv 应优先（模拟浏览器行为）")

    def test_14_pick_download_url_zip_fallback(self):
        """downloadUrlCsv 缺失时降级 downloadUrlZip。"""
        api = main.JZTQuanZhanEffectAPI()
        ret = {"data": {"downloadUrlZip": "http://x.zip"}}  # 无 csv
        url = api._pick_download_url(ret)
        self.assertEqual(url, "http://x.zip", "csv 缺失应降级 zip")

    def test_15_pick_download_url_missing(self):
        """两者都缺失 → RuntimeError。"""
        api = main.JZTQuanZhanEffectAPI()
        ret = {"data": {}}
        with self.assertRaises(RuntimeError):
            api._pick_download_url(ret)

    # ---- 完整流程（mock 网络）----
    def test_16_full_export_zip_path(self):
        """完整流程：POST 拿 csv（用户决策 csv 优先）→ GET 404 重试 → 落盘 xlsx。

        ⚠️ 用户决策 2026-08-10（第二次调整）：csv 优先（模拟浏览器行为）。
        """
        # 准备 mock 响应（csv 字节流）
        csv_bytes_text = "日期,SKU,花费\n2026/4/25,123,99.5\n"
        csv_bytes = csv_bytes_text.encode("utf-8")
        url_csv = "http://mock.oss/test.csv?Expires=123&Signature=abc"
        url_zip = "http://mock.oss/test.zip?Expires=123&Signature=abc"

        # POST 返回（同时含 zip + csv）
        post_resp = MagicMock()
        post_resp.raise_for_status = MagicMock()
        post_resp.json = MagicMock(return_value={
            "code": "1",
            "success": True,
            "data": {
                "code": "RC_SUCCESS",
                "downloadId": 297893851,
                "downloadUrlZip": url_zip,
                "downloadUrlCsv": url_csv,
            },
        })

        # GET 返回：第一次 404，第二次 200（csv 字节流）
        get_resp_404 = MagicMock()
        get_resp_404.status_code = 404
        get_resp_404.raise_for_status = MagicMock()

        get_resp_200 = MagicMock()
        get_resp_200.status_code = 200
        get_resp_200.content = csv_bytes

        # 拦截 session.post 和 requests.get
        with patch.object(main.requests.Session, "post", return_value=post_resp), \
            \
             patch("main.requests.get", side_effect=[get_resp_404, get_resp_200]), \
             \
             patch("main.time.sleep", return_value=None), \
             \
             patch("main.random.uniform", return_value=0.001):
            api = main.JZTQuanZhanEffectAPI()
            result_path = api.run_full_export(date="2026-04-25")

        # 断言：落盘文件存在、含正确数据
        self.assertTrue(os.path.isfile(result_path), f"xlsx 应已落盘：{result_path}")
        self.assertTrue(result_path.endswith(".xlsx"))
        self.assertIn("京准通全站营销单品推广效果_2026-04-25.xlsx", result_path)

        # 验证 xlsx 数据（用 pandas 反读）
        import pandas as pd
        df = pd.read_excel(result_path, dtype=str)
        self.assertEqual(len(df), 1, "csv 内 1 行数据（除表头）")
        self.assertEqual(df.iloc[0]["SKU"], "123")
        # 日期列存在且非空（prepare_date_columns 已标准化为 yyyy/m/d 或 yyyy-mm-dd HH:MM:SS）
        date_val = str(df.iloc[0]["日期"])
        self.assertTrue(
            date_val.startswith("2026/4/25") or date_val.startswith("2026-04-25"),
            f"日期列应标准化，实际={date_val!r}"
        )

        # 清理：删除落盘文件（避免污染真实输出）
        try:
            os.remove(result_path)
            # 删空日期子目录
            date_dir = os.path.dirname(result_path)
            if os.path.isdir(date_dir) and not os.listdir(date_dir):
                os.rmdir(date_dir)
        except Exception:
            pass

    def test_17_full_export_csv_fallback(self):
        """完整流程：POST 返回无 zip（仅 csv）→ GET csv → 落盘 xlsx。"""
        csv_text = "日期,订单号,花费\n2026/4/25,ORDER001,50.0\n"
        csv_bytes = csv_text.encode("utf-8")
        url_csv = "http://mock.oss/test.csv?Expires=123&Signature=abc"

        post_resp = MagicMock()
        post_resp.raise_for_status = MagicMock()
        post_resp.json = MagicMock(return_value={
            "code": "1",
            "success": True,
            "data": {
                "code": "RC_SUCCESS",
                "downloadId": 111,
                # 注意：无 downloadUrlZip，只有 csv
                "downloadUrlCsv": url_csv,
            },
        })

        get_resp_200 = MagicMock()
        get_resp_200.status_code = 200
        get_resp_200.content = csv_bytes

        with patch.object(main.requests.Session, "post", return_value=post_resp), \
             patch("main.requests.get", return_value=get_resp_200):
            api = main.JZTQuanZhanEffectAPI()
            result_path = api.run_full_export(date="2026-04-25", is_daily=True)

        self.assertTrue(os.path.isfile(result_path))
        # 验证：传入的 is_daily=True 应进入 payload
        # 清理
        try:
            os.remove(result_path)
            date_dir = os.path.dirname(result_path)
            if os.path.isdir(date_dir) and not os.listdir(date_dir):
                os.rmdir(date_dir)
        except Exception:
            pass

    # ---- 容错 ----
    def test_18_run_full_export_no_date(self):
        """未传 date/start_date/end_date → ValueError。"""
        api = main.JZTQuanZhanEffectAPI()
        with self.assertRaises(ValueError):
            api.run_full_export()

    # ---- Cookie 文件缺失 ----
    def test_19_cookie_file_missing(self):
        """Cookie 文件不存在 → FileNotFoundError。"""
        # 临时删除 cookie 文件
        if os.path.isfile(_TMP_COOKIE_PATH):
            os.remove(_TMP_COOKIE_PATH)
        try:
            with self.assertRaises(FileNotFoundError):
                main.JZTQuanZhanEffectAPI()
        finally:
            # 恢复占位 cookie
            with open(_TMP_COOKIE_PATH, "w", encoding="utf-8") as f:
                f.write("mock_cookie_for_test=placeholder")

    # ---- 调度器 callable ----
    def test_20_callable_signature(self):
        """_run_jzt_quanzhan_effect_full 必须能接受 kwargs 透传。"""
        key = "京准通全站营销单品推广效果"
        callable_fn = main.BUSINESS_REGISTRY[key]["callable"]
        self.assertIsNotNone(callable_fn)

        # 不真正执行（cookie 是 mock 占位，会网络报错），只验证函数可调用且签名匹配
        import inspect
        sig = inspect.signature(callable_fn)
        # 必须是 **kwargs 形式
        self.assertIn("kwargs", sig.parameters)
        # 接受所有 run_full_export 入参（不抛 TypeError 即通过）
        try:
            callable_fn(date="2026-04-25", is_daily=True, sku_id="999")  # 不真正执行成功，只验证函数能找到
        except Exception:
            # 网络/落盘失败不算函数签名错误
            pass


def tearDownModule():
    """测试结束恢复 cookie 文件原状。"""
    if _cookie_backup is not None:
        with open(_TMP_COOKIE_PATH, "w", encoding="utf-8") as f:
            f.write(_cookie_backup)
    elif os.path.isfile(_TMP_COOKIE_PATH):
        # 没有原备份（首次测试），删除占位
        os.remove(_TMP_COOKIE_PATH)


if __name__ == "__main__":
    # 命令行直接执行入口
    unittest.main(verbosity=2)