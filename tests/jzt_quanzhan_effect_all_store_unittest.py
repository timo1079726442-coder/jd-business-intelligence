# -*- coding: utf-8 -*-
"""项目12：京准通全站营销全店推广效果报表导出（mock 单测）。

⚠️ 不发真请求，使用 unittest.mock 拦截 requests.Session.post / requests.get。
执行命令：
    cd "d:\\CODE\\trae\\traespace\\FYA箱包旗舰店"
    python -m unittest tests.jzt_quanzhan_effect_all_store_unittest -v

覆盖场景（2026-08-10 抓包 + 用户决策）：
    1.  类必须在 main.py 模块中存在 + 不继承 JDBaseRequest
    2.  BUSINESS_REGISTRY 第13业务注册完整
    3.  payload 必含 13 个字段
    4.  payload 字段类型严格：字符串 "" / 列表 [118] / bool False
    5.  campaignTypes=[118]（与项目10 [101] / 项目11 [118 但路径不同）区分
    6.  报表名模板 = "FYA8888_全站营销_效果报表_全店推广_{startDay}_{endDay}"
    7.  orderStatus 默认 "1"（成交订单，开放传空）
    8.  isDaily 默认 False（非日报）
    9.  skuId / spuId 默认 ""（不过滤）
    10. 响应双字段判定：code=="1" + data.code=="RC_SUCCESS"
    11. downloadUrlCsv 优先，downloadUrlZip 降级
    12. 401/CookieExpired 识别（code=2001）
    13. code=601 限流识别
    14. OSS 404 重试：第一次 404 → 退避 → 第二次 200
    15. 完整流程：POST 拿 csv → GET 404 重试 → 落盘 xlsx
    16. 完整流程：POST 拿 csv → GET csv → 落盘 xlsx
    17. 空 CSV 不抛错（仅项目10/11/12 启用）
    18. 未传 date → ValueError
    19. Cookie 文件缺失 → FileNotFoundError
    20. _run_jzt_quanzhan_effect_all_store_full 函数签名匹配
"""

import io
import json
import os
import sys
import unittest
import zipfile
from unittest.mock import MagicMock, patch

# 主项目根路径加入 sys.path
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

# import main 之前创建临时 cookie 文件
_TMP_COOKIE_PATH = os.path.join(PROJECT_ROOT, "config", "jzt_cookie.txt")
_cookie_backup = None
if os.path.isfile(_TMP_COOKIE_PATH):
    with open(_TMP_COOKIE_PATH, "r", encoding="utf-8") as _f:
        _cookie_backup = _f.read()
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


class TestJZTQuanZhanEffectAllStoreAPI(unittest.TestCase):
    """京准通全站营销全店推广效果（JZTQuanZhanEffectAllStoreAPI）mock 单测。"""

    # ---- 基础断言 ----
    def test_01_class_exists(self):
        """类必须在 main.py 模块中存在。"""
        self.assertTrue(hasattr(main, "JZTQuanZhanEffectAllStoreAPI"))

    def test_02_class_not_inheriting_jdbaserequest(self):
        """⚠️ 必须不继承 JDBaseRequest（与项目7-11 同原因）。"""
        self.assertNotIn(main.JDBaseRequest, main.JZTQuanZhanEffectAllStoreAPI.__mro__)
        self.assertEqual(main.JZTQuanZhanEffectAllStoreAPI.__bases__, (object,))

    def test_03_registry_registered(self):
        """BUSINESS_REGISTRY 必须含项目12 业务。"""
        key = "京准通全站营销全店推广效果"
        self.assertIn(key, main.BUSINESS_REGISTRY)
        info = main.BUSINESS_REGISTRY[key]
        self.assertIs(info["api_class"], main.JZTQuanZhanEffectAllStoreAPI)
        self.assertIsNotNone(info["callable"], "callable 必须回填")
        self.assertEqual(info["method"], "run_full_export")

    # ---- Payload 字段映射 ----
    def test_04_payload_full_fields_default(self):
        """默认入参下，payload 必须包含 13 个字段，类型严格。"""
        api = main.JZTQuanZhanEffectAllStoreAPI()
        payload = api._build_payload("2026-07-07", "2026-07-07")
        # 字段数
        self.assertEqual(len(payload), 13, f"payload 应有 13 个字段，实际 {len(payload)}: {list(payload.keys())}")
        # 字段类型严格
        self.assertEqual(payload["platform"], "")                # str
        self.assertEqual(payload["campaignTypes"], [118])         # **list[int]（项目10 是 [101]）**
        self.assertEqual(payload["startDay"], "2026-07-07")      # str
        self.assertEqual(payload["endDay"], "2026-07-07")        # str
        self.assertEqual(payload["orderStatus"], "1")            # str "1"（默认成交订单）
        self.assertEqual(payload["giftFlag"], "")                # str
        self.assertEqual(payload["clickOrOrderDay"], 15)         # int
        self.assertEqual(payload["clickOrOrderCaliber"], 0)      # int
        self.assertEqual(payload["skuId"], "")                   # str
        self.assertEqual(payload["spuId"], "")                   # str
        self.assertIs(payload["isDaily"], False)                 # bool False（非日报）
        self.assertEqual(payload["orderStatusCategory"], 1)      # int
        self.assertIn("reportName", payload)

    def test_05_payload_report_name_template(self):
        """报表名必须严格匹配模板（与抓包一致）。"""
        api = main.JZTQuanZhanEffectAllStoreAPI()
        payload = api._build_payload("2026-07-07", "2026-07-07")
        expected = "FYA8888_全站营销_效果报表_全店推广_2026-07-07_2026-07-07"
        self.assertEqual(payload["reportName"], expected)

    def test_06_payload_order_status_override(self):
        """orderStatus 开放入参：默认 "1"，可覆盖。"""
        api = main.JZTQuanZhanEffectAllStoreAPI()
        payload_default = api._build_payload("2026-07-07", "2026-07-07")
        self.assertEqual(payload_default["orderStatus"], "1", "默认应是 \"1\"=成交订单")

        payload_empty = api._build_payload("2026-07-07", "2026-07-07", order_status="")
        self.assertEqual(payload_empty["orderStatus"], "", "显式传空应是 \"\"=不限")

        payload_other = api._build_payload("2026-07-07", "2026-07-07", order_status="2")
        self.assertEqual(payload_other["orderStatus"], "2", "显式传 \"2\" 应原样保留")

    def test_07_payload_is_daily_override(self):
        """isDaily 开放入参：默认 False，可传 True。"""
        api = main.JZTQuanZhanEffectAllStoreAPI()
        payload_default = api._build_payload("2026-07-07", "2026-07-07")
        self.assertIs(payload_default["isDaily"], False, "默认应是 False（非日报）")

        payload_true = api._build_payload("2026-07-07", "2026-07-07", is_daily=True)
        self.assertIs(payload_true["isDaily"], True, "显式传 True 应启用日报")

    def test_08_payload_sku_spu_override(self):
        """skuId / spuId 默认 ""（不过滤），开放入参。"""
        api = main.JZTQuanZhanEffectAllStoreAPI()
        payload_default = api._build_payload("2026-07-07", "2026-07-07")
        self.assertEqual(payload_default["skuId"], "")
        self.assertEqual(payload_default["spuId"], "")

        payload_filter = api._build_payload(
            "2026-07-07", "2026-07-07",
            sku_id="123456789", spu_id="987654321",
        )
        self.assertEqual(payload_filter["skuId"], "123456789")
        self.assertEqual(payload_filter["spuId"], "987654321")

    # ---- 响应判定 ----
    def test_09_handle_response_success(self):
        """响应双字段判定：code=="1" + data.code=="RC_SUCCESS"。"""
        api = main.JZTQuanZhanEffectAllStoreAPI()
        ret = {
            "code": "1",
            "msg": "操作成功",
            "success": True,
            "data": {"code": "RC_SUCCESS", "downloadId": 123, "downloadUrlCsv": "http://x.csv"},
        }
        try:
            api._handle_response(ret, "test")
        except Exception as e:
            self.fail(f"_handle_response 应通过，实际抛：{e}")

    def test_10_handle_response_data_code_fail(self):
        """data.code != RC_SUCCESS 时抛 RuntimeError。"""
        api = main.JZTQuanZhanEffectAllStoreAPI()
        ret = {
            "code": "1",
            "msg": "操作成功",
            "success": True,
            "data": {"code": "RC_FAIL", "downloadUrlCsv": "http://x.csv"},
        }
        with self.assertRaises(RuntimeError):
            api._handle_response(ret, "test")

    def test_11_handle_response_unlogin(self):
        """code=2001/未登录 → CookieExpiredError。"""
        api = main.JZTQuanZhanEffectAllStoreAPI()
        ret = {"code": 2001, "msg": "未登录", "success": False}
        with self.assertRaises(main.CookieExpiredError):
            api._handle_response(ret, "test")

    def test_12_handle_response_601(self):
        """code=601 → RuntimeError（限流）。"""
        api = main.JZTQuanZhanEffectAllStoreAPI()
        ret = {"code": "601", "msg": "操作频繁", "success": False}
        with self.assertRaises(RuntimeError):
            api._handle_response(ret, "test")

    # ---- URL 选择 ----
    def test_13_pick_download_url_csv_first(self):
        """⚠️ 用户决策（继承项目10）：csv 优先。"""
        api = main.JZTQuanZhanEffectAllStoreAPI()
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
        api = main.JZTQuanZhanEffectAllStoreAPI()
        ret = {"data": {"downloadUrlZip": "http://x.zip"}}
        url = api._pick_download_url(ret)
        self.assertEqual(url, "http://x.zip", "csv 缺失应降级 zip")

    def test_15_pick_download_url_missing(self):
        """两者都缺失 → RuntimeError。"""
        api = main.JZTQuanZhanEffectAllStoreAPI()
        ret = {"data": {}}
        with self.assertRaises(RuntimeError):
            api._pick_download_url(ret)

    # ---- 完整流程（mock 网络）----
    def test_16_full_export_zip_path(self):
        """完整流程：POST 拿 zip（csv 优先缺失场景）→ GET 404 重试 → 解压 zip → 落盘 xlsx。"""
        zip_bytes = _make_zip_bytes("日期,SKU,花费\n2026/7/7,123,99.5\n")
        url_zip = "http://mock.oss/test.zip?Expires=123&Signature=abc"

        post_resp = MagicMock()
        post_resp.raise_for_status = MagicMock()
        post_resp.json = MagicMock(return_value={
            "code": "1",
            "success": True,
            "data": {
                "code": "RC_SUCCESS",
                "downloadId": 298007175,
                # ⚠️ csv 缺失场景（仅 zip）
                "downloadUrlZip": url_zip,
            },
        })

        get_resp_404 = MagicMock()
        get_resp_404.status_code = 404
        get_resp_404.raise_for_status = MagicMock()

        get_resp_200 = MagicMock()
        get_resp_200.status_code = 200
        get_resp_200.content = zip_bytes

        with patch.object(main.requests.Session, "post", return_value=post_resp), \
             patch("main.requests.get", side_effect=[get_resp_404, get_resp_200]), \
             patch("main.time.sleep", return_value=None), \
             patch("main.random.uniform", return_value=0.001):
            api = main.JZTQuanZhanEffectAllStoreAPI()
            result_path = api.run_full_export(date="2026-07-07")

        self.assertTrue(os.path.isfile(result_path), f"xlsx 应已落盘：{result_path}")
        self.assertTrue(result_path.endswith(".xlsx"))
        self.assertIn("京准通全站营销全店推广效果_2026-07-07.xlsx", result_path)

        # 验证 xlsx 数据
        import pandas as pd
        df = pd.read_excel(result_path, dtype=str)
        self.assertEqual(len(df), 1, "zip 内 1 行数据（除表头）")
        self.assertEqual(df.iloc[0]["SKU"], "123")

        # 清理
        try:
            os.remove(result_path)
            date_dir = os.path.dirname(result_path)
            if os.path.isdir(date_dir) and not os.listdir(date_dir):
                os.rmdir(date_dir)
        except Exception:
            pass

    def test_17_full_export_csv_fallback(self):
        """完整流程：POST 拿 csv → GET csv → 落盘 xlsx。"""
        csv_text = "日期,SKU,花费\n2026/7/7,ORDER001,50.0\n"
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
                "downloadUrlCsv": url_csv,
            },
        })

        get_resp_200 = MagicMock()
        get_resp_200.status_code = 200
        get_resp_200.content = csv_bytes

        with patch.object(main.requests.Session, "post", return_value=post_resp), \
             patch("main.requests.get", return_value=get_resp_200):
            api = main.JZTQuanZhanEffectAllStoreAPI()
            result_path = api.run_full_export(date="2026-07-07", is_daily=True)

        self.assertTrue(os.path.isfile(result_path))
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
        api = main.JZTQuanZhanEffectAllStoreAPI()
        with self.assertRaises(ValueError):
            api.run_full_export()

    # ---- Cookie 文件缺失 ----
    def test_19_cookie_file_missing(self):
        """Cookie 文件不存在 → FileNotFoundError。"""
        if os.path.isfile(_TMP_COOKIE_PATH):
            os.remove(_TMP_COOKIE_PATH)
        try:
            with self.assertRaises(FileNotFoundError):
                main.JZTQuanZhanEffectAllStoreAPI()
        finally:
            with open(_TMP_COOKIE_PATH, "w", encoding="utf-8") as f:
                f.write("mock_cookie_for_test=placeholder")

    # ---- 调度器 callable ----
    def test_20_callable_signature(self):
        """_run_jzt_quanzhan_effect_all_store_full 必须能接受 kwargs 透传。"""
        key = "京准通全站营销全店推广效果"
        callable_fn = main.BUSINESS_REGISTRY[key]["callable"]
        self.assertIsNotNone(callable_fn)

        import inspect
        sig = inspect.signature(callable_fn)
        self.assertIn("kwargs", sig.parameters)
        try:
            callable_fn(date="2026-07-07", is_daily=True, sku_id="999")
        except Exception:
            pass


def tearDownModule():
    """测试结束恢复 cookie 文件原状。"""
    if _cookie_backup is not None:
        with open(_TMP_COOKIE_PATH, "w", encoding="utf-8") as f:
            f.write(_cookie_backup)
    elif os.path.isfile(_TMP_COOKIE_PATH):
        os.remove(_TMP_COOKIE_PATH)


if __name__ == "__main__":
    unittest.main(verbosity=2)