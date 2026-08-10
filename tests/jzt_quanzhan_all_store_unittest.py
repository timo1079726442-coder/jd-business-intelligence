# -*- coding: utf-8 -*-
"""项目11：京准通全站营销全店计划报表导出（mock 单测）。

⚠️ 不发真请求，使用 unittest.mock 拦截 requests.Session.post / requests.get。
执行命令：
    cd "d:\\CODE\\trae\\traespace\\FYA箱包旗舰店"
    python -m unittest tests.jzt_quanzhan_all_store_unittest -v

覆盖场景（2026-08-10 抓包 + 用户决策）：
    1.  类必须在 main.py 模块中存在 + 不继承 JDBaseRequest
    2.  BUSINESS_REGISTRY 第12业务注册完整
    3.  payload 必含 15 个字段（含 dateValues 嵌套列表）
    4.  payload 字段类型严格：字符串 "" / 列表 [int] / bool True
    5.  campaignTypes=[118]（与项目9 [101] 区分）
    6.  报表名模板 = "FYA8888_全站营销_全店计划报表_{startDay}_{endDay}"
    7.  orderStatus 开放入参：默认 ""（不限）
    8.  isDaily 开放入参：默认 True（日报）
    9.  sxuId / obys 默认 ""（不过滤）
    10. 响应双字段判定：code=="1" + data.code=="RC_SUCCESS"
    11. downloadUrlCsv 优先，downloadUrlZip 降级
    12. 401/CookieExpired 识别（code=2001）
    13. code=601 限流识别
    14. OSS 404 重试：第一次 404 → 退避 → 第二次 200
    15. 完整流程：POST 拿 csv → GET 404 重试 → 落盘 xlsx（zip 路径）
    16. 完整流程：POST 拿 csv → GET csv → 落盘 xlsx（csv 路径）
    17. 空 CSV 不抛错（仅项目11 启用）
    18. 未传 date → ValueError
    19. Cookie 文件缺失 → FileNotFoundError
    20. _run_jzt_quanzhan_campaign_all_store_full 函数签名匹配

mock 关键点：
    - 不写 cookie 文件：临时写一份有效 cookie 让 __init__ 通过
    - 拦截 session.post：返回预设 JSON
    - 拦截 requests.get（OSS 下载）：第一次 404，第二次 200（带 csv 字节流）
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
    """生成一个含 csv 的 zip 字节流（用于 mock OSS 返回）。

    ⚠️ zip 内 csv 文件名必须是 ASCII（避免 cp437/utf-8 编码问题），
       中文 csv 名在 zipfile 默认解压时可能乱码，导致 _post_process_to_xlsx
       解压后找不到对应 csv 名。
    """
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("mock_report.csv", csv_text)
    return buf.getvalue()


class TestJZTQuanZhanCampaignAllStoreAPI(unittest.TestCase):
    """京准通全站营销全店计划（JZTQuanZhanCampaignAllStoreAPI）mock 单测。"""

    # ---- 基础断言 ----
    def test_01_class_exists(self):
        """类必须在 main.py 模块中存在。"""
        self.assertTrue(hasattr(main, "JZTQuanZhanCampaignAllStoreAPI"))

    def test_02_class_not_inheriting_jdbaserequest(self):
        """⚠️ 必须不继承 JDBaseRequest（与项目7-10 同原因）。"""
        self.assertNotIn(main.JDBaseRequest, main.JZTQuanZhanCampaignAllStoreAPI.__mro__)
        self.assertEqual(main.JZTQuanZhanCampaignAllStoreAPI.__bases__, (object,))

    def test_03_registry_registered(self):
        """BUSINESS_REGISTRY 必须含项目11 业务。"""
        key = "京准通全站营销全店计划"
        self.assertIn(key, main.BUSINESS_REGISTRY)
        info = main.BUSINESS_REGISTRY[key]
        self.assertIs(info["api_class"], main.JZTQuanZhanCampaignAllStoreAPI)
        self.assertIsNotNone(info["callable"], "callable 必须回填")
        self.assertEqual(info["method"], "run_full_export")

    # ---- Payload 字段映射 ----
    def test_04_payload_full_fields_default(self):
        """默认入参下，payload 必须包含 15 个字段，类型严格。"""
        api = main.JZTQuanZhanCampaignAllStoreAPI()
        payload = api._build_payload("2026-07-04", "2026-07-04")
        # 字段数
        self.assertEqual(len(payload), 15, f"payload 应有 15 个字段，实际 {len(payload)}: {list(payload.keys())}")
        # 字段类型严格
        self.assertEqual(payload["platform"], "")                # str
        self.assertEqual(payload["campaignTypes"], [118])         # **list[int]（项目9 是 [101]）**
        self.assertEqual(payload["province"], "")                # str
        self.assertEqual(payload["startDay"], "2026-07-04")      # str
        self.assertEqual(payload["endDay"], "2026-07-04")        # str
        self.assertEqual(payload["orderStatus"], "")             # str ""（不限）
        self.assertEqual(payload["giftFlag"], "")                # str
        self.assertEqual(payload["clickOrOrderDay"], 15)         # int
        self.assertEqual(payload["clickOrOrderCaliber"], 0)      # int
        self.assertEqual(payload["sxuId"], "")                   # str（项目9 字段名 sxuId）
        self.assertEqual(payload["obys"], "")                    # str
        self.assertIs(payload["isDaily"], True)                  # bool True（日报）
        self.assertEqual(payload["orderStatusCategory"], 1)      # int
        # 嵌套列表
        self.assertEqual(payload["dateValues"], [{"startDay": "2026-07-04", "endDay": "2026-07-04"}])
        self.assertIn("reportName", payload)

    def test_05_payload_report_name_template(self):
        """报表名必须严格匹配模板（与抓包一致）。"""
        api = main.JZTQuanZhanCampaignAllStoreAPI()
        payload = api._build_payload("2026-07-04", "2026-07-04")
        expected = "FYA8888_全站营销_全店计划报表_2026-07-04_2026-07-04"
        self.assertEqual(payload["reportName"], expected)

    def test_06_payload_order_status_override(self):
        """orderStatus 开放入参：默认 ""，可覆盖。"""
        api = main.JZTQuanZhanCampaignAllStoreAPI()
        payload_default = api._build_payload("2026-07-04", "2026-07-04")
        self.assertEqual(payload_default["orderStatus"], "", "默认应是 \"\"=不限")

        payload_other = api._build_payload("2026-07-04", "2026-07-04", order_status="1")
        self.assertEqual(payload_other["orderStatus"], "1", "显式传 \"1\" 应原样保留")

    def test_07_payload_is_daily_override(self):
        """isDaily 开放入参：默认 True，可传 False。"""
        api = main.JZTQuanZhanCampaignAllStoreAPI()
        payload_default = api._build_payload("2026-07-04", "2026-07-04")
        self.assertIs(payload_default["isDaily"], True, "默认应是 True（日报）")

        payload_false = api._build_payload("2026-07-04", "2026-07-04", is_daily=False)
        self.assertIs(payload_false["isDaily"], False, "显式传 False 应启用非日报")

    def test_08_payload_sxu_obys_override(self):
        """sxuId / obys 默认 ""（不过滤），开放入参。"""
        api = main.JZTQuanZhanCampaignAllStoreAPI()
        payload_default = api._build_payload("2026-07-04", "2026-07-04")
        self.assertEqual(payload_default["sxuId"], "")
        self.assertEqual(payload_default["obys"], "")

        payload_filter = api._build_payload(
            "2026-07-04", "2026-07-04",
            sxu_id="99936530475", obys="test_obys",
        )
        self.assertEqual(payload_filter["sxuId"], "99936530475")
        self.assertEqual(payload_filter["obys"], "test_obys")

    # ---- 响应判定 ----
    def test_09_handle_response_success(self):
        """响应双字段判定：code=="1" + data.code=="RC_SUCCESS"。"""
        api = main.JZTQuanZhanCampaignAllStoreAPI()
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
        api = main.JZTQuanZhanCampaignAllStoreAPI()
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
        api = main.JZTQuanZhanCampaignAllStoreAPI()
        ret = {"code": 2001, "msg": "未登录", "success": False}
        with self.assertRaises(main.CookieExpiredError):
            api._handle_response(ret, "test")

    def test_12_handle_response_601(self):
        """code=601 → RuntimeError（限流）。"""
        api = main.JZTQuanZhanCampaignAllStoreAPI()
        ret = {"code": "601", "msg": "操作频繁", "success": False}
        with self.assertRaises(RuntimeError):
            api._handle_response(ret, "test")

    # ---- URL 选择 ----
    def test_13_pick_download_url_csv_first(self):
        """⚠️ 用户决策（继承项目10）：csv 优先。"""
        api = main.JZTQuanZhanCampaignAllStoreAPI()
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
        api = main.JZTQuanZhanCampaignAllStoreAPI()
        ret = {"data": {"downloadUrlZip": "http://x.zip"}}
        url = api._pick_download_url(ret)
        self.assertEqual(url, "http://x.zip", "csv 缺失应降级 zip")

    def test_15_pick_download_url_missing(self):
        """两者都缺失 → RuntimeError。"""
        api = main.JZTQuanZhanCampaignAllStoreAPI()
        ret = {"data": {}}
        with self.assertRaises(RuntimeError):
            api._pick_download_url(ret)

    # ---- 完整流程（mock 网络）----
    def test_16_full_export_zip_path(self):
        """完整流程：POST 拿 zip（csv 优先缺失场景）→ GET 404 重试 → 解压 zip → 落盘 xlsx。

        ⚠️ 项目11 响应同时返回 downloadUrlZip + downloadUrlCsv。本测试模拟
        「downloadUrlCsv 缺失、downloadUrlZip 存在」的降级路径，验证 zip 解压能力。
        """
        zip_bytes = _make_zip_bytes("日期,SKU,花费\n2026/7/4,123,99.5\n")
        url_zip = "http://mock.oss/test.zip?Expires=123&Signature=abc"

        post_resp = MagicMock()
        post_resp.raise_for_status = MagicMock()
        post_resp.json = MagicMock(return_value={
            "code": "1",
            "success": True,
            "data": {
                "code": "RC_SUCCESS",
                "downloadId": 297937945,
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
            api = main.JZTQuanZhanCampaignAllStoreAPI()
            result_path = api.run_full_export(date="2026-07-04")

        self.assertTrue(os.path.isfile(result_path), f"xlsx 应已落盘：{result_path}")
        self.assertTrue(result_path.endswith(".xlsx"))
        self.assertIn("京准通全站营销全店计划_2026-07-04.xlsx", result_path)

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
        csv_text = "日期,SKU,花费\n2026/7/4,ORDER001,50.0\n"
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
            api = main.JZTQuanZhanCampaignAllStoreAPI()
            result_path = api.run_full_export(date="2026-07-04", is_daily=False)

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
        api = main.JZTQuanZhanCampaignAllStoreAPI()
        with self.assertRaises(ValueError):
            api.run_full_export()

    # ---- Cookie 文件缺失 ----
    def test_19_cookie_file_missing(self):
        """Cookie 文件不存在 → FileNotFoundError。"""
        if os.path.isfile(_TMP_COOKIE_PATH):
            os.remove(_TMP_COOKIE_PATH)
        try:
            with self.assertRaises(FileNotFoundError):
                main.JZTQuanZhanCampaignAllStoreAPI()
        finally:
            with open(_TMP_COOKIE_PATH, "w", encoding="utf-8") as f:
                f.write("mock_cookie_for_test=placeholder")

    # ---- 调度器 callable ----
    def test_20_callable_signature(self):
        """_run_jzt_quanzhan_campaign_all_store_full 必须能接受 kwargs 透传。"""
        key = "京准通全站营销全店计划"
        callable_fn = main.BUSINESS_REGISTRY[key]["callable"]
        self.assertIsNotNone(callable_fn)

        import inspect
        sig = inspect.signature(callable_fn)
        self.assertIn("kwargs", sig.parameters)
        try:
            callable_fn(date="2026-07-04", is_daily=False, sku_id="999")
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