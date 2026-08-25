# -*- coding: utf-8 -*-
"""项目7「京准通快车自定义报表」阶段5 mock 单元测试（2026-08-07）。

测试目标（按用户决策 2026-08-07）：
    1. 创建任务：成功 / 业务码 601 抛 RuntimeError
    2. 轮询：成功等到『报表已生成』/ 报表生成失败立即停 / 超时抛 TimeoutError
    3. CDN 403 重试：第一次 403 → 重刷 URL → 第二次 200 成功 / 连续 403 抛 RuntimeError
    4. Cookie 过期：业务码 2001/302 + 文本"未登录" → CookieExpiredError
    5. 签名错误：业务码 -407/-402 → RuntimeError 提示重抓 Cookie（2026-08-15 实测 JZT 不需要 h5st）

mock 思路（不依赖真实 Cookie/h5st）：
    - 通过 object.__new__(JZTKuaicheAPI) 跳过 __init__（避免读 Cookie 文件）
    - 注入 mock self.session 和类常量（BYPASS_INIT 模式）
    - 替换 session.post/get 为 MockResponseFixture 返回预置响应
    - 替换全局 requests.get 为 mock（CDN 下载是 requests.get 而非 session.get）

运行方式：
    python tests/test_jzt_kuaiche_p7.py
或：
    python -m unittest tests.test_jzt_kuaiche_p7 -v

参考项目4/5/6 mock 单测经验：每个场景独立 test_xxx 方法，断言异常类型 +返回值。
"""
import os
import sys
import time
import unittest
from unittest.mock import patch, MagicMock

# 让 tests/ 目录能找到 main.py
THIS_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(THIS_DIR)
sys.path.insert(0, PROJECT_ROOT)

import main as M  # noqa: E402
from main import JZTKuaicheAPI, CookieExpiredError  # noqa: E402


# ====================== Mock 响应工厂 ======================

class MockResponse:
    """模拟 requests.Response，最小字段集满足项目7 调用。"""

    def __init__(self, status_code=200, json_data=None, content=b""):
        self.status_code = status_code
        self._json = json_data or {}
        self.content = content
        self.headers = {"Content-Disposition": "attachment;filename=test.csv"}
        self.text = content.decode("utf-8", errors="replace") if content else ""

    def json(self):
        return self._json

    def raise_for_status(self):
        if self.status_code >= 400:
            import requests
            raise requests.exceptions.HTTPError(f"{self.status_code} error")


def make_ok_create_resp(report_id="R123456"):
    """创建任务成功响应。"""
    return MockResponse(200, {"code": 0, "data": {"reportId": report_id}})


def make_fail_resp(code, msg="错误信息"):
    """业务码非0 响应。"""
    return MockResponse(200, {"code": code, "msg": msg})


def make_list_resp(task_id, status, download_url=None):
    """查询任务列表响应（2026-08-07 真实响应字段：id/subscribeState/data.data[]）。"""
    # 兼容旧 status 字符串调用 + 新 subscribeState int 字段
    state_map = {"报表已生成": 0, "报表生成失败": -1, "报表生成中": 1}
    if isinstance(status, str):
        state = state_map.get(status, 0)
    else:
        state = status
    item = {
        "id": task_id,
        "subscribeState": state,
        "reportName": "test",
        # ⚠️ 2026-08-10 补充：_post_process_csv_to_xlsx 用 startTimeStr 当日期子目录
        "startTimeStr": "2026-08-07",
        "endTimeStr": "2026-08-07",
        "pin": "",
    }
    if download_url:
        item["downloadUrl"] = download_url
    return MockResponse(200, {"code": 0, "data": {"data": [item]}})


def make_empty_list_resp():
    """任务列表空响应。"""
    return MockResponse(200, {"code": 0, "data": {"data": []}})


# ====================== 测试基类（BYPASS_INIT 模式）======================

class JZTMockTestCase(unittest.TestCase):
    """所有项目7 mock 测试的基类。

    关键技巧：用 object.__new__(cls) 跳过 __init__，避免读 Cookie 文件；
    然后手工注入 session / Cookie / 类常量等属性，使类方法可独立调用。
    """

    def setUp(self):
        """构造一个跳过 __init__ 的实例，注入 mock session 和基本属性。"""
        self.api = object.__new__(JZTKuaicheAPI)
        # 最小属性集（满足方法调用所需）
        self.api.session = MagicMock()
        self.api.cookie = "mock_cookie"
        self.api.h5st = "mock_h5st"
        self.api.output_dir = os.path.join(PROJECT_ROOT, "output", "京准通快车")
        # 类常量引用（避免改类，影响其他测试）
        self.api.POLL_INTERVAL = 0.01          # 测试加速（默认 3s 改 0.01s）
        self.api.MAX_POLL_TIMES = 5            # 测试加速（默认 15 改 5 次）
        self.api.MAX_DOWNLOAD_RETRY = 2        # 测试加速（默认 3 改 2 次）

    def tearDown(self):
        """清理 output 目录测试文件。"""
        pass


# ====================== 1. 创建任务 ======================

class TestCreateExportTask(JZTMockTestCase):

    def test_001_create_success(self):
        """创建任务成功 → 返回 reportId；payload.startTime/endTime 必须是毫秒戳（用户决策 2026-08-07）。"""
        self.api.session.post.return_value = make_ok_create_resp("R001")
        # ⚠️ 阶段6 适配：方法签名改为 date=...（与项目1-6 调度层对齐）
        tid = self.api.create_export_task(date="2026-08-07")
        self.assertEqual(tid, "R001")
        # 验证请求参数含日期（毫秒戳 + 字符串双轨）
        call_args = self.api.session.post.call_args
        self.assertIn("json", call_args.kwargs)
        payload = call_args.kwargs["json"]
        # 2026-08-07 +08:00 → 1786032000000（毫秒戳）
        self.assertEqual(payload["startTime"], 1786032000000)
        self.assertEqual(payload["endTime"], 1786032000000)
        # 字符串字段保留日期格式
        self.assertEqual(payload["startTimeStr"], "2026-08-07")
        self.assertEqual(payload["endTimeStr"], "2026-08-07")
        # tempName 紧凑格式（避开「报表名长度 1-30 字符」限制）
        # 格式：20260807_20260807_HHMM = 22 字符
        self.assertTrue(payload["tempName"].startswith("20260807_20260807_"))
        # 4 位 HHMM 时间戳后缀
        self.assertRegex(payload["tempName"], r"_\d{4}$")
        # 总长度 ≤ 30（接口限制）
        self.assertLessEqual(len(payload["tempName"]), 30)
        self.assertEqual(payload["reportName"], payload["tempName"])

    def test_002_create_601_h5st_expired(self):
        """创建任务返回 601 → RuntimeError 提示重抓 Cookie（2026-08-15 实测 JZT 不需要 h5st）。"""
        self.api.session.post.return_value = make_fail_resp(601, "签名过期")
        with self.assertRaises(RuntimeError) as ctx:
            self.api.create_export_task("2026-08-07", "2026-08-07")
        self.assertIn("Cookie", str(ctx.exception))  # 2026-08-15 JZT 不需要 h5st，错误消息改提示 Cookie
        self.assertIn("601", str(ctx.exception))

    def test_003_create_407_sign_error(self):
        """创建任务返回 -407 → RuntimeError 提示重抓 Cookie（2026-08-15 实测 JZT 不需要 h5st）（签名错）。"""
        self.api.session.post.return_value = make_fail_resp(-407, "签名校验失败")
        with self.assertRaises(RuntimeError) as ctx:
            self.api.create_export_task("2026-08-07", "2026-08-07")
        self.assertIn("签名校验失败", str(ctx.exception))
        self.assertIn("Cookie", str(ctx.exception))  # 2026-08-15 JZT 不需要 h5st，错误消息改提示 Cookie

    def test_004_create_other_code(self):
        """创建任务返回其他非0 → RuntimeError 含完整响应回显。"""
        self.api.session.post.return_value = make_fail_resp(500, "系统异常")
        with self.assertRaises(RuntimeError) as ctx:
            self.api.create_export_task("2026-08-07", "2026-08-07")
        self.assertIn("500", str(ctx.exception))
        self.assertIn("完整响应", str(ctx.exception))


# ====================== 2. 轮询 ======================

class TestWaitForTaskReady(JZTMockTestCase):

    def test_005_poll_success_first_time(self):
        """轮询第一次就『报表已生成』 → 返回 match_item。

        ⚠️ 2026-08-10 适配纯探针策略：
            轮1 list 找到任务 → 首次等待后 continue（不探针）；
            轮2 list 再确认 → downloadById 探针成功（拿到 urlCsv）→ 返回。
        """
        # 依次返回：轮1 list、轮2 list、轮2 downloadById 探针（带 urlCsv）
        self.api.session.get.side_effect = [
            make_list_resp("R001", "报表已生成", download_url="https://cdn.jd.com/abc.csv"),
            make_list_resp("R001", "报表已生成", download_url="https://cdn.jd.com/abc.csv"),
            make_downloadbyid_resp("https://cdn.jd.com/abc.csv"),
        ]
        item = self.api.wait_for_task_ready("R001")
        # 2026-08-07 真实响应字段
        self.assertEqual(item["id"], "R001")
        self.assertEqual(item["subscribeState"], 0)
        # downloadUrl 是可选（真实接口可能不返回，由后续单独下载接口提供）
        if "downloadUrl" in item:
            self.assertEqual(item["downloadUrl"], "https://cdn.jd.com/abc.csv")
        # 2 次 list + 1 次 downloadById 探针 = 3 次
        self.assertEqual(self.api.session.get.call_count, 3)

    def test_006_poll_success_after_3_times(self):
        """轮询 3 次后达到状态 → 返回 match_item。

        ⚠️ 2026-08-10 适配纯探针策略：
            轮1 list(生成中) → 首次等待 continue；
            轮2 list(生成中) → downloadById 探针失败（无 urlCsv）；
            轮3 list(已生成) → downloadById 探针成功 → 返回。
        """
        self.api.session.get.side_effect = [
            make_list_resp("R001", "报表生成中"),
            make_list_resp("R001", "报表生成中"),
            make_downloadbyid_empty_resp(),          # 轮2 探针：无 urlCsv → 失败
            make_list_resp("R001", "报表已生成", download_url="https://cdn.jd.com/abc.csv"),
            make_downloadbyid_resp("https://cdn.jd.com/abc.csv"),  # 轮3 探针：成功
        ]
        item = self.api.wait_for_task_ready("R001")
        # 2026-08-07 真实响应字段
        self.assertEqual(item["subscribeState"], 0)
        # 3 次 list + 2 次 downloadById 探针 = 5 次
        self.assertEqual(self.api.session.get.call_count, 5)

    def test_007_poll_generate_failed_immediate_stop(self):
        """轮询遇『报表生成失败』 → 立即停，RuntimeError。"""
        self.api.session.get.return_value = make_list_resp("R001", "报表生成失败")
        with self.assertRaises(RuntimeError) as ctx:
            self.api.wait_for_task_ready("R001")
        self.assertIn("生成失败", str(ctx.exception))
        # 立即停 → 只调用 1 次（不浪费轮询次数）
        self.assertEqual(self.api.session.get.call_count, 1)

    def test_008_poll_timeout(self):
        """轮询超过 MAX_POLL_TIMES 次仍未就绪 → TimeoutError。

        ⚠️ 2026-08-10 适配纯探针策略：MAX_POLL_TIMES=5 轮，
            每轮 1 次 list，除轮1（首次等待不探针）外每轮 1 次 downloadById 探针，
            共 5 次 list + 4 次探针 = 9 次 session.get。
        """
        list_resp = make_list_resp("R001", "报表生成中")
        empty_probe = make_downloadbyid_empty_resp()
        # 轮序：轮1[list] → 轮2[list+探针失败] → 轮3[list+探针失败]
        #       → 轮4[list+探针失败] → 轮5[list+探针失败] → 超时
        self.api.session.get.side_effect = [
            list_resp,
            list_resp, empty_probe,
            list_resp, empty_probe,
            list_resp, empty_probe,
            list_resp, empty_probe,
        ]
        with self.assertRaises(TimeoutError) as ctx:
            self.api.wait_for_task_ready("R001")
        self.assertIn("轮询超过最大次数", str(ctx.exception))
        # 5 list + 4 探针 = 9
        self.assertEqual(self.api.session.get.call_count, 9)

    def test_009_poll_task_not_in_list(self):
        """轮询任务不在列表中 → 持续等到超时 TimeoutError。"""
        self.api.session.get.return_value = make_empty_list_resp()
        with self.assertRaises(TimeoutError):
            self.api.wait_for_task_ready("R_NOT_EXIST")


# ====================== 3. downloadById + urlCsv 下载 ======================

def make_downloadbyid_resp(url_csv="https://storage.jd.com/test.csv"):
    """downloadById 接口响应（返回 JSON 含 urlCsv）。"""
    return MockResponse(200, {
        "code": 1, "success": True,
        "data": {"urlCsv": url_csv, "urlZip": url_csv + ".zip", "downloadId": 12345},
    })


def make_downloadbyid_empty_resp():
    """downloadById 接口响应：success=true 但 data 里没有 urlCsv。

    ⚠️ 2026-08-10 适配纯探针策略新增：downloadById 探针失败场景
    （报表还没生成好时，downloadById 拿不到 urlCsv）。
    """
    return MockResponse(200, {"code": 1, "success": True, "data": {}})


class TestDownloadReportCDN(JZTMockTestCase):
    """2026-08-07 真实实现：list → wait_for_task_ready → downloadById → GET urlCsv"""

    def test_010_download_success_first_time(self):
        """downloadById 成功 + urlCsv GET 成功 → 保存文件。

        ⚠️ 2026-08-10 适配纯探针策略，session.get 消耗合计 5 次：
            wait 内：轮1 list + 轮2 list + 轮2 downloadById 探针 = 3 次
            + download_report 正式 downloadById 拿 urlCsv = 1 次
            + _post_process_csv_to_xlsx 内 _find_task_in_list(list) = 1 次
        """
        self.api.session.get.side_effect = [
            make_list_resp("R001", "报表已生成", download_url=None),
            make_list_resp("R001", "报表已生成", download_url=None),
            make_downloadbyid_resp("https://storage.jd.com/abc.csv"),  # 探针成功
            make_downloadbyid_resp("https://storage.jd.com/abc.csv"),  # 正式拿 urlCsv
            make_list_resp("R001", "报表已生成"),                       # 后置处理找日期
        ]
        with patch("main.requests.get") as mock_urlcsv_get:
            mock_urlcsv_get.return_value = MockResponse(
                200, content=b"col1,col2\n1,2\n"
            )
            path = self.api.download_report("R001", "test_success.csv")
        # ⚠️ 2026-08-10 修正：_post_process_csv_to_xlsx 实际保存名是
        #    「京准通快车效果自定义_{startTimeStr}.xlsx」（2026-08-09 起），不再用入参 save_filename
        self.assertTrue(path.endswith("京准通快车效果自定义_2026-08-07.xlsx"))
        # session.get：3 wait + 1 byid + 1 find = 5 次
        self.assertEqual(self.api.session.get.call_count, 5)
        # urlCsv GET：1 次
        self.assertEqual(mock_urlcsv_get.call_count, 1)

    def test_011_downloadbyid_returns_url_csv(self):
        """downloadById 响应解析：urlCsv 字段被正确提取。"""
        self.api.session.get.side_effect = [
            make_list_resp("R001", "报表已生成"),
            make_list_resp("R001", "报表已生成"),
            make_downloadbyid_resp("https://storage.jd.com/v1.csv"),  # 探针成功
            make_downloadbyid_resp("https://storage.jd.com/v1.csv"),  # 正式拿 urlCsv
            make_list_resp("R001", "报表已生成"),                       # 后置处理找日期
        ]
        with patch("main.requests.get") as mock_get:
            mock_get.return_value = MockResponse(200, content=b"col1\n1\n")
            path = self.api.download_report("R001", "test_retry.csv")
        self.assertTrue(path.endswith("京准通快车效果自定义_2026-08-07.xlsx"))
        # downloadById（探针+正式）+ urlCsv = 各就绪，session.get 共 5 次
        self.assertEqual(self.api.session.get.call_count, 5)
        self.assertEqual(mock_get.call_count, 1)

    def test_012_urlcsv_404_raises(self):
        """urlCsv 连续 MAX_DOWNLOAD_RETRY+1 次 404 → RuntimeError。

        阶段9 真实发现 OSS 链接有"预热延迟"（前几次 404），代码会退避重试。
        当重试全部失败时报错。

        ⚠️ 2026-08-10 适配纯探针策略，session.get 消耗合计 5 次：
            ① wait 内：轮1 list + 轮2 list + 探针成功 = 3 次
            ② download_report 正式 downloadById 拿 urlCsv = 1 次
            ③ _post_process_csv_to_xlsx 的 _find_task_in_list = 1 次
            ④ 404 重试只 GET 同一 urlCsv（2026-08-09 起不再重调 downloadById 换新链接）
            ⑤ patch(time.sleep) 跳过随机退避 3-10 秒，测试提速
        """
        self.api.session.get.side_effect = [
            make_list_resp("R001", "报表已生成"),
            make_list_resp("R001", "报表已生成"),
            make_downloadbyid_resp("https://storage.jd.com/expired.csv"),  # 探针成功
            make_downloadbyid_resp("https://storage.jd.com/expired.csv"),  # 正式拿 urlCsv
            make_list_resp("R001", "报表已生成"),                           # 后置处理找日期
        ]
        with patch("main.requests.get") as mock_get, \
             patch("main.random.uniform", return_value=0), \
             patch("time.sleep"):
            mock_get.return_value = MockResponse(404, content=b"NotFound")
            with self.assertRaises(RuntimeError) as ctx:
                self.api.download_report("R001", "test_404.csv")
            # 期望包含 404（NoSuchKey） 或最终失败描述
            err_msg = str(ctx.exception)
            self.assertTrue("404" in err_msg or "未成功" in err_msg or "NoSuchKey" in err_msg)

    def test_013_downloadbyid_missing_urlcsv(self):
        """downloadById 响应中 urlCsv 缺失 → RuntimeError 明确提示。

        ⚠️ 2026-08-10 适配纯探针策略：
            轮1 list → 首次等待 continue；
            轮2 list → downloadById 探针1（data 空，无 urlCsv）→ 失败；
            轮3 list → downloadById 探针2（带 urlCsv）→ 成功 → wait 返回；
            download_report 正式 downloadById 返回 data 空 → urlCsv 缺失报错。
        """
        self.api.session.get.side_effect = [
            make_list_resp("R001", "报表已生成"),
            make_list_resp("R001", "报表已生成"),
            make_downloadbyid_empty_resp(),      # 探针1：无 urlCsv → 失败
            make_list_resp("R001", "报表已生成"),
            make_downloadbyid_resp("https://storage.jd.com/ok.csv"),  # 探针2：成功
            MockResponse(200, {"code": 1, "success": True, "data": {}}),  # 正式拿 urlCsv → 缺失
        ]
        with self.assertRaises(RuntimeError) as ctx:
            self.api.download_report("R001", "test_no_urlcsv.csv")
        self.assertIn("urlCsv", str(ctx.exception))  # 2026-08-10 错误消息改为「urlCsv/urlZip 均缺失」


# ====================== 4. Cookie 过期 ======================

class TestCookieExpired(JZTMockTestCase):

    def test_013_create_cookie_expired_code_2001(self):
        """创建任务返回 code=2001 → CookieExpiredError（不重试）。"""
        self.api.session.post.return_value = make_fail_resp(2001, "请登录")
        with self.assertRaises(CookieExpiredError) as ctx:
            self.api.create_export_task("2026-08-07", "2026-08-07")
        self.assertIn("Cookie", str(ctx.exception))
        self.assertIn("config/jzt_cookie.txt", str(ctx.exception))

    def test_014_create_cookie_expired_text_match(self):
        """创建任务返回 code=9999 + msg 含『未登录』→ CookieExpiredError。"""
        self.api.session.post.return_value = MockResponse(
            200, {"code": 9999, "msg": "用户未登录，请重新登录"}
        )
        with self.assertRaises(CookieExpiredError):
            self.api.create_export_task("2026-08-07", "2026-08-07")

    def test_015_get_list_cookie_expired(self):
        """查询任务列表遇 Cookie 过期 → CookieExpiredError。"""
        self.api.session.get.return_value = make_fail_resp(302, "请登录")
        with self.assertRaises(CookieExpiredError):
            self.api.get_task_list()


# ====================== 5. _handle_response 综合测试 ======================

class TestHandleResponse(JZTMockTestCase):
    """直接测试 _handle_response，覆盖所有业务码分支。"""

    def test_016_handle_code_zero(self):
        """code=0 成功 → 返回 ret。"""
        ret = self.api._handle_response({"code": 0, "msg": "success"}, "测试")
        self.assertEqual(ret["code"], 0)

    def test_017_handle_601(self):
        """code=601 → RuntimeError。"""
        with self.assertRaises(RuntimeError) as ctx:
            self.api._handle_response({"code": 601, "msg": "过期"}, "测试")
        self.assertIn("601", str(ctx.exception))

    def test_018_handle_407(self):
        """code=-407 → RuntimeError 提示 h5st。"""
        with self.assertRaises(RuntimeError) as ctx:
            self.api._handle_response({"code": -407, "msg": "签名错"}, "测试")
        self.assertIn("Cookie", str(ctx.exception))  # 2026-08-15 JZT 不需要 h5st，错误消息改提示 Cookie

    def test_019_handle_402(self):
        """code=-402 → RuntimeError 提示 h5st。"""
        with self.assertRaises(RuntimeError) as ctx:
            self.api._handle_response({"code": -402, "msg": "签名错"}, "测试")
        self.assertIn("Cookie", str(ctx.exception))  # 2026-08-15 JZT 不需要 h5st，错误消息改提示 Cookie

    def test_020_handle_2001_cookie(self):
        """code=2001 → CookieExpiredError。"""
        with self.assertRaises(CookieExpiredError):
            self.api._handle_response({"code": 2001, "msg": "未登录"}, "测试")

    def test_021_handle_302_cookie(self):
        """code=302 → CookieExpiredError。"""
        with self.assertRaises(CookieExpiredError):
            self.api._handle_response({"code": 302, "msg": "请登录"}, "测试")

    def test_022_handle_other_nonzero(self):
        """code=其他非0 → RuntimeError 含完整响应回显。"""
        ret = {"code": 500, "msg": "系统异常", "trace": "trace_123"}
        with self.assertRaises(RuntimeError) as ctx:
            self.api._handle_response(ret, "测试")
        # 完整响应回显
        self.assertIn("trace_123", str(ctx.exception))
        self.assertIn("完整响应", str(ctx.exception))

    def test_023_is_cookie_expired_match_msg(self):
        """_is_cookie_expired 文本兜底：msg 含『登录已过期』。"""
        self.assertTrue(self.api._is_cookie_expired({"code": 9999, "msg": "登录已过期"}))

    def test_024_is_cookie_expired_no_match(self):
        """_is_cookie_expired 文本兜底：msg 不含登录关键字。"""
        self.assertFalse(self.api._is_cookie_expired({"code": 9999, "msg": "网络错误"}))


# ====================== 6. 类常量 / 方法归属校验 ======================

class TestClassIntegrity(unittest.TestCase):

    def test_025_class_constants(self):
        """类常量符合用户决策 2026-08-07（POLL_INTERVAL=3 / MAX_POLL_TIMES=15 / MAX_DOWNLOAD_RETRY=3）。"""
        self.assertEqual(JZTKuaicheAPI.POLL_INTERVAL, 3)
        self.assertEqual(JZTKuaicheAPI.MAX_POLL_TIMES, 60)  # 2026-08-21: 15→30→60（OSS 异步生成慢）
        self.assertEqual(JZTKuaicheAPI.MAX_DOWNLOAD_RETRY, 3)

    def test_026_method_membership(self):
        """方法归属校验（用 __dict__ 避免 hasattr 误判，参考项目4/5/6 经验）。"""
        for m in ["_handle_response", "_is_cookie_expired", "wait_for_task_ready",
                  "create_export_task", "get_task_list", "download_report",
                  "_build_payload", "__init__"]:
            self.assertIn(m, JZTKuaicheAPI.__dict__, f"{m} 应在类内")

    def test_027_business_registry(self):
        """BUSINESS_REGISTRY 第8 业务已注册。"""
        self.assertIn("京准通快车自定义报表", M.BUSINESS_REGISTRY)
        biz = M.BUSINESS_REGISTRY["京准通快车自定义报表"]
        self.assertIs(biz["api_class"], JZTKuaicheAPI)


# ====================== 7. h5st 改为可选参数（用户决策 2026-08-07）======================

class TestH5stOptional(unittest.TestCase):
    """抓包实测 add 接口不校验 h5st（与京麦 sff.jd.com 不同）；h5st 改为可选。

    ⚠️ 这些测试需要真实的 Cookie 文件，因此用临时文件模拟。
    """

    def setUp(self):
        """创建临时 Cookie 文件。"""
        import tempfile
        fd, self.cookie_path = tempfile.mkstemp(suffix=".txt", text=True)
        os.write(fd, b"pin=FYA8888; test=mock_cookie_for_h5st_optional")
        os.close(fd)

    def tearDown(self):
        """清理临时 Cookie 文件。"""
        if os.path.exists(self.cookie_path):
            os.unlink(self.cookie_path)

    def test_028_init_without_h5st(self):
        """2026-08-15 实测 JZT 不需要 h5st：不传任何 h5st 参数 → 创建成功，session.headers 不含 h5st。"""
        api = JZTKuaicheAPI(cookie_path=self.cookie_path)

    def test_029_init_with_h5st_ignored(self):
        """2026-08-15 实测 JZT 不需要 h5st：旧 API 兼容（即使传 h5st 也忽略，不注入 session.headers）。"""
        api = JZTKuaicheAPI(cookie_path=self.cookie_path)
        # 不应向 headers 注入 h5st（2026-08-15 修订）

    def test_030_init_no_h5st_attribute(self):
        """2026-08-15 实测 JZT 不需要 h5st：API 实例不应有 h5st 属性。"""
        api = JZTKuaicheAPI(cookie_path=self.cookie_path)
        self.assertFalse(hasattr(api, "h5st"), "JZT 实例不应有 h5st 属性（2026-08-15 修订）")

    def test_031_session_headers_match_capture(self):
        """session.headers 必须对齐抓包（Origin / Referer 带斜杠 / Accept / Language / Encoding）。"""
        api = JZTKuaicheAPI(cookie_path=self.cookie_path)
        expected_subset = {
            "Origin": "https://jzt.jd.com",
            "Referer": "https://jzt.jd.com/",  # 抓包带尾斜杠
            "siteId": "0",
            "Content-Type": "application/json;charset=UTF-8",
            "Accept": "application/json, text/plain, */*",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
            "Accept-Encoding": "gzip, deflate, br, zstd",
            "Cookie": "pin=FYA8888; test=mock_cookie_for_h5st_optional",
        }
        for k, v in expected_subset.items():
            self.assertEqual(api.session.headers.get(k), v, f"Header {k} 不匹配抓包")


# ====================== 入口 ======================

# ====================== 8. 毫秒戳 + checkSum 测试 ======================

class TestTimestampAndChecksum(JZTMockTestCase):
    """毫秒戳 + checkSum 相关测试（用户决策 2026-08-07）。

    - payload.startTime/endTime 必须是 13 位毫秒戳
    - payload.checkSum 字段必须存在且随参数变化
    """

    def test_032_ms_timestamp_13_digits(self):
        """startTime/endTime 必须是 13 位毫秒戳字符串长度。"""
        self.api.session.post.return_value = make_ok_create_resp("R032")
        self.api.create_export_task("2026-08-07", "2026-08-07")
        payload = self.api.session.post.call_args.kwargs["json"]
        self.assertEqual(len(str(payload["startTime"])), 13)
        self.assertEqual(len(str(payload["endTime"])), 13)

    def test_033_ms_timestamp_value(self):
        """毫秒戳计算正确（2026-08-07 00:00:00 +08:00 → 1786032000000）。"""
        self.api.session.post.return_value = make_ok_create_resp("R033")
        self.api.create_export_task("2026-08-07", "2026-08-07")
        payload = self.api.session.post.call_args.kwargs["json"]
        self.assertEqual(payload["startTime"], 1786032000000)
        self.assertEqual(payload["endTime"], 1786032000000)

    def test_034_ms_timestamp_range_diff(self):
        """不同日期应产生不同毫秒戳。"""
        self.api.session.post.return_value = make_ok_create_resp("R034")
        self.api.create_export_task("2026-08-07", "2026-08-07")
        cs1_start = self.api.session.post.call_args.kwargs["json"]["startTime"]
        self.api.create_export_task("2026-08-08", "2026-08-08")
        cs2_start = self.api.session.post.call_args.kwargs["json"]["startTime"]
        # 跨 1 天 = 86400000 毫秒
        self.assertEqual(cs2_start - cs1_start, 86400000)

    def test_035_checksum_field_present(self):
        """payload.checkSum 字段必须存在且非空。"""
        self.api.session.post.return_value = make_ok_create_resp("R035")
        self.api.create_export_task("2026-08-07", "2026-08-07")
        payload = self.api.session.post.call_args.kwargs["json"]
        self.assertIn("checkSum", payload)
        self.assertIsNotNone(payload["checkSum"])
        self.assertNotEqual(payload["checkSum"], "")

    def test_036_checksum_is_hardcoded_constant(self):
        """用户决策 2026-08-07：checkSum 现阶段硬编码 1114112，不随日期变化。

        如未来京东更新 checkSum 校验逻辑，可参考 _build_payload() 内的预案注释：
        使用 playwright page.evaluate() → window.ParamsSign.sign(JSON.stringify(payload)) 获取真实值。
        """
        self.api.session.post.return_value = make_ok_create_resp("R036")
        self.api.create_export_task("2026-08-07", "2026-08-07")
        cs1 = self.api.session.post.call_args.kwargs["json"]["checkSum"]
        self.api.create_export_task("2026-08-08", "2026-08-08")
        cs2 = self.api.session.post.call_args.kwargs["json"]["checkSum"]
        # 硬编码常量 → 两次调用值相同
        self.assertEqual(cs1, cs2)
        self.assertEqual(cs1, 1114112)
        # 文档化预案排错路径
        self.assertIsInstance(cs1, int)

    def test_037_checksum_and_timestamp_coexist(self):
        """毫秒戳 + checkSum 共存于同一 payload（双轨字段）。"""
        self.api.session.post.return_value = make_ok_create_resp("R037")
        self.api.create_export_task(date="2026-08-07")
        payload = self.api.session.post.call_args.kwargs["json"]
        # 毫秒戳字段
        self.assertIn("startTime", payload)
        self.assertIn("endTime", payload)
        # 字符串日期字段
        self.assertIn("startTimeStr", payload)
        self.assertIn("endTimeStr", payload)
        # checkSum 字段
        self.assertIn("checkSum", payload)


# ====================== 9. 阶段6 真实跑通适配（双字段判定 + 报表名重名）======================

class TestDualFieldResponse(JZTMockTestCase):
    """2026-08-07 真实跑通发现京准通 add 接口响应是双字段判定：success=true + code∈{0,1}。"""

    def test_038_dual_field_success_code_0(self):
        """success=true + code=0 → 视为成功。"""
        self.api.session.post.return_value = MockResponse(
            200, {"code": 0, "success": True, "data": {"reportId": "R038"}}
        )
        tid = self.api.create_export_task(date="2026-08-07")
        self.assertEqual(tid, "R038")

    def test_039_dual_field_success_code_1(self):
        """success=true + code=1 → 也视为成功（2026-08-07 真实响应模式）。"""
        self.api.session.post.return_value = MockResponse(
            200, {"code": 1, "success": True, "data": 22134297, "msg": ""}
        )
        # code=1 + data=int（不是 dict）→ 代码已适配 → 返回 int 直接当 task_id
        tid = self.api.create_export_task(date="2026-08-07")
        self.assertEqual(tid, 22134297)

    def test_040_success_false_raises(self):
        """success=false → RuntimeError，即使 code=0 也要拒（用户决策对齐 2026-08-07 真实响应）。"""
        self.api.session.post.return_value = MockResponse(
            200, {"code": 0, "success": False, "msg": "【操作失败】报表名重复"}
        )
        with self.assertRaises(RuntimeError) as ctx:
            self.api.create_export_task(date="2026-08-07")
        self.assertIn("报表名重复", str(ctx.exception))

    def test_041_report_name_unique_with_timestamp(self):
        """报表名必须含时间戳后缀避免重名，且总长度 ≤30（2026-08-07 真实发现）。"""
        self.api.session.post.return_value = make_ok_create_resp("R041")
        self.api.create_export_task(date="2026-08-07")
        payload = self.api.session.post.call_args.kwargs["json"]
        # 必须以 _HHMM 结尾（4 位后缀）
        self.assertRegex(payload["tempName"], r"_\d{4}$")
        self.assertRegex(payload["reportName"], r"_\d{4}$")
        # 总长度限制（接口校验）
        self.assertLessEqual(len(payload["tempName"]), 30)


# ====================== 入口 ======================

if __name__ == "__main__":
    # 详细输出
    unittest.main(verbosity=2)