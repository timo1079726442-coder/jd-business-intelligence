# -*- coding: utf-8 -*-
"""auth_writer.py 单元测试（M-28，2026-08-24）

覆盖范围：
    1) write_cookie 写盘 + 回显（含 dict/str 两种输入）
    2) write_cookie 校验 cookies 字段缺失 → ValueError
    3) write_h5st 自动加 13 位毫秒 captured_at + h5st_key 标记
    4) write_h5st 校验 h5st_value 空 → ValueError
    5) validate_shop_id 防路径注入
    6) extract_h5st_from_list 提取 h5st（API 关键词匹配）
    7) write_from_raw_file 读 raw JSON + BOM 兼容

运行：python tests/test_auth_writer_unittest.py
"""
import os
import sys
import json
import tempfile
import unittest

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PROJECT_ROOT)

import auth_writer


class TestAuthWriterBase(unittest.TestCase):
    """基类：把 CONFIG_DIR 指向临时目录，避免污染真实 config/"""

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()
        # 备份原 CONFIG_DIR 并指向临时目录
        self._orig_config_dir = auth_writer.CONFIG_DIR
        auth_writer.CONFIG_DIR = self.tmpdir

    def tearDown(self):
        # 恢复 CONFIG_DIR
        auth_writer.CONFIG_DIR = self._orig_config_dir
        import shutil
        if os.path.exists(self.tmpdir):
            shutil.rmtree(self.tmpdir)


class TestWriteCookie(TestAuthWriterBase):
    """M-28.1/2 write_cookie 写盘 + 校验"""

    def test_write_cookie_with_dict(self):
        """dict 输入应写入 JSON 文件并返回路径"""
        file_path = auth_writer.write_cookie(
            "FYA箱包旗舰店",
            "sz",
            {"url": "https://sz.jd.com", "cookies": [{"name": "pin", "value": "FYA8888"}]},
        )
        self.assertTrue(os.path.exists(file_path))
        # 读取验证
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.assertEqual(data["cookies"][0]["name"], "pin")
        self.assertEqual(data["cookies"][0]["value"], "FYA8888")

    def test_write_cookie_with_json_string(self):
        """str（JSON 字符串）输入应同样写入"""
        file_path = auth_writer.write_cookie(
            "MIYO箱包旗舰店",
            "jzt",
            '{"cookies": [{"name": "pin", "value": "miyo-周"}]}',
        )
        self.assertTrue(os.path.exists(file_path))
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.assertEqual(data["cookies"][0]["value"], "miyo-周")

    def test_write_cookie_missing_cookies_field(self):
        """缺 cookies 字段应抛 ValueError"""
        with self.assertRaises(ValueError):
            auth_writer.write_cookie("FYA箱包旗舰店", "sz", {"url": "https://sz.jd.com"})

    def test_write_cookie_invalid_biz_type(self):
        """未知 biz_type 应抛 ValueError"""
        with self.assertRaises(ValueError):
            auth_writer.write_cookie("FYA箱包旗舰店", "unknown_type", {"cookies": []})


class TestWriteH5st(TestAuthWriterBase):
    """M-28.3/4 write_h5st 自动加时间戳"""

    def test_write_h5st_adds_millis_captured_at(self):
        """写入的 JSON 应含 13 位毫秒 captured_at + h5st_key"""
        file_path = auth_writer.write_h5st(
            "FYA箱包旗舰店",
            "abc123h5stvalue",
            h5st_key="jm_order",
        )
        self.assertTrue(os.path.exists(file_path))
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.assertEqual(data["h5st"], "abc123h5stvalue")
        self.assertEqual(data["h5st_key"], "jm_order")
        # captured_at 应为 13 位毫秒（> 1e12）
        self.assertGreater(data["captured_at"], 1_000_000_000_000)
        self.assertLess(data["captured_at"], 1_000_000_000_000_000)

    def test_write_h5st_jm_after_sale_writes_own_file(self):
        """不同 h5st_key 应写不同文件（订单/售后隔离）"""
        file_order = auth_writer.write_h5st("FYA箱包旗舰店", "order_h5st", h5st_key="jm_order")
        file_after = auth_writer.write_h5st("FYA箱包旗舰店", "after_h5st", h5st_key="jm_after_sale")
        # 文件名不同（H5ST_KEY_MAP 区分）
        self.assertNotEqual(
            os.path.basename(file_order),
            os.path.basename(file_after),
            "订单/售后 h5st 文件必须分开（项目20 实测不可跨业务复用）",
        )

    def test_write_h5st_empty_value_raises(self):
        """空 h5st 值应抛 ValueError"""
        with self.assertRaises(ValueError):
            auth_writer.write_h5st("FYA箱包旗舰店", "")

    def test_write_h5st_unknown_key_raises(self):
        """未知 h5st_key 应抛 ValueError"""
        with self.assertRaises(ValueError):
            auth_writer.write_h5st("FYA箱包旗舰店", "x", h5st_key="unknown_key")


class TestValidateShopId(unittest.TestCase):
    """M-28.5 validate_shop_id 防路径注入"""

    def test_valid_shop_ids(self):
        """正常店铺名应通过"""
        for shop in ["FYA箱包旗舰店", "MIYO箱包旗舰店", "OTA箱包旗舰店"]:
            self.assertEqual(auth_writer.validate_shop_id(shop), shop)

    def test_path_injection_rejected(self):
        """路径注入字符应抛 ValueError"""
        dangerous = ["../../etc/passwd", "..\\..\\x", "a/b", "a\\b", ""]
        for bad in dangerous:
            with self.assertRaises(ValueError, msg=f"应拒绝 {bad!r}"):
                auth_writer.validate_shop_id(bad)


class TestExtractH5stFromList(TestAuthWriterBase):
    """M-28.6 extract_h5st_from_list 提取 h5st"""

    def test_extract_from_request_list(self):
        """从请求列表中匹配 URL 含关键词的 POST 请求"""
        # 构造影刀监听结果：包含多个请求，只有一个是目标 API
        req_list = [
            {"method": "GET", "url": "https://sff.jd.com/other", "headers": {}},
            {
                "method": "POST",
                "url": "https://sff.jd.com/api?api=createdExportTask",
                "headers": {"h5st": "extracted_h5st_123", "Content-Type": "application/json"},
            },
            {"method": "POST", "url": "https://sff.jd.com/api?api=queryExportTaskInfo", "headers": {"h5st": "wrong_h5st"}},
        ]
        input_path = os.path.join(self.tmpdir, "req_list.json")
        with open(input_path, "w", encoding="utf-8") as f:
            json.dump(req_list, f)
        # 提取
        result = auth_writer.extract_h5st_from_list(
            input_path,
            "FYA箱包旗舰店",
            h5st_key="jm_order",
            api_keyword="createdExportTask",
        )
        # 验证写入文件内容
        with open(result, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.assertEqual(data["h5st"], "extracted_h5st_123")

    def test_no_matching_request_raises(self):
        """无匹配请求应抛 ValueError"""
        req_list = [{"method": "GET", "url": "https://sff.jd.com/no_match", "headers": {}}]
        input_path = os.path.join(self.tmpdir, "no_match.json")
        with open(input_path, "w", encoding="utf-8") as f:
            json.dump(req_list, f)
        with self.assertRaises(ValueError):
            auth_writer.extract_h5st_from_list(
                input_path,
                "FYA箱包旗舰店",
                h5st_key="jm_order",
                api_keyword="createdExportTask",
            )

    def test_empty_input_raises(self):
        """空输入文件应抛 ValueError"""
        input_path = os.path.join(self.tmpdir, "empty.json")
        with open(input_path, "w", encoding="utf-8") as f:
            f.write("")
        with self.assertRaises(ValueError):
            auth_writer.extract_h5st_from_list(
                input_path,
                "FYA箱包旗舰店",
                h5st_key="jm_order",
            )


class TestWriteFromRawFile(TestAuthWriterBase):
    """M-28.7 write_from_raw_file 读 raw JSON"""

    def test_write_from_raw_file_with_bom(self):
        """含 UTF-8 BOM 的 raw 文件应正常解析"""
        raw_path = os.path.join(self.tmpdir, "jm_cookie.json")
        raw_content = '﻿{"cookies": [{"name": "pin", "value": "FYA8888"}]}'
        with open(raw_path, "w", encoding="utf-8") as f:
            f.write(raw_content)
        # 显式传 biz_type（避免依赖文件名推断）
        file_path = auth_writer.write_from_raw_file(
            "FYA箱包旗舰店",
            raw_path,
            biz_type="jm",
        )
        self.assertTrue(os.path.exists(file_path))
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.assertEqual(data["cookies"][0]["value"], "FYA8888")

    def test_write_from_raw_file_no_bom(self):
        """无 BOM 的 raw 文件也应正常解析"""
        raw_path = os.path.join(self.tmpdir, "sz_cookie.json")
        raw_content = '{"cookies": [{"name": "pin", "value": "FYA8888"}]}'
        with open(raw_path, "w", encoding="utf-8") as f:
            f.write(raw_content)
        file_path = auth_writer.write_from_raw_file(
            "FYA箱包旗舰店",
            raw_path,
            biz_type="sz",
        )
        self.assertTrue(os.path.exists(file_path))

    def test_write_from_raw_file_auto_infer_biz_type(self):
        """未传 biz_type 时应按文件名前缀推断"""
        raw_path = os.path.join(self.tmpdir, "jzt_cookie.json")
        raw_content = '{"cookies": [{"name": "pin", "value": "ota8888"}]}'
        with open(raw_path, "w", encoding="utf-8") as f:
            f.write(raw_content)
        file_path = auth_writer.write_from_raw_file("OTA箱包旗舰店", raw_path)
        # 应推断为 jzt → 写入 jzt_cookie.json
        self.assertTrue(file_path.endswith("jzt_cookie.json"))
        self.assertTrue(os.path.exists(file_path))


if __name__ == "__main__":
    print("=" * 60)
    print("auth_writer.py 单元测试（M-28，2026-08-24）")
    print("=" * 60)
    unittest.main(verbosity=2)