# -*- coding: utf-8 -*-
"""runtime_config.py 单元测试（M-29，2026-08-24）

覆盖范围：
    1) get_shop_id 环境变量读取
    2) get_shop_id 缺失 → SystemExit(3)
    3) get_shop_short_name 后缀剥离（"箱包旗舰店"）
    4) get_shop_pin 环境变量优先
    5) get_app_id biz_type 路由（jm_order / jm_after_sale）
    6) get_sign_salt config 读取

运行：python tests/test_runtime_config_unittest.py
"""
import os
import sys
import unittest
from unittest.mock import patch

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PROJECT_ROOT)

import runtime_config


class TestGetShopId(unittest.TestCase):
    """M-29.1/2 get_shop_id"""

    def test_reads_env_shop_id(self):
        """SHOP_ID 环境变量应被读取"""
        with patch.dict(os.environ, {"SHOP_ID": "FYA箱包旗舰店"}, clear=True):
            self.assertEqual(runtime_config.get_shop_id(), "FYA箱包旗舰店")

    def test_missing_shop_id_raises_systemexit(self):
        """无 SHOP_ID → SystemExit(3)"""
        with patch.dict(os.environ, {}, clear=True):
            with self.assertRaises(SystemExit) as cm:
                runtime_config.get_shop_id()
            self.assertEqual(cm.exception.code, 3)

    def test_empty_shop_id_raises_systemexit(self):
        """空 SHOP_ID → SystemExit(3)"""
        with patch.dict(os.environ, {"SHOP_ID": "   "}, clear=True):
            with self.assertRaises(SystemExit) as cm:
                runtime_config.get_shop_id()
            self.assertEqual(cm.exception.code, 3)


class TestGetShopShortName(unittest.TestCase):
    """M-29.3 get_shop_short_name 后缀剥离"""

    def test_strip_suffix(self):
        """FYA箱包旗舰店 → FYA"""
        with patch.dict(os.environ, {"SHOP_ID": "FYA箱包旗舰店"}, clear=True):
            self.assertEqual(runtime_config.get_shop_short_name(), "FYA")

    def test_miyo_strip_suffix(self):
        """MIYO箱包旗舰店 → MIYO"""
        with patch.dict(os.environ, {"SHOP_ID": "MIYO箱包旗舰店"}, clear=True):
            self.assertEqual(runtime_config.get_shop_short_name(), "MIYO")


class TestGetShopPin(unittest.TestCase):
    """M-29.4 get_shop_pin"""

    def test_env_pin_priority(self):
        """SHOP_PIN 环境变量应优先"""
        with patch.dict(
            os.environ,
            {"SHOP_ID": "FYA箱包旗舰店", "SHOP_PIN": "FYA8888"},
            clear=True,
        ):
            self.assertEqual(runtime_config.get_shop_pin(), "FYA8888")

    def test_fallback_to_config_when_no_env_pin(self):
        """无 SHOP_PIN 环境变量 → 从 config.xlsx「店铺账号」查"""
        with patch.dict(
            os.environ,
            {"SHOP_ID": "FYA箱包旗舰店"},
            clear=True,
        ):
            # 用真实 config.xlsx（项目根 config/config.xlsx）
            pin = runtime_config.get_shop_pin()
            self.assertTrue(pin, "FYA 店铺应能查到 pin")

    def test_unknown_shop_raises(self):
        """config 中无此店铺 → SystemExit(3)"""
        with patch.dict(
            os.environ,
            {"SHOP_ID": "未知店铺XYZ"},
            clear=True,
        ):
            with self.assertRaises(SystemExit):
                runtime_config.get_shop_pin()


class TestGetAppId(unittest.TestCase):
    """M-29.5 get_app_id biz_type 路由"""

    def test_env_app_id_priority(self):
        """APP_ID_jm_order 环境变量应优先"""
        with patch.dict(
            os.environ,
            {"SHOP_ID": "FYA箱包旗舰店", "APP_ID_JM_ORDER": "TEST_APP_ID_123"},
            clear=True,
        ):
            self.assertEqual(
                runtime_config.get_app_id("jm_order"),
                "TEST_APP_ID_123",
            )

    def test_env_generic_app_id_fallback(self):
        """APP_ID（通用）环境变量应作为 fallback"""
        with patch.dict(
            os.environ,
            {"SHOP_ID": "FYA箱包旗舰店", "APP_ID": "GENERIC_APP_ID"},
            clear=True,
        ):
            self.assertEqual(
                runtime_config.get_app_id("jm_after_sale"),
                "GENERIC_APP_ID",
            )

    def test_read_from_config_xlsx(self):
        """config.xlsx「京麦接口/app_id_jm_order」读取"""
        with patch.dict(
            os.environ,
            {"SHOP_ID": "FYA箱包旗舰店"},
            clear=True,
        ):
            app_id = runtime_config.get_app_id("jm_order")
            # 真实 config.xlsx 应配置了 app_id（CQLEJWPYPFOVQBC8UFLQ）
            self.assertTrue(app_id, "config.xlsx 应配置 app_id_jm_order")
            self.assertGreater(len(app_id), 10)

    def test_missing_app_id_raises(self):
        """无任何 app_id 数据源 → SystemExit(3)"""
        with patch.dict(
            os.environ,
            {"SHOP_ID": "FYA箱包旗舰店"},
            clear=True,
        ):
            # mock _lookup_xlsx_value 返回 None
            with patch.object(runtime_config, "_lookup_xlsx_value", return_value=None):
                with self.assertRaises(SystemExit):
                    runtime_config.get_app_id("jm_order")


class TestGetSignSalt(unittest.TestCase):
    """M-29.6 get_sign_salt"""

    def test_read_from_config_xlsx(self):
        """config.xlsx「全局/sign_salt」或「商品流量来源/签名盐值」应可读"""
        with patch.dict(
            os.environ,
            {"SHOP_ID": "FYA箱包旗舰店"},
            clear=True,
        ):
            salt = runtime_config.get_sign_salt()
            self.assertTrue(salt, "签名盐值应可读")
            self.assertGreater(len(salt), 3)

    def test_fallback_to_old_key_with_warning(self):
        """新 key 缺失 → 读老 key + WARNING"""
        with patch.dict(
            os.environ,
            {"SHOP_ID": "FYA箱包旗舰店"},
            clear=True,
        ):
            # mock 新 key 缺失、老 key 存在
            def fake_lookup(group, var):
                if group == "全局" and var == "sign_salt":
                    return None
                if group == "商品流量来源" and var == "签名盐值":
                    return "fallback_salt_123"
                return None
            with patch.object(runtime_config, "_lookup_xlsx_value", side_effect=fake_lookup):
                self.assertEqual(
                    runtime_config.get_sign_salt(),
                    "fallback_salt_123",
                )


if __name__ == "__main__":
    print("=" * 60)
    print("runtime_config.py 单元测试（M-29，2026-08-24）")
    print("=" * 60)
    unittest.main(verbosity=2)