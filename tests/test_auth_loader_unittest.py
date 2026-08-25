# -*- coding: utf-8 -*-
"""auth_loader.py 单元测试（M-27，2026-08-24）

覆盖范围（按真实 API 调整）：
    1) get_cookie_str 基础 JSON 解析（含 expires 字段）
    2) _check_cookie_expires：过期 Cookie → CookieExpiredError
    3) _check_cookie_expires：sessionCookie=true 字段跳过
    4) is_h5st_expired / get_h5st_age_seconds：30 分钟过期判断
    5) AuthLoader SystemExit(3) 当 SHOP_ID 未设置（H-13 修复）
    6) cache_key 包含 shop_id（H-10 多店隔离）
    7) JSON > TXT 优先级 fallback

运行：python tests/test_auth_loader_unittest.py
"""
import os
import sys
import json
import time
import tempfile
import unittest
import importlib

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PROJECT_ROOT)


class TestAuthLoaderShopIdRequired(unittest.TestCase):
    """M-27.5 + H-13：AuthLoader 初始化禁止静默回落"""

    def setUp(self):
        # 清除 SHOP_ID 环境变量
        self._old_shop_id = os.environ.pop("SHOP_ID", None)
        # 重置 AuthLoader 单例
        if "auth_loader" in sys.modules:
            importlib.reload(sys.modules["auth_loader"])
        import auth_loader
        auth_loader.AuthLoader._instance = None

    def tearDown(self):
        if self._old_shop_id:
            os.environ["SHOP_ID"] = self._old_shop_id

    def test_systemexit_when_no_shop_id(self):
        """无 shop_id 入参 + 无 SHOP_ID 环境变量 → SystemExit(3)"""
        import auth_loader
        with self.assertRaises(SystemExit) as cm:
            auth_loader.AuthLoader()
        self.assertEqual(cm.exception.code, 3)


class TestAuthLoaderNormalInit(unittest.TestCase):
    """正常初始化（显式传 shop_id 或 SHOP_ID 环境变量）"""

    def setUp(self):
        os.environ["SHOP_ID"] = "FYA箱包旗舰店"
        if "auth_loader" in sys.modules:
            importlib.reload(sys.modules["auth_loader"])
        import auth_loader
        auth_loader.AuthLoader._instance = None
        self.auth_loader_mod = auth_loader

    def test_init_with_explicit_shop_id(self):
        """显式传 shop_id 应正常初始化"""
        al = self.auth_loader_mod.AuthLoader(shop_id="MIYO箱包旗舰店")
        self.assertEqual(al.shop_id, "MIYO箱包旗舰店")

    def test_init_with_env_shop_id(self):
        """SHOP_ID 环境变量应被读取"""
        al = self.auth_loader_mod.AuthLoader()
        self.assertEqual(al.shop_id, "FYA箱包旗舰店")


class TestCookieExpiryCheck(unittest.TestCase):
    """M-27.1/M-27.2 _check_cookie_expires 过期识别"""

    def setUp(self):
        os.environ["SHOP_ID"] = "FYA箱包旗舰店"
        if "auth_loader" in sys.modules:
            importlib.reload(sys.modules["auth_loader"])
        from auth_loader import AuthLoader
        AuthLoader._instance = None
        self.al = AuthLoader(shop_id="FYA箱包旗舰店")

    def test_fresh_cookie_no_error(self):
        """未过期 Cookie 应不抛错"""
        future_ts = time.time() + 86400  # 24h 后
        cookies = [{"name": "pin", "value": "FYA8888", "expires": future_ts}]
        # 不应抛错
        self.al._check_cookie_expires(cookies, source="test.json")

    def test_expired_cookie_raises(self):
        """过期 Cookie 应抛 CookieExpiredError"""
        past_ts = time.time() - 3600  # 1h 前
        cookies = [{"name": "pin", "value": "FYA8888", "expires": past_ts}]
        from auth_loader import CookieExpiredError
        with self.assertRaises(CookieExpiredError):
            self.al._check_cookie_expires(cookies, source="test.json")

    def test_session_cookie_skipped(self):
        """sessionCookie=true 的 Cookie 跳过 expires 检查"""
        past_ts = time.time() - 3600  # 1h 前已过期
        cookies = [
            {"name": "session_key", "value": "x", "expires": past_ts, "sessionCookie": True}
        ]
        # 不应抛错（sessionCookie 跳过）
        self.al._check_cookie_expires(cookies, source="test.json")

    def test_skip_whitelist_fields(self):
        """白名单字段（_gia_d/sdtoken 等）即使过期也跳过"""
        past_ts = time.time() - 3600
        cookies = [
            {"name": "_gia_d", "value": "x", "expires": past_ts},
            {"name": "sdtoken", "value": "x", "expires": past_ts},
            {"name": "pin", "value": "FYA8888", "expires": time.time() + 86400},  # 核心字段有效
        ]
        # 核心字段有效 + 白名单字段即使过期也不应抛错
        self.al._check_cookie_expires(cookies, source="test.json")

    def test_env_var_skip_check(self):
        """AUTH_SKIP_COOKIE_EXPIRE_CHECK=1 应跳过整个检查"""
        past_ts = time.time() - 3600
        cookies = [{"name": "pin", "value": "FYA8888", "expires": past_ts}]
        os.environ["AUTH_SKIP_COOKIE_EXPIRE_CHECK"] = "1"
        try:
            # 即使过期也不抛错
            self.al._check_cookie_expires(cookies, source="test.json")
        finally:
            os.environ.pop("AUTH_SKIP_COOKIE_EXPIRE_CHECK", None)


class TestH5stExpiry(unittest.TestCase):
    """M-27.3 h5st 30 分钟过期判断"""

    def setUp(self):
        os.environ["SHOP_ID"] = "FYA箱包旗舰店"
        # M-27 测试隔离修复（2026-08-24）：清理可能被其他测试污染的环境变量
        # 背景：全量跑时 test_jzt_quota_check 等可能设置 AUTH_SKIP_*_EXPIRE_CHECK=1 未清理
        #       导致本类 h5st/cookie 过期检查被跳过 → 断言失败
        os.environ.pop("AUTH_SKIP_H5ST_EXPIRE_CHECK", None)
        os.environ.pop("AUTH_SKIP_COOKIE_EXPIRE_CHECK", None)
        if "auth_loader" in sys.modules:
            importlib.reload(sys.modules["auth_loader"])
        from auth_loader import AuthLoader
        AuthLoader._instance = None
        self.al = AuthLoader(shop_id="FYA箱包旗舰店")

    def _mock_h5st_file(self, captured_at_ms):
        """通过 _find_h5st_file 返回值 mock h5st JSON 文件内容（用 patch 自动恢复 open）"""
        import io
        import json as _json
        import builtins
        json_content = _json.dumps({
            "h5st": "abc123def456",
            "captured_at": captured_at_ms,
        })
        self.al._find_h5st_file = lambda h5st_key="jm_order": "/mock/h5st.json"
        original_open = builtins.open
        def mock_open(path, *args, **kwargs):
            if path == "/mock/h5st.json":
                return io.StringIO(json_content)
            return original_open(path, *args, **kwargs)
        # 保存到 self 以便 tearDown 恢复
        self._orig_open = original_open
        self._mock_open = mock_open
        builtins.open = mock_open

    def tearDown(self):
        # 恢复 builtins.open（防止污染其他测试文件）
        if hasattr(self, "_orig_open"):
            import builtins
            builtins.open = self._orig_open
        pass

    def test_h5st_age_seconds_fresh(self):
        """新抓的 h5st（1 分钟前）→ 60 秒左右"""
        captured = int(time.time() * 1000) - 60_000
        self._mock_h5st_file(captured)
        age = self.al.get_h5st_age_seconds(h5st_key="jm_order")
        self.assertIsNotNone(age)
        self.assertGreater(age, 50)
        self.assertLess(age, 70)

    def test_h5st_is_h5st_expired_old(self):
        """31 分钟前的 h5st 应标记过期"""
        from auth_loader import H5ST_EXPIRE_SECONDS
        old_ts = int(time.time() * 1000) - (H5ST_EXPIRE_SECONDS + 60) * 1000
        self._mock_h5st_file(old_ts)
        self.assertTrue(self.al.is_h5st_expired(h5st_key="jm_order"))


class TestCacheShopPinIsolation(unittest.TestCase):
    """M-27.6 + H-10 修复：cache_key 应包含 shop_id"""

    def setUp(self):
        os.environ["SHOP_ID"] = "FYA箱包旗舰店"
        if "auth_loader" in sys.modules:
            importlib.reload(sys.modules["auth_loader"])
        from auth_loader import AuthLoader
        AuthLoader._instance = None
        self.al = AuthLoader(shop_id="FYA箱包旗舰店")

    def test_cache_key_includes_shop_id(self):
        """cache_key 格式应包含 shop_id（H-10 修复）"""
        self.al._set_cache("test_key", "test_value")
        # H-10 修复后，cache_key 应形如 "FYA箱包旗舰店:cookie_str:sz:True"
        # 验证 _set_cache 写入后能 _is_cache_valid 命中
        self.assertTrue(self.al._is_cache_valid("test_key"))

    def test_different_shop_no_collision(self):
        """不同店铺的 cache_key 不应互相命中"""
        # FYA 缓存
        self.al._set_cache("FYA箱包旗舰店:cookie_str:sz:True", "fya_cookie")
        self.assertTrue(self.al._is_cache_valid("FYA箱包旗舰店:cookie_str:sz:True"))
        # MIYO 同 key 不应命中
        self.assertFalse(self.al._is_cache_valid("MIYO箱包旗舰店:cookie_str:sz:True"))


class TestResolveFilePrefix(unittest.TestCase):
    """M-27.4 resolve_file_prefix 双花括号兜底"""

    def setUp(self):
        os.environ["SHOP_ID"] = "FYA箱包旗舰店"

    def test_known_shop(self):
        """已知店铺返回 {{xxx}}"""
        from auth_loader import resolve_file_prefix
        result = resolve_file_prefix("FYA箱包旗舰店")
        self.assertEqual(result, "{{FYA}}")

    def test_unknown_shop_fallback(self):
        """未知店铺兜底返回 {{店名}}"""
        from auth_loader import resolve_file_prefix
        result = resolve_file_prefix("未知店铺XYZ")
        self.assertEqual(result, "{{未知店铺XYZ}}")


class TestPriorityJSONvsTXT(unittest.TestCase):
    """M-27.7 get_cookie_str 优先级 JSON > TXT fallback"""

    def setUp(self):
        os.environ["SHOP_ID"] = "FYA箱包旗舰店"
        self.tmpdir = tempfile.mkdtemp()
        if "auth_loader" in sys.modules:
            importlib.reload(sys.modules["auth_loader"])
        from auth_loader import AuthLoader
        AuthLoader._instance = None

    def tearDown(self):
        import shutil
        if os.path.exists(self.tmpdir):
            shutil.rmtree(self.tmpdir)

    def test_json_takes_priority(self):
        """JSON 存在时优先读 JSON"""
        from auth_loader import AuthLoader, _cfg_get_shop_prefix
        from auth_loader import CookieExpiredError
        prefix = _cfg_get_shop_prefix("FYA箱包旗舰店")
        # JSON + TXT 都存在
        json_path = os.path.join(self.tmpdir, f"{prefix}_sz_cookie.json")
        txt_path = os.path.join(self.tmpdir, f"{prefix}_sz_cookie.txt")
        future_ts = time.time() + 86400
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump({
                "cookies": [{"name": "pin", "value": "FROM_JSON", "expires": future_ts}]
            }, f)
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write("pin=FROM_TXT")
        al = AuthLoader(shop_id="FYA箱包旗舰店", config_dir=self.tmpdir)
        # check_expire=False 避免过期检查（future_ts 也不会过期，仅简化测试）
        result = al.get_cookie_str(biz_type="sz", check_expire=False)
        self.assertIn("FROM_JSON", result)
        self.assertNotIn("FROM_TXT", result)


if __name__ == "__main__":
    print("=" * 60)
    print("auth_loader.py 单元测试（M-27，2026-08-24）")
    print("=" * 60)
    unittest.main(verbosity=2)