# -*- coding: utf-8 -*-
"""TDD 测试：京准通快车 100 下载上限预检（2026-08-21）

目标：
1. 在 create_export_task 之前预检已有任务数，避免浪费配额
2. 轮询时检测到 100 上限错误立即停，不浪费 3 分钟

运行: python -m unittest tests.test_jzt_quota_check -v
"""
import os
import sys
import unittest
from unittest.mock import patch, MagicMock

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PROJECT_ROOT)
os.chdir(PROJECT_ROOT)

os.environ.setdefault("AUTH_LOADER", "1")
os.environ.setdefault("AUTH_SKIP_COOKIE_EXPIRE_CHECK", "1")
os.environ.setdefault("AUTH_SKIP_H5ST_EXPIRE_CHECK", "1")
os.environ["SHOP_ID"] = "FYA箱包旗舰店"


class TestJZTQuotaCheck(unittest.TestCase):
    """测试京准通快车 100 下载上限检测逻辑"""

    def test_100_limit_error_detection(self):
        """能正确识别 100 下载上限错误响应"""
        from main import JZTKuaicheAPI

        # 模拟 downloadById 返回 100 上限错误
        ret_quota_exceeded = {
            "success": False,
            "code": 0,
            "msg": "【操作失败】操作受限，下载报表数超过上限(100个), 您可以到下载报表页面查看历史所有的下载记录，删除一部分下载记录再重新触发",
            "data": None,
        }
        self.assertTrue(
            JZTKuaicheAPI._is_quota_exceeded_response(ret_quota_exceeded),
            "应能识别 100 上限错误",
        )

        # 模拟正常成功响应
        ret_success = {
            "success": True,
            "code": 0,
            "msg": "成功",
            "data": {"urlCsv": "https://storage.jd.com/xxx.csv"},
        }
        self.assertFalse(
            JZTKuaicheAPI._is_quota_exceeded_response(ret_success),
            "正常响应不应被误判为上限错误",
        )

        # 模拟其他错误
        ret_other_error = {
            "success": False,
            "code": 601,
            "msg": "登录失效",
            "data": None,
        }
        self.assertFalse(
            JZTKuaicheAPI._is_quota_exceeded_response(ret_other_error),
            "其他错误不应被误判为上限错误",
        )

    def test_task_count_quota_check(self):
        """根据 get_task_list 返回的任务数判断是否接近上限"""
        from main import JZTKuaicheAPI

        api = JZTKuaicheAPI()

        # 模拟 98 个任务
        mock_resp_98 = {
            "success": True,
            "code": 0,
            "data": {"data": [{"id": i} for i in range(98)]},
        }
        count_98 = api._count_tasks_from_list(mock_resp_98)
        self.assertEqual(count_98, 98)
        self.assertTrue(
            JZTKuaicheAPI._is_quota_near_limit(count_98, threshold=95),
            "98 个任务应触发配额告警（阈值95）",
        )

        # 模拟 50 个任务
        mock_resp_50 = {
            "success": True,
            "code": 0,
            "data": {"data": [{"id": i} for i in range(50)]},
        }
        count_50 = api._count_tasks_from_list(mock_resp_50)
        self.assertEqual(count_50, 50)
        self.assertFalse(
            JZTKuaicheAPI._is_quota_near_limit(count_50, threshold=95),
            "50 个任务不应触发配额告警",
        )

        # 空列表
        mock_resp_empty = {"success": True, "code": 0, "data": {"data": []}}
        count_empty = api._count_tasks_from_list(mock_resp_empty)
        self.assertEqual(count_empty, 0)

    def test_quota_precheck_raises_early(self):
        """配额预检在接近上限时应提前报错，不创建任务"""
        from main import JZTKuaicheAPI

        api = JZTKuaicheAPI()

        # 模拟已有 99 个任务 → 接近上限应抛异常
        with patch.object(api, "get_task_list") as mock_list:
            mock_list.return_value = {
                "success": True,
                "code": 0,
                "data": {"data": [{"id": i} for i in range(99)]},
            }
            with self.assertRaises(RuntimeError) as ctx:
                api._precheck_quota(threshold=95)
            self.assertIn("下载报表数接近上限", str(ctx.exception))
            self.assertIn("99", str(ctx.exception))

    def test_quota_precheck_passes_when_safe(self):
        """配额预检在安全时应通过（不抛异常，返回当前任务数）"""
        from main import JZTKuaicheAPI

        api = JZTKuaicheAPI()

        # 模拟已有 10 个任务 → 安全通过
        mock_data = [{"id": i} for i in range(10)]
        with patch.object(api, "get_task_list") as mock_list:
            mock_list.return_value = {
                "success": True,
                "code": 0,
                "data": {"data": mock_data},
            }
            result = api._precheck_quota(threshold=95)
            self.assertEqual(result, 10, f"应返回任务数 10，实际返回 {result}")

    def test_quota_precheck_calls_get_task_list(self):
        """预检必须调用 get_task_list（验证 mock 确实被调用）"""
        from main import JZTKuaicheAPI

        api = JZTKuaicheAPI()
        mock_data = [{"id": i} for i in range(10)]
        with patch.object(api, "get_task_list") as mock_list:
            mock_list.return_value = {
                "success": True,
                "code": 0,
                "data": {"data": mock_data},
            }
            api._precheck_quota(threshold=95)
            mock_list.assert_called_once()


if __name__ == "__main__":
    unittest.main(verbosity=2)
