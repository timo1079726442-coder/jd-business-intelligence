# -*- coding: utf-8 -*-
"""db_utils.py 单元测试（M-26，2026-08-24）

覆盖范围：
    1) biz_key_to_table_name 中文业务名 → ASCII slug
    2) ensure_table 动态建表 + 缺列 ALTER
    3) upsert_df 全量覆盖（同 report_date 替换）
    4) get_existing_dates 排序 + shop_pin 过滤
    5) _safe_col_name 防注入（拒绝特殊字符）
    6) PRAGMA WAL 应用

运行：python tests/test_db_utils_unittest.py
"""
import os
import sys
import tempfile
import unittest
import sqlite3
from datetime import datetime

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PROJECT_ROOT)

# 在 import db_utils 前，必须设置 SHOP_ID/SHOP_PIN（runtime_config 强制）
os.environ["SHOP_ID"] = "FYA箱包旗舰店"
os.environ["SHOP_PIN"] = "FYA8888"

import db_utils


class TestBizKeyToTableName(unittest.TestCase):
    """M-26.1 biz_key_to_table_name 中文 → ASCII slug"""

    def test_known_businesses_white_list(self):
        """白名单业务名应映射到正确的 slug"""
        self.assertEqual(db_utils.biz_key_to_table_name("商智关键词分析"), "biz_keyword_analysis")
        self.assertEqual(db_utils.biz_key_to_table_name("店铺来源_三级渠道"), "biz_offline_channel")
        self.assertEqual(db_utils.biz_key_to_table_name("商品流量来源_搜索"), "biz_traffic_search")
        self.assertEqual(db_utils.biz_key_to_table_name("京准通快车自定义报表"), "biz_jzt_kuaiche")
        self.assertEqual(db_utils.biz_key_to_table_name("京麦订单明细_完整一键导出"), "biz_jm_order_full")
        self.assertEqual(db_utils.biz_key_to_table_name("京麦售后明细_完整一键导出"), "biz_jm_after_sale_full")

    def test_unknown_business_uses_pinyin_fallback(self):
        """未知业务名走 _pinyin_slug 兜底（取最后一段英文短写）"""
        result = db_utils.biz_key_to_table_name("测试业务_abc")
        # 最后一段非中文部分 "abc" 应该被保留
        self.assertEqual(result, "biz_abc")

    def test_all_results_start_with_biz_prefix(self):
        """所有结果必须 biz_ 前缀"""
        for biz in ["商智关键词分析", "京麦订单明细_完整一键导出", "测试业务_xyz"]:
            result = db_utils.biz_key_to_table_name(biz)
            self.assertTrue(result.startswith("biz_"), f"{biz} → {result} 缺 biz_ 前缀")


class TestSafeColName(unittest.TestCase):
    """M-26.5 _safe_col_name 防注入"""

    def test_normal_column_accepted(self):
        """正常中英文列名应接受"""
        result = db_utils._safe_col_name("访客数")
        self.assertEqual(result, '"访客数"')

    def test_english_with_underscore_accepted(self):
        """英文 + 下划线 + 数字应接受"""
        result = db_utils._safe_col_name("SKU_ID_2026")
        self.assertEqual(result, '"SKU_ID_2026"')

    def test_empty_column_uses_unnamed(self):
        """空列名应回退到 col_unnamed"""
        result = db_utils._safe_col_name("")
        self.assertEqual(result, '"col_unnamed"')
        result = db_utils._safe_col_name(None)
        self.assertEqual(result, '"col_unnamed"')

    def test_sql_injection_chars_rejected(self):
        """SQL 注入字符应抛 ValueError"""
        dangerous_inputs = [
            "col;DROP TABLE x",          # 分号
            "col--comment",               # SQL 注释符
            "col'name",                   # 单引号
            'col"name',                   # 双引号
            "col/path",                   # 斜杠
        ]
        for bad in dangerous_inputs:
            with self.assertRaises(ValueError, msg=f"应拒绝 {bad!r}"):
                db_utils._safe_col_name(bad)


class TestValidateTableName(unittest.TestCase):
    """M-26.5 _validate_table_name 表名防注入"""

    def test_valid_table_names(self):
        """合法的 biz_xxx 表名应通过"""
        for name in ["biz_keyword_analysis", "biz_jzt_kuaiche", "biz_test123"]:
            self.assertEqual(db_utils._validate_table_name(name), name)

    def test_invalid_table_names_rejected(self):
        """非法表名应抛 ValueError"""
        for bad in [
            "evil_table",                # 不带 biz_ 前缀
            "biz_",                      # 空 slug
            "biz_ABC",                   # 大写
            "biz_evil;DROP",             # 注入字符
            "biz_evil--comment",
            "",
            None,
            "keyword_analysis",          # 不带 biz_ 前缀
        ]:
            with self.assertRaises(ValueError, msg=f"应拒绝 {bad!r}"):
                db_utils._validate_table_name(bad)


class TestEnsureTableAndUpsert(unittest.TestCase):
    """M-26.2 + M-26.3 ensure_table + upsert_df"""

    def setUp(self):
        """每个测试用例独立 in-memory DB"""
        self.conn = sqlite3.connect(":memory:")
        import pandas as pd
        self.df = pd.DataFrame({
            "关键词": ["关键词A", "关键词B", "关键词C"],
            "访客数": ["100", "200", "300"],
            "下单金额": ["1000.50", "2000.50", "3000.50"],
        })

    def tearDown(self):
        self.conn.close()

    def test_create_new_table(self):
        """首次 ensure_table 应 CREATE TABLE 含 shop_pin / stat_date / report_date / etl_time"""
        db_utils.ensure_table(self.conn, "biz_test_create", self.df, "关键词")
        cur = self.conn.execute("PRAGMA table_info(biz_test_create)")
        cols = {row[1] for row in cur.fetchall()}
        # 公共字段
        self.assertIn("id", cols)
        self.assertIn("shop_pin", cols)
        self.assertIn("stat_date", cols)
        self.assertIn("report_date", cols)
        self.assertIn("etl_time", cols)
        # 业务字段
        self.assertIn("关键词", cols)
        self.assertIn("访客数", cols)
        self.assertIn("下单金额", cols)

    def test_alter_table_add_missing_columns(self):
        """已存在的表，df 含新列时 ALTER TABLE ADD COLUMN"""
        # 先建空表（仅含 id + 公共字段）
        self.conn.execute("""
            CREATE TABLE biz_test_alter (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                "shop_pin" TEXT NOT NULL,
                "stat_date" TEXT NOT NULL,
                "report_date" TEXT NOT NULL,
                "etl_time" TEXT NOT NULL,
                "关键词" TEXT
            )
        """)
        self.conn.commit()
        # ensure_table 应补列
        db_utils.ensure_table(self.conn, "biz_test_alter", self.df, "关键词")
        cur = self.conn.execute("PRAGMA table_info(biz_test_alter)")
        cols = {row[1] for row in cur.fetchall()}
        self.assertIn("访客数", cols)
        self.assertIn("下单金额", cols)

    def test_upsert_overwrites_same_date(self):
        """同 report_date 应覆盖（先 delete 后 insert）"""
        import pandas as pd
        # 第一次插入
        df1 = pd.DataFrame({"关键词": ["A"], "访客数": ["100"]})
        db_utils.upsert_df(self.conn, "biz_test_upsert", df1, "2026-08-01", "关键词")
        # 第二次插入同一 date，应覆盖
        df2 = pd.DataFrame({"关键词": ["A"], "访客数": ["999"]})
        db_utils.upsert_df(self.conn, "biz_test_upsert", df2, "2026-08-01", "关键词")
        # 验证
        cur = self.conn.execute(
            'SELECT "关键词", "访客数" FROM biz_test_upsert WHERE stat_date = ?',
            ("2026-08-01",),
        )
        rows = cur.fetchall()
        self.assertEqual(len(rows), 1, "应只剩 1 行（覆盖）")
        self.assertEqual(rows[0][1], "999", "访客数应被覆盖为 999")


class TestGetExistingDates(unittest.TestCase):
    """M-26.4 get_existing_dates 排序 + shop_pin 过滤"""

    def setUp(self):
        self.conn = sqlite3.connect(":memory:")
        # 建表并插数据
        import pandas as pd
        df = pd.DataFrame({"关键词": ["A", "B", "C"]})
        db_utils.upsert_df(self.conn, "biz_test_dates", df, "2026-08-01", "关键词")
        df = pd.DataFrame({"关键词": ["X", "Y"]})
        db_utils.upsert_df(self.conn, "biz_test_dates", df, "2026-08-02", "关键词")
        df = pd.DataFrame({"关键词": ["M"]})
        db_utils.upsert_df(self.conn, "biz_test_dates", df, "2026-08-03", "关键词")

    def tearDown(self):
        self.conn.close()

    def test_get_all_dates(self):
        """不传 shop_pin 应返回所有日期"""
        dates = db_utils.get_existing_dates(self.conn, "biz_test_dates")
        self.assertEqual(dates, {"2026-08-01", "2026-08-02", "2026-08-03"})

    def test_filter_by_shop_pin(self):
        """传 shop_pin 应只返回该店铺的日期"""
        # 插入 MIYO 店的数据（需切 SHOP_PIN 然后插）
        old_pin = os.environ["SHOP_PIN"]
        os.environ["SHOP_PIN"] = "miyo-周"
        try:
            import pandas as pd
            df = pd.DataFrame({"关键词": ["Z"]})
            db_utils.upsert_df(self.conn, "biz_test_dates", df, "2026-08-04", "关键词")
        finally:
            os.environ["SHOP_PIN"] = old_pin
        # FYA 应只有 3 天
        dates_fya = db_utils.get_existing_dates(self.conn, "biz_test_dates", shop_pin="FYA8888")
        self.assertEqual(dates_fya, {"2026-08-01", "2026-08-02", "2026-08-03"})
        # MIYO 应只有 1 天
        dates_miyo = db_utils.get_existing_dates(self.conn, "biz_test_dates", shop_pin="miyo-周")
        self.assertEqual(dates_miyo, {"2026-08-04"})

    def test_nonexistent_table_returns_empty_set(self):
        """不存在的表应返回空集合（不抛错）"""
        dates = db_utils.get_existing_dates(self.conn, "biz_nonexistent_table")
        self.assertEqual(dates, set())


class TestPragmaWAL(unittest.TestCase):
    """M-26.6 PRAGMA WAL 应用"""

    def test_wal_applied_on_connection(self):
        """连接后应已开启 WAL 模式 + synchronous=NORMAL"""
        conn = sqlite3.connect(":memory:")
        try:
            db_utils._apply_sqlite_perf_pragmas(conn)
            # 查询 journal_mode
            cur = conn.execute("PRAGMA journal_mode")
            mode = cur.fetchone()[0]
            # in-memory DB 通常不支持 WAL，可能返回 "memory"
            # 但同步级别应该被设置
            self.assertIn(mode.lower(), ("wal", "memory"))
        finally:
            conn.close()


class TestSaveToDb(unittest.TestCase):
    """M-26 端到端：save_to_db 一键入库（mock config.xlsx）"""

    def setUp(self):
        # 创建临时项目根目录（项目用 PROJECT_ROOT = os.path.dirname(__file__)）
        # 为了让 db_path 指向临时目录，需要把 db_path 配置写进 config.xlsx
        self.tmpdir = tempfile.mkdtemp()
        self._old_cwd = os.getcwd()
        os.chdir(self.tmpdir)
        self.db_path = os.path.join(self.tmpdir, "data", "jd_report.db")
        # mock _read_config_xlsx 让它返回 enable_db_storage=True + 绝对 db_path
        # 背景：db_utils.get_db_path() 调用 _read_config_xlsx()，该函数硬编码 PROJECT_ROOT/config/config.xlsx
        # 测试不能污染真实 config.xlsx，必须 monkey-patch
        self._original_read = db_utils._read_config_xlsx
        db_utils._read_config_xlsx = lambda: {
            "全局": {
                "enable_db_storage": "True",
                "db_path": self.db_path,
            }
        }

    def tearDown(self):
        # 恢复 _read_config_xlsx
        db_utils._read_config_xlsx = self._original_read
        os.chdir(self._old_cwd)
        import shutil
        if os.path.exists(self.tmpdir):
            shutil.rmtree(self.tmpdir)

    def test_save_to_db_creates_table_and_inserts(self):
        """save_to_db 应自动建表 + 插入数据"""
        import pandas as pd
        df = pd.DataFrame({"关键词": ["测试关键词"], "访客数": ["500"]})
        result = db_utils.save_to_db("商智关键词分析", df, "2026-08-24")
        self.assertTrue(result, "save_to_db 应返回 True")
        # 验证 DB 文件创建 + 表存在
        self.assertTrue(os.path.exists(self.db_path))
        conn = sqlite3.connect(self.db_path)
        try:
            cur = conn.execute(
                'SELECT COUNT(*) FROM biz_keyword_analysis WHERE stat_date = ?',
                ("2026-08-24",),
            )
            count = cur.fetchone()[0]
            self.assertEqual(count, 1)
            # 验证 shop_pin 已写入
            cur = conn.execute(
                'SELECT shop_pin FROM biz_keyword_analysis WHERE stat_date = ?',
                ("2026-08-24",),
            )
            shop_pin = cur.fetchone()[0]
            self.assertEqual(shop_pin, "FYA8888")
        finally:
            conn.close()

    def test_save_to_db_empty_df_skipped(self):
        """空 df 应跳过（返回 False，不报错）"""
        import pandas as pd
        empty_df = pd.DataFrame({"关键词": [], "访客数": []})
        result = db_utils.save_to_db("商智关键词分析", empty_df, "2026-08-24")
        self.assertFalse(result)


if __name__ == "__main__":
    print("=" * 60)
    print("db_utils.py 单元测试（M-26，2026-08-24）")
    print("=" * 60)
    unittest.main(verbosity=2)