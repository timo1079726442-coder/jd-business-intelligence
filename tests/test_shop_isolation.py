import unittest
from unittest.mock import patch

from metrics.service import MetricService


class FakeCursor:
    def __init__(self):
        self.sql = []
        self.params = []
    def execute(self, sql, params=()):
        self.sql.append(sql); self.params.append(tuple(params))
    def fetchone(self):
        sql = self.sql[-1]
        if "COUNT(*), MAX(stat_date)" in sql: return (1, "2026-08-27")
        if "COUNT(*) FROM std_biz_jm_order_full" in sql: return (1,)
        if "std_biz_jm_order_full" in sql and "SELECT" in sql: return (100.0, 5.0, 2, 1)
        return (10.0,)
    def fetchall(self): return [("2026-08-27", 1)]
    def close(self): pass


class FakeDb:
    def __init__(self, cur): self.cur = cur
    def cursor(self): return self.cur
    def close(self): pass


class ShopIsolationTests(unittest.TestCase):
    def test_miyo_queries_carry_shop_filter(self):
        cur = FakeCursor(); db = FakeDb(cur)
        with patch.object(MetricService, "_connect", return_value=db):
            result = MetricService().get_shop_overview("MIYO", "2026-08-20", "2026-08-26")
        self.assertEqual(result["shop"]["shop_key"], "MIYO")
        order_queries = [p for s, p in zip(cur.sql, cur.params) if "std_biz_jm_order_full" in s]
        self.assertTrue(order_queries)
        self.assertTrue(any("MIYO" in p and "miyo-周" in p for p in order_queries))

    def test_unknown_shop_rejected(self):
        with self.assertRaises(ValueError): MetricService().get_shop_overview("UNKNOWN", "2026-08-20", "2026-08-26")


if __name__ == "__main__": unittest.main()
