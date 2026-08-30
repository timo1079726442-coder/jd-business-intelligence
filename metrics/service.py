"""统一店铺指标 Service。

Dashboard、API 与未来 AI Tool 共用此服务；订单派生字段只从
``vw_derived_jm_order_full`` 读取，避免 Raw/Standard 层重复计算。
"""
from __future__ import annotations

import configparser
from datetime import date, datetime, timedelta
from pathlib import Path

import mysql.connector

ROOT = Path(__file__).resolve().parents[1]
CONFIG_PATH = ROOT / "config" / "mysql.local.ini"
cfg = configparser.ConfigParser()
cfg.read(CONFIG_PATH, encoding="utf-8")
DB = cfg["mysql"]

METRIC_VERSION = "legacy_miyo_v1"
SERVICE_VERSION = "shop-metrics-1.1.0"
# 展示名称 -> 数据库店铺 pin。MIYO 的 pin 使用数据库中真实值。
PINS = {"MIYO": "miyo-周", "OTA": "ota8888", "FYA": "FYA8888"}


def _qid(name: str) -> str:
    return "`" + name.replace("`", "``") + "`"


def _date_expr(field: str) -> str:
    return f"DATE(REPLACE(SUBSTRING({_qid(field)},1,10), '/', '-'))"


def _money_expr(field: str) -> str:
    """将平台导出金额文本安全转换为 DECIMAL。"""
    q = _qid(field)
    cleaned = f"REPLACE(REPLACE(REPLACE(REPLACE(TRIM({q}), ',', ''), '￥', ''), '¥', ''), '元', '')"
    return f"CAST(NULLIF({cleaned}, '') AS DECIMAL(20, 4))"


class MetricService:
    """业务指标查询边界。"""

    def __init__(self, db_config=None):
        self.db_config = db_config or DB

    def _connect(self):
        return mysql.connector.connect(
            host=self.db_config.get("host"),
            port=self.db_config.getint("port"),
            user=self.db_config.get("user"),
            password=self.db_config.get("password"),
            database=self.db_config.get("database"),
            charset=self.db_config.get("charset", "utf8mb4"),
        )

    @staticmethod
    def periods(start=None, end=None):
        end = end or date.today().isoformat()
        start = start or (date.fromisoformat(end) - timedelta(days=6)).isoformat()
        begin, finish = date.fromisoformat(start), date.fromisoformat(end)
        if begin > finish:
            raise ValueError("start_date cannot be after end_date")
        days = (finish - begin).days + 1
        return start, end, (begin - timedelta(days=days)).isoformat(), (begin - timedelta(days=1)).isoformat()

    def _metrics(self, cur, shop_key, start, end):
        pin = PINS[shop_key]
        sql = """
            SELECT
                COALESCE(SUM(order_split_amount), 0),
                COALESCE(SUM(order_refund_aggregated_amount), 0),
                COALESCE(SUM(CASE WHEN order_count_flag = 1 AND shipment_status_normalized = '已出库' AND order_split_amount > 10 THEN 1 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN order_count_flag = 1 AND shipment_status_normalized = '已出库' AND order_refund_aggregated_amount > 10 THEN 1 ELSE 0 END), 0)
            FROM vw_derived_jm_order_full
            WHERE (shop_pin = %s OR shop_pin = %s)
              AND payment_confirmed_date BETWEEN %s AND %s
        """
        cur.execute(sql, (shop_key, pin, start, end))
        gross, refund_amount, orders, return_orders = cur.fetchone()
        gross, refund_amount = float(gross or 0), float(refund_amount or 0)
        orders, return_orders = int(orders or 0), int(return_orders or 0)

        # 推广花费暂按现有三张表合计，指标状态明确标记为待确认。
        spend = 0.0
        for table in ("std_biz_jzt_kuaiche", "std_biz_jzt_quanzhan_campaign", "std_biz_jzt_quanzhan_campaign_all_store"):
            cur.execute(
                f"SELECT COALESCE(SUM({_money_expr('花费')}), 0) FROM {_qid(table)} "
                f"WHERE (shop_pin = %s OR shop_pin = %s) AND {_date_expr('日期')} BETWEEN %s AND %s",
                (shop_key, pin, start, end),
            )
            spend += float(cur.fetchone()[0] or 0)

        net = gross - refund_amount
        return {
            "gmv": gross,
            "refund_amount": refund_amount,
            "net_gmv": net,
            "refund_rate": refund_amount / gross if gross else None,
            "order_count": orders,
            "refund_order_count": return_orders,
            "return_rate": return_orders / orders if orders else None,
            "ad_cost": spend,
            "roi": net / spend if spend else None,
            "estimated_profit": net * (1 - 0.09 - 0.03 - 0.7) - spend,
        }

    def _quality(self, cur, shop_key, start, end):
        pin = PINS[shop_key]
        cur.execute("SELECT COUNT(*), MAX(stat_date) FROM vw_dashboard_shop_daily WHERE shop_code = %s AND stat_date BETWEEN %s AND %s", (shop_key, start, end))
        rows, latest = cur.fetchone()
        cur.execute("SELECT COUNT(*) FROM std_biz_jm_order_full WHERE (shop_pin = %s OR shop_pin = %s) AND stat_date BETWEEN %s AND %s", (shop_key, pin, start, end))
        order_rows = cur.fetchone()[0]
        return {"complete": bool(rows and order_rows), "missing": not bool(rows), "empty": bool(rows == 0 and order_rows == 0), "failed": False, "stale": False, "latest_data_date": str(latest) if latest else None, "latest_update_at": None, "missing_businesses": [], "missing_dates": []}

    def get_shop_overview(self, shop_key, start=None, end=None):
        if shop_key not in PINS:
            raise ValueError("invalid shop_key")
        start, end, compare_start, compare_end = self.periods(start, end)
        db = self._connect(); cur = db.cursor()
        try:
            current = self._metrics(cur, shop_key, start, end)
            comparison = self._metrics(cur, shop_key, compare_start, compare_end)
            quality = self._quality(cur, shop_key, start, end)
            cur.execute("SELECT stat_date, SUM(row_count) FROM vw_dashboard_shop_daily WHERE shop_code = %s AND stat_date BETWEEN %s AND %s GROUP BY stat_date ORDER BY stat_date DESC", (shop_key, start, end))
            daily = [{"date": str(r[0]), "record_count": int(r[1] or 0)} for r in cur.fetchall()]
        finally:
            cur.close(); db.close()

        metrics = {}
        for key, value in current.items():
            comp = comparison.get(key)
            change = (value / comp - 1) if isinstance(value, (int, float)) and comp not in (None, 0) else None
            unit = "ratio" if key in ("refund_rate", "return_rate", "roi") else ("count" if "count" in key else "CNY")
            status = "pending_confirmation" if key in ("ad_cost", "roi") else ("pending_business_semantics" if key == "estimated_profit" else "legacy_reference")
            metrics[key] = {"value": value, "comparison_value": comp, "change": change, "unit": unit, "status": status}

        now = datetime.now().astimezone().isoformat()
        return {"shop": {"shop_key": shop_key, "shop_pin": PINS[shop_key]}, "period": {"start": start, "end": end}, "comparison_period": {"start": compare_start, "end": compare_end}, "metrics": metrics, "data_quality": quality, "daily": daily, "meta": {"metric_version": METRIC_VERSION, "derived_rule_version": METRIC_VERSION, "generated_at": now, "data_updated_at": quality["latest_data_date"], "service_version": SERVICE_VERSION}}
