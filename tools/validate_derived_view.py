"""检查 Derived 视图的行数、店铺隔离和关键派生字段。"""
import configparser
from pathlib import Path

import mysql.connector

ROOT = Path(__file__).resolve().parents[1]
cfg = configparser.ConfigParser(); cfg.read(ROOT / "config" / "mysql.local.ini", encoding="utf-8")
c = cfg["mysql"]
db = mysql.connector.connect(host=c.get("host"), port=c.getint("port"), user=c.get("user"), password=c.get("password"), database=c.get("database"), charset=c.get("charset", "utf8mb4"))
cur = db.cursor(dictionary=True)
cur.execute("SELECT COUNT(*) AS n, COUNT(DISTINCT shop_pin) AS shops, SUM(order_count_flag) AS orders, SUM(order_split_amount) AS amount, SUM(order_refund_aggregated_amount) AS refunds FROM vw_derived_jm_order_full")
print(cur.fetchone())
cur.execute("SELECT shop_pin, COUNT(*) AS n, SUM(order_count_flag) AS orders, ROUND(SUM(order_split_amount),2) AS amount, ROUND(SUM(order_refund_aggregated_amount),2) AS refunds FROM vw_derived_jm_order_full GROUP BY shop_pin ORDER BY shop_pin")
for row in cur.fetchall(): print(row)
cur.execute("SELECT id, shop_pin, `订单号`, `付款确认时间`, `订单状态`, `订单类型`, order_split_amount, order_count_flag, shipment_status_normalized FROM vw_derived_jm_order_full WHERE shop_pin='miyo-周' ORDER BY id LIMIT 5")
for row in cur.fetchall(): print(row)
cur.close(); db.close()
