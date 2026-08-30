import configparser
from pathlib import Path
import mysql.connector

root = Path(__file__).resolve().parents[1]
cfg = configparser.ConfigParser()
cfg.read(root / 'config/mysql.local.ini', encoding='utf-8')
c = cfg['mysql']
db = mysql.connector.connect(host=c.get('host'), port=c.getint('port'), user=c.get('user'), password=c.get('password'), database=c.get('database'), charset=c.get('charset','utf8mb4'))
cur = db.cursor()
for view in ('vw_dashboard_shop_daily','vw_dashboard_orders_daily','vw_dashboard_promotion_daily'):
    cur.execute(f"SELECT * FROM `{view}` WHERE stat_date='2026-08-27'")
    rows = cur.fetchall()
    print(view, len(rows))
    for row in rows:
        print(row)
cur.close()
db.close()
