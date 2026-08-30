"""Promote high-confidence Excel rows into a queryable, reversible unified layer."""
import configparser
import hashlib
import json
import re
import sqlite3
from pathlib import Path

import mysql.connector

ROOT = Path(__file__).resolve().parents[1]
TARGETS = {
    "订单列表": "biz_jm_order_full", "售后": "biz_jm_after_sale_full",
    "推广自定义报表": "biz_jzt_kuaiche", "全站": "biz_jzt_quanzhan_campaign",
    "全站-全店": "biz_jzt_quanzhan_campaign_all_store",
    "搜索": "biz_traffic_search", "推荐": "biz_traffic_recommend", "购物车": "biz_traffic_cart",
}

def norm(value):
    return re.sub(r"[^0-9a-zA-Z\u4e00-\u9fff]", "", str(value or "")).lower()

cfg = configparser.ConfigParser(); cfg.read(ROOT / "config" / "mysql.local.ini", encoding="utf-8")
c = cfg["mysql"]
sqlite_db = sqlite3.connect(f"file:{(ROOT / 'data' / 'jd_report.db').as_posix()}?mode=ro", uri=True)
columns = {t: [r[1] for r in sqlite_db.execute(f'PRAGMA table_info("{t}")')] for t in set(TARGETS.values())}
sqlite_db.close()
db = mysql.connector.connect(host=c.get("host"), port=c.getint("port"), user=c.get("user"), password=c.get("password"), database=c["database"], charset=c.get("charset", "utf8mb4"))
writer_db = mysql.connector.connect(host=c.get("host"), port=c.getint("port"), user=c.get("user"), password=c.get("password"), database=c["database"], charset=c.get("charset", "utf8mb4"))
cur = db.cursor()
cur.execute("""CREATE TABLE IF NOT EXISTS unified_legacy_rows (
 row_hash CHAR(64) PRIMARY KEY, business_type VARCHAR(128) NOT NULL, shop_code VARCHAR(64) NOT NULL,
 stat_date DATE NOT NULL, source_file VARCHAR(128) NOT NULL, source_sheet VARCHAR(128) NOT NULL,
 source_row INT NOT NULL, matched_field_count INT NOT NULL, data_json LONGTEXT NOT NULL,
 created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
 KEY ix_unified_type_shop_date(business_type, shop_code, stat_date)
) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci""")
cur.execute("SELECT source_file,source_sheet,headers_json FROM legacy_excel_sheet_meta")
meta = {(f,s): json.loads(h) for f,s,h in cur.fetchall()}
cur.execute("SELECT source_file,source_sheet,source_shop,source_row,inferred_date,row_hash,payload FROM legacy_excel_rows ORDER BY row_id")
insert = ("INSERT IGNORE INTO unified_legacy_rows "
          "(row_hash,business_type,shop_code,stat_date,source_file,source_sheet,source_row,matched_field_count,data_json) "
          "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)")
batch=[]; total=0; skipped=0
for source_file,sheet,shop,row_no,stat_date,row_hash,payload in cur:
    target = TARGETS.get(sheet)
    if not target:
        skipped += 1; continue
    headers = meta[(source_file,sheet)]
    values = json.loads(payload)
    lookup = {norm(x): x for x in columns[target]}
    data = {lookup[norm(h)]: values[i] for i,h in enumerate(headers) if h and i < len(values) and norm(h) in lookup}
    if len(data) / max(1, len([h for h in headers if h])) < 0.70:
        skipped += 1; continue
    batch.append((row_hash,target,shop,stat_date,source_file,sheet,row_no,len(data),json.dumps(data,ensure_ascii=False,separators=(",",":"))))
    if len(batch) >= 500:
        cur2 = writer_db.cursor(); cur2.executemany(insert,batch); writer_db.commit(); total += cur2.rowcount; cur2.close(); batch=[]
if batch:
    cur2 = writer_db.cursor(); cur2.executemany(insert,batch); writer_db.commit(); total += cur2.rowcount; cur2.close()
print(f"UNIFIED_INSERTED={total} SKIPPED_LOW_CONFIDENCE={skipped}")
cur.close(); db.close(); writer_db.close()
