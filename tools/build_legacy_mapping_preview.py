"""Build a read-only mapping preview from Excel headers to SQLite business fields."""
import configparser
import json
import re
import sqlite3
from pathlib import Path

import mysql.connector

ROOT = Path(__file__).resolve().parents[1]
CFG = configparser.ConfigParser()
CFG.read(ROOT / "config" / "mysql.local.ini", encoding="utf-8")
MYSQL = CFG["mysql"]

SHEET_TARGETS = {
    "订单列表": "biz_jm_order_full",
    "售后": "biz_jm_after_sale_full",
    "推广自定义报表": "biz_jzt_kuaiche",
    "快车订单": "biz_jzt_kuaiche_order_effect",
    "全站": "biz_jzt_quanzhan_campaign",
    "全站-全店": "biz_jzt_quanzhan_campaign_all_store",
    "全站订单": "biz_jzt_quanzhan_effect",
    "搜索": "biz_traffic_search",
    "推荐": "biz_traffic_recommend",
    "购物车": "biz_traffic_cart",
}


def normalize(value):
    return re.sub(r"[^0-9a-zA-Z\u4e00-\u9fff]", "", str(value or "")).lower()


sqlite_db = sqlite3.connect(f"file:{(ROOT / 'data' / 'jd_report.db').as_posix()}?mode=ro", uri=True)
source_columns = {}
for target in set(SHEET_TARGETS.values()):
    source_columns[target] = [r[1] for r in sqlite_db.execute(f'PRAGMA table_info("{target}")')]
sqlite_db.close()

mysql_db = mysql.connector.connect(
    host=MYSQL.get("host"), port=MYSQL.getint("port"), user=MYSQL.get("user"),
    password=MYSQL.get("password"), database=MYSQL.get("database"),
    charset=MYSQL.get("charset", "utf8mb4"),
)
cursor = mysql_db.cursor()
cursor.execute("SELECT source_file, source_sheet, headers_json FROM legacy_excel_sheet_meta ORDER BY source_file, source_sheet")
items = []
for source_file, source_sheet, headers_json in cursor.fetchall():
    target = SHEET_TARGETS.get(source_sheet)
    if not target:
        continue
    headers = json.loads(headers_json)
    lookup = {normalize(column): column for column in source_columns[target]}
    matches = []
    unmatched = []
    for index, header in enumerate(headers):
        if not header:
            continue
        source_column = lookup.get(normalize(header))
        if source_column:
            matches.append({"excel_index": index, "excel_header": header, "target_column": source_column})
        else:
            unmatched.append(header)
    cursor.execute(
        "SELECT COUNT(*), MIN(inferred_date), MAX(inferred_date) FROM legacy_excel_rows "
        "WHERE source_file=%s AND source_sheet=%s",
        (source_file, source_sheet),
    )
    row_count, date_min, date_max = cursor.fetchone()
    items.append({
        "source_file": source_file,
        "source_sheet": source_sheet,
        "target_table": target,
        "historical_rows": row_count,
        "date_min": str(date_min) if date_min else None,
        "date_max": str(date_max) if date_max else None,
        "matched_columns": len(matches),
        "total_headers": len([h for h in headers if h]),
        "matches": matches,
        "unmatched_headers": unmatched,
    })
cursor.close()
mysql_db.close()

out = ROOT / "data" / "legacy_mapping_preview.json"
out.write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")
for item in items:
    print(f"{item['source_file']} | {item['source_sheet']} -> {item['target_table']} | "
          f"{item['matched_columns']}/{item['total_headers']} fields | {item['historical_rows']} rows | "
          f"{item['date_min']}~{item['date_max']}")
print(f"preview={out}")
