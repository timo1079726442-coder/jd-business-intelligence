"""将 SQLite 业务表迁移为字段清晰的 MySQL std_ 标准表。"""
import configparser
import re
import sqlite3
from pathlib import Path
import mysql.connector

ROOT = Path(__file__).resolve().parents[1]
cfg = configparser.ConfigParser(); cfg.read(ROOT / "config/mysql.local.ini", encoding="utf-8")
c = cfg["mysql"]
sqlite_path = ROOT / "data/jd_report.db"

def qi(name):
    return "`" + str(name).replace("`", "``") + "`"

sq = sqlite3.connect(sqlite_path)
my = mysql.connector.connect(host=c.get("host"), port=c.getint("port"), user=c.get("user"), password=c.get("password"), database=c.get("database"), charset=c.get("charset", "utf8mb4"))
mc = my.cursor()
tables = [r[0] for r in sq.execute("SELECT name FROM sqlite_master WHERE type='table' AND name LIKE 'biz_%' ORDER BY name")]
total = 0
for source in tables:
    target = "std_" + source
    info = sq.execute(f"PRAGMA table_info({qi(source)})").fetchall()
    columns = [row[1] for row in info if row[1].lower() != "id"]
    if not columns:
        continue
    mc.execute(f"DROP TABLE IF EXISTS {qi(target)}")
    defs = ["`id` BIGINT NOT NULL AUTO_INCREMENT PRIMARY KEY"]
    for col in columns:
        defs.append(f"{qi(col)} LONGTEXT NULL")
    if "shop_pin" in columns:
        defs.append("KEY `ix_shop_pin` (`shop_pin`(128))")
    if "stat_date" in columns:
        defs.append("KEY `ix_stat_date` (`stat_date`(32))")
    mc.execute(f"CREATE TABLE {qi(target)} ({', '.join(defs)}) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci")
    placeholders = ",".join(["%s"] * len(columns))
    insert = f"INSERT INTO {qi(target)} ({','.join(qi(x) for x in columns)}) VALUES ({placeholders})"
    cur = sq.execute(f"SELECT {','.join(qi(x) for x in columns)} FROM {qi(source)}")
    batch = []
    count = 0
    for row in cur:
        batch.append(tuple(None if v is None else str(v) for v in row))
        if len(batch) >= 1000:
            mc.executemany(insert, batch); my.commit(); count += len(batch); batch = []
    if batch:
        mc.executemany(insert, batch); my.commit(); count += len(batch)
    total += count
    print(f"{source} -> {target}: {count}", flush=True)
print(f"STANDARD_TABLES_OK tables={len(tables)} rows={total}", flush=True)
mc.close(); my.close(); sq.close()
