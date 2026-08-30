"""建立可回滚的 Derived/Enriched v1 MySQL 视图。"""
import configparser
from pathlib import Path

import mysql.connector

ROOT = Path(__file__).resolve().parents[1]
CONFIG_PATH = ROOT / "config" / "mysql.local.ini"
SQL_PATH = ROOT / "sql" / "create_derived_views.sql"


def main():
    cfg = configparser.ConfigParser()
    cfg.read(CONFIG_PATH, encoding="utf-8")
    db = cfg["mysql"]
    sql = SQL_PATH.read_text(encoding="utf-8")
    connection = mysql.connector.connect(
        host=db.get("host"),
        port=db.getint("port"),
        user=db.get("user"),
        password=db.get("password"),
        database=db.get("database"),
        charset=db.get("charset", "utf8mb4"),
    )
    try:
        cursor = connection.cursor()
        # 该文件只有一个 CREATE VIEW 语句；不启用 multi-statement，便于失败时定位。
        cursor.execute(sql)
        connection.commit()
        cursor.execute("SELECT COUNT(*) FROM `vw_derived_jm_order_full`")
        print(f"DERIVED_VIEW_OK rows={cursor.fetchone()[0]}", flush=True)
    finally:
        cursor.close()
        connection.close()


if __name__ == "__main__":
    main()
