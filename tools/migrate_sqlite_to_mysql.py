"""Stream the SQLite business database into MySQL raw tables.

The source database is never modified. Raw tables use ASCII table/column names
and retain the original SQLite column names in mysql_column_map.
"""
from __future__ import annotations

import configparser
import hashlib
import json
import sqlite3
import time
from pathlib import Path

import mysql.connector

ROOT = Path(__file__).resolve().parents[1]
SQLITE_PATH = ROOT / "data" / "jd_report.db"
CONFIG_PATH = ROOT / "config" / "mysql.local.ini"
BATCH_SIZE = 1000


def qi(value: str) -> str:
    return "`" + value.replace("`", "``") + "`"


def load_config() -> dict[str, str | int]:
    parser = configparser.ConfigParser()
    if not CONFIG_PATH.exists():
        raise SystemExit(f"missing config: {CONFIG_PATH}")
    parser.read(CONFIG_PATH, encoding="utf-8")
    section = parser["mysql"]
    return {
        "host": section.get("host", "127.0.0.1"),
        "port": section.getint("port", 3306),
        "database": section["database"],
        "user": section["user"],
        "password": section["password"],
        "charset": section.get("charset", "utf8mb4"),
    }


def raw_name(table: str) -> str:
    return "raw_" + table


def column_name(index: int) -> str:
    return f"c{index:04d}"


def main() -> None:
    cfg = load_config()
    sqlite_db = sqlite3.connect(f"file:{SQLITE_PATH.as_posix()}?mode=ro", uri=True)
    sqlite_db.row_factory = sqlite3.Row
    mysql_db = mysql.connector.connect(
        host=cfg["host"], port=cfg["port"], user=cfg["user"],
        password=cfg["password"], database=cfg["database"], charset=cfg["charset"],
    )
    mysql_db.autocommit = False
    cur = mysql_db.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS mysql_migration_runs (
            run_id BIGINT AUTO_INCREMENT PRIMARY KEY,
            source_path VARCHAR(512) NOT NULL,
            source_size BIGINT NOT NULL,
            source_sha256 CHAR(64) NOT NULL,
            started_at DATETIME NOT NULL,
            finished_at DATETIME NULL,
            status VARCHAR(32) NOT NULL,
            table_count INT NOT NULL DEFAULT 0,
            row_count BIGINT NOT NULL DEFAULT 0,
            error_text TEXT NULL
        ) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS mysql_column_map (
            source_table VARCHAR(128) NOT NULL,
            target_table VARCHAR(128) NOT NULL,
            source_index INT NOT NULL,
            source_column TEXT NOT NULL,
            target_column VARCHAR(16) NOT NULL,
            source_type VARCHAR(64) NULL,
            PRIMARY KEY (source_table, source_index)
        ) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci
    """)
    digest = hashlib.sha256()
    with SQLITE_PATH.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            digest.update(chunk)
    cur.execute(
        "INSERT INTO mysql_migration_runs "
        "(source_path,source_size,source_sha256,started_at,status) VALUES (%s,%s,%s,NOW(),'RUNNING')",
        (str(SQLITE_PATH), SQLITE_PATH.stat().st_size, digest.hexdigest()),
    )
    run_id = cur.lastrowid
    mysql_db.commit()
    total_rows = 0
    table_count = 0
    started = time.monotonic()
    try:
        tables = [r[0] for r in sqlite_db.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name"
        )]
        for source_table in tables:
            cols = sqlite_db.execute(f"PRAGMA table_info({qi(source_table)})").fetchall()
            target_table = raw_name(source_table)
            cur.execute(f"DROP TABLE IF EXISTS {qi(target_table)}")
            definitions = [f"{qi(column_name(i))} LONGTEXT NULL" for i, _ in enumerate(cols)]
            definitions.extend([
                "_source_table VARCHAR(128) NOT NULL",
                "_migration_run_id BIGINT NOT NULL",
            ])
            cur.execute(
                f"CREATE TABLE {qi(target_table)} ({','.join(definitions)}) "
                "CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci"
            )
            for i, col in enumerate(cols):
                cur.execute(
                    "REPLACE INTO mysql_column_map "
                    "(source_table,target_table,source_index,source_column, target_column,source_type) "
                    "VALUES (%s,%s,%s,%s,%s,%s)",
                    (source_table, target_table, i, col[1], column_name(i), col[2]),
                )
            placeholders = ",".join(["%s"] * (len(cols) + 2))
            insert_sql = f"INSERT INTO {qi(target_table)} VALUES ({placeholders})"
            batch = []
            table_rows = 0
            for row in sqlite_db.execute(f"SELECT * FROM {qi(source_table)}"):
                batch.append(tuple(None if v is None else str(v) for v in row) + (source_table, run_id))
                if len(batch) >= BATCH_SIZE:
                    cur.executemany(insert_sql, batch)
                    table_rows += len(batch)
                    total_rows += len(batch)
                    batch.clear()
            if batch:
                cur.executemany(insert_sql, batch)
                table_rows += len(batch)
                total_rows += len(batch)
            if any(c[1] == "shop_pin" for c in cols):
                idx_name = f"ix_{target_table}_shop"
                cur.execute(f"CREATE INDEX {qi(idx_name)} ON {qi(target_table)} ({qi(column_name(next(i for i,c in enumerate(cols) if c[1]=='shop_pin')))[:]}(128))")
            if any(c[1] in ("stat_date", "report_date") for c in cols):
                date_i = next(i for i,c in enumerate(cols) if c[1] in ("stat_date", "report_date"))
                idx_name = f"ix_{target_table}_date"
                cur.execute(f"CREATE INDEX {qi(idx_name)} ON {qi(target_table)} ({qi(column_name(date_i))}(32))")
            mysql_db.commit()
            table_count += 1
            print(f"migrated {source_table}: {table_rows}")
        cur.execute(
            "UPDATE mysql_migration_runs SET finished_at=NOW(),status='SUCCESS',table_count=%s,row_count=%s WHERE run_id=%s",
            (table_count, total_rows, run_id),
        )
        mysql_db.commit()
        print(json.dumps({"status":"SUCCESS","run_id":run_id,"tables":table_count,"rows":total_rows,"seconds":round(time.monotonic()-started,1)}, ensure_ascii=False))
    except Exception as exc:
        mysql_db.rollback()
        cur.execute("UPDATE mysql_migration_runs SET finished_at=NOW(),status='FAILED',error_text=%s WHERE run_id=%s", (repr(exc), run_id))
        mysql_db.commit()
        raise
    finally:
        cur.close()
        mysql_db.close()
        sqlite_db.close()


if __name__ == "__main__":
    main()
