"""Finance pending/settled export cleaner and fact loader.

The finance export has two independent sheets.  The cleaner never joins or
aggregates those sheets; it keeps transaction and fee detail at their source
grain and adds only lineage/control columns.  No existing business table is
modified by this module.
"""
from __future__ import annotations

import hashlib
import json
import re
import sqlite3
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Iterable

SHEET_TRANSACTION = "交易汇总"
SHEET_FEE = "费用明细"

HEADER_MAP = {
    "商家ID": "vendor_id",
    "公司ID": "company_id",
    "公司名称": "company_name",
    "钱包账户": "wallet_account",
    "订单编号": "order_id",
    "业务单据编号": "rf_busi_id",
    "父单号": "parent_id",
    "订单状态": "order_status",
    "下单时间": "date_submit",
    "完成时间": "order_finish_time",
    "商品编号": "sku_id",
    "商品名称": "sku_name",
    "商品单价": "sku_price",
    "商品数量": "sku_num",
    "预计结算时间": "expected_settlement_time",
    "结算时间": "financial_settled_time",
    "结算金额合计": "settle_amount",
    "收入金额合计": "income_sum",
    "支出金额合计": "outcome_sum",
    "用户支付金额": "user_paid_amount",
    "平台扣费基数": "platform_fee_base",
    "费率": "platform_fee_rate",
    "费用名称": "fee_name",
    "费用分类": "fee_category",
    "金额": "fee_amount",
    "对账公式": "bill_formula_desc",
    "结算主体": "settle_org",
    "商户订单号": "bank_flow",
    "资金动账备注": "remark",
    "结算单类型": "detail_type",
    "结算单号": "statement_id",
    "条款规则": "charge_provision",
    "费用备注": "fee_remark",
    "商户承担金额": "merchant_bear_amount",
    "平台承担金额": "platform_bear_amount",
    "政府承担金额": "government_bear_amount",
    "单独结算资产": "separate_settlement_asset",
}

DECIMAL_FIELDS = {
    "sku_price", "sku_num", "settle_amount", "income_sum", "outcome_sum",
    "user_paid_amount", "platform_fee_base", "platform_fee_rate", "fee_amount",
    "merchant_bear_amount", "platform_bear_amount", "government_bear_amount",
    "separate_settlement_asset",
}
DATETIME_FIELDS = {
    "date_submit", "order_finish_time", "expected_settlement_time", "financial_settled_time"
}
KEY_FIELDS = ("order_id", "rf_busi_id", "parent_id", "sku_id", "statement_id", "bank_flow")


@dataclass
class FinanceIngestResult:
    shop_key: str
    report_type: str
    bill_status: int
    task_type: int
    run_id: str
    source_file: str
    file_hash: str
    transaction_rows: int = 0
    fee_rows: int = 0
    duplicate_rows: int = 0
    quarantined_rows: int = 0
    missing_required_keys: int = 0
    parse_errors: int = 0
    schema_drift: list[str] = field(default_factory=list)
    sheets: list[str] = field(default_factory=list)
    sqlite_tables: list[str] = field(default_factory=list)


def _text(value: Any) -> str | None:
    if value is None:
        return None
    if isinstance(value, float) and value != value:
        return None
    result = str(value).strip()
    return result or None


def _decimal_text(value: Any) -> str | None:
    value = _text(value)
    if value is None:
        return None
    try:
        value = value.replace(",", "")
        parsed = Decimal(value)
    except (InvalidOperation, ValueError):
        raise ValueError(f"invalid decimal: {value[:40]}")
    return format(parsed, "f")


def _datetime_text(value: Any) -> str | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.isoformat(sep=" ", timespec="seconds")
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time()).isoformat(sep=" ", timespec="seconds")
    raw = _text(value)
    if not raw:
        return None
    # Keep source timezone/precision semantics but normalize common Excel text.
    raw = raw.replace("T", " ").replace("/", "-")
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            return datetime.strptime(raw[:19], fmt).isoformat(sep=" ", timespec="seconds")
        except ValueError:
            continue
    return raw


def _canonical_headers(headers: Iterable[Any]) -> tuple[list[str], list[str]]:
    canonical: list[str] = []
    drift: list[str] = []
    seen: dict[str, int] = {}
    for header in headers:
        source = _text(header) or "unnamed"
        target = HEADER_MAP.get(source)
        if target is None:
            safe = re.sub(r"[^0-9A-Za-z_]+", "_", source).strip("_").lower() or "source_col"
            target = f"source_{safe}"
            drift.append(source)
        count = seen.get(target, 0) + 1
        seen[target] = count
        canonical.append(target if count == 1 else f"{target}_{count}")
    return canonical, drift


def _fingerprint(row: dict[str, Any]) -> str:
    # Lineage controls identify the download, not the business row.  Excluding
    # them makes a re-archived identical workbook idempotent across runs.
    body = {k: row.get(k) for k in row if k not in {"source_row", "run_id", "updated_at", "source_file", "source_sheet", "file_hash"}}
    return hashlib.sha256(json.dumps(body, ensure_ascii=False, sort_keys=True, default=str).encode("utf-8")).hexdigest()


def _row_from_values(headers: list[str], values: tuple[Any, ...], *, report_type: str, shop_key: str,
                    bill_status: int, task_type: int, source_file: str, source_row: int,
                    source_sheet: str, run_id: str, file_hash: str) -> dict[str, Any]:
    row: dict[str, Any] = {
        "shop_key": shop_key,
        "report_type": report_type,
        "bill_status": bill_status,
        "task_type": task_type,
        "source_file": source_file,
        "source_sheet": source_sheet,
        "source_row": source_row,
        "run_id": run_id,
        "file_hash": file_hash,
    }
    for header, value in zip(headers, values):
        if header in DECIMAL_FIELDS:
            row[header] = _decimal_text(value)
        elif header in DATETIME_FIELDS:
            row[header] = _datetime_text(value)
        else:
            row[header] = _text(value)
    row["row_fingerprint"] = _fingerprint(row)
    return row


def _hash_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _safe_table_name(report_type: str, kind: str) -> str:
    state = "pending" if report_type == "pending" else "settled"
    return f"finance_{state}_{kind}"


def _ensure_sqlite_table(conn: sqlite3.Connection, table: str, columns: list[str]) -> None:
    fixed = [
        "id INTEGER PRIMARY KEY AUTOINCREMENT",
        "shop_key TEXT NOT NULL",
        "report_type TEXT NOT NULL",
        "bill_status INTEGER NOT NULL",
        "task_type INTEGER NOT NULL",
        "source_file TEXT NOT NULL",
        "source_sheet TEXT NOT NULL",
        "source_row INTEGER NOT NULL",
        "run_id TEXT NOT NULL",
        "file_hash TEXT NOT NULL",
        "row_fingerprint TEXT NOT NULL UNIQUE",
    ]
    for col in columns:
        if col not in {x.split()[0] for x in fixed}:
            fixed.append(f'"{col}" TEXT')
    conn.execute(f'CREATE TABLE IF NOT EXISTS "{table}" ({", ".join(fixed)})')
    # Existing P1 tables may predate the source_sheet lineage column.  Extend
    # only these finance task tables in place; never rewrite accepted facts.
    existing = {row[1] for row in conn.execute(f'PRAGMA table_info("{table}")').fetchall()}
    if "source_sheet" not in existing:
        conn.execute(f'ALTER TABLE "{table}" ADD COLUMN source_sheet TEXT')
        sheet = SHEET_TRANSACTION if table.endswith("_transaction") else SHEET_FEE
        conn.execute(f'UPDATE "{table}" SET source_sheet=? WHERE source_sheet IS NULL', (sheet,))
    conn.execute(f'CREATE INDEX IF NOT EXISTS "ix_{table}_shop_order" ON "{table}"(shop_key, order_id)')
    conn.execute(f'CREATE INDEX IF NOT EXISTS "ix_{table}_file" ON "{table}"(file_hash)')


def _ensure_quarantine_table(conn: sqlite3.Connection) -> None:
    conn.execute(
        """CREATE TABLE IF NOT EXISTS finance_quarantine (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            quarantine_key TEXT NOT NULL UNIQUE,
            shop_key TEXT NOT NULL,
            report_type TEXT NOT NULL,
            bill_status INTEGER NOT NULL,
            task_type INTEGER NOT NULL,
            source_file TEXT NOT NULL,
            source_sheet TEXT NOT NULL,
            source_row INTEGER NOT NULL,
            run_id TEXT NOT NULL,
            file_hash TEXT NOT NULL,
            reason TEXT NOT NULL,
            error_type TEXT,
            row_fingerprint TEXT,
            created_at TEXT NOT NULL
        )"""
    )
    conn.execute("CREATE INDEX IF NOT EXISTS ix_finance_quarantine_source ON finance_quarantine(file_hash, source_sheet, source_row)")


def ingest_workbook(*, workbook_path: str | Path, sqlite_path: str | Path, shop_key: str,
                    report_type: str, bill_status: int, task_type: int, run_id: str,
                    source_file: str | None = None, file_hash: str | None = None) -> FinanceIngestResult:
    """Clean one immutable workbook into two state-specific SQLite tables."""
    if report_type not in {"pending", "settled"}:
        raise ValueError("report_type must be pending or settled")
    path = Path(workbook_path)
    source_file = source_file or path.name
    file_hash = file_hash or _hash_file(path)
    import openpyxl

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    result = FinanceIngestResult(shop_key, report_type, bill_status, task_type, run_id, source_file, file_hash)
    result.sheets = list(wb.sheetnames)
    missing = [name for name in (SHEET_TRANSACTION, SHEET_FEE) if name not in wb.sheetnames]
    if missing:
        result.schema_drift.extend(missing)
        raise ValueError(f"missing finance sheets: {','.join(missing)}")
    db = Path(sqlite_path)
    db.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db)
    try:
        _ensure_quarantine_table(conn)
        for sheet_name, kind in ((SHEET_TRANSACTION, "transaction"), (SHEET_FEE, "fee_detail")):
            ws = wb[sheet_name]
            rows = ws.iter_rows(values_only=True)
            headers = next(rows, None)
            headers = next(rows, None) if headers else None
            if not headers:
                result.schema_drift.append(f"{sheet_name}:empty_header")
                continue
            canonical, drift = _canonical_headers(headers)
            result.schema_drift.extend(f"{sheet_name}:{x}" for x in drift)
            table = _safe_table_name(report_type, kind)
            _ensure_sqlite_table(conn, table, canonical)
            result.sqlite_tables.append(table)
            seen: set[str] = set()
            for source_row, values in enumerate(rows, start=3):
                if not any(value not in (None, "") for value in values):
                    continue
                try:
                    row = _row_from_values(canonical, values, report_type=report_type, shop_key=shop_key,
                                           bill_status=bill_status, task_type=task_type, source_file=source_file,
                                           source_sheet=sheet_name, source_row=source_row, run_id=run_id, file_hash=file_hash)
                except ValueError as exc:
                    result.parse_errors += 1
                    result.quarantined_rows += 1
                    qkey = hashlib.sha256(f"{file_hash}|{sheet_name}|{source_row}|parse_error".encode()).hexdigest()
                    conn.execute(
                        "INSERT OR IGNORE INTO finance_quarantine "
                        "(quarantine_key,shop_key,report_type,bill_status,task_type,source_file,source_sheet,source_row,run_id,file_hash,reason,error_type,row_fingerprint,created_at) "
                        "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,datetime('now'))",
                        (qkey, shop_key, report_type, bill_status, task_type, source_file, sheet_name, source_row, run_id, file_hash,
                         "parse_error", type(exc).__name__, None),
                    )
                    continue
                if not row.get("order_id") and not row.get("rf_busi_id"):
                    result.missing_required_keys += 1
                    result.quarantined_rows += 1
                    qkey = hashlib.sha256(f"{file_hash}|{sheet_name}|{source_row}|missing_key".encode()).hexdigest()
                    conn.execute(
                        "INSERT OR IGNORE INTO finance_quarantine "
                        "(quarantine_key,shop_key,report_type,bill_status,task_type,source_file,source_sheet,source_row,run_id,file_hash,reason,error_type,row_fingerprint,created_at) "
                        "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,datetime('now'))",
                        (qkey, shop_key, report_type, bill_status, task_type, source_file, sheet_name, source_row, run_id, file_hash,
                         "missing_key", "MissingFinanceKey", row["row_fingerprint"]),
                    )
                    continue
                if row["row_fingerprint"] in seen:
                    result.duplicate_rows += 1
                    result.quarantined_rows += 1
                    qkey = hashlib.sha256(f"{file_hash}|{sheet_name}|{source_row}|duplicate".encode()).hexdigest()
                    conn.execute(
                        "INSERT OR IGNORE INTO finance_quarantine "
                        "(quarantine_key,shop_key,report_type,bill_status,task_type,source_file,source_sheet,source_row,run_id,file_hash,reason,error_type,row_fingerprint,created_at) "
                        "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,datetime('now'))",
                        (qkey, shop_key, report_type, bill_status, task_type, source_file, sheet_name, source_row, run_id, file_hash,
                         "duplicate", "DuplicateRow", row["row_fingerprint"]),
                    )
                    continue
                seen.add(row["row_fingerprint"])
                cols = list(row)
                placeholders = ",".join("?" for _ in cols)
                quoted = ",".join(f'"{c}"' for c in cols)
                cursor = conn.execute(f'INSERT OR IGNORE INTO "{table}" ({quoted}) VALUES ({placeholders})', tuple(row[c] for c in cols))
                if cursor.rowcount == 0:
                    result.duplicate_rows += 1
                    # A replay of the same immutable workbook is expected to
                    # hit the fingerprint UNIQUE key.  It is idempotent, not
                    # a new source-row quarantine.  Only a duplicate arriving
                    # from a different file hash is persisted as quarantine.
                    same_source = conn.execute(
                        f'SELECT 1 FROM "{table}" WHERE row_fingerprint=? AND file_hash=? LIMIT 1',
                        (row["row_fingerprint"], file_hash),
                    ).fetchone()
                    if same_source is None:
                        result.quarantined_rows += 1
                        qkey = hashlib.sha256(f"{file_hash}|{sheet_name}|{source_row}|duplicate".encode()).hexdigest()
                        conn.execute(
                            "INSERT OR IGNORE INTO finance_quarantine "
                            "(quarantine_key,shop_key,report_type,bill_status,task_type,source_file,source_sheet,source_row,run_id,file_hash,reason,error_type,row_fingerprint,created_at) "
                            "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,datetime('now'))",
                            (qkey, shop_key, report_type, bill_status, task_type, source_file, sheet_name, source_row, run_id, file_hash,
                             "duplicate", "DuplicateRow", row["row_fingerprint"]),
                        )
                    continue
                if kind == "transaction":
                    result.transaction_rows += 1
                else:
                    result.fee_rows += 1
        conn.commit()
    finally:
        conn.close()
        wb.close()
    return result


def sync_mysql_facts(*, sqlite_path: str | Path, mysql_config_path: str | Path,
                     report_type: str, dry_run: bool = False) -> dict[str, Any]:
    """Copy finance SQLite rows to new MySQL fact tables, idempotently.

    This function only creates/updates ``fact_finance_*`` tables.  It never
    touches existing order, after-sale, metric or derived tables.
    """
    state = "pending" if report_type == "pending" else "settled"
    sqlite_conn = sqlite3.connect(sqlite_path)
    tables = [f"finance_{state}_transaction", f"finance_{state}_fee_detail"]
    rows_by_table = {table: sqlite_conn.execute(f'SELECT * FROM "{table}"').fetchall() for table in tables}
    columns_by_table = {
        table: [r[1] for r in sqlite_conn.execute(f'PRAGMA table_info("{table}")').fetchall()]
        for table in tables
    }
    sqlite_conn.close()
    if dry_run:
        return {"report_type": report_type, "tables": {t: len(rows_by_table[t]) for t in tables}, "dry_run": True}
    import configparser
    import mysql.connector

    parser = configparser.ConfigParser()
    if not parser.read(mysql_config_path, encoding="utf-8") or "mysql" not in parser:
        raise RuntimeError("MySQL config missing")
    cfg = parser["mysql"]
    conn = mysql.connector.connect(host=cfg.get("host"), port=cfg.getint("port"), database=cfg.get("database"),
                                   user=cfg.get("user"), password=cfg.get("password"), charset="utf8mb4")
    counts: dict[str, int] = {}
    try:
        cur = conn.cursor()
        for source_table in tables:
            # Pending and settled exports intentionally remain separate facts:
            # their column contracts differ (expected vs actual settlement,
            # statement/bank-flow metadata) and one state must never overwrite
            # the other.
            target = f"fact_finance_{state}_transaction" if source_table.endswith("transaction") else f"fact_finance_{state}_fee_detail"
            columns = columns_by_table[source_table]
            defs = []
            for col in columns:
                if col == "id":
                    continue
                if col in {"bill_status", "task_type", "source_row"}:
                    typ = "INT"
                elif col in {"sku_price", "sku_num", "settle_amount", "income_sum", "outcome_sum", "user_paid_amount",
                             "platform_fee_base", "platform_fee_rate", "fee_amount", "merchant_bear_amount",
                             "platform_bear_amount", "government_bear_amount", "separate_settlement_asset"}:
                    typ = "DECIMAL(24,8)"
                elif col == "row_fingerprint":
                    typ = "CHAR(64) NOT NULL"
                else:
                    typ = "LONGTEXT"
                defs.append(f'`{col}` {typ}')
            defs.extend(["`id` BIGINT AUTO_INCREMENT PRIMARY KEY", "UNIQUE KEY `uq_finance_row_fingerprint` (`row_fingerprint`)",
                         "KEY `ix_finance_shop_order` (`shop_key`(32), `order_id`(64))",
                         "KEY `ix_finance_state` (`report_type`(16), `bill_status`, `task_type`)"])
            cur.execute(f"CREATE TABLE IF NOT EXISTS `{target}` ({', '.join(defs)}) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4")
            cur.execute(f"SHOW COLUMNS FROM `{target}`")
            existing_mysql = {row[0] for row in cur.fetchall()}
            for col in columns:
                if col not in existing_mysql:
                    cur.execute(f"ALTER TABLE `{target}` ADD COLUMN `{col}` LONGTEXT NULL")
            sheet = SHEET_TRANSACTION if source_table.endswith("_transaction") else SHEET_FEE
            if "source_sheet" in columns:
                cur.execute(f"UPDATE `{target}` SET `source_sheet`=%s WHERE `source_sheet` IS NULL OR `source_sheet`=''", (sheet,))
            insert_cols = [c for c in columns if c != "id"]
            quoted_cols = ",".join(f'`{c}`' for c in insert_cols)
            placeholders = ",".join("%s" for _ in insert_cols)
            sql = f"INSERT IGNORE INTO `{target}` ({quoted_cols}) VALUES ({placeholders})"
            values = [tuple(row[columns.index(c)] for c in insert_cols) for row in rows_by_table[source_table]]
            if values:
                cur.executemany(sql, values)
            counts[target] = len(values)
        conn.commit()
    finally:
        conn.close()
    return {"report_type": report_type, "tables": counts, "dry_run": False}
