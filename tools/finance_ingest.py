"""Archive and load verified finance export workbooks.

This is a separate P1 entry point.  It does not call the order/after-sale
collectors and does not alter existing business tables.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from finance.etl import ingest_workbook, sync_mysql_facts


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _archive(source: Path, *, shop: str, report_type: str, bill_status: int, task_type: int,
             requested_start: str, requested_end: str, run_id: str) -> tuple[Path, dict]:
    file_hash = _sha256(source)
    target_dir = ROOT / "data" / "raw" / "finance" / shop / report_type
    target_dir.mkdir(parents=True, exist_ok=True)
    target = target_dir / f"{run_id}_{source.name}"
    if target.exists() and _sha256(target) != file_hash:
        raise RuntimeError(f"archive target hash conflict: {target.name}")
    if not target.exists():
        shutil.copy2(source, target)
    metadata = {
        "shop_key": shop, "report_type": report_type, "bill_status": bill_status,
        "task_type": task_type, "requested_start": requested_start, "requested_end": requested_end,
        "downloaded_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "source_filename": source.name, "file_hash": file_hash, "run_id": run_id,
    }
    (target_dir / f"{run_id}_{source.stem}.metadata.json").write_text(
        json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return target, metadata


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--pending", required=True)
    parser.add_argument("--settled", required=True)
    parser.add_argument("--shop", default="MIYO")
    parser.add_argument("--start", required=True)
    parser.add_argument("--end", required=True)
    parser.add_argument("--run-id", default=None)
    parser.add_argument("--mysql-sync", action="store_true")
    args = parser.parse_args()
    run_id = args.run_id or f"finance-p1-{datetime.now():%Y%m%d%H%M%S}"
    outputs = {"run_id": run_id, "shop": args.shop, "raw_archive": [], "ingest": [], "mysql": []}
    for report_type, path_text, bill_status, task_type in (
        ("pending", args.pending, 1, 1), ("settled", args.settled, 2, 2)
    ):
        source = Path(path_text)
        archived, metadata = _archive(source, shop=args.shop, report_type=report_type, bill_status=bill_status,
                                      task_type=task_type, requested_start=args.start, requested_end=args.end,
                                      run_id=run_id)
        outputs["raw_archive"].append(metadata)
        result = ingest_workbook(workbook_path=archived, sqlite_path=ROOT / "data" / "jd_report.db",
                                 shop_key=args.shop, report_type=report_type, bill_status=bill_status,
                                 task_type=task_type, run_id=run_id, source_file=archived.name)
        outputs["ingest"].append(result.__dict__)
        if args.mysql_sync:
            outputs["mysql"].append(sync_mysql_facts(sqlite_path=ROOT / "data" / "jd_report.db",
                                                      mysql_config_path=ROOT / "config" / "mysql.local.ini",
                                                      report_type=report_type))
    out_dir = ROOT / "runtime" / "finance_etl_p1" / run_id
    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / "result.json").write_text(json.dumps(outputs, ensure_ascii=False, indent=2), encoding="utf-8")
    # No IDs, amounts, credentials or signed URLs are printed.
    print(json.dumps({
        "run_id": run_id,
        "raw_archive": [{"report_type": x["report_type"], "file_hash": x["file_hash"]} for x in outputs["raw_archive"]],
        "ingest": [{"report_type": x["report_type"], "transaction_rows": x["transaction_rows"],
                     "fee_rows": x["fee_rows"], "duplicate_rows": x["duplicate_rows"],
                     "quarantined_rows": x["quarantined_rows"], "schema_drift": x["schema_drift"]}
                    for x in outputs["ingest"]],
        "mysql": outputs["mysql"],
    }, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
