from __future__ import annotations

import sqlite3
import unittest
from pathlib import Path

from finance.etl import _decimal_text, _datetime_text, ingest_workbook


class FinanceETLTests(unittest.TestCase):
    def test_finance_value_normalization(self) -> None:
        self.assertEqual(_decimal_text("1,234.50"), "1234.50")
        self.assertIsNone(_decimal_text(None))
        self.assertEqual(_datetime_text("2026/09/01 12:30:00"), "2026-09-01 12:30:00")


    def test_finance_workbook_is_state_separated_and_idempotent(self) -> None:
        source = Path("runtime/finance_auth_smoke_test/20260905_180525_finance_auth/pending.xlsx")
        if not source.exists():
            self.skipTest("dynamic finance smoke output not present")
        import tempfile
        with tempfile.TemporaryDirectory() as td:
            db = Path(td) / "finance.db"
            first = ingest_workbook(workbook_path=source, sqlite_path=db, shop_key="MIYO", report_type="pending",
                                    bill_status=1, task_type=1, run_id="test-run", source_file=source.name)
            second = ingest_workbook(workbook_path=source, sqlite_path=db, shop_key="MIYO", report_type="pending",
                                     bill_status=1, task_type=1, run_id="test-run-2", source_file=source.name)
            self.assertGreater(first.transaction_rows, 0)
            self.assertGreater(first.fee_rows, 0)
            self.assertEqual(second.transaction_rows, 0)
            self.assertEqual(second.fee_rows, 0)
            conn = sqlite3.connect(db)
            self.assertEqual(conn.execute("select count(*) from finance_pending_transaction").fetchone()[0], first.transaction_rows)
            self.assertEqual(conn.execute("select count(*) from finance_pending_fee_detail").fetchone()[0], first.fee_rows)
            self.assertEqual(conn.execute("select count(*) from finance_quarantine").fetchone()[0], 0)
            self.assertIn("source_sheet", [row[1] for row in conn.execute("pragma table_info(finance_pending_transaction)")])
            conn.close()

    def test_settled_fee_duplicate_rows_are_traceable(self) -> None:
        source = Path("runtime/finance_auth_smoke_test/20260905_180525_finance_auth/settled.xlsx")
        if not source.exists():
            self.skipTest("dynamic finance smoke output not present")
        import tempfile
        with tempfile.TemporaryDirectory() as td:
            db = Path(td) / "finance.db"
            result = ingest_workbook(workbook_path=source, sqlite_path=db, shop_key="MIYO", report_type="settled",
                                     bill_status=2, task_type=2, run_id="test-settled", source_file=source.name)
            self.assertEqual(result.transaction_rows, 3583)
            self.assertEqual(result.fee_rows, 6829)
            self.assertEqual(result.quarantined_rows, 23)
            conn = sqlite3.connect(db)
            self.assertEqual(conn.execute("select reason, count(*) from finance_quarantine group by reason").fetchall(), [("duplicate", 23)])
            self.assertEqual(conn.execute("select count(*) from finance_settled_fee_detail where source_sheet is not null and source_row is not null and run_id is not null").fetchone()[0], 6829)
            conn.close()
