"""Finance export ingestion boundary.

This package is intentionally isolated from the existing order, after-sale,
traffic, advertising and metric pipelines.  It only turns verified finance
export workbooks into traceable staging/fact rows.
"""

from .etl import FinanceIngestResult, ingest_workbook, sync_mysql_facts

__all__ = ["FinanceIngestResult", "ingest_workbook", "sync_mysql_facts"]
