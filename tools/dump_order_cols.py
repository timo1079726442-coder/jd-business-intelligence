import sqlite3
from pathlib import Path
db=sqlite3.connect(Path(__file__).resolve().parents[1]/"data/jd_report.db")
for i, row in enumerate(db.execute("PRAGMA table_info(biz_jm_order_full)")):
    print(i, repr(row[1]))
db.close()
