import sqlite3
from pathlib import Path

root = Path(__file__).resolve().parents[1]
db = sqlite3.connect(root / 'data/jd_report.db')
cols = [r[1] for r in db.execute('PRAGMA table_info(biz_jm_order_full)')]
for i in (77, 78, 79, 80, 81, 82):
    print(i + 1, cols[i] if i < len(cols) else None)
db.close()
