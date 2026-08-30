import configparser
from pathlib import Path
import mysql.connector
root=Path(__file__).resolve().parents[1]
p=configparser.ConfigParser(); p.read(root/'config/mysql.local.ini',encoding='utf-8'); c=p['mysql']
db=mysql.connector.connect(host=c.get('host'),port=c.getint('port'),user=c.get('user'),password=c.get('password'),database=c.get('database'),charset=c.get('charset','utf8mb4'))
cur=db.cursor()
cur.execute('''CREATE TABLE IF NOT EXISTS legacy_excel_summary (
 source_shop VARCHAR(64) NOT NULL, source_sheet VARCHAR(128) NOT NULL,
 inferred_date DATE NOT NULL, row_count BIGINT NOT NULL,
 PRIMARY KEY(source_shop,source_sheet,inferred_date)
) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci''')
cur.execute('TRUNCATE TABLE legacy_excel_summary')
cur.execute('''INSERT INTO legacy_excel_summary(source_shop,source_sheet,inferred_date,row_count)
 SELECT source_shop,source_sheet,inferred_date,COUNT(*) FROM legacy_excel_rows
 WHERE inferred_date IS NOT NULL GROUP BY source_shop,source_sheet,inferred_date''')
db.commit()
cur.execute('SELECT source_shop,source_sheet,MIN(inferred_date),MAX(inferred_date),SUM(row_count) FROM legacy_excel_summary GROUP BY source_shop,source_sheet ORDER BY source_shop,source_sheet')
for row in cur.fetchall(): print(row)
cur.execute('SELECT COUNT(*),COUNT(DISTINCT row_hash) FROM legacy_excel_rows')
print('TOTAL',cur.fetchone())
cur.close(); db.close()
