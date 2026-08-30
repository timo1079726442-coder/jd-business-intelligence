import configparser
from pathlib import Path
import mysql.connector

root=Path(__file__).resolve().parents[1]
p=configparser.ConfigParser(); p.read(root/'config/mysql.local.ini',encoding='utf-8'); c=p['mysql']
db=mysql.connector.connect(host=c.get('host'),port=c.getint('port'),user=c.get('user'),password=c.get('password'),database=c['database'])
cur=db.cursor()
cur.execute('SELECT business_type,shop_code,COUNT(*),MIN(stat_date),MAX(stat_date) FROM unified_legacy_rows GROUP BY business_type,shop_code ORDER BY business_type,shop_code')
for row in cur.fetchall(): print(row)
cur.execute('SELECT COUNT(*),COUNT(DISTINCT row_hash) FROM unified_legacy_rows')
print('TOTAL',cur.fetchone())
cur.close();db.close()
