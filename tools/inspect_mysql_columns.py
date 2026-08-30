import configparser
from pathlib import Path
import mysql.connector

root = Path(__file__).resolve().parents[1]
cfg = configparser.ConfigParser()
cfg.read(root / 'config/mysql.local.ini', encoding='utf-8')
c = cfg['mysql']
db = mysql.connector.connect(host=c.get('host'), port=c.getint('port'), user=c.get('user'), password=c.get('password'), database=c.get('database'), charset=c.get('charset','utf8mb4'))
cur = db.cursor()
cur.execute("SELECT table_name,column_name FROM information_schema.columns WHERE table_schema=%s AND table_name LIKE 'raw\\_%' ORDER BY table_name,ordinal_position", (c.get('database'),))
last = ''
for table, column in cur.fetchall():
    if table != last:
        print('\n' + table + ':', end=' ')
        last = table
    print(column, end=' | ')
print()
cur.close()
db.close()
