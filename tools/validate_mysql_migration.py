import configparser, sqlite3
from pathlib import Path
import mysql.connector

root=Path(__file__).resolve().parents[1]
p=configparser.ConfigParser()
p.read(root/'config/mysql.local.ini',encoding='utf-8')
c=p['mysql']
my=mysql.connector.connect(host=c.get('host'),port=c.getint('port'),user=c.get('user'),password=c.get('password'),database=c.get('database'),charset=c.get('charset','utf8mb4'))
sq=sqlite3.connect(f"file:{(root/'data/jd_report.db').as_posix()}?mode=ro",uri=True)
mc=my.cursor(); sc=sq.cursor(); mismatches=[]
tables=[r[0] for r in sc.execute("select name from sqlite_master where type='table' and name not like 'sqlite_%' order by name")]
for t in tables:
    s=sc.execute(f'SELECT COUNT(*) FROM "{t}"').fetchone()[0]
    mc.execute(f'SELECT COUNT(*) FROM `raw_{t}`')
    m=mc.fetchone()[0]
    print(t,s,m,'OK' if s==m else 'MISMATCH')
    if s!=m: mismatches.append(t)
print('RESULT', 'OK' if not mismatches else 'MISMATCH '+','.join(mismatches))
mc.close(); my.close(); sc.close(); sq.close()
