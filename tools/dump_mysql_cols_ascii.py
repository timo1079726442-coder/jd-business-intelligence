import configparser
import mysql.connector
from pathlib import Path
root=Path(__file__).resolve().parents[1]; c=configparser.ConfigParser(); c.read(root/"config/mysql.local.ini",encoding="utf-8"); d=c["mysql"]
db=mysql.connector.connect(host=d.get("host"),port=d.getint("port"),user=d.get("user"),password=d.get("password"),database=d.get("database")); cur=db.cursor()
for table in ("std_biz_jm_order_full","std_biz_jzt_kuaiche","std_biz_jzt_quanzhan_campaign","std_biz_jzt_quanzhan_campaign_all_store"):
    cur.execute("SELECT column_name FROM information_schema.columns WHERE table_schema=%s AND table_name=%s ORDER BY ordinal_position",(d.get("database"),table))
    print(table)
    for i,(name,) in enumerate(cur.fetchall()): print(i,ascii(name))
cur.close(); db.close()
