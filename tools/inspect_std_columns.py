import configparser
from pathlib import Path
import mysql.connector

root = Path(__file__).resolve().parents[1]
cfg = configparser.ConfigParser()
cfg.read(root / 'config/mysql.local.ini', encoding='utf-8')
c = cfg['mysql']
db = mysql.connector.connect(host=c.get('host'), port=c.getint('port'), user=c.get('user'), password=c.get('password'), database=c.get('database'), charset=c.get('charset','utf8mb4'))
cur = db.cursor()
for table in ('std_biz_jm_order_full','std_biz_jm_after_sale_full','std_biz_jzt_kuaiche_order_effect','std_biz_jzt_quanzhan_effect','std_biz_jzt_kuaiche','std_biz_jzt_quanzhan_campaign','std_biz_jzt_quanzhan_campaign_all_store'):
    cur.execute("SELECT column_name FROM information_schema.columns WHERE table_schema=%s AND table_name=%s ORDER BY ordinal_position", (c.get('database'), table))
    cols = [x[0] for x in cur.fetchall()]
    print(table + ': ' + ' | '.join(repr(x) for x in cols if any(k in x for k in ('金额','订单','出库','花费','应收'))))
cur.close()
db.close()
