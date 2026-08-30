import configparser, json, traceback
from datetime import date, timedelta, datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse
import mysql.connector
from metrics.service import MetricService, METRIC_VERSION, SERVICE_VERSION

ROOT = Path(__file__).resolve().parent
HTML = ROOT / "dashboard" / "index_v1.html"
cfg = configparser.ConfigParser(); cfg.read(ROOT / "config" / "mysql.local.ini", encoding="utf-8")
DB = cfg["mysql"]; SHOPS = {"FYA", "MIYO", "OTA"}; PINS = {"MIYO": "miyo-周", "OTA": "ota8888", "FYA": "FYA"}
STARTED_AT = datetime.now().astimezone().isoformat()

def qid(name): return "`" + name.replace("`", "``") + "`"
def mysql_columns(cur, table):
    conn = mysql.connector.connect(host=DB.get("host"), port=DB.getint("port"), user=DB.get("user"), password=DB.get("password"), database=DB.get("database"))
    meta = conn.cursor()
    meta.execute("SELECT COLUMN_NAME FROM information_schema.columns WHERE table_schema=%s AND table_name=%s ORDER BY ordinal_position", (DB.get("database"), table))
    rows = meta.fetchall(); meta.close(); conn.close()
    return [r[0] for r in rows]

def query_summary(shop, start=None, end=None):
    if shop not in SHOPS: raise ValueError("invalid shop")
    end = end or date.today().isoformat(); start = start or (date.fromisoformat(end) - timedelta(days=6)).isoformat()
    begin, finish = date.fromisoformat(start), date.fromisoformat(end); days = (finish - begin).days + 1
    compare_start = (begin - timedelta(days=days)).isoformat(); compare_end = (begin - timedelta(days=1)).isoformat()
    db = mysql.connector.connect(host=DB.get("host"), port=DB.getint("port"), user=DB.get("user"), password=DB.get("password"), database=DB.get("database"), charset=DB.get("charset", "utf8mb4")); cur = db.cursor(dictionary=True)
    order_cols = mysql_columns(cur, "std_biz_jm_order_full"); ad_cols = mysql_columns(cur, "std_biz_jzt_kuaiche"); campaign_cols = mysql_columns(cur, "std_biz_jzt_quanzhan_campaign")
    def pick(needle, fallback_index):
        return next((x for x in order_cols if needle in x), order_cols[fallback_index] if len(order_cols) > fallback_index else None)
    gross_col = pick("订单号拆分金额", 83)
    refund_col = pick("订单号汇总退款金额", 82)
    flag_col = pick("订单数", 84)
    status_col = pick("出库状态", 85)
    if not all((gross_col, refund_col, flag_col, status_col)):
        raise RuntimeError("订单表字段映射不完整")
    spend_col = next((x for x in ad_cols if "花费" in x), None); campaign_spend_col = next((x for x in campaign_cols if "投放成本" in x), None); pin = PINS[shop]
    def metrics(a,b):
        g,r,f,s = map(qid,(gross_col,refund_col,flag_col,status_col))
        sql = f"SELECT COALESCE(SUM(CAST(REPLACE(NULLIF({g},''),',','') AS DECIMAL(18,2))),0) gross, COALESCE(SUM(CAST(REPLACE(NULLIF({r},''),',','') AS DECIMAL(18,2))),0) refund, COALESCE(SUM(CASE WHEN {f}='1' AND {s}='已出库' AND CAST(REPLACE(NULLIF({g},''),',','') AS DECIMAL(18,2))>10 THEN 1 ELSE 0 END),0) orders, COALESCE(SUM(CASE WHEN {f}='1' AND {s}='已出库' AND CAST(REPLACE(NULLIF({r},''),',','') AS DECIMAL(18,2))>10 THEN 1 ELSE 0 END),0) return_orders FROM std_biz_jm_order_full WHERE (shop_pin=%s OR shop_pin=%s) AND stat_date BETWEEN %s AND %s"
        cur.execute(sql,(shop,pin,a,b)); row=cur.fetchone(); spend=0.0
        if spend_col:
            cur.execute(f"SELECT COALESCE(SUM(CAST(REPLACE(NULLIF({qid(spend_col)},''),',','') AS DECIMAL(18,2))),0) spend FROM std_biz_jzt_kuaiche WHERE (shop_pin=%s OR shop_pin=%s) AND stat_date BETWEEN %s AND %s",(shop,pin,a,b)); spend += float(cur.fetchone()["spend"] or 0)
        if campaign_spend_col:
            for table in ("std_biz_jzt_quanzhan_campaign","std_biz_jzt_quanzhan_campaign_all_store"):
                cur.execute(f"SELECT COALESCE(SUM(CAST(REPLACE(NULLIF({qid(campaign_spend_col)},''),',','') AS DECIMAL(18,2))),0) spend FROM {table} WHERE (shop_pin=%s OR shop_pin=%s) AND stat_date BETWEEN %s AND %s",(shop,pin,a,b)); spend += float(cur.fetchone()["spend"] or 0)
        gross,refund=float(row["gross"] or 0),float(row["refund"] or 0); orders,returns=int(row["orders"] or 0),int(row["return_orders"] or 0); net=gross-refund
        return {"gross":gross,"refund":refund,"net":net,"refund_rate":refund/gross if gross else None,"spend":spend,"orders":orders,"return_orders":returns,"return_rate":returns/orders if orders else None,"roi":net/spend if spend else None}
    current,compare=metrics(start,end),metrics(compare_start,compare_end)
    for key,value in list(current.items()): current[key+"_change"] = (value/compare[key]-1) if isinstance(value,(int,float)) and compare.get(key) not in (None,0) else None
    cur.execute("SELECT stat_date, SUM(row_count) row_count FROM vw_dashboard_shop_daily WHERE shop_code=%s AND stat_date BETWEEN %s AND %s GROUP BY stat_date ORDER BY stat_date DESC",(shop,start,end)); daily=cur.fetchall(); cur.close(); db.close()
    return {"shop":shop,"start":start,"end":end,"compare_start":compare_start,"compare_end":compare_end,"current":current,"compare":compare,"daily":daily}

_metric_service = MetricService()

def query_summary(shop, start=None, end=None):
    """Stable API facade; metric logic lives in metrics.service."""
    return _metric_service.get_shop_overview(shop, start, end)

class Handler(BaseHTTPRequestHandler):
    def do_GET(self):
        parsed=urlparse(self.path)
        if parsed.path=="/": body,ctype=HTML.read_bytes(),"text/html; charset=utf-8"; self.send_response(200)
        elif parsed.path=="/api/health":
            body=json.dumps({"service":"jd-dashboard","pid":__import__("os").getpid(),"port":18766,"started_at":STARTED_AT,"service_version":SERVICE_VERSION,"metric_version":METRIC_VERSION},ensure_ascii=False).encode(); ctype="application/json; charset=utf-8"; self.send_response(200)
        elif parsed.path=="/api/summary":
            q=parse_qs(parsed.query)
            try: body=json.dumps(query_summary(q.get("shop",["MIYO"])[0],q.get("start",[None])[0],q.get("end",[None])[0]),ensure_ascii=False).encode(); status=200
            except ValueError: body=b'{"error":"invalid shop"}'; status=400
            except Exception as exc: body=json.dumps({"error":{"code":"INTERNAL_ERROR","message":"看板数据读取失败"}},ensure_ascii=False).encode(); status=500
            self.send_response(status); ctype="application/json; charset=utf-8"
        else: body,ctype=b"Not Found","text/plain; charset=utf-8"; self.send_response(404)
        self.send_header("Content-Type",ctype); self.send_header("Content-Length",str(len(body))); self.end_headers(); self.wfile.write(body)
    def log_message(self,*_): return

if __name__ == "__main__":
    print("请使用唯一启动入口：python dashboard_start.py", flush=True)
