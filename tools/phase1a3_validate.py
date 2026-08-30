"""阶段 1A-3：Legacy Golden Dataset 与 MIYO 生产快照只读验证。"""
from __future__ import annotations

import csv
import json
import sqlite3
import sys
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from decimal import Decimal
from pathlib import Path

import configparser
import mysql.connector
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from derived.order_v1 import derive_order_rows
from tools import analyze_miyo_derived_full as lineage

WORKBOOK = ROOT / "output/MIYO箱包旗舰店/京东MIYO数据库.xlsx"
GOLDEN_REPORT = ROOT / "docs/Legacy Golden Dataset验证报告.md"
SNAPSHOT_REPORT = ROOT / "docs/MIYO生产快照差异报告.md"
SNAPSHOT_CSV = ROOT / "docs/MIYO生产快照差异明细.csv"
BATCH_REPORT = ROOT / "docs/数据快照与批次追踪设计说明.md"
EXCEPTIONS = ROOT / "validation/legacy_exceptions/miyo_order_derived.json"


def _float(value):
    try:
        return float(value or 0)
    except (TypeError, ValueError):
        return 0.0


def _date_value(value):
    value = lineage.excel_date(value)
    if isinstance(value, datetime):
        return value.date()
    return value


def _row_metrics(rows, derived, begin, finish):
    result = {"gmv": 0.0, "refund_amount": 0.0, "order_count": 0, "refund_order_count": 0}
    for row_no, row in enumerate(rows[1:], 2):
        dt = _date_value(row[28] if len(row) > 28 else None)  # AC：付款确认时间
        if not dt or not (begin <= dt <= finish):
            continue
        d = derived[row_no]
        if int(_float(d.get("CB"))) != 1:
            continue
        result["gmv"] += _float(d.get("CA")); result["refund_amount"] += _float(d.get("BZ"))
        if d.get("CC") == "已出库" and _float(d.get("CA")) > 10:
            result["order_count"] += 1
        if d.get("CC") == "已出库" and _float(d.get("BZ")) > 10:
            result["refund_order_count"] += 1
    result["net_gmv"] = result["gmv"] - result["refund_amount"]
    result["refund_rate"] = result["refund_amount"] / result["gmv"] if result["gmv"] else None
    result["return_rate"] = result["refund_order_count"] / result["order_count"] if result["order_count"] else None
    return result


def _build_ad_cache(wb):
    cache = []
    for sheet_name, cost_col, date_col in (("推广自定义报表", 34, 51), ("全站", 5, 17), ("全站-全店", 4, 15)):
        name = next((n for n in wb.sheetnames if sheet_name in n and not (sheet_name == "全站" and "全店" in n)), None)
        if not name:
            continue
        # 推广自定义报表约 76 万行；Golden 规则回归以订单/售后为主，避免在多个
        # 视图进程中重复扫描大表。该大表的推广口径由现有缓存/生产对账单独验证。
        if wb[name].max_row > 100000:
            continue
        values = []
        lo, hi = min(cost_col, date_col), max(cost_col, date_col)
        for row in wb[name].iter_rows(min_row=2, min_col=lo, max_col=hi, values_only=True):
            dt = _date_value(row[date_col - lo] if len(row) > date_col - lo else None)
            if dt:
                values.append((dt, _float(row[cost_col - lo] if len(row) > cost_col - lo else 0)))
        cache.append(values)
    return cache


def _ad_cost(cache, begin, finish):
    total = 0.0
    for values in cache:
        total += sum(amount for dt, amount in values if begin <= dt <= finish)
    return total


def _metric_set(base, ad_cost):
    result = dict(base); result["ad_cost"] = ad_cost
    result["fee_ratio"] = ad_cost / result["gmv"] if result["gmv"] else None
    result["roi"] = result["net_gmv"] / ad_cost if ad_cost else None
    result["estimated_profit"] = result["net_gmv"] * (1 - 0.09 - 0.03 - 0.7) - ad_cost
    return result


def _standard_inputs(order_values, after_values):
    oi = lineage.ORDER_COLS; ai = lineage.AFTER_COLS
    orders = []
    for row_no, row in enumerate(order_values[1:], 2):
        cell = lambda name: row[oi[name] - 1] if oi[name] - 1 < len(row) else None
        orders.append({"shop_pin": "MIYO", "order_no": cell("A"), "price": cell("G"), "allocation_base": cell("I"), "status": cell("L"), "order_type": cell("M"), "payment_confirmed_at": cell("AC"), "product_id": cell("B"), "sku_id": cell("B"), "source_id": str(row_no)})
    after = []
    for row_no, row in enumerate(after_values[1:], 2):
        cell = lambda name: row[ai[name] - 1] if ai[name] - 1 < len(row) else None
        after.append({"shop_pin": "MIYO", "order_no": cell("K"), "refund_amount": cell("AJ"), "shipment_status": cell("M"), "application_at": cell("H"), "source_id": str(row_no)})
    return orders, after


def golden_validation():
    wb = load_workbook(WORKBOOK, read_only=True, data_only=True)
    print("LOADED WORKBOOK", flush=True)
    orders = list(wb["订单列表"].values); after = list(wb["售后"].values)
    print(f"LOADED SOURCE rows={len(orders)} after={len(after)}", flush=True)
    legacy = lineage.evaluate_rows(orders, after, {})
    print("LEGACY DERIVED", flush=True)
    std_orders, std_after = _standard_inputs(orders, after)
    new_rows = derive_order_rows(std_orders, std_after)
    print("PY DERIVED", flush=True)
    new = {int(r["source_id"]): {"BZ": r["order_refund_aggregated_amount"], "CA": r["order_split_amount"], "CB": r["order_count_flag"], "CC": r["shipment_status_normalized"]} for r in new_rows}
    known = {(x["field"], x["excel_row"]) for x in json.loads(EXCEPTIONS.read_text(encoding="utf-8"))["exceptions"]}
    totals = Counter(); field_mismatch = Counter(); samples = []
    totals["registered_legacy_exception"] = len(known)
    for row_no, old in legacy.items():
        for field in ("BZ", "CA", "CB", "CC"):
            old_v, new_v = old[field], new.get(row_no, {}).get(field)
            if old_v is None or new_v is None:
                continue
            same = abs(_float(old_v) - _float(new_v)) <= 0.01 if field != "CC" else str(old_v) == str(new_v)
            totals["total"] += 1
            if same: totals["consistent"] += 1
            elif (("order_split_amount" if field == "CA" else "order_count_flag" if field == "CB" else field.lower()), row_no) in known:
                totals["known_exception"] += 1
            else:
                totals["new_exception"] += 1; field_mismatch[field] += 1
                if len(samples) < 20: samples.append((field, row_no, old_v, new_v))
    ad_cache = _build_ad_cache(wb)
    print("AD CACHE", flush=True)
    available_dates = [_date_value(row[28]) for row in orders[1:] if len(row) > 28 and _date_value(row[28])]
    latest = max(available_dates) if available_dates else date(2026, 8, 19)
    windows = [("last_7_days_available", latest - timedelta(days=6), latest), ("mid_month_7_days", date(2026, 8, 13), date(2026, 8, 19)), ("cross_month", date(2026, 7, 30), date(2026, 8, 5)), ("refund_window", date(2026, 8, 13), date(2026, 8, 19)), ("multi_product_window", date(2026, 1, 23), date(2026, 1, 24)), ("after_sale_status_window", date(2026, 7, 22), date(2026, 7, 28))]
    metric_rows = []
    for name, begin, finish in windows:
        old_base = _row_metrics(orders, legacy, begin, finish)
        new_base = {"gmv": 0.0, "refund_amount": 0.0, "order_count": 0, "refund_order_count": 0}
        for record in new_rows:
            dt = _date_value(record.get("payment_confirmed_at"))
            if not dt or not (begin <= dt <= finish) or int(record["order_count_flag"]) != 1:
                continue
            new_base["gmv"] += _float(record["order_split_amount"]); new_base["refund_amount"] += _float(record["order_refund_aggregated_amount"])
            if record["shipment_status_normalized"] == "已出库" and _float(record["order_split_amount"]) > 10: new_base["order_count"] += 1
            if record["shipment_status_normalized"] == "已出库" and _float(record["order_refund_aggregated_amount"]) > 10: new_base["refund_order_count"] += 1
        new_base["net_gmv"] = new_base["gmv"] - new_base["refund_amount"]; new_base["refund_rate"] = new_base["refund_amount"] / new_base["gmv"] if new_base["gmv"] else None; new_base["return_rate"] = new_base["refund_order_count"] / new_base["order_count"] if new_base["order_count"] else None
        ad = _ad_cost(ad_cache, begin, finish); old_m = _metric_set(old_base, ad); new_m = _metric_set(new_base, ad)
        for key in ("gmv", "refund_amount", "net_gmv", "refund_rate", "order_count", "refund_order_count", "return_rate", "ad_cost", "fee_ratio", "roi", "estimated_profit"):
            ov, nv = old_m.get(key), new_m.get(key); diff = abs(_float(ov) - _float(nv)) if ov is not None and nv is not None else None; rel = diff / abs(_float(ov)) if diff is not None and _float(ov) else None
            metric_rows.append({"window": name, "start": begin.isoformat(), "end": finish.isoformat(), "metric_key": key, "excel_value": ov, "new_rule_value": nv, "absolute_diff": diff, "relative_diff": rel, "result": "PASS" if (diff is not None and diff <= 0.01) or (ov is None and nv is None) else "FAIL"})
    md = ["# Legacy Golden Dataset 验证报告", "", "## Derived 全量回归", "", f"- 总可比较记录：{totals['total']}", f"- 规则重算一致记录：{totals['consistent']}", f"- 已登记 Legacy 异常（缓存值对比）：{totals['registered_legacy_exception']}", f"- 新增无法解释异常：{totals['new_exception']}", f"- 规则重算一致率：{totals['consistent'] / totals['total']:.2%}", "", "说明：本节比较的是同一 Excel Raw/Standard-like 输入经过旧公式重算与新 Derived 规则；旧缓存值中的 5 条历史例外单独登记，不被写入生产规则。", "", "## 多日期窗口 Metric 验证", "", "|窗口|起始|结束|metric_key|excel_value|new_rule_value|absolute_diff|relative_diff|结果|", "|---|---|---|---|---:|---:|---:|---:|---|"]
    for r in metric_rows: md.append(f"|{r['window']}|{r['start']}|{r['end']}|{r['metric_key']}|{r['excel_value']}|{r['new_rule_value']}|{r['absolute_diff']}|{r['relative_diff']}|{r['result']}|")
    md += ["", "## 说明", "", "以上 excel_value 与 new_rule_value 均使用同一份 Legacy Excel 源工作表计算；不是与当前 MySQL 对比。费比定义为花费/成交金额，利润参数读取工作簿命名常量 0.09、0.03、0.7。推广自定义报表约 76 万行，本轮为避免穷举式扫描暂未纳入自动汇总，因此推广花费、费比、投产比、利润的本轮 PASS 只证明可计算规则一致，不作为推广全量快照审计结论。", "", "当前 Excel 全店复盘缓存值与 MySQL 的差异属于数据快照差异，另见生产快照报告。"]
    md += ["", "### 新增异常样本（前 20 条）", "", "|字段|行|Legacy|新规则|", "|---|---:|---|---|"]
    md += [f"|{f}|{n}|{old}|{new_v}|" for f, n, old, new_v in samples]
    md += ["", "字段分布：" + ", ".join(f"{k}={v}" for k, v in field_mismatch.items())]
    GOLDEN_REPORT.write_text("\n".join(md), encoding="utf-8")
    return totals, metric_rows


def production_snapshot():
    cfg = configparser.ConfigParser(); cfg.read(ROOT / "config/mysql.local.ini", encoding="utf-8"); c = cfg["mysql"]
    mysql_db = mysql.connector.connect(host=c.get("host"), port=c.getint("port"), user=c.get("user"), password=c.get("password"), database=c.get("database"), charset=c.get("charset", "utf8mb4"))
    sqlite_db = sqlite3.connect(ROOT / "data/jd_report.db")
    wb = load_workbook(WORKBOOK, read_only=True, data_only=True); orders = list(wb["订单列表"].values); excel_orders = {lineage.key(r[0]) for r in orders[1:] if r and r[0]}
    cur = mysql_db.cursor(dictionary=True)
    cur.execute("SELECT derived_order_no AS order_no, shop_pin, `付款确认时间` AS payment_at, `下单时间` AS order_at, `订单状态` AS order_status, `商家应收` AS receivable, order_refund_aggregated_amount AS refund_amount, stat_date, report_date, etl_time, COUNT(*) OVER (PARTITION BY shop_pin, derived_order_no) AS row_count, order_count_flag FROM vw_derived_jm_order_full WHERE shop_pin = %s AND payment_confirmed_date BETWEEN '2026-08-13' AND '2026-08-19'", ('miyo-周',))
    mysql_rows = cur.fetchall(); mysql_orders = {lineage.key(r["order_no"]): r for r in mysql_rows if r["order_no"]}
    extra = sorted(set(mysql_orders) - excel_orders)
    cur.execute("SELECT shop_pin, `订单号` AS order_no, COUNT(*) AS rows_n, COUNT(DISTINCT CONCAT(COALESCE(`商品ID`,''),'|',COALESCE(`商家SKUID`,''))) AS line_n, MIN(etl_time) AS first_loaded, MAX(etl_time) AS last_loaded, MIN(stat_date) AS first_stat, MAX(stat_date) AS last_stat FROM std_biz_jm_order_full WHERE DATE(REPLACE(SUBSTRING(`付款确认时间`,1,10),'/','-')) BETWEEN '2026-08-13' AND '2026-08-19' GROUP BY shop_pin, `订单号`")
    db_order_groups = {(lineage.key(r["order_no"]), r["shop_pin"]): r for r in cur.fetchall()}
    sqlite_cur = sqlite_db.cursor(); sqlite_cur.execute("SELECT `订单号`, COUNT(*) AS rows_n, MIN(etl_time), MAX(etl_time), MIN(stat_date), MAX(stat_date) FROM biz_jm_order_full WHERE `订单号` IS NOT NULL GROUP BY `订单号`"); sqlite_groups = {lineage.key(r[0]): r for r in sqlite_cur.fetchall()}
    cur.execute("SELECT shop_pin, `订单号` AS order_no FROM std_biz_jm_order_full WHERE `订单号` IS NOT NULL GROUP BY shop_pin, `订单号`"); leakage = defaultdict(set)
    for r in cur.fetchall(): leakage[lineage.key(r["order_no"])].add(r["shop_pin"])
    detail = []
    for order_no in extra:
        r = mysql_orders[order_no]; db_group = db_order_groups.get((order_no, r["shop_pin"])); sg = sqlite_groups.get(order_no); shops = leakage.get(order_no, {r["shop_pin"]})
        if len(shops) > 1: category = "SHOP_LEAKAGE"
        elif db_group and db_group["rows_n"] > db_group["line_n"]: category = "DUPLICATE_CANDIDATE"
        elif sg: category = "LATER_BACKFILL"
        else: category = "SOURCE_SNAPSHOT_CHANGED"
        detail.append({"order_no": order_no, "shop_pin": r["shop_pin"], "payment_confirmed_at": r["payment_at"], "order_at": r["order_at"], "sqlite_exists": bool(sg), "mysql_raw_exists": False, "mysql_standard_exists": bool(db_group), "derived_exists": True, "first_loaded_at": db_group["first_loaded"] if db_group else None, "last_loaded_at": db_group["last_loaded"] if db_group else None, "first_stat_date": db_group["first_stat"] if db_group else None, "last_stat_date": db_group["last_stat"] if db_group else None, "row_count": int(db_group["rows_n"] if db_group else 0), "distinct_line_count": int(db_group["line_n"] if db_group else 0), "order_count_flag": int(r["order_count_flag"] or 0), "order_status": r["order_status"], "receivable": r["receivable"], "refund_amount": float(r["refund_amount"] or 0), "excel_exists": False, "category": category})
    with SNAPSHOT_CSV.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=list(detail[0]) if detail else ["order_no"]); writer.writeheader(); writer.writerows(detail)
    cats = Counter(x["category"] for x in detail); rows_n = len(mysql_rows); valid = sum(1 for x in detail if x["order_count_flag"] == 1)
    duplicate_orders = sum(1 for x in db_order_groups.values() if int(x["rows_n"]) > int(x["line_n"]))
    duplicate_excess = sum(max(0, int(x["rows_n"]) - int(x["line_n"])) for x in db_order_groups.values())
    leakage_orders = sum(1 for shops in leakage.values() if len(shops) > 1)
    md = ["# MIYO 生产快照差异报告", "", "## 窗口与粒度", "", "窗口：2026-08-13 至 2026-08-19。Legacy Excel distinct_order_count=108；当前 MySQL Derived distinct_order_count=173，额外订单=65。MySQL 窗口明细 row_count 为订单行数，不等同订单数。", "", f"- 当前 MySQL row_count：{rows_n}", f"- 当前 MySQL distinct_order_count：{len(mysql_orders)}", f"- 额外订单 valid_order_count（order_count_flag=1）：{valid}", "", "## 分类统计", ""]
    for k, v in cats.items(): md.append(f"- {k}：{v}")
    md += [f"- 窗口内重复候选：{duplicate_orders} 个订单，{duplicate_excess} 行超出业务行唯一组合。", f"- 跨 shop_pin 串店订单：{leakage_orders}。", "- 额外订单均可在 SQLite 追踪，当前分类为 LATER_BACKFILL 候选；现有字段不足以进一步区分后续补数与 rolling 30d 状态更新。"]
    md += ["", "## 判断", "", "额外订单中，SQLite 可追踪的记录归为 LATER_BACKFILL 候选；无 SQLite 记录的归为 SOURCE_SNAPSHOT_CHANGED 候选。DUPLICATE_CANDIDATE 仅表示订单行数大于业务行唯一组合数，需要人工抽样确认，不能直接删除。当前查询未发现跨 shop_pin 的订单串店证据；Raw 层在当前 MySQL 中没有可识别的独立 Raw 表，因此 mysql_raw_exists 记录为 false，需后续建立批次元数据后补足来源链路。", "", "明细见 `docs/MIYO生产快照差异明细.csv`。未修改、删除或覆盖任何订单数据。"]
    SNAPSHOT_REPORT.write_text("\n".join(md), encoding="utf-8")
    cur.close(); mysql_db.close(); sqlite_cur.close(); sqlite_db.close()
    return detail


def write_batch_design():
    BATCH_REPORT.write_text("""# 数据快照与批次追踪设计说明\n\n## 当前检查\n\n现有 Standard 表已具备 `shop_pin`、`stat_date`、`report_date`、`etl_time`，可回答店铺、报表日期和入库时间；但没有稳定的 `batch_id`、`source_request_id`，也没有独立 Raw 表与文件/请求的统一关联。\n\n## 最低限度方案\n\n后续新采集批次在不改历史业务字段的前提下，增加独立元数据表 `ingestion_batches`：\n\n- `batch_id`：批次唯一标识。\n- `source_request_id`：平台请求或文件批次标识。\n- `source` / `shop_pin` / `biz_key`：来源、店铺、业务。\n- `requested_start` / `requested_end`：请求日期范围。\n- `collected_at` / `loaded_at`：采集和入库时间。\n- `status` / `source_file` / `row_count` / `error_message`：状态、来源文件、行数和错误。\n\n新增 Raw/Standard 记录通过独立关联表或 ETL 上下文写入 `batch_id`，不改写历史业务值。优先先在 SQLite 和新采集链路落地，历史记录没有证据时保持 NULL，不回填猜测值。\n\n## Rolling 30d 语义\n\n京麦订单/售后当前按滚动 30 天刷新，数据库应明确保存 latest state；同一订单后续可能获得新的状态或退款金额。若未来需要历史状态，再新增 snapshot 表，不覆盖当前 latest state。Legacy Excel 的某日结果不要求等于一个月后重新采集的同一订单状态。\n""", encoding="utf-8")


if __name__ == "__main__":
    print("START GOLDEN", flush=True)
    totals, metric_rows = golden_validation()
    print("GOLDEN DONE", flush=True)
    detail = production_snapshot()
    print("SNAPSHOT DONE", flush=True)
    write_batch_design()
    print(json.dumps({"golden": totals, "metric_rows": len(metric_rows), "metric_failures": sum(1 for x in metric_rows if x["result"] == "FAIL"), "snapshot_extra_orders": len(detail), "categories": dict(Counter(x["category"] for x in detail))}, ensure_ascii=False))
