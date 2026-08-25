# -*- coding: utf-8 -*-
"""backfill_master.py — 从已生成 xlsx 补入 DB + 重建总表（京准通/京麦，2026-08-25）

背景：run_recent_30d.py 直调 handler 绕过 main.run_business 自动入库钩子，
     京准通/京麦 xlsx 已生成但没进 DB / 总表。

方案：pandas 读 xlsx（大文件 40 万行/个，约 90 秒/文件，C 实现不 OOM）
     → 按日期列拆分 → save_to_db（内部联动 collect）
     → rebuild_from_db 重建三域总表

用法：
    python backfill_master.py --only MIYO            # 单店京准通+京麦
    python backfill_master.py --only MIYO --biz jzt  # 只京准通
    python backfill_master.py --only MIYO --biz jm   # 只京麦

注意：京准通/京麦数据量大（单业务 40 万行），后台跑约 10-25 分钟。
"""
import os
import sys
import glob
import argparse
import logging
from datetime import datetime

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
os.chdir(PROJECT_ROOT)
sys.path.insert(0, PROJECT_ROOT)

from biz_config_loader import list_shop_ids

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger("backfill_master")

# 京准通：单文件 30 天，按日期列拆分
JZT_BIZ = [
    ("京准通快车自定义报表",     "京准通快车效果自定义",     "京准通快车效果自定义_*.xlsx",       "日期"),
    ("京准通快车订单效果明细",   "京准通快车订单效果明细",   "京准通快车订单效果明细_*.xlsx",     "下单时间"),
    ("京准通全站营销单品计划",   "京准通全站营销单品计划",   "京准通全站营销单品计划_*.xlsx",     "日期"),
    ("京准通全站营销单品推广效果", "京准通全站营销单品推广效果", "京准通全站营销单品推广效果_*.xlsx", "下单日期"),
    ("京准通全站营销全店计划",   "京准通全站营销全店计划",   "京准通全站营销全店计划_*.xlsx",     "日期"),
    ("京准通全站营销全店推广效果", "京准通全站营销全店推广效果", "京准通全站营销全店推广效果_*.xlsx", "下单日期"),
]
# 京麦：订单 header=0，售后 header=1（3 层结构：行1分组名/行2字段/行3+数据）
JM_BIZ = [
    ("京麦订单明细_完整一键导出", "京麦订单明细", "订单明细_*.xlsx", "下单时间", 0),
    ("京麦售后明细_完整一键导出", "京麦售后明细", "售后明细_*.xlsx", "售后申请时间", 1),
]


def _shop_pin(shop_id: str) -> str:
    """shop_id → shop_pin（从 config.xlsx「店铺账号」sheet 读）"""
    import openpyxl
    wb = openpyxl.load_workbook("config/config.xlsx", read_only=True, data_only=True)
    ws = wb["店铺账号"]
    short = shop_id.replace("箱包旗舰店", "")
    for row in ws.iter_rows(values_only=True):
        if not row or len(row) < 3:
            continue
        if str(row[0]).strip() == short:
            wb.close()
            return str(row[1]).strip()
    wb.close()
    raise ValueError(f"❌ 在 config.xlsx「店铺账号」找不到 {short}")


def _save(biz_key, df, report_date):
    """入库 DB（save_to_db 内部已联动 excel_master.collect）。"""
    from db_utils import save_to_db
    try:
        ok = save_to_db(biz_key=biz_key, df=df, report_date=report_date)
        if ok:
            log.debug(f"    入库 {biz_key}/{report_date}: {len(df)} 行")
        else:
            log.warning(f"    ⚠️ {biz_key}/{report_date} 入库跳过")
    except Exception as e:
        log.warning(f"    ⚠️ {biz_key}/{report_date} 入库失败: {type(e).__name__}: {e}")


def _collect_single_split(biz_key, folder, fname_pattern, date_col, header, out_dir) -> int:
    """逐个文件读 xlsx → 按文件内「日期」列拆 N 组 → 逐组入库（2026-08-25 简化）。

    每个 xlsx 文件本身就是「近 30 天」完整快照（每天 ~13,000 行京准通快车自定义明细），
    跨文件 concat + drop_duplicates 没意义（用户决策 2026-08-25）。
    多次入库同 stat_date：依赖 upsert 语义（先 delete 同 shop_pin+stat_date 后 insert）保证幂等。

    入参:
        biz_key       业务 key
        folder        xlsx 所在子目录名（如「京准通快车效果自定义」）
        fname_pattern 文件名 glob（如「京准通快车效果自定义_*.xlsx」）
        date_col      文件内日期列名（如「日期」），按此拆 stat_date
        header        xlsx 表头行（0=首行；京麦售后用 1）
        out_dir       output 父目录（含 shop_id 子目录）
    """
    import pandas as pd
    d = os.path.join(out_dir, folder)
    matches = sorted(glob.glob(os.path.join(d, fname_pattern)), key=os.path.getmtime)
    if not matches:
        log.warning(f"  ⚠️ {biz_key}: 无文件匹配 {fname_pattern}")
        return 0

    log.info(f"  📦 {biz_key}: 发现 {len(matches)} 个文件，逐文件按「{date_col}」列拆入对应 stat_date")
    total_groups = 0
    for p in matches:
        try:
            df = pd.read_excel(p, header=header, dtype=str, na_filter=False)
        except Exception as e:
            log.warning(f"  ⚠️ {biz_key}/{os.path.basename(p)} 读失败: {type(e).__name__}: {e}")
            continue
        if df is None or df.empty:
            continue
        log.info(f"    已读 {os.path.basename(p)}: {len(df)} 行")

        # 按文件内「日期」列拆 N 组入库
        if not date_col or date_col not in df.columns:
            log.warning(f"    ⚠️ 无 {date_col} 列，整文件按文件名日期单次入库")
            _save(biz_key, df, datetime.now().strftime("%Y-%m-%d"))
            total_groups += 1
            continue

        s = pd.to_datetime(df[date_col], errors="coerce")
        valid = s.notna()
        if valid.sum() == 0:
            log.warning(f"    ⚠️ 日期列全部解析失败，整文件按今天单次入库")
            _save(biz_key, df, datetime.now().strftime("%Y-%m-%d"))
            total_groups += 1
            continue

        for dt, g in df[valid].groupby(s[valid].dt.date):
            if pd.isna(dt):
                continue
            _save(biz_key, g.reset_index(drop=True), dt.isoformat())
            total_groups += 1

    log.info(f"  ✅ {biz_key}: 按日期拆分入库 {total_groups} 组")
    return total_groups


def _dedup_by_primary_key(biz_key, df):
    """按业务主键去重（2026-08-25 临时禁用）。

    根因：京准通快车自定义每行是「日期×计划×单元×SKU×地域」组合明细，
    但 infer_primary_key 首列匹配「计划ID」就返回，按计划ID 去重后
    每计划只剩 1 行/天 → 每天 1 万行变 5-11 行 → 1000 倍数据丢失。

    替代方案：依赖外层 _collect_single_split 已做的全字段 drop_duplicates（合并阶段去重跨文件重复），
    按文件内日期 groupby 时不再去重，避免误删跨日期但同主键的合法行。
    如果后续需要业务级精确去重，请用 config.xlsx 的复合主键配置，而非默认推断。
    """
    return df


def run_shop(shop_id: str, biz_filter: str = "all") -> bool:
    """补入一家店的京准通/京麦到 DB + 重建总表。"""
    pin = _shop_pin(shop_id)
    out_dir = os.path.join("output", shop_id)
    os.environ["SHOP_PIN"] = pin
    os.environ["SHOP_ID"] = shop_id

    log.info(f"━━━ {shop_id} ({pin}) 补入 DB ━━━")
    total = 0
    missing = []

    if biz_filter in ("all", "jzt"):
        log.info("  ├─ 京准通 6 业务")
        for biz, folder, fname, dcol in JZT_BIZ:
            n = _collect_single_split(biz, folder, fname, dcol, 0, out_dir)
            if n == 0:
                missing.append(biz)
            total += n

    if biz_filter in ("all", "jm"):
        log.info("  └─ 京麦 2 业务")
        for biz, folder, fname, dcol, hdr in JM_BIZ:
            n = _collect_single_split(biz, folder, fname, dcol, hdr, out_dir)
            if n == 0:
                missing.append(biz)
            total += n

    log.info(f"  → 共入库 {total} 组，缺失: {missing or '无'}")

    if total == 0:
        log.warning(f"⚠️ {shop_id} 无数据可补入，跳过重建")
        return False

    # 数据已进 DB，重建三域总表
    log.info(f"  → rebuild_from_db({shop_id}, {pin})")
    import excel_master as em
    em.rebuild_from_db(shop_id, pin)
    return True


def main():
    ap = argparse.ArgumentParser(description="从已生成 xlsx 补入 DB + 重建总表")
    ap.add_argument("--only", required=True, help="店铺前缀 FYA/MIYO/OTA")
    ap.add_argument("--biz", default="all", choices=["all", "jzt", "jm"], help="业务域")
    args = ap.parse_args()

    all_shops = list_shop_ids()
    targets = [s for s in all_shops if s.startswith(args.only)]
    if not targets:
        log.error(f"❌ --only {args.only} 没有匹配的店")
        sys.exit(3)

    log.info(f"━━━ 补入 {len(targets)} 店: {targets}（biz={args.biz}）━━━")
    ok = 0
    for sid in targets:
        if run_shop(sid, args.biz):
            ok += 1
    log.info(f"━━━ 完成：{ok}/{len(targets)} ━━━")
    sys.exit(0 if ok == len(targets) else 1)


if __name__ == "__main__":
    main()