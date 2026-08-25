# -*- coding: utf-8 -*-
"""
run_3shops_feed_master.py — 三店真实 xlsx → 总表（2026-08-24 项目26-3）

背景：
    run_recent_30d.py 直调 handler，绕过了 main.run_business 的 _auto_save_db_from_xlsx 钩子，
    导致京准通/京麦的 xlsx 没进总表。本脚本作为兜底，从 output/<店>/ 下读取真实 xlsx → collect → flush。

特点：
    - 自动识别 3 店（从 config.xlsx「店铺清单」读 shop_id + file_prefix）
    - 商智 7 业务逐日平铺 + 京准通/京麦 30 天单文件按日期列拆分
    - 每个店依次 flush，写入各自总表（{SHOP_PIN}_商智总表 / _京准通总表 / _京麦总表）
    - 性能优化已就位（大批量删除走一次性重写 < 5 秒）

使用：
    python run_3shops_feed_master.py              # 三店全跑
    python run_3shops_feed_master.py --only OTA  # 单店
"""
import os
import sys
import glob
import argparse
import logging
from datetime import date

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
os.chdir(PROJECT_ROOT)
sys.path.insert(0, PROJECT_ROOT)

import pandas as pd
import excel_master as em
from biz_config_loader import list_shop_ids  # noqa

# 读 xlsx 统一 dtype=str + na_filter=False：
#   ① 订单号/服务单号等长数字保持字符串，防精度丢失（AGENTS.md Excel规则3）
#   ② 空单元格保留为 "" 而非 NaN，便于后续判断
_READ_XLSX = {"dtype": str, "na_filter": False}


def _read_xlsx(path, header=None):
    """读 xlsx（订单号等长数字列强制文本；支持多层表头 header=1）"""
    if header is None:
        return pd.read_excel(path, **_READ_XLSX)
    return pd.read_excel(path, header=header, **_READ_XLSX)

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger("feed_master")

# 商智 7 业务逐日平铺文件
SZ_BIZ = [
    ("商品流量来源_搜索", "搜索流量", "搜索流量_{date:%Y-%m-%d}.xlsx", None),
    ("商品流量来源_推荐", "推荐流量", "推荐流量_{date:%Y-%m-%d}.xlsx", None),
    ("商品流量来源_购物车", "购物车流量", "购物车流量_{date:%Y-%m-%d}.xlsx", None),
    ("店铺来源_三级渠道", "店铺来源_三级渠道", "店铺来源_三级渠道_{date:%Y-%m-%d}.xlsx", None),
    ("商品明细导出", "商品明细", "*_{date:%Y%m%d}_全部渠道_商品明细_分天下载.xlsx", None),
    ("商品流失分析", "商品流失分析", "商品流失分析_全部渠道_SPU_{date:%Y%m%d}_{date:%Y%m%d}.xlsx", None),
    ("商智关键词分析", "商智关键词分析", "商智关键词分析_{date:%Y-%m-%d}_day.xlsx", "day"),
]

# 京准通：单文件 30 天（通配符匹配多代文件，取最新；2026-08-24）
JZT_BIZ = [
    ("京准通快车自定义报表", "京准通快车效果自定义", "京准通快车效果自定义_*.xlsx", "日期"),
    ("京准通快车订单效果明细", "京准通快车订单效果明细", "京准通快车订单效果明细_*.xlsx", "下单时间"),
    ("京准通全站营销单品计划", "京准通全站营销单品计划", "京准通全站营销单品计划_*.xlsx", "日期"),
    ("京准通全站营销单品推广效果", "京准通全站营销单品推广效果", "京准通全站营销单品推广效果_*.xlsx", "下单日期"),
    ("京准通全站营销全店计划", "京准通全站营销全店计划", "京准通全站营销全店计划_*.xlsx", "日期"),
    ("京准通全站营销全店推广效果", "京准通全站营销全店推广效果", "京准通全站营销全店推广效果_*.xlsx", "下单日期"),
]

# 京麦：单文件 30 天（取最新文件；2026-08-24）
JM_BIZ = [
    ("京麦订单明细_完整一键导出", "京麦订单明细", "订单明细_*.xlsx", "下单时间", None),
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


def _shop_out(shop_id: str) -> str:
    return os.path.join("output", shop_id)


def _collect_daily(biz_key, folder, pattern, granularity, dates, out_dir) -> int:
    """逐日 collect"""
    n = 0
    d = os.path.join(out_dir, folder)
    for dt in dates:
        pat = os.path.join(d, pattern.format(date=dt))
        matches = glob.glob(pat)
        if not matches:
            continue
        try:
            df = _read_xlsx(matches[0])
        except Exception as e:
            log.warning(f"  ⚠️ {biz_key}/{dt} 读失败: {e}")
            continue
        if df.empty:
            continue
        em.collect(biz_key, df, dt.isoformat(), granularity)
        log.info(f"  ✅ {biz_key}/{dt}: {len(df)} 行")
        n += 1
    return n


def _collect_single_split(biz_key, folder, fname_pattern, date_col, header, out_dir) -> int:
    """单文件按日期列拆分 collect。

    2026-08-24 修正关键 bug：
        同目录可能存在多代【分段】拉取文件（如 京麦订单明细_07-22 / _08-22 / _08-24，
        每个都是 25-30 天窗口），【只能取最新】会丢掉历史窗口，造成总表行数小于单表合计。
        正确做法：合并所有同名前缀的 xlsx → 按日期列去重 → 按日期 group 拆分 collect。

    售后明细等无显式日期列的文件：整文件 collect 一次（stat_date 用文件 mtime 的前一日）。
    """
    d = os.path.join(out_dir, folder)
    matches = sorted(glob.glob(os.path.join(d, fname_pattern)), key=os.path.getmtime)
    if not matches:
        return 0
    if len(matches) > 1:
        log.info(f"  📦 {biz_key}: 发现 {len(matches)} 个文件，合并去重: " +
                 ", ".join(os.path.basename(m) for m in matches))

    # 1. 合并所有文件
    frames = []
    bases = []
    for p in matches:
        base = os.path.basename(p)
        bases.append(base)
        try:
            df = _read_xlsx(p, header)
        except Exception as e:
            log.warning(f"  ⚠️ {biz_key}/{base} 读失败: {e}")
            continue
        if df is None or (hasattr(df, "empty") and df.empty):
            continue
        frames.append(df)
    if not frames:
        log.warning(f"  ⚠️ {biz_key}: 合并后无数据（{len(matches)} 文件全空）")
        return 0

    # 2. 行级去重（不同窗口的同一天可能重叠）
    merged = pd.concat(frames, ignore_index=True)
    before = len(merged)
    # 用全字段做去重；如果日期列存在，先把日期转字符串对齐
    if date_col and date_col in merged.columns:
        merged[date_col] = merged[date_col].astype(str).str.strip()
    merged = merged.drop_duplicates()
    after = len(merged)
    if before != after:
        log.info(f"    🔁 {biz_key}: 合并 {before} 行 → 去重 {after} 行（删 {before - after}）")

    # 3. 按日期列 split collect（每日期一组）
    if not date_col or date_col not in merged.columns:
        log.warning(f"  ⚠️ {biz_key}: 无 {date_col} 列，整文件单次 collect")
        em.collect(biz_key, merged.reset_index(drop=True), "2026-08-22")
        return 1

    s = pd.to_datetime(merged[date_col], errors="coerce")
    valid_mask = s.notna()
    if valid_mask.sum() == 0:
        log.warning(f"  ⚠️ {biz_key}: 日期列[{date_col}] 全部解析失败，整文件单次 collect")
        em.collect(biz_key, merged.reset_index(drop=True), "2026-08-22")
        return 1

    n = 0
    n_groups = 0
    for dt, g in merged[valid_mask].groupby(s[valid_mask].dt.date):
        if pd.isna(dt):
            continue
        g2 = g.reset_index(drop=True)
        if g2.empty:
            continue
        em.collect(biz_key, g2, dt.isoformat())
        n += len(g2)
        n_groups += 1
    log.info(f"  ✅ {biz_key}: 合并 {len(matches)} 文件 → {after} 行 / {n_groups} 个日期组")
    return n_groups


def _summary(shop_pin: str, label: str = ""):
    """打印总表 sheet 行数与分布"""
    from excel_master import list_master_files
    for p in list_master_files(shop_pin):
        if not os.path.exists(p):
            print(f"  ⚠️ 总表未生成: {p}")
            continue
        print(f"\n  ━━━ {os.path.basename(p)} ━━━")
        for sn in pd.ExcelFile(p).sheet_names:
            df = pd.read_excel(p, sheet_name=sn)
            if "stat_date" in df.columns and len(df):
                st = pd.to_datetime(df["stat_date"], errors="coerce")
                cnt = st.dt.strftime("%Y-%m-%d").value_counts().sort_index()
                n_dates = len(cnt)
                dist = f"{n_dates} 天/共 {len(df)} 行"
            else:
                dist = "(空 sheet)"
            print(f"     [{sn}] {dist}")


def run_shop(shop_id: str) -> bool:
    """喂一家店。返回是否成功（至少有一组 collect 数据）"""
    pin = _shop_pin(shop_id)
    out_dir = _shop_out(shop_id)
    os.environ["SHOP_PIN"] = pin
    os.environ["SHOP_ID"] = shop_id

    # 27 天窗口（2026-08-24 用户决策：从 7/28 起补齐到昨天）
    # AGENTS.md：商智按近 7 天是历史默认；本次按用户要求补齐到 7/28 起 27 天
    end_d = date.today() - __import__("datetime").timedelta(days=1)
    start_d = date(2026, 7, 28)  # 固定起始日期（用户决策 2026-08-24）
    sz_dates = []
    d = start_d
    while d <= end_d:
        sz_dates.append(d)
        d += __import__("datetime").timedelta(days=1)
    log.info(f"  商智窗口: {start_d} ~ {end_d} ({len(sz_dates)} 天)")

    log.info(f"━━━ {shop_id} ({pin}) 开始 feed master ━━━")
    total = 0
    missing = []

    # 1. 商智逐日
    log.info("  ┌─ 商智 7 业务 × 7 天")
    for biz, folder, pat, gran in SZ_BIZ:
        n = _collect_daily(biz, folder, pat, gran, sz_dates, out_dir)
        if n == 0:
            missing.append(biz)
        total += n

    # 2. 京准通单文件
    log.info("  ├─ 京准通 6 业务 × 30 天")
    for biz, folder, fname, dcol in JZT_BIZ:
        n = _collect_single_split(biz, folder, fname, dcol, None, out_dir)
        if n == 0:
            missing.append(biz)
        total += n

    # 3. 京麦单文件
    log.info("  └─ 京麦 2 业务 × 30 天")
    for biz, folder, fname, dcol, hdr in JM_BIZ:
        n = _collect_single_split(biz, folder, fname, dcol, hdr, out_dir)
        if n == 0:
            missing.append(biz)
        total += n

    log.info(f"  → 共 collect {total} 组，缺失: {missing or '无'}")

    if total == 0:
        log.warning(f"⚠️ {shop_id} 无任何 xlsx 可 collect，跳过 flush")
        return False

    log.info(f"  → flush_shop({shop_id}, {pin})")
    em.flush_shop(shop_id, pin)
    _summary(pin)
    return True


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--only", help="只跑指定店前缀 (FYA/MIYO/OTA)，不传跑三店")
    args = ap.parse_args()

    all_shops = list_shop_ids()  # ['FYA箱包旗舰店', 'MIYO箱包旗舰店', 'OTA箱包旗舰店']
    targets = all_shops
    if args.only:
        targets = [s for s in all_shops if s.startswith(args.only)]

    if not targets:
        log.error(f"❌ --only {args.only} 没有匹配的店")
        sys.exit(3)

    log.info(f"━━━ 准备跑 {len(targets)} 店: {targets} ━━━")
    ok = 0
    for sid in targets:
        if run_shop(sid):
            ok += 1
        # 店间冷却（防风控），最末店不冷却
        if sid != targets[-1]:
            log.info("  ⏳ 店间冷却 60 秒…")
            import time
            time.sleep(60)

    log.info(f"━━━ feed master 完成：{ok}/{len(targets)} ━━━")
    sys.exit(0 if ok == len(targets) else 1)


if __name__ == "__main__":
    main()