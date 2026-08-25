# -*- coding: utf-8 -*-
"""refresh_reports.py
============================================================
报表刷新统一入口（2026-08-21 新增，项目23）

核心策略（区分两类业务）：
    1. **区间类业务** (supports_range=True)
       - 京麦订单明细、京麦售后明细、京准通快车自定义、京准通快车订单效果明细、
         京准通全站营销 4 个、商品明细、商品流失分析、店铺来源_三级渠道
       - 行为：按"近 30 天"区间**全量覆盖**调用（区间数据，过期日会刷新）
       - xlsx：直接用 run_business 重写同日期文件
       - DB ：save_to_db 走 Upsert（先删同 report_date 再插，覆盖语义）

    2. **逐日类业务** (supports_range=False)
       - 商智关键词分析 (day 粒度)、搜索流量、购物车流量、推荐流量
       - 行为：检查缺失日期，**只补缺失**，不重复下载已有日期（防风控）
       - xlsx：list_existing_dates 扫文件，按日补齐
       - DB ：get_existing_dates 扫表，按日补齐

与现有工具关系（互补，不重复造轮子）：
    - daily_update.py      → 全量覆盖，DB 入库（不带巡检）
    - fill_missing.py      → 仅 DB 补齐（不支持区间类业务）
    - rpa_run.py            → 全量覆盖（30 天），不查缺失
    - refresh_reports.py   → **本工具**：xlsx + DB 双向巡检，按业务类型分策略

用法：
    # 默认所有启用店铺 + 启用业务 + 近 30 天
    python refresh_reports.py

    # 单店
    python refresh_reports.py --shop "MIYO箱包旗舰店"

    # 指定日期窗口
    python refresh_reports.py --start_date 2026-07-22 --end_date 2026-08-20

    # 指定业务
    python refresh_reports.py --biz_keys "京麦订单明细_完整一键导出,商智关键词分析"

    # 只跑 xlsx（不入库）
    python refresh_reports.py --skip-db

    # 只跑 DB（不下 xlsx）
    python refresh_reports.py --skip-xlsx

    # 干跑（只打印计划，不实际请求）
    python refresh_reports.py --dry-run

退出码（与影刀约定一致）：
    0 - 全部成功
    1 - 部分业务失败
    2 - 鉴权过期
    3 - 参数错误
"""
import os
import sys
import argparse
import logging
import traceback
import time
from datetime import datetime, timedelta

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, PROJECT_ROOT)

# 鉴权检查跳过（由接口返回码判定）
os.environ.setdefault("AUTH_LOADER", "1")
os.environ.setdefault("AUTH_RPA_CLI", "0")
os.environ.setdefault("AUTH_SKIP_COOKIE_EXPIRE_CHECK", "1")
os.environ.setdefault("AUTH_SKIP_H5ST_EXPIRE_CHECK", "1")

from biz_config_loader import (
    list_shop_ids,
    list_biz_keys,
    get_output_dirname,
    get_biz_feature,
)

# rpa_run 的 list_existing_dates（xlsx 巡检核心）
from rpa_run import list_existing_dates as _list_xlsx_dates
# db_utils 的 DB 巡检
from db_utils import (
    get_db_path,
    get_existing_dates as _list_db_dates,
    biz_key_to_table_name,
)

# ============================ 退出码 ============================
EXIT_SUCCESS = 0
EXIT_BIZ_FAIL = 1
EXIT_AUTH_EXPIRED = 2
EXIT_PARAM_ERROR = 3

# ============================ 日志 ============================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)


# ============================ 辅助 ============================

def _gen_date_range(start_date: str, end_date: str) -> list:
    """生成 [start_date, end_date] 区间所有日期（含两端）"""
    start = datetime.strptime(start_date, "%Y-%m-%d")
    end = datetime.strptime(end_date, "%Y-%m-%d")
    dates = []
    cur = start
    while cur <= end:
        dates.append(cur.strftime("%Y-%m-%d"))
        cur += timedelta(days=1)
    return dates


def _compute_missing_dates(expected: list, existing: set) -> list:
    """expected 中不在 existing 的日期（升序）"""
    return [d for d in expected if d not in existing]


def _is_auth_expired_exception(e: Exception) -> bool:
    """鉴权过期异常识别"""
    name = type(e).__name__
    return name in ("CookieExpiredError", "H5stExpiredError", "AuthFileNotFound")


# ============================ 巡检：xlsx + DB ============================

def _inspect_xlsx(shop_id: str, biz_key: str, expected_dates: list) -> dict:
    """巡检 xlsx 已存在日期

    返回:
        {"missing": [...], "existing_count": int}
    """
    output_root = os.path.join(PROJECT_ROOT, "output")
    existing = _list_xlsx_dates(output_root, shop_id, biz_key)
    missing = _compute_missing_dates(expected_dates, existing)
    return {"missing": missing, "existing_count": len(existing)}


def _inspect_db(biz_key: str, expected_dates: list, granularity: str = None) -> dict:
    """巡检 DB 已存在日期

    返回:
        {"missing": [...], "existing_count": int, "db_enabled": bool}
    """
    db_path = get_db_path()
    if not os.path.exists(db_path):
        return {"missing": expected_dates, "existing_count": 0, "db_enabled": False}

    table_name = biz_key_to_table_name(biz_key)
    import sqlite3
    conn = sqlite3.connect(db_path, timeout=30)
    try:
        try:
            existing = _list_db_dates(conn, table_name, granularity)
        except Exception as e:
            logging.debug(f"DB 表 {table_name} 不存在：{e}")
            existing = set()
    finally:
        conn.close()

    missing = _compute_missing_dates(expected_dates, existing)
    return {"missing": missing, "existing_count": len(existing), "db_enabled": True}


# ============================ 业务执行 ============================

def _run_range_biz(shop_id: str, biz_key: str, start_date: str, end_date: str,
                   skip_xlsx: bool, skip_db: bool, dry_run: bool) -> int:
    """区间类业务（supports_range=True）：按近 30 天区间全量覆盖调用

    策略：
        - xlsx：直接重跑，覆盖原日期文件（save_to_xlsx 同名覆盖）
        - DB ：save_to_db 走 Upsert（先删同 report_date 再插）
        - 串接执行：先 xlsx，再 DB（DB save_to_db 通常在 xlsx 流程末尾触发，无需手动调）

    参数:
        shop_id     店铺名
        biz_key     业务 key
        start_date  开始日期 YYYY-MM-DD
        end_date    结束日期 YYYY-MM-DD
        skip_xlsx   是否跳过 xlsx 下载
        skip_db     是否跳过 DB 入库
        dry_run     干跑模式

    返回:
        int - 退出码
    """
    print(f"\n{'='*70}")
    print(f"📊 [区间类-全量覆盖] {biz_key} | {shop_id}")
    print(f"   日期窗口: {start_date} ~ {end_date}")
    print(f"   xlsx: {'跳过' if skip_xlsx else '覆盖'} | DB: {'跳过' if skip_db else '覆盖'}")
    print(f"{'='*70}")

    os.environ["SHOP_ID"] = shop_id

    try:
        import main as _main
    except ImportError as e:
        print(f"❌ import main 失败：{e}")
        return EXIT_BIZ_FAIL

    if dry_run:
        print(f"   [DRY-RUN] main.run_business('{biz_key}', start_date='{start_date}', end_date='{end_date}')")
        return EXIT_SUCCESS

    try:
        # 区间类业务：直接传 start_date/end_date，由 main.run_business 内部完成 xlsx + DB 落地
        _main.run_business(biz_key, start_date=start_date, end_date=end_date)
        print(f"✅ {biz_key} 完成")
        return EXIT_SUCCESS
    except Exception as e:
        if _is_auth_expired_exception(e):
            print(f"🔄 [鉴权过期] {e}")
            return EXIT_AUTH_EXPIRED
        print(f"❌ {biz_key} 失败：{type(e).__name__}: {e}")
        traceback.print_exc()
        return EXIT_BIZ_FAIL


def _run_daily_biz(shop_id: str, biz_key: str, start_date: str, end_date: str,
                   expected_dates: list, granularity: str,
                   skip_xlsx: bool, skip_db: bool, dry_run: bool,
                   interval: int = 0) -> int:
    """逐日类业务（supports_range=False）：检查缺失，按日补齐

    策略：
        - xlsx 缺失：list_existing_dates 扫，按日补齐
        - DB 缺失：get_existing_dates 扫，按日补齐
        - 合并：xlsx ∪ DB 任一缺失则补（一次性跑，run_business 同时写 xlsx + DB）

    参数:
        shop_id        店铺名
        biz_key        业务 key
        start_date     开始日期
        end_date       结束日期
        expected_dates 预期应有日期列表（_gen_date_range 已生成）
        granularity    粒度（None / "day" / "month"）
        skip_xlsx      跳过 xlsx
        skip_db        跳过 DB
        dry_run        干跑

    返回:
        int - 退出码
    """
    print(f"\n{'='*70}")
    print(f"📅 [逐日类-缺失补齐] {biz_key} | {shop_id}")
    print(f"   日期窗口: {start_date} ~ {end_date}（{len(expected_dates)} 天）")
    print(f"   粒度: {granularity or 'N/A'}")

    os.environ["SHOP_ID"] = shop_id

    # 巡检：xlsx + DB
    xlsx_status = _inspect_xlsx(shop_id, biz_key, expected_dates) if not skip_xlsx else {"missing": [], "existing_count": "-"}
    db_status = _inspect_db(biz_key, expected_dates, granularity) if not skip_db else {"missing": [], "existing_count": "-", "db_enabled": True}

    # 合并缺失：xlsx ∪ DB
    missing_set = set()
    if not skip_xlsx:
        missing_set.update(xlsx_status["missing"])
    if not skip_db:
        missing_set.update(db_status["missing"])
    missing = sorted(missing_set)

    print(f"   xlsx: 已 {xlsx_status['existing_count']} / 缺 {len(xlsx_status['missing'])}")
    if not skip_db:
        print(f"   DB :  已 {db_status['existing_count']} / 缺 {len(db_status['missing'])}")
    print(f"   合并待补: {len(missing)} 天: {missing[:5]}{'...' if len(missing) > 5 else ''}")
    print(f"{'='*70}")

    if not missing:
        print(f"   ✅ 无缺失，跳过")
        return EXIT_SUCCESS

    if dry_run:
        for d in missing:
            print(f"   [DRY-RUN] main.run_business('{biz_key}', date='{d}')")
        return EXIT_SUCCESS

    try:
        import main as _main
    except ImportError as e:
        print(f"❌ import main 失败：{e}")
        return EXIT_BIZ_FAIL

    success_n = 0
    failed_n = 0
    auth_expired = False
    for i, d in enumerate(missing, 1):
        print(f"   [{i}/{len(missing)}] 补 {d} ...", end=" ")
        try:
            kwargs = {"date": d}
            if granularity:
                kwargs["granularity"] = granularity
            _main.run_business(biz_key, **kwargs)
            print("✅")
            success_n += 1
        except Exception as e:
            if _is_auth_expired_exception(e):
                print(f"🔄 [鉴权过期] {e}")
                auth_expired = True
                break
            print(f"❌ {type(e).__name__}: {e}")
            failed_n += 1

        # 间隔控制（防风控）
        if interval > 0 and i < len(missing):
            time.sleep(interval)

    print(f"   📊 {biz_key}：成功 {success_n} / 失败 {failed_n} / 总 {len(missing)}")
    if auth_expired:
        return EXIT_AUTH_EXPIRED
    if failed_n > 0:
        return EXIT_BIZ_FAIL
    return EXIT_SUCCESS


# ============================ 单店调度 ============================

def run_shop_refresh(shop_id: str, biz_keys: list, start_date: str, end_date: str,
                     expected_dates: list, skip_xlsx: bool, skip_db: bool,
                     dry_run: bool, interval: int) -> int:
    """跑一个店的所有启用业务（按 supports_range 分流）"""
    print(f"\n{'#'*70}")
    print(f"# [SHOP] {shop_id}")
    print(f"{'#'*70}")

    overall_exit = EXIT_SUCCESS

    for biz_key in biz_keys:
        feature = get_biz_feature(biz_key)
        if not feature:
            print(f"\n⚠️ 业务 {biz_key} 无特性配置，跳过")
            continue

        supports_range = feature.get("supports_range", True)
        granularity = feature.get("default_granularity")

        if supports_range:
            code = _run_range_biz(
                shop_id=shop_id, biz_key=biz_key,
                start_date=start_date, end_date=end_date,
                skip_xlsx=skip_xlsx, skip_db=skip_db, dry_run=dry_run,
            )
        else:
            code = _run_daily_biz(
                shop_id=shop_id, biz_key=biz_key,
                start_date=start_date, end_date=end_date,
                expected_dates=expected_dates, granularity=granularity,
                skip_xlsx=skip_xlsx, skip_db=skip_db, dry_run=dry_run,
                interval=interval,
            )

        # 退出码优先级：鉴权过期(2) > 业务失败(1) > 成功(0)
        if code == EXIT_AUTH_EXPIRED:
            return EXIT_AUTH_EXPIRED  # 立即停所有后续店/业务
        if code != EXIT_SUCCESS:
            overall_exit = EXIT_BIZ_FAIL

    return overall_exit


# ============================ CLI ============================

def main():
    parser = argparse.ArgumentParser(
        description="报表刷新统一入口（2026-08-21 新增）：区间类业务覆盖，逐日类业务补齐。xlsx + DB 双向同步",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--shop", default=None,
                        help="指定单个店铺；不传则遍历 config.xlsx 启用店铺")
    parser.add_argument("--biz_keys", default=None,
                        help="指定业务key列表（逗号分隔），默认遍历启用业务")
    parser.add_argument("--start_date", default=None,
                        help="起始日期 YYYY-MM-DD（默认：今天-30天）")
    parser.add_argument("--end_date", default=None,
                        help="结束日期 YYYY-MM-DD（默认：今天-1天）")
    parser.add_argument("--interval", type=int, default=0,
                        help="逐日循环间隔秒数（默认 0）")
    parser.add_argument("--skip-xlsx", action="store_true",
                        help="跳过 xlsx 巡检与下载（只跑 DB）")
    parser.add_argument("--skip-db", action="store_true",
                        help="跳过 DB 巡检与入库（只跑 xlsx）")
    parser.add_argument("--dry-run", action="store_true",
                        help="干跑：只打印计划，不实际请求京东接口")
    args = parser.parse_args()

    if args.skip_xlsx and args.skip_db:
        print("❌ --skip-xlsx 与 --skip-db 不能同时使用")
        return EXIT_PARAM_ERROR

    # 日期窗口
    today = datetime.now()
    end_date = args.end_date or (today - timedelta(days=1)).strftime("%Y-%m-%d")
    start_date = args.start_date or (today - timedelta(days=30)).strftime("%Y-%m-%d")
    expected_dates = _gen_date_range(start_date, end_date)

    # 店铺列表
    if args.shop:
        shop_ids = [args.shop]
    else:
        shop_ids = list_shop_ids(enabled_only=True)

    # 业务列表
    if args.biz_keys:
        biz_keys = [b.strip() for b in args.biz_keys.split(",") if b.strip()]
    else:
        biz_keys = list_biz_keys(enabled_only=True)

    # 启动 banner
    print("\n" + "=" * 70)
    print(f"🔄 refresh_reports.py | 报表刷新统一入口")
    print(f"   店铺: {shop_ids}")
    print(f"   业务: {biz_keys}")
    print(f"   日期窗口: {start_date} ~ {end_date}（{len(expected_dates)} 天）")
    print(f"   模式: {'DRY-RUN' if args.dry_run else '实际执行'}")
    print(f"   xlsx: {'跳过' if args.skip_xlsx else '✅'} | DB: {'跳过' if args.skip_db else '✅'}")
    print("=" * 70)

    # 店间间隔（防风控）
    from rpa_run import _read_shop_interval_seconds
    shop_interval = _read_shop_interval_seconds(None)  # 走 xlsx 配置
    print(f"   店间冷却: {shop_interval} 秒")

    # IMAP 健康检查（与 rpa_run 一致）
    try:
        from imap_config_loader import load_imap_config, check_imap_health as _imap_check
        _cfg = load_imap_config()
        _imap_check(_cfg)
    except FileNotFoundError:
        pass  # ini 不存在静默跳过
    except RuntimeError as e:
        if not args.dry_run:
            print(f"\n[BLOCK] {e}\n")
            return EXIT_AUTH_EXPIRED

    overall_exit = EXIT_SUCCESS

    for i, shop_id in enumerate(shop_ids):
        # 店间冷却
        if i > 0 and shop_interval > 0 and not args.dry_run:
            print(f"\n⏳ 店间冷却 {shop_interval} 秒...")
            time.sleep(shop_interval)

        code = run_shop_refresh(
            shop_id=shop_id, biz_keys=biz_keys,
            start_date=start_date, end_date=end_date,
            expected_dates=expected_dates,
            skip_xlsx=args.skip_xlsx, skip_db=args.skip_db,
            dry_run=args.dry_run, interval=args.interval,
        )

        # 项目24（2026-08-22 新增）：每店刷新完后刷自己的 Excel 总表
        # refresh_reports 是 subprocess 调 main.run_business()，不触发 main.py main() 入口
        # 所以这里必须手动调 flush；run_business 内 save_to_db 已 collect 到 _CACHE
        if not args.dry_run and code != EXIT_AUTH_EXPIRED:
            _try_flush_excel_master_safe_for_shop(shop_id)

        if code == EXIT_AUTH_EXPIRED:
            return EXIT_AUTH_EXPIRED
        if code != EXIT_SUCCESS:
            overall_exit = EXIT_BIZ_FAIL

    # 汇总
    print("\n" + "=" * 70)
    print(f"📊 三店刷新汇总: {'✅ 全部成功' if overall_exit == 0 else '⚠️ 有失败'}")
    print("=" * 70)

    # 项目24（2026-08-22 新增）：每店刷新完后刷自己的 Excel 总表
    # 注意：refresh_reports 是多店遍历，必须每店单独调一次
    # main.run_business 已通过 save_to_db → collect 暂存数据，所以这里直接调 flush
    # 但 main.run_business 内已经调了 _try_flush_excel_master_safe（每次调 main 都触发）
    # 所以这里 refresh_reports 实际上是冗余调用，但幂等无害
    # （幂等保证：flush_shop 后 _CACHE 已清空，再次 flush 是空数据直接返回）

    return overall_exit


def _try_flush_excel_master_safe_for_shop(shop_id: str) -> None:
    """单店 flush 兜底（与 main.py 的 _try_flush_excel_master_safe 类似）。"""
    try:
        import excel_master as _em
        from runtime_config import get_shop_id as _get_sid, get_shop_pin as _get_pin
        # 注意：runtime_config 读环境变量 SHOP_ID/SHOP_PIN，需要在外层 set
        shop_pin = _get_pin()
        _em.flush_shop(shop_id, shop_pin)
    except Exception as e:
        print(f"⚠️ [ExcelMaster] refresh_reports.flush_shop({shop_id}) 失败：{type(e).__name__}: {e}")


if __name__ == "__main__":
    sys.exit(main())