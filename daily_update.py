# -*- coding: utf-8 -*-
"""
daily_update.py
============================================================
京东数据报表每日自动更新脚本（2026-08-14 上线骨架）。

用途：
    - 影刀 RPA 每日定时调用（如每天凌晨 2 点）
    - 自动拉取**最近 30 天**（不含当天）的数据，按 (report_date, 业务主键) Upsert 到 SQLite
    - 与现有 Excel 导出**双保险**：同时保留 Excel 备份 + 数据库入库

策略（关键）：
    1. 日期窗口：[今天-30天, 今天-1天]（含两端，共 30 天；不含今天，因为今天数据未生成）
    2. 业务分类（必须区分）：
       - supports_range=True   → 直接 start_date/end_date 区间拉取（如京准通快车、京麦订单）
       - supports_range=False  → 逐日循环调用（商智 4 个流量来源、关键词分析等）
    3. 失败容忍：单业务单日失败不影响其他业务（try/except 包裹）

CLI 用法：
    # 默认近 30 天，MVP 2 个业务
    python daily_update.py

    # 指定业务
    python daily_update.py --biz_keys "商智关键词分析,店铺来源_三级渠道"

    # 指定日期窗口（影刀定时可传）
    python daily_update.py --start_date 2026-07-15 --end_date 2026-08-13

    # 自定义间隔（避开风控）
    python daily_update.py --interval 30

    # 干跑（只打印，不入库）
    python daily_update.py --dry-run

退出码（影刀约定）：
    0 - 全部成功
    1 - 部分失败
    2 - 鉴权过期（Cookie / h5st）
    3 - 系统错误
"""
import sys
import os
import argparse
import logging
import traceback
from datetime import datetime, timedelta

# 项目根目录加入 sys.path
PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, PROJECT_ROOT)

# Phase 2.4（2026-08-20）：业务清单 + 业务特性从 config.xlsx 读取
# 替代原硬编码 MVP_BIZ_KEYS / BIZ_FEATURES（AGENTS.md 第3条禁止硬编码业务列表）
from biz_config_loader import list_biz_keys, get_biz_feature

# ====================================================================
#  业务配置：从 config.xlsx「业务清单」sheet 读取（Phase 2.4 改造，2026-08-20）
# ====================================================================
# 替代原硬编码 MVP_BIZ_KEYS / BIZ_FEATURES（AGENTS.md 第3条禁止硬编码业务列表）
# 调用方式：
#   list_biz_keys()                → 启用的全部业务（默认 --biz_keys 入参）
#   get_biz_feature(biz_key)       → 业务特性 dict（supports_range/default_granularity/extra_kwargs）
# 数据源：config.xlsx「业务清单」sheet（由 create_config_sheets.py 生成）


# ====================================================================
#  核心逻辑
# ====================================================================

def _gen_date_range(start_date: str, end_date: str):
    """生成 [start_date, end_date] 区间内所有日期（含两端）。"""
    start = datetime.strptime(start_date, "%Y-%m-%d")
    end = datetime.strptime(end_date, "%Y-%m-%d")
    dates = []
    cur = start
    while cur <= end:
        dates.append(cur.strftime("%Y-%m-%d"))
        cur += timedelta(days=1)
    return dates


# 退出码（与 AGENTS.md 第47-51条约定对齐）
EXIT_SUCCESS = 0       # 全部成功
EXIT_BIZ_FAIL = 1      # 部分业务失败
EXIT_AUTH_EXPIRED = 2  # 鉴权过期（Cookie / h5st，触发影刀重抓）
EXIT_SYSTEM_ERROR = 3  # 系统错误


def _is_auth_expired_exception(e: Exception) -> bool:
    """判断异常是否是鉴权过期类（兼容 main.py 与 auth_loader.py 两套定义）

    背景：main.py:68 自定义了 CookieExpiredError，auth_loader.py:110 也定义了一套。
    两套类互不继承，except (auth_loader.CookieExpiredError, ...) 会漏掉 main.py 抛的。
    按类名判断最稳健，跨模块兼容。
    """
    name = type(e).__name__
    return name in ("CookieExpiredError", "H5stExpiredError", "AuthFileNotFound")


def _run_one_biz(biz_key: str, start_date: str, end_date: str, interval: int = 0, dry_run: bool = False) -> int:
    """跑一个业务，返回退出码（0=成功 / 1=业务失败 / 2=鉴权过期）。

    策略：
        - supports_range=True：直接区间调用
        - supports_range=False：逐日循环
    """
    import main  # 延迟导入，避免循环引用
    import time

    # Phase 2.4：业务特性从 config.xlsx 读取（替代 BIZ_FEATURES 字典）
    feature = get_biz_feature(biz_key)
    if not feature:
        logging.warning(f"⚠️ 业务 {biz_key} 无特性配置，跳过")
        return EXIT_SUCCESS  # 无配置视为跳过，不当作失败

    supports_range = feature.get("supports_range", True)
    granularity = feature.get("default_granularity")
    extra_kwargs = feature.get("extra_kwargs", {})

    print(f"\n{'='*60}")
    print(f"📊 业务：{biz_key}")
    print(f"   日期窗口：{start_date} ~ {end_date}")
    print(f"   区间支持：{supports_range}（True=一次性区间，False=逐日循环）")
    print(f"   粒度：{granularity or 'N/A'}")
    print(f"{'='*60}")

    try:
        if supports_range:
            # 一次性区间调用
            kwargs = {
                "start_date": start_date,
                "end_date": end_date,
                **extra_kwargs,
            }
            if granularity:
                kwargs["granularity"] = granularity

            if dry_run:
                print(f"   [DRY-RUN] 调用：main.run_business('{biz_key}', **{kwargs})")
            else:
                main.run_business(biz_key, **kwargs)
        else:
            # 逐日循环
            dates = _gen_date_range(start_date, end_date)
            print(f"   共 {len(dates)} 天，逐日循环")
            for i, d in enumerate(dates, 1):
                print(f"   [{i}/{len(dates)}] 日期：{d}")
                kwargs = {"date": d, **extra_kwargs}
                if granularity:
                    kwargs["granularity"] = granularity

                if dry_run:
                    print(f"      [DRY-RUN] main.run_business('{biz_key}', **{kwargs})")
                else:
                    main.run_business(biz_key, **kwargs)
                # 间隔控制（避开风控）
                if interval > 0 and i < len(dates):
                    print(f"      等待 {interval} 秒...")
                    time.sleep(interval)

        print(f"✅ {biz_key} 完成")
        return EXIT_SUCCESS
    except Exception as e:
        # 鉴权过期优先识别（按类名匹配，兼容 main.py 与 auth_loader.py 两套定义）
        if _is_auth_expired_exception(e):
            print(f"🔄 [鉴权过期] {biz_key}：{type(e).__name__}: {e}")
            return EXIT_AUTH_EXPIRED
        print(f"❌ {biz_key} 失败：{type(e).__name__}: {e}")
        traceback.print_exc()
        return EXIT_BIZ_FAIL


def main():
    parser = argparse.ArgumentParser(description="京东数据报表每日自动更新（2026-08-14）")
    parser.add_argument(
        "--biz_keys",
        type=str,
        default=",".join(list_biz_keys()),
        help=f"业务key列表（逗号分隔），默认全部启用业务：{','.join(list_biz_keys())}",
    )
    parser.add_argument(
        "--start_date",
        type=str,
        default=None,
        help="开始日期 YYYY-MM-DD（默认：今天-30天）",
    )
    parser.add_argument(
        "--end_date",
        type=str,
        default=None,
        help="结束日期 YYYY-MM-DD（默认：今天-1天，因为今天数据未生成）",
    )
    parser.add_argument(
        "--interval",
        type=int,
        default=0,
        help="逐日循环时的间隔秒数（默认 0 = 不间隔）",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="干跑：只打印计划，不实际调用业务",
    )
    args = parser.parse_args()

    # 计算日期窗口
    today = datetime.now()
    end_date = args.end_date or (today - timedelta(days=1)).strftime("%Y-%m-%d")
    start_date = args.start_date or (today - timedelta(days=30)).strftime("%Y-%m-%d")

    biz_keys = [b.strip() for b in args.biz_keys.split(",") if b.strip()]

    # 启动 banner
    print()
    print("=" * 70)
    print(f"📅 京东数据报表 - 每日自动更新")
    print(f"   时间：{today.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"   日期窗口：{start_date} ~ {end_date}（共 {(datetime.strptime(end_date, '%Y-%m-%d') - datetime.strptime(start_date, '%Y-%m-%d')).days + 1} 天）")
    print(f"   业务清单：{biz_keys}")
    print(f"   间隔：{args.interval} 秒")
    print(f"   干跑：{args.dry_run}")
    print("=" * 70)

    # 跑每个业务
    success_count = 0
    failed_count = 0
    auth_expired_count = 0  # 鉴权过期计数（独立于业务失败）
    for biz_key in biz_keys:
        code = _run_one_biz(biz_key, start_date, end_date, args.interval, args.dry_run)
        if code == EXIT_SUCCESS:
            success_count += 1
        elif code == EXIT_AUTH_EXPIRED:
            auth_expired_count += 1
        else:
            failed_count += 1

    # 总结
    print()
    print("=" * 70)
    print(f"📊 总结：成功 {success_count} / 失败 {failed_count} / 鉴权过期 {auth_expired_count} / 总 {len(biz_keys)}")
    print("=" * 70)

    # 退出码优先级：鉴权过期(2) > 业务失败(1) > 全部成功(0)
    # AGENTS.md 第49条：鉴权过期触发影刀重抓，必须优先返回 2
    if auth_expired_count > 0:
        _try_flush_excel_master_safe()  # 项目24（2026-08-22）：鉴权过期也尝试刷总表
        return EXIT_AUTH_EXPIRED
    if failed_count > 0:
        _try_flush_excel_master_safe()  # 部分失败也要刷（成功部分已 collect）
        return EXIT_BIZ_FAIL

    # 全部成功路径
    _try_flush_excel_master_safe()
    return EXIT_SUCCESS


def _try_flush_excel_master_safe() -> None:
    """daily_update.py 收尾的兜底 flush。

    场景：
        - daily_update.py 是 subprocess 调 main.run_business() 触发 collect
        - main.run_business 不会主动调 flush（避免 daily_update 里重复调）
        - 所以必须由 daily_update 自己负责收尾刷
    """
    try:
        import excel_master as _em
        from runtime_config import get_shop_id, get_shop_pin
        shop_id = get_shop_id()
        shop_pin = get_shop_pin()
        _em.flush_shop(shop_id, shop_pin)
    except Exception as e:
        # flush 失败仅警告（不影响主流程退出码）
        print(f"⚠️ [ExcelMaster] daily_update.flush_shop 失败（不影响 DB）：{type(e).__name__}: {e}")


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )
    sys.exit(main())