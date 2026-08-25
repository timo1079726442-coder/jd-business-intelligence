# -*- coding: utf-8 -*-
"""
fill_missing.py
============================================================
京东数据报表 - 缺失日期巡查与补录脚本（2026-08-14 上线骨架）。

用途：
    - 扫描数据库中各业务已有 report_date
    - 与预期日期列表比对（默认最近 30 天，可指定更早起始日）
    - 自动触发缺失日期的下载补录
    - 特别适用于周末/节假日停跑后周一补齐

与 daily_update.py 的区别：
    - daily_update.py   → 全量覆盖（30 天都重跑）
    - fill_missing.py   → 只补缺失日期（已有的不重复下载）

CLI 用法：
    # 默认扫描最近 30 天缺失
    python fill_missing.py

    # 指定业务
    python fill_missing.py --biz_keys "商智关键词分析,店铺来源_三级渠道"

    # 扫描更早历史（如项目上线前的基础数据）
    python fill_missing.py --start_date 2026-06-01

    # 仅显示缺失日期，不下载
    python fill_missing.py --dry-run

    # 仅补关键词分析的 day 粒度
    python fill_missing.py --biz_keys "商智关键词分析" --granularity day

退出码：
    0 - 全部完成（含无缺失）
    1 - 部分失败
    2 - 鉴权过期
    3 - 系统错误
"""
import sys
import os
import argparse
import sqlite3
import logging
import traceback
from datetime import datetime, timedelta

# 项目根目录加入 sys.path
PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, PROJECT_ROOT)

from db_utils import (
    get_db_path,
    biz_key_to_table_name,
    get_existing_dates,
)
# Phase 2.4（2026-08-20）：业务清单从 config.xlsx「业务清单」sheet 读取
# 替代原硬编码 MVP_BIZ_KEYS（AGENTS.md 第3条禁止硬编码业务列表）
from biz_config_loader import list_biz_keys


# 退出码（与 AGENTS.md 第47-51条约定对齐，与 daily_update.py 一致）
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


def _find_missing_dates(biz_key: str, expected_dates: list, granularity=None):
    """查 DB 中已有日期，返回 expected 中**不在 DB**的日期集合。

    参数:
        biz_key        业务名
        expected_dates 预期应有日期列表
        granularity    粒度（None / "day" / "month"）
    返回:
        list - 缺失日期（按日期升序）
    """
    db_path = get_db_path()
    if not os.path.exists(db_path):
        return expected_dates  # DB 不存在 → 全缺失

    table_name = biz_key_to_table_name(biz_key)
    # M-03 修复（2026-08-24 审计）：按当前店铺 shop_pin 过滤已有日期
    # 背景：跨店聚合后，MIYO 已有的日期会被认为 FYA 也有 → FYA 永远不会补录
    shop_pin = os.getenv("SHOP_PIN", "").strip()
    conn = sqlite3.connect(db_path, timeout=30)
    try:
        existing = get_existing_dates(conn, table_name, granularity, shop_pin=shop_pin or None)
    except sqlite3.OperationalError as e:
        # 表不存在 → 全缺失
        logging.debug(f"表 {table_name} 不存在：{e}")
        existing = set()
    finally:
        conn.close()

    return [d for d in expected_dates if d not in existing]


def _fill_one_biz(biz_key: str, missing_dates: list, granularity=None, interval: int = 0, dry_run: bool = False) -> int:
    """补录一个业务的缺失日期，返回退出码（0=成功 / 1=业务失败 / 2=鉴权过期）。

    鉴权过期时立即 break 后续日期（避免无意义重复触发风控）。
    """
    import main

    if not missing_dates:
        print(f"   ✅ {biz_key} 无缺失")
        return EXIT_SUCCESS

    print(f"\n   📥 {biz_key} 缺失 {len(missing_dates)} 天：{missing_dates}")

    success = 0
    failed = 0
    auth_expired = False  # 鉴权过期标志（用于优先返回 2）
    for i, d in enumerate(missing_dates, 1):
        print(f"   [{i}/{len(missing_dates)}] 补录 {d} ...", end=" ")
        try:
            kwargs = {"date": d}
            if granularity:
                kwargs["granularity"] = granularity

            if dry_run:
                print("[DRY-RUN]")
            else:
                main.run_business(biz_key, **kwargs)
                print("✅")
                success += 1
        except Exception as e:
            # 鉴权过期优先识别（按类名匹配，兼容 main.py 与 auth_loader.py 两套定义）
            if _is_auth_expired_exception(e):
                print(f"🔄 [鉴权过期] {e}")
                auth_expired = True
                break  # 鉴权过期 → 后续日期不再尝试，避免重复触发风控
            print(f"❌ {e}")
            failed += 1

        if interval > 0 and i < len(missing_dates):
            import time
            time.sleep(interval)

    print(f"   📊 {biz_key}：成功 {success} / 失败 {failed} / 总 {len(missing_dates)}")
    # 退出码优先级：鉴权过期(2) > 业务失败(1) > 成功(0)
    if auth_expired:
        return EXIT_AUTH_EXPIRED
    if failed > 0:
        return EXIT_BIZ_FAIL
    return EXIT_SUCCESS


def main():
    parser = argparse.ArgumentParser(description="京东数据报表 - 缺失日期巡查与补录（2026-08-14）")
    parser.add_argument(
        "--biz_keys",
        type=str,
        default=",".join(list_biz_keys()),
        help=f"业务key列表（逗号分隔），默认全部启用业务",
    )
    parser.add_argument(
        "--start_date",
        type=str,
        default=None,
        help="起始日期 YYYY-MM-DD（默认：今天-30天）",
    )
    parser.add_argument(
        "--end_date",
        type=str,
        default=None,
        help="结束日期 YYYY-MM-DD（默认：今天-1天）",
    )
    parser.add_argument(
        "--granularity",
        type=str,
        choices=["day", "month"],
        default=None,
        help="粒度筛选（仅对支持粒度的业务生效，如关键词分析）",
    )
    parser.add_argument(
        "--interval",
        type=int,
        default=0,
        help="逐日循环时的间隔秒数（默认 0）",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="仅扫描缺失日期，不下载",
    )
    args = parser.parse_args()

    today = datetime.now()
    end_date = args.end_date or (today - timedelta(days=1)).strftime("%Y-%m-%d")
    start_date = args.start_date or (today - timedelta(days=30)).strftime("%Y-%m-%d")
    expected_dates = _gen_date_range(start_date, end_date)

    biz_keys = [b.strip() for b in args.biz_keys.split(",") if b.strip()]

    # 启动 banner
    print()
    print("=" * 70)
    print(f"🔍 京东数据报表 - 缺失日期巡查")
    print(f"   时间：{today.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"   扫描窗口：{start_date} ~ {end_date}（共 {len(expected_dates)} 天）")
    print(f"   业务清单：{biz_keys}")
    print(f"   粒度：{args.granularity or 'N/A'}")
    print(f"   干跑：{args.dry_run}")
    print("=" * 70)

    # 查每个业务的缺失日期
    missing_map = {}
    for biz_key in biz_keys:
        missing = _find_missing_dates(biz_key, expected_dates, args.granularity)
        missing_map[biz_key] = missing
        print(f"\n📋 {biz_key}：缺失 {len(missing)} / 共 {len(expected_dates)} 天")
        if missing and len(missing) <= 10:
            for d in missing:
                print(f"   - {d}")
        elif missing:
            for d in missing[:5]:
                print(f"   - {d}")
            print(f"   ... 还有 {len(missing) - 5} 天")

    total_missing = sum(len(m) for m in missing_map.values())
    print()
    print(f"📊 总缺失：{total_missing} 条业务-日期")

    if total_missing == 0:
        print("🎉 无缺失，无需补录")
        return 0

    if args.dry_run:
        print()
        print("[DRY-RUN] 仅扫描，不实际补录")
        return 0

    # 补录
    success_count = 0
    failed_count = 0
    auth_expired_count = 0  # 鉴权过期计数（独立于业务失败）
    for biz_key in biz_keys:
        code = _fill_one_biz(biz_key, missing_map[biz_key], args.granularity, args.interval, args.dry_run)
        if code == EXIT_SUCCESS:
            success_count += 1
        elif code == EXIT_AUTH_EXPIRED:
            auth_expired_count += 1
        else:
            failed_count += 1

    # 总结
    print()
    print("=" * 70)
    print(f"📊 补录总结：业务成功 {success_count} / 失败 {failed_count} / 鉴权过期 {auth_expired_count} / 总 {len(biz_keys)}")
    print("=" * 70)

    # 退出码优先级：鉴权过期(2) > 业务失败(1) > 全部成功(0)
    # AGENTS.md 第49条：鉴权过期触发影刀重抓，必须优先返回 2
    if auth_expired_count > 0:
        _try_flush_excel_master_safe()  # 项目24（2026-08-22）
        return EXIT_AUTH_EXPIRED
    if failed_count > 0:
        _try_flush_excel_master_safe()  # 项目24（2026-08-22）
        return EXIT_BIZ_FAIL

    # 全部成功路径
    _try_flush_excel_master_safe()
    return EXIT_SUCCESS


def _try_flush_excel_master_safe() -> None:
    """fill_missing.py 收尾的兜底 flush（同 daily_update）。"""
    try:
        import excel_master as _em
        from runtime_config import get_shop_id, get_shop_pin
        shop_id = get_shop_id()
        shop_pin = get_shop_pin()
        _em.flush_shop(shop_id, shop_pin)
    except Exception as e:
        print(f"⚠️ [ExcelMaster] fill_missing.flush_shop 失败（不影响 DB）：{type(e).__name__}: {e}")


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )
    sys.exit(main())