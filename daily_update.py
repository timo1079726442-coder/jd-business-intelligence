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

# ====================================================================
#  业务配置：MVP 先做 2 个
# ====================================================================

# 业务清单 + 是否支持区间
# MVP 阶段只做这 2 个；全量业务后续扩展
MVP_BIZ_KEYS = [
    "商智关键词分析",       # day 粒度：逐日；month 粒度：直接区间
    "店铺来源_三级渠道",     # 单日/区间都支持
]

# 业务特点（区分 supports_range）
BIZ_FEATURES = {
    "商智关键词分析": {
        # supports_range 标记是否支持直接 start_date/end_date 区间调用
        # day 粒度下接口其实只支持单日，30 天必须逐日循环；
        # month 粒度下接口支持 start_date/end_date 区间聚合
        # daily_update 默认用 day 粒度（避免误用 month 引起空数据），所以 supports_range=False
        "supports_range": False,
        "default_granularity": "day",
        "extra_kwargs": {},
    },
    "店铺来源_三级渠道": {
        # 店铺来源_三级渠道支持单日和区间（业务实现）
        "supports_range": True,
        "default_granularity": None,
        "extra_kwargs": {},
    },
    # 未来扩展：把项目1-20 业务都加上
    # "商品流量来源_搜索": {"supports_range": False, "default_granularity": None, "extra_kwargs": {}},
    # ...
}


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


def _run_one_biz(biz_key: str, start_date: str, end_date: str, interval: int = 0, dry_run: bool = False):
    """跑一个业务。

    策略：
        - supports_range=True：直接区间调用
        - supports_range=False：逐日循环
    """
    import main  # 延迟导入，避免循环引用
    import time

    feature = BIZ_FEATURES.get(biz_key, {})
    if not feature:
        logging.warning(f"⚠️ 业务 {biz_key} 无特性配置，跳过")
        return False

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
        return True
    except Exception as e:
        print(f"❌ {biz_key} 失败：{type(e).__name__}: {e}")
        traceback.print_exc()
        return False


def main():
    parser = argparse.ArgumentParser(description="京东数据报表每日自动更新（2026-08-14）")
    parser.add_argument(
        "--biz_keys",
        type=str,
        default=",".join(MVP_BIZ_KEYS),
        help=f"业务key列表（逗号分隔），默认 MVP 2 个：{','.join(MVP_BIZ_KEYS)}",
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
    for biz_key in biz_keys:
        ok = _run_one_biz(biz_key, start_date, end_date, args.interval, args.dry_run)
        if ok:
            success_count += 1
        else:
            failed_count += 1

    # 总结
    print()
    print("=" * 70)
    print(f"📊 总结：成功 {success_count} / 失败 {failed_count} / 总 {len(biz_keys)}")
    print("=" * 70)

    if failed_count == 0:
        return 0
    elif failed_count < len(biz_keys):
        return 1
    else:
        return 1


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )
    sys.exit(main())