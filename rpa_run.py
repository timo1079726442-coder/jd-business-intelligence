# -*- coding: utf-8 -*-
"""多店铺数据导出手动触发入口（项目22，2026-08-17）

用户决策 2026-08-17：
    - **手动触发**：本脚本是手动工具，不接影刀定时调度。你想跑就执行一次。
    - **不要凌晨自动跑**：鉴权过期时人工判断重抓更稳。
    - **3 个店一次跑完**：传 --shop 跑单店；不传则遍历 3 个店。

职责：
    1. 每个店按"近 30 天"拉一次：京麦订单 + 京麦售后 + 京准通全 6 业务
    2. 商智（搜索/购物车/推荐 + 4 个其他业务）走缺日补齐：
       - 先扫 output/{店名}/{业务}/{date}/ 是否存在 → 缺哪些日期
       - 只补缺的那些日期（避免重复请求触发风控）
    3. IMAP 收件统一用一份 config/imap_config.ini（QQ:1079726442@qq.com）
    4. 按退出码约定返回：
       0 = 全部成功
       1 = 部分业务失败
       2 = 鉴权过期（你需手动触发 RPA 重抓 cookie/h5st）
       3 = 参数错误 / 系统错误

调用示例：
    # 全部3 个店全跑：
    python rpa_run.py
    # 单店跑：
    python rpa_run.py --shop "FYA箱包旗舰店"
    # Dry-run（仅预览，不实际下载）：
    python rpa_run.py --dry-run
"""
import os
import sys
import argparse
import traceback
from datetime import datetime, timedelta

# 必须在 import main 之前设环境变量（启用 AuthLoader 接管 cookie）
if "AUTH_LOADER" not in os.environ:
    os.environ["AUTH_LOADER"] = "1"
if "AUTH_RPA_CLI" not in os.environ:
    os.environ["AUTH_RPA_CLI"] = "0"  # 手动触发模式：不调 RPA
# Cookie 过期检查跳过（AGENTS.md 用户决策：由接口返回码判定）
os.environ["AUTH_SKIP_COOKIE_EXPIRE_CHECK"] = "1"

# ============================ 退出码 ============================
EXIT_SUCCESS = 0
EXIT_BIZ_FAIL = 1
EXIT_AUTH_EXPIRED = 2
EXIT_SYSTEM_ERROR = 3

# ============================ 店铺清单 ============================
# 与 auth_loader.py 的 SHOP_ID_TO_PREFIX 保持一致
KNOWN_SHOPS = [
    "FYA箱包旗舰店",
    "MIYO箱包旗舰店",
    "OTA箱包旗舰店",
]

# ============================ 业务清单 ============================
# 京麦 + 京准通（按"近 30 天"全量覆盖，无需补齐）
JM_JZT_BIZ_KEYS = [
    # === 京麦 2 个（h5st 必需，订单完整链路需 IMAP）===
    "京麦订单明细_完整一键导出",      # 项目14，含 IMAP 短信密码链路
    "京麦售后明细_完整一键导出",      # 项目16，4 步一键无 IMAP

    # === 京准通全 6 个业务（不需要 h5st，2026-08-15 实测确认）===
    "京准通快车自定义报表",            # 项目7，异步 3 步
    "京准通快车订单效果明细",          # 项目8，同步 2 步
    "京准通全站营销单品计划",          # 项目9，同步 2 步
    "京准通全站营销单品推广效果",      # 项目10，同步 2 步 + zip/csv 智能识别
    "京准通全站营销全店计划",          # 项目11，靠 campaignTypes=[118] 区分
    "京准通全站营销全店推广效果",      # 项目12，靠 campaignTypes=[118] 区分
]

# 商智业务（按"缺日补齐"，先扫再下）
# ⚠️ 2026-08-17 用户决策：商智全量 7 个业务都跑
# - 商品流量来源_自主访问 已在 BUSINESS_REGISTRY 里 enabled=False，调度层自动跳过
# - 商智关键词分析 用 day 粒度跑日表（与近 30 天窗口对齐）
商智_BIZ_KEYS = [
    # 渠道 3 个（最常用，区间自动逐日拆分）
    "商品流量来源_搜索",
    "商品流量来源_购物车",
    "商品流量来源_推荐",

    # 其他商智报表 4 个
    "店铺来源_三级渠道",        # 离线流量报表（项目4）
    "商品明细导出",              # GET 同步 xlsx（项目5）
    "商品流失分析",              # xls → xlsx 转存（项目6）
    "商智关键词分析",            # day/month 双粒度（项目13）
]


def compute_last_30_days():
    """计算「今天-30天 ~ 今天-1天」日期区间（AGENTS.md 每日定时约定）

    返回:
        (start_date, end_date) - "YYYY-MM-DD" 格式字符串
    """
    today = datetime.now()
    end = today - timedelta(days=1)
    start = today - timedelta(days=30)
    return start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")


def list_existing_dates(output_root: str, shop_id: str, biz_key: str) -> set:
    """扫描 output/{店名}/{业务名}/ 下的日期子目录，返回已有日期集合

    参数:
        output_root - 项目根/output
        shop_id     - 店铺名
        biz_key - 业务 key（用来定位子目录名）

    返回:
        set - 已存在的日期集合 {"2026-07-18", "2026-07-19", ...}
    """
    # 业务 key → 输出目录名映射（按项目惯例，覆盖全部业务）
    biz_to_dirname = {
        # === 京麦 2 个 ===
        "京麦订单明细_完整一键导出":  "京麦订单明细",
        "京麦售后明细_完整一键导出":  "京麦售后明细",

        # === 京准通 6 个 ===
        "京准通快车自定义报表":        "京准通快车效果自定义",
        "京准通快车订单效果明细":      "京准通快车订单效果明细",
        "京准通全站营销单品计划":      "京准通全站营销单品计划",
        "京准通全站营销单品推广效果":  "京准通全站营销单品推广效果",
        "京准通全站营销全店计划":      "京准通全站营销全店计划",
        "京准通全站营销全店推广效果":  "京准通全站营销全店推广效果",

        # === 商智 7 个 ===
        "商品流量来源_搜索":          "搜索流量",
        "商品流量来源_购物车":        "购物车流量",
        "商品流量来源_推荐":          "推荐流量",
        "商品流量来源_自主访问":      "自主访问流量",   # 停用业务，无目录
        "店铺来源_三级渠道":          "店铺来源_三级渠道",
        "商品明细导出":                "商品明细",
        "商品流失分析":                "商品流失分析",
        "商智关键词分析":              "商智关键词分析",
    }
    dirname = biz_to_dirname.get(biz_key)
    if not dirname:
        return set()

    biz_dir = os.path.join(output_root, shop_id, dirname)
    if not os.path.isdir(biz_dir):
        return set()

    existing = set()
    for entry in os.listdir(biz_dir):
        full = os.path.join(biz_dir, entry)
        if os.path.isdir(full) and len(entry) == 10 and entry[4] == "-":
            existing.add(entry)
    return existing


def compute_missing_dates(start_date: str, end_date: str, existing: set) -> list:
    """计算 [start_date, end_date] 区间里缺失的日期列表（按时间正序）

    返回:
        list - 缺失日期字符串列表 ["2026-08-15", "2026-08-16", ...]
    """
    start_dt = datetime.strptime(start_date, "%Y-%m-%d")
    end_dt = datetime.strptime(end_date, "%Y-%m-%d")
    missing = []
    cur = start_dt
    while cur <= end_dt:
        s = cur.strftime("%Y-%m-%d")
        if s not in existing:
            missing.append(s)
        cur += timedelta(days=1)
    return missing


def run_shop(shop_id: str, dry_run: bool = False) -> int:
    """跑单个店的所有业务

    返回:
        int - 该店退出码（0=全成功 / 1=部分失败 / 2=鉴权过期）
    """
    print(f"\n{'='*70}")
    print(f"[SHOP] {shop_id} 开始执行")
    print(f"{'='*70}")

    os.environ["SHOP_ID"] = shop_id  # 确保 AuthLoader 读到正确店

    # 延迟 import（环境变量设完后再 import main）
    try:
        from main import run_business
        from auth_loader import CookieExpiredError, H5stExpiredError
    except ImportError as e:
        print(f"[ERR] import main 失败：{e}")
        return EXIT_SYSTEM_ERROR

    start_date, end_date = compute_last_30_days()
    print(f"[DATE] 近30 天区间: {start_date} ~ {end_date}")

    overall_exit = EXIT_SUCCESS
    has_auth_expired = False

    # === 第 1 批：京麦 + 京准通（按"近 30 天"全量跑，覆盖原日期）===
    print(f"\n--- [BATCH 1/2] 京麦 + 京准通：近 30 天全量覆盖 ---")
    for biz_key in JM_JZT_BIZ_KEYS:
        if dry_run:
            print(f"[DRY-RUN] {biz_key} | {start_date} ~ {end_date}")
            continue
        print(f"\n>>> {biz_key} ({start_date} ~ {end_date})")
        try:
            run_business(biz_key, start_date=start_date, end_date=end_date)
            print(f"[OK] {biz_key}")
        except (CookieExpiredError, H5stExpiredError) as e:
            print(f"[AUTH_EXPIRED] {biz_key}: {e}")
            has_auth_expired = True
        except Exception as e:
            print(f"[BIZ_FAIL] {biz_key}: {e}")
            overall_exit = EXIT_BIZ_FAIL

    # === 第 2 批：商智 7 业务（按"缺日补齐"）===
    print(f"\n--- [BATCH 2/2] 商智 7 业务：缺日补齐 ---")
    output_root = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")

    for biz_key in 商智_BIZ_KEYS:
        existing = list_existing_dates(output_root, shop_id, biz_key)
        missing = compute_missing_dates(start_date, end_date, existing)
        print(f"\n>>> {biz_key} | 已存在 {len(existing)} 天 | 缺 {len(missing)} 天: {missing[:5]}{'...' if len(missing) > 5 else ''}")
        if not missing:
            print(f"[SKIP] {biz_key} 已覆盖，无需补齐")
            continue
        if dry_run:
            print(f"[DRY-RUN] {biz_key} | 需补 {len(missing)} 天")
            continue

        # 逐日循环补齐（商智服务端不支持多日区间导出，必须逐日调）
        for date_str in missing:
            print(f"\n -> {date_str}")
            try:
                # ⚠️ 商智关键词分析需传 granularity=day（与近 30 天窗口对齐）
                if biz_key == "商智关键词分析":
                    run_business(biz_key, date=date_str, granularity="day")
                else:
                    run_business(biz_key, date=date_str)
                print(f"  [OK] {biz_key} | {date_str}")
            except (CookieExpiredError, H5stExpiredError) as e:
                print(f"  [AUTH_EXPIRED] {biz_key} | {date_str}: {e}")
                has_auth_expired = True
                break  # 鉴权过期 → 该业务后续日期不再尝试
            except Exception as e:
                print(f"  [BIZ_FAIL] {biz_key} | {date_str}: {e}")
                overall_exit = EXIT_BIZ_FAIL

    # 鉴权过期优先于业务失败返回
    if has_auth_expired:
        return EXIT_AUTH_EXPIRED
    return overall_exit


def main():
    """rpa_run.py 主入口"""
    parser = argparse.ArgumentParser(
        prog="rpa_run",
        description="多店铺数据导入手动触发入口（项目22，2026-08-17）",
    )
    parser.add_argument("--shop", default=None, help="指定单个店铺；不传则遍历全部 3 个店")
    parser.add_argument("--dry-run", action="store_true", help="仅预览计划，不实际下载/入库")
    args = parser.parse_args()

    # 决定本次跑哪些店
    shops = [args.shop] if args.shop else KNOWN_SHOPS
    print(f"[INIT] 准备跑 {len(shops)} 个店: {shops}")
    print(f"[INIT] DRY-RUN 模式: {args.dry_run}")

    if args.dry_run:
        print(f"\n{'='*70}\n[DRY-RUN 模式] 仅预览计划，不实际下载\n{'='*70}")

    # 累计所有店的退出码
    shop_results = {}
    for shop_id in shops:
        if shop_id not in KNOWN_SHOPS:
            print(f"[WARN] 未知店铺 {shop_id}（应在 KNOWN_SHOPS 中），跳过")
            shop_results[shop_id] = EXIT_SYSTEM_ERROR
            continue
        try:
            shop_results[shop_id] = run_shop(shop_id, dry_run=args.dry_run)
        except Exception as e:
            print(f"[SYSTEM_ERROR] {shop_id}: {e}")
            traceback.print_exc()
            shop_results[shop_id] = EXIT_SYSTEM_ERROR

    # 输出汇总
    print(f"\n{'='*70}")
    print(f"[SUMMARY] {len(shops)} 个店执行结果")
    print(f"{'='*70}")
    for shop_id, code in shop_results.items():
        tag = {
            0: "✅ 成功",
            1: "⚠️ 部分失败",
            2: "🔄 鉴权过期（需RPA重抓）",
            3: "❌ 系统错误",
        }.get(code, f"?({code})")
        print(f"  {shop_id}: {tag}")

    # 决定最终退出码（任一店鉴权过期 → 整体返 2，触发影刀重抓）
    if any(c == EXIT_AUTH_EXPIRED for c in shop_results.values()):
        return EXIT_AUTH_EXPIRED
    if any(c == EXIT_SYSTEM_ERROR for c in shop_results.values()):
        return EXIT_SYSTEM_ERROR
    if any(c == EXIT_BIZ_FAIL for c in shop_results.values()):
        return EXIT_BIZ_FAIL
    return EXIT_SUCCESS


if __name__ == "__main__":
    sys.exit(main())