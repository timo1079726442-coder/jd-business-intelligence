# -*- coding: utf-8 -*-
"""
run_recent_30d.py — 单店近30天/7天批量导出（项目26，2026-08-22）

执行计划：
  - 京准通 6 个业务 × 30 天（区间 start_date/end_date）
  - 京麦 2 个业务 × 30 天（区间）
  - 商智 7 个业务 × 7 天（关键词分析按 day 粒度逐日；其他支持区间则区间）

环境要求：
  - SHOP_PIN / SHOP_ID 必须已设置（脚本会自动从 config.xlsx 读取）
  - config/{SHOP_PIN}_*_cookie.json 必须存在
  - 京麦业务必须 config/{SHOP_PIN}_jm_*_h5st.json 已就绪

退出码：
  - 0: 全部成功
  - 1: 部分失败（汇总日志输出）
  - 2: 鉴权过期（人工重抓）

使用：
    set SHOP_PIN=FYA8888
    set SHOP_ID=FYA箱包旗舰店
    set AUTH_LOADER=1
    python run_recent_30d.py
"""
import os
import sys
import json
import time
import logging
from datetime import datetime, date, timedelta

# ──────────────────────────────── 配置区 ────────────────────────────────
PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
os.chdir(PROJECT_ROOT)
sys.path.insert(0, PROJECT_ROOT)

# M-11 修复（2026-08-24 审计）：业务天数从 config.xlsx「全局配置」sheet「全局/批次天数_xxx」读取
# 缺失时走下方 _get_window_days() 兜底（30/30/7）+ WARN


def _get_window_days(name: str, default: int) -> int:
    """读取业务窗口天数（2026-08-24 M-11 新增）。

    数据源: config.xlsx「全局配置」sheet → 项目名「全局」→ 变量名「批次天数_<name>」
    兜底: <default>（如 30 / 7）
    """
    try:
        import openpyxl
        wb = openpyxl.load_workbook("config/config.xlsx", read_only=True, data_only=True)
        if "全局配置" not in wb.sheetnames:
            wb.close()
            return default
        ws = wb["全局配置"]
        target_key = f"批次天数_{name}"
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row and len(row) >= 3 and str(row[0] or "").strip() == "全局" and str(row[1] or "").strip() == target_key:
                value = str(row[2] or "").strip()
                wb.close()
                if value:
                    return int(value)
                break
        wb.close()
    except Exception as e:
        print(f"[WARN] [M-11] 批次天数_{name} 读取失败：{type(e).__name__}: {e}")
    print(f"[WARN] [M-11] config.xlsx「全局/批次天数_{name}」未配置，使用兜底 {default}")
    return default


# 京准通 30 天区间（M-11：从 config 读取）
JZT_DAYS = _get_window_days("jzt", 30)
# 京麦 30 天区间（M-11：从 config 读取）
JM_DAYS = _get_window_days("jm", 30)
# 商智 7 天（M-11：从 config 读取）
SZ_DAYS = _get_window_days("sz", 7)
# 商智关键词分析逐日（day 粒度不支持区间）
SZ_KEYWORD_DAILY = True

# M-10 修复（2026-08-24 审计）：业务清单从 biz_config_loader 读取，替代硬编码
# 业务变更只需改 config.xlsx「业务清单」sheet，不用改 Python 代码
from biz_config_loader import list_biz_keys

# 京准通 6 业务（batch="jm_jzt" 全集 = 京准通 + 京麦；本节筛京准通 key 前缀）
_JZT_JM_KEYS = list_biz_keys(batch="jm_jzt", enabled_only=True)
JZT_BIZ = [k for k in _JZT_JM_KEYS if k.startswith("京准通")]
JM_BIZ = [k for k in _JZT_JM_KEYS if k.startswith("京麦")]
# 商智 7 业务
SZ_BIZ = list_biz_keys(batch="sz", enabled_only=True)
# 商智全部走逐日（handler 无参实例化只接受 date 单日，区间在 API 内部自动拆解）
SZ_DAILY_BIZ = set(SZ_BIZ)

# 配置一致性警告（若读到的业务数与预期不符）
if len(JZT_BIZ) == 0 or len(JM_BIZ) == 0 or len(SZ_BIZ) == 0:
    print(f"[WARN] [M-10] 业务清单为空：JZT={len(JZT_BIZ)}, JM={len(JM_BIZ)}, SZ={len(SZ_BIZ)}")
    print(f"       检查 config.xlsx「业务清单」sheet 的 enabled 列是否勾选")

LOG_DIR = os.path.join(PROJECT_ROOT, "logs")
os.makedirs(LOG_DIR, exist_ok=True)
LOG_FILE = os.path.join(LOG_DIR, f"run_recent_{datetime.now():%Y%m%d_%H%M%S}.log")

# ──────────────────────────────── 日志 ────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("run_recent_30d")


# ──────────────────────────────── 业务执行 ────────────────────────────────

def _today() -> date:
    return date.today()


def _date_range(n_days: int) -> tuple[str, str]:
    """生成最近 n 天的区间（不含今天，今天数据未生成）"""
    end = _today() - timedelta(days=1)
    start = end - timedelta(days=n_days - 1)
    return start.isoformat(), end.isoformat()


def _import_business_runner(biz_key: str):
    """从 main.py 获取业务 handler（优先 callable，否则标准 api_class.method 实例方法）"""
    import main
    return main.get_business_handler(biz_key), main


_AUTH_EXPIRED_TYPES = ("CookieExpiredError", "H5stExpiredError", "AuthFileNotFound")


def _classify_exception(e: Exception) -> str:
    """分类异常：auth_expired / biz_error / unknown"""
    name = type(e).__name__
    if name in _AUTH_EXPIRED_TYPES or "expired" in str(e).lower():
        return "auth_expired"
    return "biz_error"


def run_business_one_day(biz_key: str, day: str, h5st: str = "") -> tuple[bool, str]:
    """跑一天的单业务（day 粒度，逐日循环）"""
    handler, _ = _import_business_runner(biz_key)
    try:
        kwargs = {"date": day}
        if h5st:
            kwargs["h5st"] = h5st
        result = handler(**kwargs)
        log.info(f"✅ {biz_key} ({day}) → {result}")
        return True, str(result) if result else ""
    except SystemExit as e:
        # CookieExpiredError 等被业务类转成 sys.exit(2) 时
        if e.code == 2:
            log.error(f"🔐 {biz_key} ({day}) → 鉴权过期 (exit 2)")
            raise
        raise
    except Exception as e:
        kind = _classify_exception(e)
        if kind == "auth_expired":
            log.error(f"🔐 {biz_key} ({day}) → 鉴权过期: {e}")
            raise
        log.error(f"❌ {biz_key} ({day}) → {type(e).__name__}: {e}")
        return False, str(e)


def run_business_range(biz_key: str, start: str, end: str, h5st: str = "") -> tuple[bool, str]:
    """跑区间业务（start_date/end_date）"""
    handler, _ = _import_business_runner(biz_key)
    try:
        kwargs = {"start_date": start, "end_date": end}
        if h5st:
            kwargs["h5st"] = h5st
        result = handler(**kwargs)
        log.info(f"✅ {biz_key} ({start}~{end}) → {result}")
        return True, str(result) if result else ""
    except SystemExit as e:
        if e.code == 2:
            log.error(f"🔐 {biz_key} ({start}~{end}) → 鉴权过期 (exit 2)")
            raise
        raise
    except Exception as e:
        kind = _classify_exception(e)
        if kind == "auth_expired":
            log.error(f"🔐 {biz_key} ({start}~{end}) → 鉴权过期: {e}")
            raise
        log.error(f"❌ {biz_key} ({start}~{end}) → {type(e).__name__}: {e}")
        return False, str(e)


def run_business_daily(biz_key: str, days: int, h5st: str = "") -> tuple[int, int]:
    """逐日循环跑业务（关键词分析等 day 粒度业务）"""
    end_d = _today() - timedelta(days=1)
    start_d = end_d - timedelta(days=days - 1)
    ok = fail = 0
    d = start_d
    while d <= end_d:
        s = d.isoformat()
        is_ok, _ = run_business_one_day(biz_key, s, h5st)
        if is_ok:
            ok += 1
        else:
            fail += 1
        d += timedelta(days=1)
    return ok, fail


def run_jzt_range(days: int) -> tuple[int, int]:
    """京准通所有业务 × 区间"""
    start, end = _date_range(days)
    log.info(f"━━━ 京准通 6 业务 × {days} 天区间: {start} ~ {end} ━━━")
    ok = fail = 0
    for biz in JZT_BIZ:
        is_ok, _ = run_business_range(biz, start, end)
        if is_ok:
            ok += 1
        else:
            fail += 1
        time.sleep(2)  # 业务间间隔，避免触发限流
    return ok, fail


def run_jm_range(days: int) -> tuple[int, int]:
    """京麦 2 业务 × 区间（需要 h5st，自动从 AuthLoader 读）"""
    start, end = _date_range(days)
    log.info(f"━━━ 京麦 2 业务 × {days} 天区间: {start} ~ {end} ━━━")
    ok = fail = 0

    # 京麦订单 h5st key=jm_order
    h5st_order = ""
    try:
        from auth_loader import AuthLoader
        auth = AuthLoader()
        h5st_order = auth.get_h5st(check_expire=False, h5st_key="jm_order")
    except Exception as e:
        log.warning(f"⚠️ 读京麦订单 h5st 失败：{e}")

    h5st_shouhou = ""
    try:
        from auth_loader import AuthLoader
        auth = AuthLoader()
        h5st_shouhou = auth.get_h5st(check_expire=False, h5st_key="jm_after_sale")
    except Exception as e:
        log.warning(f"⚠️ 读京麦售后 h5st 失败：{e}")

    biz_h5st = {
        "京麦订单明细_完整一键导出": h5st_order,
        "京麦售后明细_完整一键导出": h5st_shouhou,
    }
    for biz in JM_BIZ:
        h5st = biz_h5st.get(biz, "")
        if not h5st:
            log.error(f"❌ {biz} 缺少 h5st，跳过")
            fail += 1
            continue
        is_ok, _ = run_business_range(biz, start, end, h5st=h5st)
        if is_ok:
            ok += 1
        else:
            fail += 1
        time.sleep(2)
    return ok, fail


def run_sz_mixed(days: int) -> tuple[int, int]:
    """商智 7 业务：关键词分析逐日，其他走区间"""
    start, end = _date_range(days)
    log.info(f"━━━ 商智 7 业务 × {days} 天: {start} ~ {end} ━━━")
    ok = fail = 0
    for biz in SZ_BIZ:
        if biz in SZ_DAILY_BIZ and SZ_KEYWORD_DAILY:
            log.info(f"  → {biz} 逐日模式（{days} 天）")
            d_ok, d_fail = run_business_daily(biz, days)
            ok += d_ok
            fail += d_fail
        else:
            is_ok, _ = run_business_range(biz, start, end)
            if is_ok:
                ok += 1
            else:
                fail += 1
        time.sleep(2)
    return ok, fail


# ──────────────────────────────── 主流程 ────────────────────────────────

def main():
    shop_pin = os.getenv("SHOP_PIN", "")
    shop_id = os.getenv("SHOP_ID", "")
    if not shop_pin or not shop_id:
        log.error("❌ 环境变量 SHOP_PIN / SHOP_ID 未设置")
        sys.exit(3)
    log.info(f"🏪 店铺: {shop_id} ({shop_pin})")
    log.info(f"📋 日志: {LOG_FILE}")

    total_ok = total_fail = 0

    # 1. 京准通 6 业务 × 30 天
    ok, fail = run_jzt_range(JZT_DAYS)
    total_ok += ok
    total_fail += fail

    # 2. 京麦 2 业务 × 30 天
    ok, fail = run_jm_range(JM_DAYS)
    total_ok += ok
    total_fail += fail

    # 3. 商智 7 业务 × 7 天
    ok, fail = run_sz_mixed(SZ_DAYS)
    total_ok += ok
    total_fail += fail

    # 汇总
    log.info(f"━━━ 汇总: 成功 {total_ok} / 失败 {total_fail} ━━━")
    if total_fail == 0:
        sys.exit(0)
    sys.exit(1)


if __name__ == "__main__":
    main()