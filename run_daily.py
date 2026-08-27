# -*- coding: utf-8 -*-
"""run_daily.py — 每日调度器（2026-08-26 项目27 新增）

功能：
    读 config.xlsx「业务下载配置」sheet → 按配置驱动每个报表下载 → DB → 总表全自动。

覆盖策略:
    - 月度: 拉配置整段 → DB upsert（整月覆盖）
    - 30天: 拉近 30 天 → DB upsert（订单状态最新）
    - 单日: 查 DB 已有日期 → 缺失日期单日补 → 入 DB

每个业务通过 main.run_business 统一入口调用，业务类内部：
    - df.to_excel → 删除（不落盘单表）
    - 显式 db_utils.save_to_db → DB
    - save_to_db 内部已联动 excel_master.collect → 总表

收尾: 调 excel_master.rebuild_from_db 重建三域总表。

退出码:
    0 - 全部成功
    1 - 部分失败
    2 - 鉴权过期（人工重抓）
    3 - 系统错误

使用:
    set SHOP_PIN=miyo-周
    set SHOP_ID=MIYO箱包旗舰店
    set AUTH_LOADER=1
    python run_daily.py

    # 单店指定
    python run_daily.py --shop MIYO箱包旗舰店

    # 干跑（不真的调业务，只读配置 + 打印计划）
    python run_daily.py --dry-run
"""
import os
import sys
import time
import argparse
import logging
from datetime import datetime, date, timedelta
from typing import Dict, List, Set, Tuple

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
os.chdir(PROJECT_ROOT)
sys.path.insert(0, PROJECT_ROOT)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger("run_daily")

# 退出码约定（与 fill_missing.py 一致）
EXIT_SUCCESS = 0
EXIT_BIZ_FAIL = 1
EXIT_AUTH_EXPIRED = 2
EXIT_SYSTEM_ERROR = 3

# 鉴权过期异常类名（main.py 与 auth_loader.py 都定义了）
_AUTH_EXPIRED_NAMES = ("CookieExpiredError", "H5stExpiredError", "AuthFileNotFound")


def _is_auth_expired(e: Exception) -> bool:
    """判定异常是否为鉴权过期"""
    name = type(e).__name__
    if name in _AUTH_EXPIRED_NAMES or "expired" in str(e).lower():
        return True
    return False


def _expected_dates(start: str, end: str) -> List[str]:
    """生成 [start, end] 区间所有日期（YYYY-MM-DD）"""
    from datetime import datetime as _dt, timedelta as _td
    s = _dt.strptime(start, "%Y-%m-%d").date()
    e = _dt.strptime(end, "%Y-%m-%d").date()
    out = []
    cur = s
    while cur <= e:
        out.append(cur.strftime("%Y-%m-%d"))
        cur += _td(days=1)
    return out


def _find_missing_dates(biz_key: str, expected_dates: List[str], shop_pin: str) -> List[str]:
    """查 DB 已有日期 vs 期望日期，返回缺失列表（仅单日策略用）"""
    import sqlite3
    from db_utils import get_db_path, biz_key_to_table_name

    db_path = get_db_path()
    if not os.path.exists(db_path):
        return expected_dates  # DB 不存在，全部缺失

    table = biz_key_to_table_name(biz_key)
    conn = sqlite3.connect(db_path)
    try:
        try:
            rows = conn.execute(
                f"SELECT DISTINCT report_date FROM \"{table}\" WHERE shop_pin=?",
                (shop_pin,),
            ).fetchall()
        except sqlite3.OperationalError as e:
            if "no such table" in str(e).lower():
                return expected_dates  # 表不存在，全部缺失
            raise
        existing = {r[0] for r in rows}
    finally:
        conn.close()
    return [d for d in expected_dates if d not in existing]


def _get_existing_dates(biz_key: str, shop_pin: str) -> set:
    """查 DB 该业务该店已有 report_date 集合（增量缺日扫描用）"""
    import sqlite3
    from db_utils import get_db_path, biz_key_to_table_name
    db_path = get_db_path()
    if not os.path.exists(db_path):
        return set()
    table = biz_key_to_table_name(biz_key)
    conn = sqlite3.connect(db_path)
    try:
        try:
            rows = conn.execute(
                f"SELECT DISTINCT report_date FROM \"{table}\" WHERE shop_pin=?",
                (shop_pin,),
            ).fetchall()
        except sqlite3.OperationalError as e:
            if "no such table" in str(e).lower():
                return set()
            raise
        return {r[0] for r in rows}
    finally:
        conn.close()


def _scan_missing_dates(biz_key: str, cfg, shop_pin: str) -> List[str]:
    """扫描缺失日期（2026-08-26 增量模式，用户需求）。

    规则：
        - 已有数据：缺口 = 已有最大日期+1 ~ 今天-1（查缺补日，往前推进）
        - 无数据（首次）：缺口 = config 配置的 start_date ~ end_date

    背景：今天=8-26，总表最新到 8-23 → 补 8-24、8-25，如此逐日推进。
    """
    today = datetime.now().date()
    yesterday = today - timedelta(days=1)

    existing = _get_existing_dates(biz_key, shop_pin)
    parsed = set()
    for d in existing:
        try:
            parsed.add(datetime.strptime(d, "%Y-%m-%d").date())
        except ValueError:
            pass

    if parsed:
        latest = max(parsed)
        if latest >= yesterday:
            return []  # 已最新，无缺口
        out = []
        cur = latest + timedelta(days=1)
        while cur <= yesterday:
            out.append(cur.strftime("%Y-%m-%d"))
            cur += timedelta(days=1)
        return out

    # 首次无数据 → 用配置起止（京准通/商智首次会整段拉）
    s = datetime.strptime(cfg.start_date, "%Y-%m-%d").date() if cfg.start_date else yesterday
    e = datetime.strptime(cfg.end_date, "%Y-%m-%d").date() if cfg.end_date else yesterday
    out = []
    cur = s
    while cur <= e:
        out.append(cur.strftime("%Y-%m-%d"))
        cur += timedelta(days=1)
    return out


def _get_jm_h5st(biz_key: str) -> str:
    """京麦业务需要的 h5st（30 分钟过期，从 AuthLoader 读）。

    京麦订单明细 → jm_order；京麦售后明细 → jm_after_sale。
    h5st 过期时返回空字符串（由调用方抛鉴权错误）。
    """
    try:
        from auth_loader import AuthLoader
        h5st_key = "jm_order" if "订单" in biz_key else "jm_after_sale"
        auth = AuthLoader()
        return auth.get_h5st(check_expire=False, h5st_key=h5st_key)
    except Exception as e:
        log.warning(f"⚠️ 读京麦 h5st 失败（{biz_key}）：{type(e).__name__}: {e}")
        return ""


def _backfill_xlsx_file(biz_key: str, xlsx_path: str, shop_id: str) -> bool:
    """对单个 xlsx 文件按文件内日期列拆分入库（京准通/京麦，复用 backfill 逻辑）。

    背景（2026-08-26）：京麦订单/售后 handler 返回 dict（含 xlsx_path），
    且 xlsx 是近 30 天合并（文件名日期 ≠ 内容日期范围），
    必须按文件内日期列拆成 N 组逐日入库（upsert 幂等）。
    不能用 _auto_save_db_from_xlsx（单 report_date 会把 30 天数据全塞到一天）。

    返回:
        True 已入库 / False 未处理（biz_key 不在京准通/京麦映射）
    """
    if not os.path.exists(xlsx_path):
        log.warning(f"  ⚠️ xlsx 不存在：{xlsx_path}")
        return False
    import pandas as pd
    from db_utils import save_to_db
    from backfill_master import JZT_BIZ, JM_BIZ

    # 找 biz_key 配置（JZT 4 元组 / JM 5 元组）
    date_col = None
    header = 0
    found = False
    for cfg in JZT_BIZ:
        if cfg[0] == biz_key:
            date_col, header = cfg[3], 0
            found = True
            break
    if not found:
        for cfg in JM_BIZ:
            if cfg[0] == biz_key:
                date_col, header = cfg[3], cfg[4]
                found = True
                break
    if not found:
        log.warning(f"  ⚠️ {biz_key} 不在京准通/京麦配置，跳过 xlsx 拆分入库")
        return False

    try:
        df = pd.read_excel(xlsx_path, header=header, dtype=str, na_filter=False)
    except Exception as e:
        log.error(f"  ❌ 读 xlsx 失败 {os.path.basename(xlsx_path)}：{type(e).__name__}: {e}")
        return False
    if df is None or df.empty:
        log.info(f"    ⚠️ xlsx 为空，跳过拆分入库：{os.path.basename(xlsx_path)}")
        return False

    # 按日期列拆分入库
    if date_col not in df.columns:
        log.warning(f"    ⚠️ 无 {date_col} 列，按文件名日期单次入库")
        save_to_db(biz_key=biz_key, df=df, report_date=datetime.now().strftime("%Y-%m-%d"))
        return True

    s = pd.to_datetime(df[date_col], errors="coerce")
    valid = s.notna()
    if valid.sum() == 0:
        save_to_db(biz_key=biz_key, df=df, report_date=datetime.now().strftime("%Y-%m-%d"))
        return True

    n = 0
    for dt, g in df[valid].groupby(s[valid].dt.date):
        if pd.isna(dt):
            continue
        save_to_db(biz_key=biz_key, df=g.reset_index(drop=True), report_date=dt.isoformat())
        n += 1
    log.info(f"    ✅ {biz_key} 拆分入库 {n} 天（{os.path.basename(xlsx_path)}）")
    return True


def _extract_and_backfill(cfg, result, shop_id: str) -> None:
    """从 handler 返回值提取 xlsx 路径并拆分入库。

    handler 返回类型（2026-08-26 混合模式 A）：
        - pd.DataFrame（商智改造类）→ 内部已 save_to_db，无需处理
        - str xlsx 路径（京准通 run_full_export 返回）→ 拆分入库
        - dict 含 xlsx_path（京麦 run_full_export 返回）→ 拆分入库
    """
    xlsx_path = None
    if isinstance(result, str) and result.lower().endswith(".xlsx"):
        xlsx_path = result
    elif isinstance(result, dict) and result.get("xlsx_path"):
        xlsx_path = result["xlsx_path"]
    if xlsx_path:
        _backfill_xlsx_file(cfg.biz_key, xlsx_path, shop_id)
    else:
        log.info(f"    ✅ {cfg.biz_key} → df returned (save_to_db 内已自动入库)")


def _run_single_config(cfg, shop_id: str, shop_pin: str, dry_run: bool = False) -> bool:
    """跑单个报表（单店）

    返回:
        True = 成功, False = 失败（鉴权过期立即抛 SystemExit(2)）
    """
    import main

    handler = main.get_business_handler(cfg.biz_key)
    if handler is None:
        log.error(f"❌ [{cfg.module}/{cfg.report_name}] 未注册业务 handler：{cfg.biz_key}")
        return False

    # 京麦业务需要 h5st（30 分钟过期，自动从 AuthLoader 读并透传）
    jm_h5st = _get_jm_h5st(cfg.biz_key) if cfg.module == "京麦" else ""
    if cfg.module == "京麦" and not jm_h5st:
        log.error(f"❌ [{cfg.module}/{cfg.report_name}] h5st 读取失败/过期，跳过")
        return False

    # 策略分流
    if cfg.strategy == "单日":
        # 1. 扫描缺失日期（2026-08-26 增量模式：最新日期+1 ~ 今天-1）
        missing = _scan_missing_dates(cfg.biz_key, cfg, shop_pin)
        if not missing:
            log.info(f"  ✅ [{cfg.module}/{cfg.report_name}] 无缺失日期（已最新），跳过")
            return True
        log.info(f"  🔄 [{cfg.module}/{cfg.report_name}] 扫描到缺失 {len(missing)} 天：{missing[0]} ~ {missing[-1]}")
        # 2. 逐日补录
        for d in missing:
            log.info(f"    ↳ 补录 {d}")
            if dry_run:
                log.info(f"    [DRY-RUN] 跳过实际调用")
                continue
            try:
                extra = {'h5st': jm_h5st} if cfg.module=='京麦' else {}; result = handler(date=d, **extra)
                # 混合模式：handler 可能返回 df（商智）、xlsx 路径（京准通）、dict(xlsx_path)（京麦）
                _extract_and_backfill(cfg, result, shop_id)
            except SystemExit as e:
                if e.code == 2:
                    log.error(f"  🔐 [{cfg.module}/{cfg.report_name}] 鉴权过期")
                    raise
                raise
            except Exception as e:
                if _is_auth_expired(e):
                    log.error(f"  🔐 [{cfg.module}/{cfg.report_name}] 鉴权过期：{e}")
                    raise SystemExit(2)
                log.error(f"  ❌ [{cfg.module}/{cfg.report_name}] {d} 失败：{type(e).__name__}: {e}")
                return False
        return True

    elif cfg.strategy in ("近3天", "近7天", "近30天", "30天"):
        s, e = cfg.resolve_dates()
        log.info(f"  🔄 [{cfg.module}/{cfg.report_name}] {cfg.strategy}覆盖 {s} ~ {e}")
        if dry_run:
            log.info(f"    [DRY-RUN] 跳过实际调用")
            return True
        try:
            extra = {'h5st': jm_h5st} if cfg.module=='京麦' else {}; result = handler(start_date=s, end_date=e, **extra)
            _extract_and_backfill(cfg, result, shop_id)
            return True
        except SystemExit as ex:
            if ex.code == 2:
                log.error(f"  🔐 [{cfg.module}/{cfg.report_name}] 鉴权过期")
                raise
            raise
        except Exception as ex:
            if _is_auth_expired(ex):
                log.error(f"  🔐 [{cfg.module}/{cfg.report_name}] 鉴权过期：{ex}")
                raise SystemExit(2)
            log.error(f"  ❌ [{cfg.module}/{cfg.report_name}] 失败：{type(ex).__name__}: {ex}")
            return False

    elif cfg.strategy == "月度":
        log.info(f"  🔄 [{cfg.module}/{cfg.report_name}] 月度覆盖 {cfg.start_date} ~ {cfg.end_date}")
        if dry_run:
            log.info(f"    [DRY-RUN] 跳过实际调用")
            return True
        try:
            extra = {'h5st': jm_h5st} if cfg.module=='京麦' else {}; result = handler(start_date=cfg.start_date, end_date=cfg.end_date, **extra)
            _extract_and_backfill(cfg, result, shop_id)
            return True
        except SystemExit as ex:
            if ex.code == 2:
                log.error(f"  🔐 [{cfg.module}/{cfg.report_name}] 鉴权过期")
                raise
            raise
        except Exception as ex:
            if _is_auth_expired(ex):
                log.error(f"  🔐 [{cfg.module}/{cfg.report_name}] 鉴权过期：{ex}")
                raise SystemExit(2)
            log.error(f"  ❌ [{cfg.module}/{cfg.report_name}] 失败：{type(ex).__name__}: {ex}")
            return False
    else:
        log.error(f"❌ 未知策略：{cfg.strategy}")
        return False


def run_shop(shop_id: str, configs: list, dry_run: bool = False) -> Tuple[int, int]:
    """跑单店所有 config，返回 (成功数, 失败数)"""
    # 获取 shop_pin：先尝试 runtime_config（env → 店铺账号 sheet），否则回退 env
    shop_pin = os.getenv("SHOP_PIN", "").strip()
    if not shop_pin:
        try:
            from runtime_config import get_shop_pin
            shop_pin = get_shop_pin()
        except SystemExit:
            log.error(f"❌ [{shop_id}] SHOP_PIN 未设置（runtime_config 拒绝兜底）")
            return 0, 1

    log.info(f"{'='*70}")
    log.info(f"🏪 店铺：{shop_id} ({shop_pin})")
    log.info(f"{'='*70}")

    ok = 0
    fail = 0
    for cfg in configs:
        try:
            if _run_single_config(cfg, shop_id, shop_pin, dry_run=dry_run):
                ok += 1
            else:
                fail += 1
        except SystemExit as e:
            if e.code == EXIT_AUTH_EXPIRED:
                log.error(f"⛔ 鉴权过期中断，停止 [{shop_id}]")
                return ok, fail
            raise

    # 收尾：增量写总表（2026-08-26 用户需求：不再全量 rebuild，只写本次 collect 的数据）
    if not dry_run:
        try:
            import excel_master as em
            log.info(f"📊 [{shop_id}] 增量写总表...")
            em.flush_shop(shop_id, shop_pin)
            log.info(f"✅ [{shop_id}] 总表增量更新完成")
        except Exception as e:
            log.error(f"⚠️ [{shop_id}] 总表写入失败：{type(e).__name__}: {e}")

    log.info(f"📈 [{shop_id}] 汇总：成功 {ok} / 失败 {fail}")
    return ok, fail


def main():
    ap = argparse.ArgumentParser(description="每日调度器（读 config.xlsx 驱动）")
    ap.add_argument("--shop", help="指定单店（默认遍历店铺清单）")
    ap.add_argument("--dry-run", action="store_true", help="只读配置 + 打印计划，不真跑")
    args = ap.parse_args()

    # 读配置
    from config_download_reader import load_download_configs
    try:
        configs = load_download_configs()
    except (FileNotFoundError, ValueError) as e:
        log.error(str(e))
        sys.exit(EXIT_SYSTEM_ERROR)

    log.info(f"📋 加载 {len(configs)} 条启用配置")

    # 多店
    if args.shop:
        shops = [(args.shop, os.getenv("SHOP_PIN", "").strip())]
        if not shops[0][1]:
            log.error("❌ --shop 需要 SHOP_PIN 环境变量")
            sys.exit(EXIT_SYSTEM_ERROR)
    else:
        # 遍历店铺清单
        try:
            from biz_config_loader import list_shop_ids
            shop_ids = list_shop_ids(enabled_only=True)
            if not shop_ids:
                log.error("❌ 店铺清单为空")
                sys.exit(EXIT_SYSTEM_ERROR)
        except Exception as e:
            log.error(f"❌ 读店铺清单失败：{e}")
            sys.exit(EXIT_SYSTEM_ERROR)
        shops = [(s, os.getenv("SHOP_PIN", "").strip()) for s in shop_ids]

    # 店间冷却（复用 rpa_run）
    shop_interval = 60
    try:
        from rpa_run import _read_shop_interval_seconds
        shop_interval = _read_shop_interval_seconds(None)
    except Exception:
        pass

    grand_ok = 0
    grand_fail = 0
    for idx, (shop_id, shop_pin) in enumerate(shops):
        if idx > 0:
            log.info(f"⏸ 店间冷却 {shop_interval}s...")
            if not args.dry_run:
                time.sleep(shop_interval)
        os.environ["SHOP_ID"] = shop_id
        os.environ["SHOP_PIN"] = shop_pin
        ok, fail = run_shop(shop_id, configs, dry_run=args.dry_run)
        grand_ok += ok
        grand_fail += fail
        if fail > 0 and not args.dry_run:
            log.warning(f"⚠️ [{shop_id}] 有失败，继续下一店")

    log.info(f"{'='*70}")
    log.info(f"🏁 总汇总：成功 {grand_ok} / 失败 {grand_fail}")
    log.info(f"{'='*70}")

    sys.exit(EXIT_SUCCESS if grand_fail == 0 else EXIT_BIZ_FAIL)


if __name__ == "__main__":
    main()
