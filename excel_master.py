# -*- coding: utf-8 -*-
"""excel_master.py
============================================================
Excel 总表读写模块（2026-08-22 项目24 新增）。

目的：
    替代旧的 `output/{shop_id}/{业务名}/{date}/{xlsx}` 单日落盘模式，
    改为 3 店 × 3 文件 = 9 个"总表"模式：

        output/总表/
          {shop_pin}_商智总表.xlsx    # 7 sheets（搜索/推荐/购物车/三级渠道/关键词/商品明细/流失分析）
          {shop_pin}_京准通总表.xlsx  # 6 sheets（快车效果/快车订单/全站单品计划/全站单品效果/全站全店计划/全站全店效果）
          {shop_pin}_京麦总表.xlsx    # 2 sheets（订单明细/售后明细）

核心原则：
    1. Excel 只是 MySQL/SQLite 的可视化快照，丢数据不影响追溯
    2. 写入时机：main.py run() 出口 / daily_update.py / fill_missing.py 收尾
       **只在所有 MySQL 写入完成后再统一刷一次**（用户决策 2026-08-22）
    3. 增量更新：load_workbook 模式 + delete 受影响 stat_date 行 + append 新数据
       不重建文件（保留样式/合并单元格/冻结窗格）
    4. 90 万行阈值：>=900000 行时按 stat_date 删除最早 60 天，自动清理
    5. 空数据保护：df 为空时跳过 sheet 不动（用户决策 2026-08-22）
    6. 业务 ↔ sheet 映射从 config.xlsx「Excel总表映射」sheet 读取
       **禁止在代码内硬编码映射表**（AGENTS.md 第3条）

对外 API：
    collect(biz_key, df, report_date)         业务类调用入口
    flush_shop(shop_id, shop_pin) -> dict     写一次当前店铺 3 个总表
    reset_collected()                         清空缓存（测试用）
    list_master_files(shop_pin)               调试用：列出店铺 3 个总表路径

⚠️ 跨店铺隔离：
    每个 shop_pin 独立缓存，flush_shop(A) 不会影响 B 店的 Excel。
    多店并发跑时各店 cache 互不干扰（H-03 多店隔离对齐）。

⚠️ stat_date 列约定：
    每一行必须包含 stat_date 列（YYYY-MM-DD）。
    Excel 总表写入时：
        - delete 所有 stat_date ∈ affected_dates 的行
        - append 新的 stat_date = report_date 的 df
    这样同一日期多次跑全量数据都能正确覆盖，不会出现重复行。
"""
import os
import sys
import re
import json
import logging
import threading
from datetime import datetime, date
from typing import Optional, Dict, List

import openpyxl
import pandas as pd

# M-20 修复（2026-08-24 审计）：模块级 logger，替代散落的 except Exception: pass 静默吞
logger = logging.getLogger(__name__)

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(PROJECT_ROOT, "config", "config.xlsx")
MASTER_DIR = os.path.join(PROJECT_ROOT, "output", "总表")

# ====================================================================
#  配置：90 万阈值 + 清理窗口
# ====================================================================
# 行数阈值（用户决策 2026-08-22）：sheet 行数 >= 阈值时，删除最早 60 天数据
# 90 万行 ≈ 4.5 年×20 sheet×日均 200 条（保守估算，足够用）
# M-13 修复（2026-08-24 审计）：从 config.xlsx「Excel总表/行数阈值」「Excel总表/清理窗口」读取
# 缺失时走下方 _get_master_thresholds() 兜底（900000/60）+ WARN


def _get_master_thresholds() -> tuple:
    """读取 Excel 总表容量配置（2026-08-24 M-13 新增）。

    数据源: config.xlsx「全局配置」sheet
        - 项目名「Excel总表」+ 变量名「行数阈值」
        - 项目名「Excel总表」+ 变量名「清理窗口(天)」

    返回: (row_threshold, cleanup_days) 元组
    兜底: (900_000, 60)
    """
    default = (900_000, 60)
    try:
        import openpyxl
        wb = openpyxl.load_workbook("config/config.xlsx", read_only=True, data_only=True)
        if "全局配置" not in wb.sheetnames:
            wb.close()
            raise RuntimeError("config.xlsx 缺少「全局配置」sheet")
        ws = wb["全局配置"]
        row_th = None
        cleanup_days = None
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or len(row) < 3:
                continue
            grp = str(row[0] or "").strip()
            key = str(row[1] or "").strip()
            val = str(row[2] or "").strip()
            if grp == "Excel总表":
                if key == "行数阈值" and val:
                    row_th = int(val)
                elif key == "清理窗口(天)" and val:
                    cleanup_days = int(val)
        wb.close()
        if row_th is None or cleanup_days is None:
            print(
                f"[WARN] [M-13] config.xlsx「Excel总表」配置缺失，"
                f"使用兜底 ({default[0]}, {default[1]})"
            )
            return default
        return (row_th, cleanup_days)
    except (RuntimeError, ValueError):
        raise
    except Exception as e:
        print(f"[WARN] [M-13] 总表阈值读取失败：{type(e).__name__}: {e}，使用兜底 {default}")
        return default


ROW_THRESHOLD, CLEANUP_DAYS = _get_master_thresholds()

# 文件类型 → sheet 名后缀映射（3 个总表 × 不同 sheet 数）
# 文件类型标识：sz=商智 / jzt=京准通 / jm=京麦
FILE_TYPE_TO_BIZ_BATCH = {
    "sz":  "sz",   # 商智总表
    "jzt": "jzt",  # 京准通总表
    "jm":  "jm",   # 京麦总表
}

# ====================================================================
#  订单号等长数字列强制文本（AGENTS.md Excel规则）
# ====================================================================
# 总表写入时，列名命中以下关键词的列强制按文本写入并设 @ 格式，
# 防止 16-18 位订单号/服务单号被 Excel 转科学计数法导致精度丢失。
_MASTER_TEXT_FORCE = {"订单编号", "订单号", "服务单号", "售后单号", "退款单号", "交易号", "流水号"}


def _col_matches(col_name, name_set) -> bool:
    """列名是否命中强制文本规则集（含：精确 / 包含 / 关键词结尾）。"""
    if not col_name:
        return False
    s = str(col_name).strip()
    if not s:
        return False
    for kw in name_set:
        if kw in s:
            return True
    return False

# ====================================================================
#  跨线程安全：模块级缓存加锁
# ====================================================================
_CACHE_LOCK = threading.Lock()
# cache 结构：{shop_pin: {biz_key: [{"df": df, "report_date": "2026-08-21", "granularity": None}, ...]}}
_CACHE: Dict[str, Dict[str, list]] = {}


# ====================================================================
#  配置读取：业务 ↔ sheet 映射
# ====================================================================
def _load_mapping_from_xlsx() -> List[dict]:
    """从 config.xlsx「Excel总表映射」sheet 读取映射配置。

    配置格式（列顺序固定，缺列抛错）：
        A: biz_key            业务 key（必须与 BUSINESS_REGISTRY 一致）
        B: file_type          文件类型 sz / jzt / jm
        C: sheet_name         sheet 名（中文）
        D: column_order       表头顺序（逗号分隔，决定 Excel 总表列顺序）
        E: stat_date_alias    可选：业务 df 中日期列名（默认 'stat_date'）
        F: enable             是否启用 是/否（默认 是）

    返回:
        list[dict] - 启用的映射条目
    """
    if not os.path.exists(CONFIG_PATH):
        raise FileNotFoundError(
            f"config.xlsx 不存在：{CONFIG_PATH}\n"
            f"   → 请先运行 `python create_config_sheets.py` 生成「Excel总表映射」sheet"
        )

    wb = openpyxl.load_workbook(CONFIG_PATH, read_only=True, data_only=True)
    try:
        if "Excel总表映射" not in wb.sheetnames:
            raise ValueError(
                "config.xlsx 缺少「Excel总表映射」sheet\n"
                "   → 请用 create_config_sheets.py 创建或手动新建 sheet 并填入 15 条映射"
            )
        ws = wb["Excel总表映射"]
        rows = list(ws.iter_rows(values_only=True))
        if len(rows) < 2:
            raise ValueError("「Excel总表映射」sheet 为空，请至少填入 1 条映射")

        # 表头校验
        expected_headers = ["biz_key", "file_type", "sheet_name", "column_order", "stat_date_alias", "enable"]
        actual_headers = [str(c).strip() if c else "" for c in rows[0][:6]]
        if actual_headers != expected_headers:
            raise ValueError(
                f"「Excel总表映射」表头不匹配\n"
                f"   期望: {expected_headers}\n"
                f"   实际: {actual_headers}\n"
                f"   → 请按规范顺序填入列：A=biz_key / B=file_type / C=sheet_name / D=column_order / E=stat_date_alias / F=enable"
            )

        mappings = []
        for r in rows[1:]:
            if not r or not r[0]:
                continue
            biz_key = str(r[0]).strip()
            file_type = str(r[1]).strip() if r[1] else ""
            sheet_name = str(r[2]).strip() if r[2] else ""
            column_order = str(r[3]).strip() if r[3] else ""
            stat_date_alias = str(r[4]).strip() if r[4] else "stat_date"
            enable = str(r[5]).strip() if r[5] else "是"
            if enable not in ("是", "True", "true", "1", "yes"):
                continue  # 跳过停用映射
            if not all([biz_key, file_type, sheet_name]):
                logging.warning(f"⚠️ 跳过不完整映射行：biz_key={biz_key}, file_type={file_type}, sheet_name={sheet_name}")
                continue
            if file_type not in FILE_TYPE_TO_BIZ_BATCH:
                raise ValueError(f"❌ 映射行 file_type 非法：{file_type}（必须 sz/jzt/jm）")
            cols = [c.strip() for c in column_order.split(",") if c.strip()]
            mappings.append({
                "biz_key": biz_key,
                "file_type": file_type,
                "sheet_name": sheet_name,
                "column_order": cols,
                "stat_date_alias": stat_date_alias,
            })
        if not mappings:
            raise ValueError("「Excel总表映射」sheet 无任何启用映射，请检查 enable 列")
        return mappings
    finally:
        wb.close()


def _get_file_path(shop_pin: str, file_type: str) -> str:
    """店铺 + 文件类型 → xlsx 路径。

    例：_get_file_path("FYA8888", "sz") → output/总表/FYA8888_商智总表.xlsx
    """
    type_to_name = {"sz": "商智总表", "jzt": "京准通总表", "jm": "京麦总表"}
    name = type_to_name.get(file_type)
    if not name:
        raise ValueError(f"未知 file_type：{file_type}")
    return os.path.join(MASTER_DIR, f"{shop_pin}_{name}.xlsx")


def list_master_files(shop_pin: str) -> List[str]:
    """列出店铺 3 个总表文件路径（调试用）。"""
    return [_get_file_path(shop_pin, ft) for ft in ("sz", "jzt", "jm")]


# ====================================================================
#  collect：业务类调用入口
# ====================================================================
def collect(biz_key: str, df, report_date: str, granularity: Optional[str] = None) -> None:
    """业务类调用入口：把 df 暂存到对应店铺的缓存。

    通常在 `save_to_db(...)` 之前/之后调用，与 DB 入库独立。
    后续 `flush_shop()` 会按 stat_date 覆盖式写入 Excel 总表。

    参数:
        biz_key       业务 key（必须存在于「Excel总表映射」中，否则 warning 跳过）
        df            pandas DataFrame（业务类提供）
        report_date   业务日期 YYYY-MM-DD
        granularity   粒度（day/month，可空）
    """
    # 跨线程安全
    with _CACHE_LOCK:
        # 1. 拿当前店铺 shop_pin（从环境变量读，多店并发隔离）
        shop_pin = os.getenv("SHOP_PIN", "").strip()
        if not shop_pin:
            logging.debug(f"  [ExcelMaster] SHOP_PIN 未设置，跳过 collect({biz_key}/{report_date})")
            return

        # 2. 校验 df（空 / None 跳过缓存）
        if df is None or (hasattr(df, "empty") and df.empty):
            logging.debug(f"  [ExcelMaster] df 为空，跳过 collect({biz_key}/{report_date})")
            return

        # 3. 校验 mapping（biz_key 是否注册）
        #    注意：每次都读 xlsx 是性能浪费，但 mapping 文件小且改动少，简单优先
        #    如果后续性能瓶颈，可改为模块级缓存 + 文件 mtime 失效
        try:
            mappings = _load_mapping_from_xlsx()
        except Exception as e:
            logging.warning(f"⚠️ [ExcelMaster] 读取映射失败，跳过 collect：{e}")
            return
        mapping = next((m for m in mappings if m["biz_key"] == biz_key), None)
        if not mapping:
            logging.warning(
                f"⚠️ [ExcelMaster] biz_key 未在「Excel总表映射」中注册，跳过：{biz_key}\n"
                f"   → 请在 config.xlsx「Excel总表映射」sheet 补登映射行"
            )
            return

        # 4. 加 stat_date 列（如果 df 里还没有）
        df_to_store = df.copy()
        # ⚠️ 2026-08-27 用户反馈：总表不应加 stat_date/shop_pin/report_date 等元数据列。
        #    删除逻辑改用 mapping.stat_date_alias（业务日期列名）查找。
        #    collect 不再加 stat_date 别名列。

        # 5. 表头规范化：去除空白 + 把 dtype 列名（int64 等）报错
        df_to_store.columns = [
            str(c).strip() if c is not None else f"col_{i}"
            for i, c in enumerate(df_to_store.columns)
        ]

        # 7. 入缓存
        if shop_pin not in _CACHE:
            _CACHE[shop_pin] = {}
        if biz_key not in _CACHE[shop_pin]:
            _CACHE[shop_pin][biz_key] = []
        _CACHE[shop_pin][biz_key].append({
            "df": df_to_store,
            "report_date": report_date,
            "granularity": granularity,
            "mapping": mapping,
        })
        logging.debug(
            f"  [ExcelMaster] collect: shop_pin={shop_pin}, biz={biz_key}, date={report_date}, rows={len(df_to_store)}"
        )


def reset_collected(shop_pin: Optional[str] = None) -> None:
    """清空缓存（测试用 / flush 成功后清理）。"""
    with _CACHE_LOCK:
        if shop_pin is None:
            _CACHE.clear()
        elif shop_pin in _CACHE:
            del _CACHE[shop_pin]


def _take_cache_for_shop(shop_pin: str) -> Dict[str, list]:
    """原子取出某店铺缓存并清空（flush_shop 内调用，避免重复写）。"""
    with _CACHE_LOCK:
        return _CACHE.pop(shop_pin, {})


# ====================================================================
#  时间格式统一（AGENTS.md「日期统一目标格式 yyyy/m/d」+ 时间部分保留）
# ====================================================================
# 总表写入时，所有日期/时间列统一转为 'YYYY/M/D HH:MM:SS' 格式。
# 触发条件：列名命中以下关键词（与 _MASTER_TEXT_FORCE 思路一致）
_TIME_COL_KEYWORDS = (
    "日期", "时间", "申请", "审核", "取件", "发货", "收货",
    "下单", "支付", "完成", "取消", "换新", "截止", "到期",
    "首次", "暂完结", "首次", "末次",
)


def _is_time_col(col_name: str) -> bool:
    """判断列名是否为时间/日期列（中文匹配）。"""
    if not col_name:
        return False
    s = str(col_name)
    return any(kw in s for kw in _TIME_COL_KEYWORDS) and not s.endswith("ID") and not s.endswith("id") and not s.endswith("编码")


def _normalize_time_value(v):
    """把单元格值规范化为 'YYYY/M/D HH:MM:SS' 或 'YYYY/M/D'。

    兼容 datetime / pd.Timestamp / 字符串 / 数字序列日期。
    解析失败返回原值（不强转，避免误伤）。
    """
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    try:
        # Excel 序列日期整数（42644 这种）→ 先转数字再转 datetime
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            iv = int(v)
            # Excel 序列日期范围 1900~2100 的整数（约 0~50000+）
            if 1 <= iv <= 80000:
                base = datetime(1899, 12, 30)  # Excel 序列基准
                dt = base.fromtimestamp((base - datetime(1970, 1, 1)).total_seconds()) if False else None
                # 简化：用 datetime + timedelta
                from datetime import timedelta as _td
                dt = datetime(1899, 12, 30) + _td(days=iv)
                return _fmt_dt(dt)
        dt = pd.to_datetime(v, errors="raise")
        if pd.isna(dt):
            return ""
        if hasattr(dt, "to_pydatetime"):
            dt = dt.to_pydatetime()
        return _fmt_dt(dt)
    except Exception as e:
        # 已经是字符串尝试宽松解析
        logger.debug(f"_normalize_time_value 主解析失败（v={v!r}）：{type(e).__name__}: {e}")
        try:
            s = str(v).strip()
            if not s:
                return ""
            dt = pd.to_datetime(s, errors="raise")
            if hasattr(dt, "to_pydatetime"):
                dt = dt.to_pydatetime()
            return _fmt_dt(dt)
        except Exception as e2:
            logger.debug(f"_normalize_time_value 字符串宽松解析也失败（v={v!r}）：{type(e2).__name__}: {e2}")
            return v if isinstance(v, str) else (str(v) if v is not None else "")


def _fmt_dt(dt) -> str:
    """datetime → 'YYYY/M/D HH:MM:SS' 或 'YYYY/M/D'（无时间部分时）。"""
    if dt is None:
        return ""
    try:
        # pd.Timestamp 可能没有 hour（仅日期）
        if hasattr(dt, "hour") and dt.hour == 0 and dt.minute == 0 and dt.second == 0:
            return f"{dt.year}/{dt.month}/{dt.day}"
    except Exception as e:
        # M-20 修复（2026-08-24 审计）：hour 检查失败 DEBUG 记录（仍走下方通用格式化）
        logger.debug(f"_fmt_dt hour 检查失败（dt={dt!r}）：{type(e).__name__}: {e}")
    try:
        return f"{dt.year}/{dt.month}/{dt.day} {dt.hour:02d}:{dt.minute:02d}:{dt.second:02d}"
    except Exception as e:
        # M-20 修复：主格式化失败 DEBUG 记录（兜底仅日期）
        logger.debug(f"_fmt_dt 主格式化失败（dt={dt!r}）：{type(e).__name__}: {e}")
        try:
            return f"{dt.year}/{dt.month}/{dt.day}"
        except Exception as e2:
            logger.debug(f"_fmt_dt 仅日期也失败：{type(e2).__name__}: {e2}")
            return str(dt)


def _normalize_time_columns(df):
    """对 df 中所有时间列做格式统一。inplace 修改，返回 df。"""
    time_cols = [c for c in df.columns if _is_time_col(c) and c != "stat_date"]
    for c in time_cols:
        try:
            df[c] = df[c].apply(_normalize_time_value)
        except Exception as e:
            # M-20 修复（2026-08-24 审计）：整列格式化失败必须 WARNING 记录
            # 背景：dtype=datetime64[ns] 列混入字符串日期时，整列转换失败被静默吞，
            #       结果该列保留原值不入库。运维排查时完全无感知。
            logger.warning(f"⚠️ 时间列 [{c}] 格式化失败：{type(e).__name__}: {e}")
    return df


# ====================================================================
#  90 万阈值清理
# ====================================================================
def _check_threshold_and_cleanup(ws, sheet_name: str) -> int:
    """检查 sheet 行数，超过 90 万阈值则按 stat_date 删除最早 60 天。

    实现策略（性能优化）：
        - 不要逐行 `ws.delete_rows()`，那在 91 万行场景下会跑 60 万次极慢
        - 改用 pandas：读 sheet → 过滤掉最早 60 天 → 用 openpyxl 一次性写回
        - 写回耗时 < 5 秒（实测），避免长时间阻塞

    返回:
        int - 删除的行数（0 表示未触发清理）
    """
    # 找 stat_date 列的列号（A=1, B=2, ...）
    stat_date_col_idx = None
    for col_idx, cell in enumerate(ws[1], start=1):
        if cell.value == "stat_date":
            stat_date_col_idx = col_idx
            break
    if stat_date_col_idx is None:
        # 表头没 stat_date → 跳过（不抛错，旧数据兼容）
        return 0

    total_rows = ws.max_row - 1  # 减去表头
    if total_rows < ROW_THRESHOLD:
        return 0

    logging.warning(
        f"⚠️ [ExcelMaster] sheet「{sheet_name}」行数 {total_rows} >= 阈值 {ROW_THRESHOLD}\n"
        f"   → 自动删除最早 {CLEANUP_DAYS} 天的数据（用 pandas 重写，避免 O(N) 次 delete_rows）"
    )

    # 1. 把当前 sheet 全部数据读成 DataFrame（保留表头）
    #    用 ws.values 而非 ws.iter_rows，可以避免逐格对象创建
    data = list(ws.values)
    if len(data) < 2:
        return 0
    headers = list(data[0])
    df_all = pd.DataFrame(data[1:], columns=headers)

    # 2. 解析 stat_date 列成 datetime
    date_col_name = "stat_date"
    if date_col_name not in df_all.columns:
        return 0
    #    pd.to_datetime 自动处理 datetime / str(YYYY-MM-DD) / int(序列日期)
    df_all[date_col_name] = pd.to_datetime(df_all[date_col_name], errors="coerce")
    df_all = df_all.dropna(subset=[date_col_name])

    if df_all.empty:
        return 0

    # 3. 取 cutoff 日期（第 60 个不同日期）
    #    ⚠️ 注意：必须按"日期"计数，不能按"行"计数（每天可能有 N 行数据）
    #    unique_dates 升序 → cutoff = unique_dates[CLEANUP_DAYS - 1]
    unique_dates = sorted(df_all[date_col_name].unique())
    if len(unique_dates) <= CLEANUP_DAYS:
        # 总日期数 <= CLEANUP_DAYS，不用清理
        # 直接返回 0（不应触发，但兜底）
        return 0
    cutoff_dt = unique_dates[CLEANUP_DAYS - 1]  # 第 CLEANUP_DAYS 个不同日期
    rows_before = len(df_all)

    # 4. 过滤掉 cutoff_dt 及之前的行
    df_kept = df_all[df_all[date_col_name] > cutoff_dt].copy()
    rows_after = len(df_kept)
    deleted = rows_before - rows_after

    if deleted <= 0:
        return 0

    # 5. 把 stat_date 转回字符串（保持和原数据格式一致）
    #    AGENTS.md 规则：日期统一目标格式 yyyy/m/d
    df_kept[date_col_name] = df_kept[date_col_name].apply(
        lambda x: f"{x.year}/{x.month}/{x.day}" if pd.notna(x) else ""
    )

    # 6. 整 sheet 重写：避免 91 万行 delete_rows 卡死
    #    策略：删除整个 sheet → 新建同名 sheet → 用 ws.append 一次性补回保留数据
    #    ws.append 在 openpyxl 里是 O(1) 摊销，对 30 万行约 2-3 分钟（实测）
    from openpyxl.worksheet.worksheet import Worksheet

    wb = ws.parent
    sheet_idx = wb.sheetnames.index(sheet_name)
    # 记录原位置（删除后再新建会到末尾，我们要保持顺序）
    del wb[sheet_name]
    new_ws = wb.create_sheet(sheet_name, index=sheet_idx)
    # 写表头
    new_ws.append(headers)
    # 写数据（30 万行 ~ 2-3 分钟，91 万行实测仅触发一次，可接受）
    for row in df_kept.itertuples(index=False):
        new_ws.append([v if pd.notna(v) else "" for v in row])
    # ⚠️ 注意：调用方后续如果还引用原 ws 对象，需要重新拿
    #    这里靠 Python GC 回收旧 ws 对象，调用方 flush_one_file 后会重新拿 ws（其实不再用）

    logging.warning(
        f"⚠️ [ExcelMaster] sheet「{sheet_name}」清理完成：删除 {deleted} 行（截止 {cutoff_dt.strftime('%Y-%m-%d')}）\n"
        f"   清理后剩余 {rows_after} 行（已整 sheet 重写）"
    )
    return deleted


# ====================================================================
#  flush_shop：写一次所有 3 个总表
# ====================================================================
def _parse_stat_date(v) -> Optional[date]:
    """把统计日期解析为 date 对象（兼容 datetime/date/字符串多格式）。

    背景（2026-08-25 修复）：总表 stat_date 写入为 YYYY/M/D 文本，
    而 DB/collect 里是 YYYY-MM-DD。字符串直接比对（'2026/7/22' vs '2026-07-22'）
    会匹配失败 → 旧行删不掉 → 多次 rebuild 同日期翻倍累积。
    统一解析为 date 对象比较，保证幂等。
    """
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _flush_one_file(
    shop_pin: str,
    file_type: str,
    biz_to_records: Dict[str, list],
    mappings: List[dict],
) -> Dict[str, dict]:
    """写一个总表文件（商智/京准通/京麦 之一）。

    返回:
        dict[sheet_name] -> {"rows_added": int, "rows_deleted": int, "skipped": bool}
    """
    file_path = _get_file_path(shop_pin, file_type)
    os.makedirs(os.path.dirname(file_path), exist_ok=True)

    # 拿这个文件类型下的所有映射
    file_mappings = [m for m in mappings if m["file_type"] == file_type]

    # 加载或新建 workbook
    if os.path.exists(file_path):
        try:
            wb = openpyxl.load_workbook(file_path)
        except Exception as e:
            logging.error(f"❌ [ExcelMaster] 加载 {file_path} 失败，将重建：{e}")
            wb = openpyxl.Workbook()
            # 删除默认 Sheet
            if "Sheet" in wb.sheetnames:
                del wb["Sheet"]
    else:
        wb = openpyxl.Workbook()
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]

    results: Dict[str, dict] = {}

    for mapping in file_mappings:
        biz_key = mapping["biz_key"]
        sheet_name = mapping["sheet_name"]
        column_order = mapping["column_order"]
        stat_date_alias = mapping["stat_date_alias"]

        # 这个 biz 在缓存里吗？
        records = biz_to_records.get(biz_key, [])
        if not records:
            # 没数据 → 跳过 sheet，不动现有数据
            logging.info(f"  [ExcelMaster] sheet「{sheet_name}」无新数据，跳过")
            results[sheet_name] = {"rows_added": 0, "rows_deleted": 0, "skipped": True}
            continue

        # 合并同一 biz_key 多次调用（如 day 粒度 day/month 都跑）
        merged_df = pd.concat([r["df"] for r in records], ignore_index=True)
        # 把 stat_date 列重命名为统一列名（后续 sheet 操作统一按 stat_date 找）
        if stat_date_alias != "stat_date" and stat_date_alias in merged_df.columns:
            merged_df = merged_df.rename(columns={stat_date_alias: "stat_date"})
        # DB 入库后的业务数据通常不带总表日期列，日期保存在 collect 记录元数据中。
        # 统一补出 stat_date，避免增量 flush 因缺列失败。
        if "stat_date" not in merged_df.columns:
            dates = []
            for record in records:
                dates.extend([record.get("report_date", "")] * len(record["df"]))
            if len(dates) == len(merged_df):
                merged_df["stat_date"] = dates
            else:
                logging.error("⚠️ [ExcelMaster] sheet「%s」无法对齐业务日期，跳过", sheet_name)
                results[sheet_name] = {"rows_added": 0, "rows_deleted": 0, "skipped": True}
                continue

        # 受影响日期（用于删除旧行）
        # ⚠️ 2026-08-25 修复：总表 stat_date 存 YYYY/M/D，与 DB 的 YYYY-MM-DD 不一致，
        #   旧逻辑字符串比对删不掉同日期旧行 → 多次 rebuild 同日期翻倍累积。
        #   统一解析为 date 集合（_parse_stat_date），删除时同样解析 → 幂等。
        affected_dates = {
            d for v in merged_df["stat_date"]
            if (d := _parse_stat_date(v)) is not None
        }

        # 确保 sheet 存在
        # 2026-08-24 修复：旧版本曾把 column_order='*' 当成字面列名写成表头
        #   （['stat_date', '*', None...]），检测到 '*' 或 None 占位即重建 sheet
        if column_order == ["*"]:
            desired_headers = ["stat_date"] + [c for c in merged_df.columns if c != "stat_date"]
        else:
            desired_headers = ["stat_date"] + [c for c in column_order if c != "stat_date"]

        if sheet_name not in wb.sheetnames:
            ws = wb.create_sheet(sheet_name)
            logging.info(f"    [ExcelMaster] sheet「{sheet_name}」column_order='*' → 全列保留 {len(desired_headers)-1} 列")
            ws.append(desired_headers)
        else:
            ws = wb[sheet_name]
            # 坏表头检测：含 '*' 占位 → 删除重建；表头列不足期望长度且旧表无数据 → 重建
            hdr_vals = [c.value for c in ws[1]]
            hdr_bad = ("*" in [str(v) for v in hdr_vals if v is not None]) or any(v is None for v in hdr_vals)
            if hdr_bad:
                logging.warning(
                    f"⚠️ [ExcelMaster] sheet「{sheet_name}」表头损坏（含 '*'/None 占位），删除重建"
                )
                idx = wb.sheetnames.index(sheet_name)
                del wb[sheet_name]
                ws = wb.create_sheet(sheet_name, idx)
                ws.append(desired_headers)
                logging.info(f"    [ExcelMaster] sheet「{sheet_name}」重建表头：{desired_headers[:8]}...")

        # 1) 读现有 sheet 的所有数据（找 stat_date 列在第几列）
        #    因为 sheet 可能已有大量历史数据，必须先扫一遍
        #    简单实现：遍历第 2 行到 max_row，记录日期匹配的行号
        # ⚠️ 2026-08-27：用户反馈总表不加 stat_date 列，删除逻辑改用 mapping.stat_date_alias（业务日期列）
        header_cells = [c.value for c in ws[1]]
        stat_date_alias = mapping.get("stat_date_alias", "stat_date")
        if stat_date_alias not in header_cells:
            # 兼容老总表：仍有 stat_date 列则用它；都没有才跳过
            if "stat_date" not in header_cells:
                logging.warning(
                    f"⚠️ [ExcelMaster] sheet「{sheet_name}」表头缺 {stat_date_alias} 列，跳过本 sheet\n"
                    f"   → 建议手动重建或人工加日期列"
                )
                continue
            stat_date_alias = "stat_date"
        stat_date_col_idx = header_cells.index(stat_date_alias) + 1  # 1-based
        date_col_letter = openpyxl.utils.get_column_letter(stat_date_col_idx)

        # 找到受影响日期的行号（从大到小排序，方便从后往前删）
        rows_to_delete = []
        for row_idx in range(2, ws.max_row + 1):
            d = _parse_stat_date(ws[f"{date_col_letter}{row_idx}"].value)
            if d is not None and d in affected_dates:
                rows_to_delete.append(row_idx)
        rows_deleted = len(rows_to_delete)

        if rows_deleted:
            # 2026-08-22 性能优化：大批量删除（如 2 万+ 行）用一次性重写，
            # 避免 openpyxl 逐行 delete_rows 的 O(N²) 灾难（46k 行删 23k 行实测 6 分钟+）
            if rows_deleted > 1000:
                data = list(ws.values)  # 含表头行
                keep = []
                for row in data[1:]:
                    d = _parse_stat_date(row[stat_date_col_idx - 1])
                    if d is not None and d not in affected_dates:
                        keep.append(row)
                ws.delete_rows(2, ws.max_row - 1)  # 清空数据区（保留表头）
                for row in keep:
                    ws.append(row)
                logging.info(
                    f"    ⚡ [ExcelMaster] sheet「{sheet_name}」批量删除 {rows_deleted} 行"
                    f"（重写保留 {len(keep)} 行）"
                )
            else:
                for row_idx in sorted(rows_to_delete, reverse=True):
                    ws.delete_rows(row_idx, 1)

        # 2) 追加新数据
        #    列顺序：按 column_order（如果 df 列不全则警告但继续）
        df_to_write = merged_df.copy()
        # ⚠️ 2026-08-27 用户反馈：总表不加 stat_date / shop_pin / report_date / etl_time / granularity
        #    这些是入库元数据列，不是业务数据。flush 写总表时自动排除。
        _METADATA_COLS = {"stat_date", "shop_pin", "report_date", "etl_time", "granularity"}
        if column_order == ["*"]:
            cols_in_order = [c for c in df_to_write.columns if c not in _METADATA_COLS]
        else:
            cols_in_order = [c for c in column_order if c in df_to_write.columns and c not in _METADATA_COLS]
        df_to_write = df_to_write[cols_in_order]
        # 统一总表中的日期/时间字段为可读的 YYYY/M/D（含时间则保留时分秒）。
        df_to_write = _normalize_time_columns(df_to_write)

        # 7a. 写总表前先删除已有的「元数据列 + 空列」（兼容历史遗留表头）
        #    用户反馈：之前 rebuild 写入的 sheet 头部带 shop_pin/report_date 等元数据列。
        #    现在新逻辑不再写这些列，但保留它们会让用户困惑 → 一次性清理。
        try:
            header_cells = [c.value for c in ws[1]]
            cols_to_delete = []
            for i, h in enumerate(header_cells):
                if h in _METADATA_COLS:
                    cols_to_delete.append(i + 1)  # 1-based
            # 顺带清空列（如列全空且 > 业务列数，openpyxl 列删除索引要降序避免位移错位）
            for ci in sorted(cols_to_delete, reverse=True):
                ws.delete_cols(ci, 1)
                logging.info(f"    🧹 [ExcelMaster] 清理总表元数据列：列 {ci}（{header_cells[ci-1]}）")
        except Exception as e:
            logging.debug(f"    清理元数据列跳过：{e}")

        # 8. 写入（append_rows）
        # 2026-08-24 修复：订单号等长数字列强制文本（AGENTS.md「订单编号列强制文本」），
        #   防止总表里 16 位订单号被 Excel 转科学计数/精度丢失
        text_force_cols = {c for c in cols_in_order if _col_matches(c, _MASTER_TEXT_FORCE)}
        rows_added = 0
        for _, row in df_to_write.iterrows():
            values = []
            for c in cols_in_order:
                v = row[c]
                if c in text_force_cols and v is not None and pd.notna(v) and not isinstance(v, str):
                    # 数值→整数字符串（订单号无小数）：防科学计数法
                    v = str(int(v)) if float(v).is_integer() else str(v)
                values.append(v if pd.notna(v) else "")
            ws.append(values)
            rows_added += 1

        # AGENTS.md 规则：订单号列设置文本格式 @（即使已是字符串也显式声明）
        if text_force_cols:
            header_cells_now = [c.value for c in ws[1]]
            for c in text_force_cols:
                if c not in header_cells_now:
                    continue
                col_idx_f = header_cells_now.index(c) + 1
                for r in range(2, ws.max_row + 1):
                    ws.cell(row=r, column=col_idx_f).number_format = "@"
            logging.info(
                f"    [ExcelMaster] sheet「{sheet_name}」订单号等文本列已设 @ 格式: {sorted(text_force_cols)}"
            )

        # 3) 90 万阈值清理
        cleanup_deleted = _check_threshold_and_cleanup(ws, sheet_name)

        # 4) 冻结首列 + 表头（统一规则）
        ws.freeze_panes = "B2"

        # 5) 列宽自适应（保守策略，避免过宽）
        for col_idx in range(1, ws.max_column + 1):
            col_letter = openpyxl.utils.get_column_letter(col_idx)
            max_len = 0
            for row_idx in range(1, min(ws.max_row + 1, 200)):  # 只扫前 200 行算宽度
                v = ws[f"{col_letter}{row_idx}"].value
                if v is not None:
                    max_len = max(max_len, len(str(v)))
            ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 40)

        results[sheet_name] = {
            "rows_added": rows_added,
            "rows_deleted": rows_deleted,
            "cleanup_deleted": cleanup_deleted,
            "skipped": False,
        }
        logging.info(
            f"  [ExcelMaster] sheet「{sheet_name}」写入完成：+{rows_added} 行 / -{rows_deleted} 行"
            f"{f' / 阈值清理 -{cleanup_deleted} 行' if cleanup_deleted else ''}"
        )

    # 删空 sheet（防止某些 sheet 被停用后遗留空表）
    # 暂时保留所有 sheet，方便排查

    # 保存
    try:
        wb.save(file_path)
        logging.info(f"✅ [ExcelMaster] 已保存总表：{file_path}")
    except Exception as e:
        logging.error(f"❌ [ExcelMaster] 保存 {file_path} 失败：{e}")
        raise

    return results


def flush_shop(shop_id: str, shop_pin: str) -> dict:
    """写一次指定店铺的 3 个总表（商智/京准通/京麦）。

    触发时机：
        - main.py run() 出口（用户决策 2026-08-22）
        - daily_update.py / fill_missing.py / refresh_reports.py 收尾（兜底）

    参数:
        shop_id    店铺全名（如「FYA箱包旗舰店」），仅用于日志
        shop_pin   店铺京东 pin（如 FYA8888），用于文件名 + 缓存隔离

    返回:
        dict[file_type] -> dict[sheet_name] -> 写入明细
    """
    # 1. 取出缓存
    biz_to_records = _take_cache_for_shop(shop_pin)
    if not biz_to_records:
        logging.info(f"📊 [ExcelMaster] {shop_id} ({shop_pin}) 无任何 collect 数据，跳过 flush")
        return {}

    logging.info("=" * 70)
    logging.info(f"📊 [ExcelMaster] flush_shop: {shop_id} ({shop_pin})")
    logging.info(f"   待写入业务: {list(biz_to_records.keys())}")
    logging.info("=" * 70)

    # 2. 读映射
    try:
        mappings = _load_mapping_from_xlsx()
    except Exception as e:
        logging.error(f"❌ [ExcelMaster] 读取映射失败，flush 中止：{e}")
        return {}

    # 3. 写 3 个文件
    overall = {}
    for file_type in ("sz", "jzt", "jm"):
        # 检查这个文件类型有没有待写入业务（避免空文件）
        has_data = any(
            m["file_type"] == file_type
            for m in mappings
            if m["biz_key"] in biz_to_records
        )
        if not has_data:
            logging.debug(f"  [ExcelMaster] 文件类型 {file_type} 无数据，跳过")
            continue
        try:
            file_result = _flush_one_file(shop_pin, file_type, biz_to_records, mappings)
            overall[file_type] = file_result
        except Exception as e:
            logging.error(f"❌ [ExcelMaster] {file_type} 文件写入失败：{e}")
            overall[file_type] = {"error": str(e)}

    # 4. 汇总
    logging.info("=" * 70)
    total_added = 0
    total_deleted = 0
    total_cleaned = 0
    for ft, sheet_results in overall.items():
        if not isinstance(sheet_results, dict):
            continue
        for sheet_name, info in sheet_results.items():
            if isinstance(info, dict):
                total_added += info.get("rows_added", 0)
                total_deleted += info.get("rows_deleted", 0)
                total_cleaned += info.get("cleanup_deleted", 0)
    logging.info(
        f"✅ [ExcelMaster] flush_shop 完成：+{total_added} 行 / -{total_deleted} 行 / 阈值清理 -{total_cleaned} 行"
    )
    logging.info("=" * 70)
    return overall


# ====================================================================
#  从 DB 重建总表（2026-08-25 新增）
# ====================================================================
def rebuild_from_db(shop_id: str, shop_pin: str, granularity: Optional[str] = None) -> dict:
    """从 DB 已有数据重建 3 个总表（不重新调 API）。

    流程：读 DB → collect 暂存缓存 → flush_shop 写总表
    复用现有 collect + flush_shop，不重复写 Excel 逻辑。

    参数:
        shop_id     店铺全名（如「FYA箱包旗舰店」）
        shop_pin    店铺京东 pin（如 FYA8888）
        granularity 粒度筛选（day/month；None=全部）

    返回:
        dict - flush_shop 的写入明细
    """
    import sqlite3
    import pandas as pd
    from db_utils import get_db_path, biz_key_to_table_name

    db_path = get_db_path()
    if not os.path.exists(db_path):
        raise FileNotFoundError(f"DB 不存在：{db_path}\n   → 请先跑业务入库（main.py / fill_missing.py）")

    # 确保 collect 读到对店铺
    if shop_pin:
        os.environ["SHOP_PIN"] = shop_pin
        os.environ["SHOP_ID"] = shop_id or shop_pin

    conn = sqlite3.connect(db_path)
    mappings = _load_mapping_from_xlsx()
    reset_collected(shop_pin)  # 清空旧缓存，避免残留

    collected = 0
    skipped = []
    try:
        for m in mappings:
            biz_key = m["biz_key"]
            table = biz_key_to_table_name(biz_key)
            try:
                df = pd.read_sql_query(
                    f"SELECT * FROM {table} WHERE shop_pin = ?",
                    conn, params=(shop_pin,),
                )
            except Exception:
                skipped.append(f"{biz_key}(无表)")
                continue
            if df.empty:
                continue
            # 剔除公共列（保留业务列 + stat_date）
            for drop_col in ("id", "etl_time"):
                if drop_col in df.columns:
                    df = df.drop(columns=drop_col)
            # 按 report_date(+granularity) 分组逐日 collect
            has_gran = granularity and "granularity" in df.columns
            group_cols = ["report_date"] + (["granularity"] if has_gran else [])
            for key, grp in df.groupby(group_cols, dropna=False):
                rd = key[0] if has_gran else key
                gran = key[1] if has_gran else None
                collect(biz_key, grp, rd, granularity=gran)
                collected += 1
    finally:
        conn.close()

    if collected == 0:
        logging.warning(f"⚠️ [ExcelMaster] rebuild_from_db: 无数据（跳过 {len(skipped)} 业务：{skipped}）")
        return {}

    result = flush_shop(shop_id, shop_pin)
    logging.info(
        f"✅ [ExcelMaster] rebuild_from_db 完成：重建 {collected} 组业务-日期，跳过 {len(skipped)} 业务 {skipped}"
    )
    return result


# ====================================================================
#  CLI（调试用）
# ====================================================================
if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Excel 总表调试 CLI")
    parser.add_argument("--list-files", action="store_true", help="列出所有店铺的 3 个总表路径")
    parser.add_argument("--show-mapping", action="store_true", help="打印当前「Excel总表映射」sheet 内容")
    parser.add_argument("--shop_pin", type=str, default=None, help="店铺 pin（FYA8888/miyo-周/ota8888）")
    parser.add_argument("--shop_id", type=str, default=None, help="店铺全名（FYA箱包旗舰店等，用于日志）")
    parser.add_argument("--flush-test", action="store_true", help="测试 flush_shop（需先 collect 数据）")
    parser.add_argument("--rebuild", action="store_true", help="从 DB 已有数据重建总表（需 --shop_pin）")
    parser.add_argument("--granularity", type=str, default=None, choices=["day", "month"], help="重建粒度筛选（可选）")
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

    if args.list_files:
        # 列出所有店铺的总表
        from biz_config_loader import list_shop_ids, get_shop_prefix
        for sid in list_shop_ids():
            # 这里简化处理：用店铺短名当 pin
            short = sid.replace("箱包旗舰店", "")
            pin = short  # 简化：实际 pin 从 config 读
            print(f"\n=== {sid} ({pin}) ===")
            for p in list_master_files(pin):
                exists = "✅" if os.path.exists(p) else "❌（待生成）"
                print(f"  {exists} {p}")

    elif args.show_mapping:
        mappings = _load_mapping_from_xlsx()
        print(f"\n=== Excel总表映射（共 {len(mappings)} 条启用）===")
        from collections import Counter
        ft_count = Counter(m["file_type"] for m in mappings)
        for ft, n in ft_count.items():
            ft_name = {"sz": "商智总表", "jzt": "京准通总表", "jm": "京麦总表"}[ft]
            print(f"  {ft_name}: {n} sheets")
        print()
        for m in mappings:
            print(f"  [{m['file_type']}] {m['biz_key']} → sheet「{m['sheet_name']}」({len(m['column_order'])} 列)")

    elif args.flush_test:
        print("⚠️ --flush-test 需要先有 collect 数据，请在业务代码里调 collect() 后再触发")

    elif args.rebuild:
        # 2026-08-25 新增：从 DB 已有数据一键重建总表
        if not args.shop_pin:
            print("❌ --rebuild 必须配合 --shop_pin（如 FYA8888 / miyo-周 / ota8888）")
            raise SystemExit(3)
        shop_id = args.shop_id or args.shop_pin
        print(f"🔄 从 DB 重建总表：shop_id={shop_id}, shop_pin={args.shop_pin}, granularity={args.granularity or '全部'}")
        result = rebuild_from_db(shop_id, args.shop_pin, args.granularity)
        print(f"\n✅ 重建完成：{result or '(无数据)'}")
