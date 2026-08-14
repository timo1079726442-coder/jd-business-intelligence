# -*- coding: utf-8 -*-
"""
db_utils.py
============================================================
京东数据报表 SQLite 入库工具集（2026-08-14 上线骨架）。

目的：
    - 把现有 20 个业务的 Excel 导出同时入库到 SQLite（双保险）
    - 提供「每日自动更新」「缺失日期补录」两个上层入口

设计原则：
    - 业务类零侵入：业务类只调 save_to_db(biz_key, df, report_date) 即可
    - 业务主键手动配置（config.xlsx），自动推断仅作 fallback
    - 全量覆盖：当天数据全部替换（不是 merge）

表结构通用模板：
    CREATE TABLE biz_{slug} (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        report_date TEXT NOT NULL,
        etl_time TEXT NOT NULL,
        <业务动态列>...,
        <业务唯一键> TEXT NOT NULL,
        UNIQUE(report_date, <业务唯一键>)
    );

模块入口（供 daily_update.py / fill_missing.py 调用）：
    - get_db_path()                                  -> 读取 config.xlsx 中 db_path
    - is_db_enabled()                                -> 读取 config.xlsx 中 enable_db_storage
    - biz_key_to_table_name(biz_key)                 -> 中文业务名 → 表名 biz_xxx
    - ensure_table(conn, table_name, df, pk_col)     -> 动态建表 + 缺列补齐
    - upsert_df(conn, table_name, df, report_date,
                 pk_col, granularity=None)           -> 整 df upsert
    - get_existing_dates(conn, table_name,
                         granularity=None)           -> 查已有日期集合
    - get_existing_count(conn, table_name)           -> 查总行数
    - save_to_db(biz_key, df, report_date,
                 granularity=None)                   -> 一键入库入口

⚠️ 业务主键自动推断 fallback 规则：
    - 优先用 config.xlsx 中「<业务名> primary_key_column」配置
    - 否则按列名包含「ID / 名称 / 编号 / 关键词 / 渠道」关键词推断
    - 否则取第一列文本字段
"""
import os
import re
import sqlite3
import logging
from datetime import datetime
from typing import List, Optional, Union

import pandas as pd

# ====================================================================
#  常量与工具
# ====================================================================

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))


def _pinyin_slug(biz_key: str) -> str:
    """中文业务名 → ASCII 短写（用于 SQLite 表名）。

    简单策略：去掉空格/标点，转 pinyin 首字母 or 直接用英文短写。
    本函数先用中文业务名内嵌的英文短写回退（业务名中已有「_」分隔符的部分）；
    实在没有则用拼音首字母或保留原字符。
    """
    # 业务名通常已经是「中文_英文短写」格式，取最后一段英文短写
    parts = re.split(r"[_\-\s]+", biz_key.strip())
    # 找最后一段非中文的
    for part in reversed(parts):
        if part and re.match(r"^[A-Za-z]+$", part):
            return part.lower()
    # 兜底：保留中文 unicode（SQLite 支持中文表名，但建议避免）
    return biz_key


# ====================================================================
#  config 读取（DB 相关 3 项）
# ====================================================================

def _read_config_xlsx() -> dict:
    """读取 config.xlsx「全局配置」sheet，返回 {业务名/全局: {key: value}}。

    复用 main.py 同样的读取逻辑，但只读 DB 相关 3 项。
    """
    config_path = os.path.join(PROJECT_ROOT, "config", "config.xlsx")
    if not os.path.exists(config_path):
        return {}
    try:
        import openpyxl
        wb = openpyxl.load_workbook(config_path, read_only=True, data_only=True)
        result = {}
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or len(row) < 3:
                    continue
                group = str(row[0]) if row[0] is not None else ""
                key = str(row[1]) if row[1] is not None else ""
                val = row[2]
                if not (group and key and val is not None):
                    continue
                result.setdefault(group, {})[key] = str(val).strip()
        wb.close()
        return result
    except Exception as e:
        logging.warning(f"⚠️ config.xlsx 读取失败：{e}")
        return {}


def get_db_path() -> str:
    """DB 路径（默认 data/jd_report.db）。

    配置位置：config.xlsx「全局配置」sheet → 「全局」组 → `db_path`
    """
    cfg = _read_config_xlsx()
    rel = cfg.get("全局", {}).get("db_path", "data/jd_report.db")
    if os.path.isabs(rel):
        return rel
    return os.path.join(PROJECT_ROOT, rel)


def is_db_enabled() -> bool:
    """是否启用数据库入库。

    配置位置：config.xlsx「全局」组 → `enable_db_storage`（默认 True）
    """
    cfg = _read_config_xlsx()
    val = cfg.get("全局", {}).get("enable_db_storage", "True").strip().lower()
    return val in ("true", "1", "yes", "on", "是")


# ====================================================================
#  表名生成 + 业务主键推断
# ====================================================================

# 业务名 → 表名 slug 的映射（白名单/黑名单）
# 白名单：业务名含英文短写时直接用
_BIZ_SLUG_OVERRIDE = {
    "商智关键词分析": "keyword_analysis",
    "店铺来源_三级渠道": "offline_channel",
    "商品流量来源_搜索": "traffic_search",
    "商品流量来源_推荐": "traffic_recommend",
    "商品流量来源_购物车": "traffic_cart",
    "商品流量来源_自主访问": "traffic_selfvisit",
    "商品明细导出": "product_detail",
    "商品流失分析": "loss_product",
    "京准通快车自定义报表": "jzt_kuaiche",
    "京准通快车订单效果明细": "jzt_kuaiche_order_effect",
    "京准通全站营销单品计划": "jzt_quanzhan_campaign",
    "京准通全站营销单品推广效果": "jzt_quanzhan_effect",
    "京准通全站营销全店计划": "jzt_quanzhan_campaign_all_store",
    "京准通全站营销全店推广效果": "jzt_quanzhan_effect_all_store",
    "京麦订单明细_创建任务": "jm_order_create",
    "京麦订单明细_创建并轮询": "jm_order_wait",
    "京麦订单明细_创建轮询并下载zip": "jm_order_dl_zip",
    "京麦订单明细_创建轮询下载并申请密码": "jm_order_dl_pwd",
    "京麦订单明细_完整一键导出": "jm_order_full",
    "京麦售后明细_完整一键导出": "jm_after_sale_full",
}


def biz_key_to_table_name(biz_key: str) -> str:
    """业务名 → 表名 `biz_<slug>`。

    优先用 _BIZ_SLUG_OVERRIDE 白名单，否则按 _pinyin_slug 兜底。
    """
    slug = _BIZ_SLUG_OVERRIDE.get(biz_key, _pinyin_slug(biz_key))
    return f"biz_{slug}"


def infer_primary_key(df: pd.DataFrame, biz_key: str = "") -> Optional[Union[str, list]]:
    """自动推断业务主键列名。

    优先级：
        1. config.xlsx 中「<业务短写> primary_key_columns」配置（复合主键，复数）
        2. config.xlsx 中「<业务短写> primary_key_column」配置（单主键）
        3. 列名包含关键词（ID / 编号 / 名称 / 关键词 / SKU）
        4. 第一列文本字段

    ⚠️ 复合主键配置示例（一级来源,二级来源,三级来源）：「三级渠道 primary_key_columns = 一级来源,二级来源,三级来源」

    返回:
        str  - 单主键列名
        list - 复合主键列名列表
        None - 推断失败
    """
    # 1. config 配置
    cfg = _read_config_xlsx()
    biz_short = biz_key.split("_")[-1] if "_" in biz_key else biz_key

    # 1.1 优先查复合主键（primary_key_columns 复数）
    pk_composite = cfg.get(biz_short, {}).get("primary_key_columns", "").strip()
    if pk_composite:
        cols = [c.strip() for c in pk_composite.split(",") if c.strip()]
        if cols and all(c in df.columns for c in cols):
            return cols

    # 1.2 单主键
    pk = cfg.get(biz_short, {}).get("primary_key_column", "").strip()
    if pk and pk in df.columns:
        return pk

    # 2. 列名关键词推断
    keywords = ["ID", "id", "编号", "名称", "名字", "关键词", "SKU", "计划ID"]
    for col in df.columns:
        for kw in keywords:
            if kw in str(col):
                return col

    # 3. 第一列文本字段
    if len(df.columns) > 0:
        first_col = df.columns[0]
        if df[first_col].dtype == object:
            return first_col

    return None


# ====================================================================
#  动态建表 + 列类型推断
# ====================================================================

def _infer_sqlite_type(series: pd.Series) -> str:
    """根据 pandas Series 的 dtype 推断 SQLite 列类型。

    简化策略：所有非空列一律 TEXT（与 Excel 读入 dtype=str 一致），
    数值类自动转 REAL（但在 Excel 已 safe_convert_numeric 后，存为 TEXT 安全）。
    """
    return "TEXT"


def _safe_col_name(col: str) -> str:
    """安全的列名（防 SQL 关键字 + 特殊字符）。"""
    # 中文/英文/数字都保留；保留关键字需加双引号
    if not col:
        return "col_unnamed"
    # SQLite 关键字（精简版）
    keywords = {
        "order", "group", "select", "from", "where", "table", "index",
        "primary", "key", "date", "datetime", "time", "user", "name",
    }
    if col.lower() in keywords:
        return f'"{col}"'  # 加双引号转义
    return f'"{col}"'


def ensure_table(
    conn: sqlite3.Connection,
    table_name: str,
    df: pd.DataFrame,
    pk_col: Union[str, list],
) -> None:
    """动态建表 + 缺列补齐。

    行为：
        - 如果表不存在：CREATE TABLE，含 id / report_date / etl_time / <df cols> / UNIQUE(report_date, pk_col(s))
        - 如果表已存在：检查 df 列是否都存在，缺则 ALTER TABLE ADD COLUMN

    ⚠️ 不删除列、不删除表（避免误操作丢数据）。

    参数:
        pk_col  - str (单主键) 或 list[str] (复合主键)
    """
    pk_cols = pk_col if isinstance(pk_col, list) else [pk_col]

    # 校验所有主键列都在 df 中
    for c in pk_cols:
        if c not in df.columns:
            raise ValueError(f"❌ 主键列 [{c}] 不在 df 列中：{list(df.columns)}")

    df_cols = [c for c in df.columns if c != "id"]

    # 1. 查表是否存在
    cur = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
        (table_name,),
    )
    table_exists = cur.fetchone() is not None

    if not table_exists:
        col_defs = [
            "id INTEGER PRIMARY KEY AUTOINCREMENT",
            '"report_date" TEXT NOT NULL',
            '"etl_time" TEXT NOT NULL',
        ]
        for col in df_cols:
            col_defs.append(f"{_safe_col_name(col)} {_infer_sqlite_type(df[col])}")

        # UNIQUE(report_date, pk_col1, pk_col2, ...)
        unique_keys = ['"report_date"'] + [_safe_col_name(c) for c in pk_cols]
        col_defs.append(f"UNIQUE({', '.join(unique_keys)})")

        sql = f"CREATE TABLE IF NOT EXISTS {table_name} (\n  " + ",\n  ".join(col_defs) + "\n)"
        conn.execute(sql)
        conn.commit()
        logging.info(
            f"✅ 创建表 {table_name}（{len(df_cols)} 列 + report_date + etl_time + "
            f"UNIQUE(report_date, {','.join(pk_cols)})）"
        )
    else:
        # ALTER TABLE ADD COLUMN（缺啥补啥）
        cur = conn.execute(f"PRAGMA table_info({table_name})")
        existing_cols = {row[1] for row in cur.fetchall()}
        added = []
        for col in df_cols:
            if col not in existing_cols:
                sql = f"ALTER TABLE {table_name} ADD COLUMN {_safe_col_name(col)} {_infer_sqlite_type(df[col])}"
                conn.execute(sql)
                added.append(col)
        if added:
            conn.commit()
            logging.info(f"✅ 表 {table_name} 新增列：{added}")


# ====================================================================
#  Upsert DataFrame
# ====================================================================

def upsert_df(
    conn: sqlite3.Connection,
    table_name: str,
    df: pd.DataFrame,
    report_date: str,
    pk_col: str,
    granularity: Optional[str] = None,
) -> int:
    """整 DataFrame upsert 到指定表。

    行为：
        - 先 delete 同 report_date（且同 granularity，如果提供）的全部记录
        - 然后 executemany 插入（避免冲突报错，且覆盖最快）
        - 返回插入行数

    ⚠️ 同 report_date + granularity 的所有记录会被先清掉再插入（Upsert 语义 = 覆盖更新）。

    参数:
        conn          sqlite3 连接
        table_name    业务表名
        df            数据（首列必须是主键列）
        report_date   业务日期 YYYY-MM-DD
        pk_col        主键列名
        granularity   粒度标识（day/month；None 表示业务本身无粒度区分，如三级渠道）
    返回:
        int - 实际插入的行数
    """
    if df.empty:
        logging.warning(f"⚠️ df 为空，跳过 {table_name}/{report_date}/{granularity or '-'} 入库")
        return 0

    # 1. 确保表存在
    ensure_table(conn, table_name, df, pk_col)

    # 2. 如果传了 granularity 但表里没这列，ALTER TABLE ADD COLUMN（一次）
    if granularity:
        cur = conn.execute(f"PRAGMA table_info({table_name})")
        existing_cols = {row[1] for row in cur.fetchall()}
        if "granularity" not in existing_cols:
            conn.execute(f'ALTER TABLE {table_name} ADD COLUMN "granularity" TEXT')
            conn.commit()
            logging.info(f"  └─ 给 {table_name} 加 granularity 列")

    # 3. 删除同 report_date + granularity 的全部记录（全覆盖）
    if granularity:
        cur = conn.execute(
            f"DELETE FROM {table_name} WHERE report_date = ? AND granularity = ?",
            (report_date, granularity),
        )
    else:
        cur = conn.execute(
            f"DELETE FROM {table_name} WHERE report_date = ?",
            (report_date,),
        )
    deleted = cur.rowcount
    if deleted > 0:
        logging.debug(f"  └─ 覆盖：删除 {table_name}/{report_date}/{granularity or '-'} 的 {deleted} 行")

    # 4. 构造插入数据
    etl_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    df_to_insert = df.copy()

    # 业务列转字符串（防 pandas 类型与 SQLite 类型不匹配）
    for col in df_to_insert.columns:
        df_to_insert[col] = df_to_insert[col].astype(str).replace({"nan": "", "<NA>": "", "None": ""})

    # 5. executemany 批量插入
    cols = list(df_to_insert.columns)
    col_names_sql = ", ".join([_safe_col_name(c) for c in cols])

    # 加 report_date / etl_time / (可选 granularity)
    extra_cols = ['"report_date"', '"etl_time"']
    extra_vals = [report_date, etl_time]
    if granularity:
        extra_cols.append('"granularity"')
        extra_vals.append(granularity)

    placeholders = ", ".join(["?"] * (len(cols) + len(extra_cols)))
    all_cols_sql = col_names_sql + ", " + ", ".join(extra_cols)

    rows = []
    for _, row in df_to_insert.iterrows():
        rows.append(tuple(list(row) + extra_vals))

    sql = f"INSERT INTO {table_name} ({all_cols_sql}) VALUES ({placeholders})"
    conn.executemany(sql, rows)
    conn.commit()

    logging.info(f"✅ 入库 {table_name}/{report_date}/{granularity or '-'}：{len(rows)} 行（覆盖 {deleted} 行）")
    return len(rows)


# ====================================================================
#  缺失日期查询
# ====================================================================

def get_existing_dates(
    conn: sqlite3.Connection,
    table_name: str,
    granularity: Optional[str] = None,
) -> set:
    """返回该表已有 report_date 的集合。

    granularity 提供时只查对应粒度的日期。
    """
    cur = conn.execute(
        f"SELECT DISTINCT report_date FROM {table_name}"
        + (" WHERE granularity = ?" if granularity else ""),
        (granularity,) if granularity else (),
    )
    return {row[0] for row in cur.fetchall()}


def get_existing_count(conn: sqlite3.Connection, table_name: str) -> int:
    """返回该表的总行数（调试用）。"""
    cur = conn.execute(f"SELECT COUNT(*) FROM {table_name}")
    return cur.fetchone()[0]


# ====================================================================
#  一键入口（业务类调用）
# ====================================================================

def save_to_db(
    biz_key: str,
    df: pd.DataFrame,
    report_date: str,
    granularity: Optional[str] = None,
) -> bool:
    """业务类调用入口：一键入库（推荐用法）。

    行为：
        - 自动判断是否启用 DB（is_db_enabled()）
        - 自动建表（ensure_table）
        - 自动推断主键列（infer_primary_key）
        - 自动 upsert（upsert_df）

    参数:
        biz_key       业务名（中文，如「商智关键词分析」）
        df            业务 DataFrame（已 Excel 后置干净）
        report_date   业务日期 YYYY-MM-DD
        granularity   粒度（day/month；None 表示业务本身无粒度区分）
    返回:
        bool - 成功入库 True / 跳过 False
    """
    if not is_db_enabled():
        logging.debug(f"  DB 未启用，跳过 {biz_key}/{report_date}")
        return False

    if df is None or df.empty:
        logging.warning(f"⚠️ df 为空，跳过 {biz_key}/{report_date}")
        return False

    # 主键推断
    pk_col = infer_primary_key(df, biz_key)
    if not pk_col:
        logging.error(f"❌ 无法推断主键列：{biz_key}，df.columns={list(df.columns)}")
        return False

    pk_cols = pk_col if isinstance(pk_col, list) else [pk_col]

    # 准备 df（确保所有主键列都非空）
    df_clean = df.copy()
    for c in pk_cols:
        if c in df_clean.columns:
            df_clean = df_clean[df_clean[c].astype(str).str.strip() != ""]
            df_clean = df_clean[df_clean[c].astype(str) != "nan"]

    if df_clean.empty:
        logging.warning(f"⚠️ 主键列 [{pk_cols}] 全空，跳过 {biz_key}/{report_date}")
        return False

    # 入库
    table_name = biz_key_to_table_name(biz_key)
    db_path = get_db_path()
    os.makedirs(os.path.dirname(db_path), exist_ok=True)

    try:
        conn = sqlite3.connect(db_path, timeout=30)
        n = upsert_df(conn, table_name, df_clean, report_date, pk_col, granularity)
        conn.close()
        return n > 0
    except sqlite3.Error as e:
        logging.error(f"❌ DB 入库失败 [{biz_key}/{report_date}]：{e}")
        return False


# ====================================================================
#  CLI（调试用）
# ====================================================================

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="db_utils 调试 CLI")
    parser.add_argument("--db-info", action="store_true", help="查看 DB 路径 + 是否启用")
    parser.add_argument("--list-tables", action="store_true", help="列出所有表 + 行数")
    parser.add_argument("--list-dates", type=str, help="查指定表的 report_date 集合（参数：业务名）")
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

    if args.db_info:
        print(f"DB 路径: {get_db_path()}")
        print(f"DB 启用: {is_db_enabled()}")
        print(f"DB 文件存在: {os.path.exists(get_db_path())}")

    elif args.list_tables:
        db_path = get_db_path()
        if not os.path.exists(db_path):
            print(f"DB 文件不存在：{db_path}")
        else:
            conn = sqlite3.connect(db_path)
            cur = conn.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
            tables = cur.fetchall()
            print(f"共 {len(tables)} 个表：")
            for (name,) in tables:
                cnt = get_existing_count(conn, name)
                print(f"  {name}: {cnt} 行")
            conn.close()

    elif args.list_dates:
        db_path = get_db_path()
        if not os.path.exists(db_path):
            print(f"DB 文件不存在：{db_path}")
        else:
            table_name = biz_key_to_table_name(args.list_dates)
            conn = sqlite3.connect(db_path)
            dates = sorted(get_existing_dates(conn, table_name))
            print(f"{args.list_dates} ({table_name}) 已有 {len(dates)} 个日期：")
            for d in dates:
                print(f"  {d}")
            conn.close()