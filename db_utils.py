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


# 业务白名单（2026-08-25）：infer_primary_key 返回 None → ensure_table 不加 UNIQUE
# 原因：这些业务每行已是「日期+计划+单元+SKU+地域」完整粒度，
# 单主键推断会过度约束（每天每计划多行被 UNIQUE 拦掉 → 1000 倍数据丢失）。
# 靠 save_to_db 的 upsert 语义（先 delete 同 shop_pin+stat_date 后 insert）保证幂等。
_BIZ_NO_UNIQUE = {
    "京准通快车自定义报表",       # 每天每计划多行（计划×单元×SKU×地域）
    "京准通快车订单效果明细",     # 每天每计划/订单多行
    "京准通全站营销单品计划",     # 每天每个商品计划多行
    "京准通全站营销单品推广效果", # 每天每计划多行
    "京麦订单明细_完整一键导出",  # 每订单含多商品（订单号+商品ID 复合，但简化不加 UNIQUE）
    "京麦售后明细_完整一键导出",  # 每售后单含多商品
}


def infer_primary_key(df: pd.DataFrame, biz_key: str = "") -> Optional[Union[str, list]]:
    """自动推断业务主键列名。

    优先级：
        1. 业务白名单（_BIZ_NO_UNIQUE）→ 返回 None（不加 UNIQUE，靠 upsert 兜底）
        2. config.xlsx 中「<业务短写> primary_key_columns」配置（复合主键，复数）
        3. config.xlsx 中「<业务短写> primary_key_column」配置（单主键）
        4. 列名包含关键词（ID / 编号 / 名称 / 关键词 / SKU）
        5. 第一列文本字段

    ⚠️ 复合主键配置示例（一级来源,二级来源,三级来源）：「三级渠道 primary_key_columns = 一级来源,二级来源,三级来源」

    返回:
        str  - 单主键列名
        list - 复合主键列名列表
        None - 推断失败 或 业务在白名单（不加 UNIQUE）
    """
    # 0. 白名单：返回 None（让 ensure_table 跳过 UNIQUE）
    if biz_key in _BIZ_NO_UNIQUE:
        return None

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


# 安全标识符正则（中英文+数字+下划线+中文括号+点号）
# 兼容 SQLite 表名 `biz_xxx` 和列名（含中文，如「一级来源」「加购客户数（SPU）」「平均停留时长(秒)」）
# 放宽规则（2026-08-21 项目23）：允许中文括号（）和英文括号()，全角空格，禁止 ; -- " ' 等 SQL 注入字符
_SAFE_IDENT_PATTERN = re.compile(r'^[A-Za-z0-9_\u4e00-\u9fa5（）()\.·\s]+$')


def _safe_col_name(col: str) -> str:
    """安全的列名（防 SQL 关键字 + 特殊字符）。

    规则：
        1. 空值 → 返回固定占位列名 "col_unnamed"
        2. 校验只含 中英文/数字/下划线，含其他字符（如 ; -- " '）抛 ValueError
        3. SQLite 关键字（order/group/select 等）和普通列名统一用双引号包裹
    """
    if not col:
        return '"col_unnamed"'
    # 安全校验：拒绝注入字符（AGENTS.md 第40条 不允许手动拼接 SQL 字符串的兜底防护）
    if not _SAFE_IDENT_PATTERN.match(str(col)):
        raise ValueError(
            f"❌ 列名含非法字符（仅允许中英文/数字/下划线）：{col!r}\n"
            f"   → 这可能是 SQL 注入风险，请检查 df.columns 来源"
        )
    return f'"{col}"'


# 非法列名字符 → 下划线（2026-08-25 M-33 修复）
# 京东报表列名含 / % - 等业务字符（如「省/直辖市」「点击率(%)」「等级1-5」），
# 直接入库会触发 _safe_col_name 白名单拒绝。此处仅替换非白名单字符，保留中英文/数字/括号。
_ILLEGAL_COL_CHARS = re.compile(r'[^A-Za-z0-9_一-龥（）()\.·\s]')

# SQLite 保留列（ensure_table/upsert_df 自动追加，列名大小写不敏感）
# ⚠️ 京东报表裸「ID」列与自动主键 id 视为同名 → CREATE 报 duplicate column name: ID
_SQLITE_RESERVED_COLS = {"id", "shop_pin", "stat_date", "report_date", "etl_time", "granularity"}


def _sanitize_columns(df: pd.DataFrame, biz_key: str = "") -> pd.DataFrame:
    """清洗 DataFrame 列名：非法字符替换为下划线。

    设计原则：
        - 复制 df（不改原对象），清洗列名后返回
        - SQL 层 _safe_col_name 仍严格校验（清洗后必通过）
        - 清洗后列名改变时打 WARNING（提示列名归一化）
        - 若清洗后列名重复（如「A/B」「A-B」都变 A_B），后续列名加序号避免冲突

    参数:
        df      原始 DataFrame（业务后置后的）
        biz_key 业务名（用于日志定位）

    返回:
        清洗列名后的 DataFrame（副本）
    """
    renamed = []
    seen = {}
    new_cols = []
    for c in df.columns:
        raw = str(c).strip()
        if not raw:
            base = "col_unnamed"
        else:
            base = _ILLEGAL_COL_CHARS.sub("_", raw)
        # SQLite 列名大小写不敏感：京东裸「ID」列与自动主键 id 撞名 → 加后缀保留业务列
        if base.lower() in _SQLITE_RESERVED_COLS:
            base = base + "_"
        # 处理重复列名（多文件合并时同名列，如「ID」「ID」）
        # 方案：第二个及以后加序号后缀（ID → ID_2 → ID_3）
        cnt = seen.get(base, 0)
        seen[base] = cnt + 1
        if cnt == 0:
            clean = base
        else:
            clean = f"{base}_{cnt + 1}"
        new_cols.append(clean)
        if clean != raw:
            renamed.append((raw, clean))
    if renamed:
        logging.warning(
            f"⚠️ [M-33] {biz_key} 列名清洗/去重 {len(renamed)} 处：{[(a, b) for a, b in renamed[:5]]}"
            + ("..." if len(renamed) > 5 else "")
        )
    # 直接按位置赋值列名（df.columns 可能重复，按位置最稳妥，返回副本）
    df2 = df.copy()
    df2.columns = new_cols
    return df2


def _validate_table_name(table_name: str) -> str:
    """校验表名是否符合规范（防 SQL 注入）。

    规则：
        - 必须匹配 `^biz_[a-z0-9_]+$`（小写英文+数字+下划线，biz_ 前缀）
        - 不符合抛 ValueError
        - 校验通过后原样返回（SQLite 表名不需要双引号）

    AGENTS.md 第37条：表名规范 `biz_{业务英文短名}`；第40条：不允许手动拼接 SQL。
    """
    if not isinstance(table_name, str) or not re.match(r'^biz_[a-z0-9_]+$', table_name):
        raise ValueError(
            f"❌ 表名不合规（必须匹配 biz_[a-z0-9_]+）：{table_name!r}\n"
            f"   → 表名应来自 biz_key_to_table_name() 的白名单 _BIZ_SLUG_OVERRIDE"
        )
    return table_name


def _apply_sqlite_perf_pragmas(conn: sqlite3.Connection):
    """开启 SQLite 写性能优化（2026-08-21 项目23 新增，针对京准通快车单日 4936 行大表）。

    ⚠️ WAL 模式 + synchronous=NORMAL 适合"批量写、可容忍丢最后 1 秒数据"的场景：
        - 我们的脚本是单进程跑业务，写完就 commit
        - 磁盘断电丢 1 秒数据是可接受的（可以从京东重新拉）
        - WAL 模式让读不阻塞写（虽然我们当前是单进程，但未来扩展友好）

    适用业务：
        - jzt_kuaiche（快车自定义报表，单日 4936 行）
        - jzt_kuaiche_order_effect（快车订单效果，单日 ~2000 行）
        - 其他行数 < 1000 的业务影响不大但开启无害
    """
    try:
        conn.execute("PRAGMA journal_mode = WAL")          # WAL 模式（不阻塞读）
        conn.execute("PRAGMA synchronous = NORMAL")       # 折中：写性能 ↑，崩溃丢 < 1 秒
        conn.execute("PRAGMA temp_store = MEMORY")        # 临时表放内存
        conn.execute("PRAGMA cache_size = -64000")        # 64MB 缓存
    except Exception as e:
        logging.warning(f"⚠️ SQLite PRAGMA 优化失败（不影响功能）：{e}")


def ensure_table(
    conn: sqlite3.Connection,
    table_name: str,
    df: pd.DataFrame,
    pk_col: Union[str, list, None],
) -> None:
    """动态建表 + 缺列补齐。

    行为：
        - 如果表不存在：CREATE TABLE，含 id / shop_pin / stat_date / report_date / etl_time / <df cols> /
          UNIQUE(shop_pin, stat_date, pk_col(s))
        - 如果表已存在：检查 df 列是否都存在，缺则 ALTER TABLE ADD COLUMN
          同时检查 shop_pin / stat_date 列是否存在（H-03 修复），缺则补齐

    ⚠️ H-03 多店隔离修复（2026-08-22）：所有表强制携带 shop_pin + stat_date，
        联合唯一键 (shop_pin, stat_date, 主键列) 支持 MySQL 迁移时建立
        UNIQUE KEY (shop_pin, stat_date)。

    ⚠️ 不删除列、不删除表（避免误操作丢数据）。

    参数:
        pk_col  - str (单主键) 或 list[str] (复合主键)
    """
    # 安全校验：表名必须合规（防 SQL 注入，AGENTS.md 第40条兜底）
    table_name = _validate_table_name(table_name)

    # 主键可能为 None（业务白名单 _BIZ_NO_UNIQUE：每行已是完整粒度，靠 upsert 兜底）
    pk_cols: list = []
    if pk_col:
        pk_cols = pk_col if isinstance(pk_col, list) else [pk_col]
        # 校验所有主键列都在 df 中
        for c in pk_cols:
            if c not in df.columns:
                raise ValueError(f"❌ 主键列 [{c}] 不在 df 列中：{list(df.columns)}")

    # 排除自动主键 id（大小写不敏感，SQLite 认为 id/ID 同名；兜底 _sanitize_columns 已重命名）
    df_cols = [c for c in df.columns if str(c).lower() != "id"]

    # 1. 查表是否存在
    cur = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
        (table_name,),
    )
    table_exists = cur.fetchone() is not None

    if not table_exists:
        col_defs = [
            "id INTEGER PRIMARY KEY AUTOINCREMENT",
            '"shop_pin" TEXT NOT NULL',     # H-03 新增：店铺主账号
            '"stat_date" TEXT NOT NULL',     # H-03 新增：业务统计日期（与 report_date 等价，独立列方便 MySQL 索引）
            '"report_date" TEXT NOT NULL',
            '"etl_time" TEXT NOT NULL',
        ]
        for col in df_cols:
            col_defs.append(f"{_safe_col_name(col)} {_infer_sqlite_type(df[col])}")

        # H-03 修复：联合唯一键改为 (shop_pin, stat_date, pk_col1, pk_col2, ...)
        # 多店并发跑同一业务时，通过 shop_pin 隔离，绝不串库
        # MySQL 迁移时此约束可直接对应 UNIQUE KEY (shop_pin, stat_date)
        # ⚠️ 2026-08-25：pk_cols 为空时跳过 UNIQUE（业务白名单，每行已是完整粒度，
        # 靠 upsert 语义 [先 delete 同 shop_pin+stat_date 后 insert] 保证幂等）
        if pk_cols:
            unique_keys = ['"shop_pin"', '"stat_date"'] + [_safe_col_name(c) for c in pk_cols]
            col_defs.append(f"UNIQUE({', '.join(unique_keys)})")

        sql = f"CREATE TABLE IF NOT EXISTS {table_name} (\n  " + ",\n  ".join(col_defs) + "\n)"
        conn.execute(sql)
        conn.commit()
        logging.info(
            f"✅ 创建表 {table_name}（{len(df_cols)} 列 + shop_pin + stat_date + report_date + etl_time + "
            f"UNIQUE(shop_pin, stat_date, {','.join(pk_cols)})）"
        )
    else:
        # ALTER TABLE ADD COLUMN（缺啥补啥）
        cur = conn.execute(f"PRAGMA table_info({table_name})")
        existing_cols = {row[1] for row in cur.fetchall()}
        added = []
        # H-03：确保 shop_pin / stat_date 列存在（旧表升级）
        for must_col in ("shop_pin", "stat_date"):
            if must_col not in existing_cols:
                sql = f'ALTER TABLE {table_name} ADD COLUMN "{must_col}" TEXT NOT NULL DEFAULT ""'
                try:
                    conn.execute(sql)
                    added.append(must_col)
                except Exception as e:
                    # SQLite 老版本不支持 NOT NULL DEFAULT，可降级为 TEXT
                    logging.warning(f"⚠️ 加 {must_col} 列失败（{e}），降级为 TEXT")
                    conn.execute(f'ALTER TABLE {table_name} ADD COLUMN "{must_col}" TEXT')
                    added.append(must_col)
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
    pk_col: Optional[Union[str, list]],
    granularity: Optional[str] = None,
) -> int:
    """整 DataFrame upsert 到指定表。

    行为：
        - 先 delete 同 report_date（且同 granularity，如果提供）的全部记录
        - 然后 executemany 插入（避免冲突报错，且覆盖最快）
        - 返回插入行数

    ⚠️ 同 report_date + granularity 的所有记录会被先清掉再插入（Upsert 语义 = 覆盖更新）。
    ⚠️ 2026-08-25：pk_col 可为 None（业务白名单），ensure_table 跳过 UNIQUE，靠 delete+insert 兜底。

    参数:
        conn          sqlite3 连接
        table_name    业务表名
        df            数据（首列必须是主键列）
        report_date   业务日期 YYYY-MM-DD
        pk_col        主键列名（str / list / None）
        granularity   粒度标识（day/month；None 表示业务本身无粒度区分，如三级渠道）
    返回:
        int - 实际插入的行数
    """
    # 安全校验：表名必须合规（防 SQL 注入，AGENTS.md 第40条兜底）
    table_name = _validate_table_name(table_name)

    if df.empty:
        logging.warning(f"⚠️ df 为空，跳过 {table_name}/{report_date}/{granularity or '-'} 入库")
        return 0

    # 1. 确保表存在（ensure_table 内部也会校验 table_name）
    ensure_table(conn, table_name, df, pk_col)

    # 2. 如果传了 granularity 但表里没这列，ALTER TABLE ADD COLUMN（一次）
    if granularity:
        cur = conn.execute(f"PRAGMA table_info({table_name})")
        existing_cols = {row[1] for row in cur.fetchall()}
        if "granularity" not in existing_cols:
            conn.execute(f'ALTER TABLE {table_name} ADD COLUMN "granularity" TEXT')
            conn.commit()
            logging.info(f"  └─ 给 {table_name} 加 granularity 列")

    # 3. 删除同 shop_pin + stat_date（+ granularity）的全部记录（全覆盖，H-03 多店隔离）
    #    H-03 修复：从环境变量 SHOP_PIN 读取当前店铺主账号
    #    多店并发跑同一业务时，按 shop_pin 隔离，A 店 DELETE 不会影响 B 店数据
    shop_pin = os.getenv("SHOP_PIN", "").strip()
    if not shop_pin:
        # 容错：从环境变量 SHOP_ID 推导（店铺短名兜底，但可能与京东 pin 不同）
        shop_id = os.getenv("SHOP_ID", "")
        shop_pin = shop_id.replace("箱包旗舰店", "") if shop_id else "UNKNOWN"
        logging.warning(
            f"⚠️ [H-03] SHOP_PIN 未设置，临时用 SHOP_ID 推导（{shop_pin}）。\n"
            f"       建议 set SHOP_PIN=FYA8888 / miyo-周 / ota8888"
        )
    if granularity:
        cur = conn.execute(
            f"DELETE FROM {table_name} WHERE shop_pin = ? AND stat_date = ? AND granularity = ?",
            (shop_pin, report_date, granularity),
        )
    else:
        cur = conn.execute(
            f"DELETE FROM {table_name} WHERE shop_pin = ? AND stat_date = ?",
            (shop_pin, report_date),
        )
    deleted = cur.rowcount
    if deleted > 0:
        logging.debug(
            f"  └─ 覆盖：删除 {table_name}/shop_pin={shop_pin}/stat_date={report_date}"
            f"/{granularity or '-'} 的 {deleted} 行"
        )

    # 4. 构造插入数据
    etl_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    df_to_insert = df.copy()

    # 业务列转字符串（防 pandas 类型与 SQLite 类型不匹配）
    for col in df_to_insert.columns:
        df_to_insert[col] = df_to_insert[col].astype(str).replace({"nan": "", "<NA>": "", "None": ""})

    # 5. executemany 批量插入
    cols = list(df_to_insert.columns)
    col_names_sql = ", ".join([_safe_col_name(c) for c in cols])

    # H-03 修复：加 shop_pin / stat_date / report_date / etl_time / (可选 granularity)
    extra_cols = ['"shop_pin"', '"stat_date"', '"report_date"', '"etl_time"']
    extra_vals = [shop_pin, report_date, report_date, etl_time]
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

    logging.info(
        f"✅ 入库 {table_name}/shop_pin={shop_pin}/stat_date={report_date}/{granularity or '-'}："
        f"{len(rows)} 行（覆盖 {deleted} 行）"
    )
    return len(rows)


# ====================================================================
#  缺失日期查询
# ====================================================================

def get_existing_dates(
    conn: sqlite3.Connection,
    table_name: str,
    granularity: Optional[str] = None,
    shop_pin: Optional[str] = None,
) -> set:
    """返回该表已有 report_date 的集合。

    参数:
        conn          sqlite3 连接
        table_name    业务表名（biz_xxx）
        granularity   粒度（day/month；None 表示业务本身无粒度区分）
        shop_pin      店铺主账号（H-03 多店隔离）；None 表示不按店铺过滤（兼容旧调用方）

    返回:
        set - 该表在指定店铺+粒度下的 report_date 集合

    ⚠️ H-03 多店隔离 + M-03 修复（2026-08-24 审计）：fill_missing 必须按 shop_pin 过滤
        背景：原版跨店聚合后，MIYO 已有的日期会被认为 FYA 也有 → FYA 永远不会补录
        修复：fill_missing.py 显式传入当前店铺的 shop_pin
    """
    # 安全校验：表名必须合规（防 SQL 注入，AGENTS.md 第40条兜底）
    table_name = _validate_table_name(table_name)

    # 构造 WHERE 条件（按优先级叠加）
    where_clauses = []
    params = []
    if shop_pin:
        where_clauses.append("shop_pin = ?")
        params.append(shop_pin)
    if granularity:
        where_clauses.append("granularity = ?")
        params.append(granularity)

    sql = f"SELECT DISTINCT report_date FROM {table_name}"
    if where_clauses:
        sql += " WHERE " + " AND ".join(where_clauses)

    try:
        cur = conn.execute(sql, tuple(params))
        return {row[0] for row in cur.fetchall()}
    except sqlite3.OperationalError as e:
        # 表不存在 → 返回空集合（与 fill_missing.py 行为一致）
        # 背景：首次跑某业务时表还不存在，按"全缺失"处理是合理的
        if "no such table" in str(e).lower():
            return set()
        raise


def mark_empty_date(biz_key: str, report_date: str, shop_pin: Optional[str] = None) -> None:
    """记录已成功请求但无业务行的日期，避免空报表被重复拉取。"""
    pin = (shop_pin or os.getenv("SHOP_PIN", "").strip() or os.getenv("SHOP_ID", "UNKNOWN")).strip()
    db_path = get_db_path()
    os.makedirs(os.path.dirname(db_path), exist_ok=True)
    conn = sqlite3.connect(db_path, timeout=30)
    try:
        conn.execute(
            "CREATE TABLE IF NOT EXISTS biz_empty_dates ("
            "biz_key TEXT NOT NULL, shop_pin TEXT NOT NULL, report_date TEXT NOT NULL, "
            "etl_time TEXT NOT NULL, PRIMARY KEY (biz_key, shop_pin, report_date))"
        )
        conn.execute(
            "INSERT OR REPLACE INTO biz_empty_dates "
            "(biz_key, shop_pin, report_date, etl_time) VALUES (?, ?, ?, ?)",
            (biz_key, pin, report_date, datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        )
        conn.commit()
    finally:
        conn.close()


def get_empty_dates(biz_key: str, shop_pin: Optional[str] = None) -> set:
    """返回指定业务/店铺已确认为空的日期集合。"""
    db_path = get_db_path()
    if not os.path.exists(db_path):
        return set()
    pin = (shop_pin or os.getenv("SHOP_PIN", "").strip() or os.getenv("SHOP_ID", "UNKNOWN")).strip()
    conn = sqlite3.connect(db_path, timeout=30)
    try:
        try:
            rows = conn.execute(
                "SELECT report_date FROM biz_empty_dates WHERE biz_key=? AND shop_pin=?",
                (biz_key, pin),
            ).fetchall()
        except sqlite3.OperationalError as e:
            if "no such table" in str(e).lower():
                return set()
            raise
        return {r[0] for r in rows}
    finally:
        conn.close()


def get_existing_count(conn: sqlite3.Connection, table_name: str) -> int:
    """返回该表的总行数（调试用）。"""
    # 安全校验：表名必须合规（防 SQL 注入，AGENTS.md 第40条兜底）
    table_name = _validate_table_name(table_name)
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
        - 自动建表（ensure_table，含 shop_pin / stat_date 字段）
        - 自动推断主键列（infer_primary_key）
        - 自动 upsert（upsert_df，按 shop_pin + stat_date 维度覆盖，H-03 多店隔离）

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

    # 列名清洗（2026-08-25 M-33 修复）：京东真实列名可能含 / % - 等字符
    # 如「省/直辖市」「点击率(%)」「商家备注等级（等级1-5...）」
    # 这些字符不在 _SAFE_IDENT_PATTERN 白名单内，直接入库会抛 ValueError
    # 方案：复制 df，把非法字符替换为下划线（防注入 + 兼容京东列名双满足）
    # SQL 层 _safe_col_name 仍严格校验（清洗后必然通过）
    df = _sanitize_columns(df, biz_key)

    # 主键推断
    pk_col = infer_primary_key(df, biz_key)
    if pk_col is None:
        if biz_key in _BIZ_NO_UNIQUE:
            # 业务白名单：每行已是「日期+计划+单元+SKU+地域」完整粒度，
            # 加 UNIQUE 会过度约束丢失数据；靠 upsert 语义保证幂等。
            logging.info(f"  ℹ️ {biz_key} 主键推断返回 None（业务白名单），不加 UNIQUE，靠 upsert 兜底")
            pk_cols = []
        else:
            # 非白名单业务确实推断失败（如 df 全是数值列） → 报错
            logging.error(f"❌ 无法推断主键列：{biz_key}，df.columns={list(df.columns)}")
            return False
    else:
        pk_cols = pk_col if isinstance(pk_col, list) else [pk_col]

    # 准备 df（仅当有主键列时过滤空值）
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

    # 项目24（2026-08-22 新增）：DB 入库成功后自动联动 Excel 总表缓存
    # 业务类零侵入：save_to_db 内部自动调 excel_master.collect
    # 这样所有调 save_to_db 的地方都自动同步到总表，不用改业务类
    _excel_collect_hook = None
    try:
        import excel_master as _em
        _excel_collect_hook = _em.collect
    except ImportError:
        pass  # excel_master 不存在时不阻塞 DB 入库

    conn = None
    try:
        conn = sqlite3.connect(db_path, timeout=30)
        # 开启 SQLite 写性能优化（2026-08-21 项目23）：WAL + NORMAL synchronous
        _apply_sqlite_perf_pragmas(conn)
        n = upsert_df(conn, table_name, df_clean, report_date, pk_col, granularity)

        # 项目24：DB 入库成功后联动 Excel 总表（df 用清理后的 df_clean）
        if n > 0 and _excel_collect_hook is not None:
            try:
                _excel_collect_hook(
                    biz_key=biz_key,
                    df=df_clean,
                    report_date=report_date,
                    granularity=granularity,
                )
            except Exception as e:
                # collect 失败不影响 DB 入库结果（Excel 只是快照）
                logging.warning(
                    f"⚠️ [ExcelMaster] collect 失败（不影响 DB 入库）：{biz_key}/{report_date}：{e}"
                )

        return n > 0
    except sqlite3.Error as e:
        logging.error(f"❌ DB 入库失败 [{biz_key}/{report_date}]：{e}")
        return False
    except Exception as e:
        # H-08 / M-01 修复：非 sqlite3.Error 也要记录日志，避免吞报错
        logging.error(f"❌ DB 入库异常 [{biz_key}/{report_date}]：{type(e).__name__}: {e}")
        return False
    finally:
        # M-01 修复：所有路径都关闭连接，防止泄漏
        if conn is not None:
            try:
                conn.close()
            except Exception as e:
                logging.warning(f"⚠️ conn.close() 失败：{e}")


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
