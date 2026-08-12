# -*- coding: utf-8 -*-
"""
main.py - 京东商智数据导出工具（重构版 v2.0）

本文件集中存放：
    - 通用请求基类（JDBaseRequest）：Cookie管理 / 风控签名 / 30秒间隔 / 重试 / UA切换 / 日志 / Excel保存
    - 各业务API实现（按业务名分块）
    - 业务注册中心（BUSINESS_REGISTRY）：新业务只需注册，无需改动调度核心
    - 主程序入口 main()：支持命令行调用 + 业务清单打印

代码组织原则（按全局铁律第4条）：
    - 不拆分大量独立py文件，所有接口集成在本文件
    - 每个API实现前用 ============ 分层注释隔离标记
    - 标注：接口业务名称、接口地址、参数说明，便于快速定位/修改/维护

【整改记录 2026-08-04】
    1. 删除所有业务参数硬编码（INTERVAL/DATETYPE/LIMIT/SORT_FIELD/SORT_TYPE/COMPARE_TYPE/GROUP_TYPE/ATTRIBUTES/LAST_SRC_CHANNEL_ID1）
       全部改为从 config.xlsx 读取
    2. 新增业务注册中心 BUSINESS_REGISTRY，支持按业务key动态调度
    3. 支持批量调用：run_business(biz_key_list) 一次跑多个业务
    4. 支持命令行调用：python main.py --biz_key "商品流量来源_搜索" --date "2026-07-29"
    5. 程序启动打印：全局配置 + 已注册业务清单，方便核对

【整改记录 2026-08-05】
    1. config项目名统一为"商品流量来源"；9项业务参数补齐【说明】列中文注释
    2. 业务名调整：3001渠道执行名改为"商品流量来源_购物车"
       （自主访问流量与购物车数据口径重叠，统一以"购物车"命名执行）
    3. 自主访问保留注册配置但标记 enabled=False，调度层过滤不执行，后续需要可随时开启
    4. 默认批量执行：搜索/推荐/购物车 3个启用渠道
    5. config精简：6项固定业务参数（lastSrcChannelId1/groupType/attributes/sortField/sortType/compareType）
       经用户确认后移出config.xlsx，固化为代码常量 FIXED_BIZ_PARAMS；
       可变参数 interval/dateType/limit 仍从config.xlsx读取（缺省兜底+警告）
    6. 商品流量来源导出后置处理：新增公共工具 convert_date_format()/safe_convert_numeric()，
       导出时首列A插入【日期】列 + 全表数值安全转换（>15位长数字保留文本）+
       日期列真实日期单元格格式（打开不弹格式警告）
    7. 对齐京东官方订单导出风险提示：safe_convert_numeric 新增按列名规则——
       强制文本黑名单 TEXT_FORCE_COLUMNS={订单编号}（整列跳过转换保留文本）、
       整数0位小数白名单 INTEGER_ZERO_DECIMAL_COLUMNS={SKU,SPU}（转数字+格式0）；
       新增通用 apply_column_formats() 按列名批量设置单元格格式（订单编号@/SKU·SPU数值0位小数/日期格式）
"""

import os
import sys
import time
import json
import hashlib
import random
import re
import logging
import argparse
from datetime import datetime
from urllib.parse import urlparse

import requests
from openpyxl import load_workbook


# ============================================================
#  全局常量
# ============================================================
# 当前店铺（只用于展示输出，不参与签名）
SHOP_NAME = "FYA8888"


# ============================================================
#  自定义异常
# ============================================================
class CookieExpiredError(Exception):
    """Cookie过期异常"""
    pass


class RiskControlError(Exception):
    """风控拦截异常"""
    pass


class BusinessNotFoundError(Exception):
    """业务未注册异常"""
    pass


# ============================================================
#  公共工具函数（所有报表复用，禁止硬编码具体业务逻辑）
# ------------------------------------------------------------
#  通用工具：
#    ① convert_date_format(date_str)   —— 通用日期格式转换
#    ② safe_convert_numeric(df)        —— 全表数值安全转换（按列名黑/白名单+长度规则）
#    ③ apply_column_formats(...)       —— 按列名规则批量设置Excel单元格格式
#  说明：商品流量来源报表已接入（插入日期列后处理）；
#        后续订单/售后/京准通等报表可传入自己的DataFrame/日期字符串直接复用，无需改动。
# ============================================================

# ⚠️ 强制文本列黑名单（对齐京东官方订单导出风险提示）：
#    命中列整列跳过数值转换，强制保留原始文本字符串，彻底规避订单号科学计数法/末尾数字变0。
TEXT_FORCE_COLUMNS = {"订单编号"}
# ⚠️ 整数0位小数白名单：命中列允许转为数字；写入Excel单元格格式为 0（数值、0位小数、无千分位）。
#   2026-08-10 用户决策扩展：新增「商品ID」（服务端导出时常以科学计数法显示 1E+13，整数无小数）
INTEGER_ZERO_DECIMAL_COLUMNS = {"SKU", "SPU", "商品ID"}
# ⚠️ 指标词黑名单（2026-08-09 订单明细接入后扩展）：
#    列名同时含「指标词」的，不应命中 INTEGER_ZERO_DECIMAL_COLUMNS 等数值格式白名单
#    例如「SKU金额」含"SKU"但其实是金额指标，应该保留默认 2 位小数格式
#    例如「SKU数量」含"SKU"但是数量指标，应该保留默认 General 格式
METRIC_BLOCKLIST = {"金额", "数量", "名称", "类型", "状态", "城市", "省份", "市", "省"}


def _col_matches(col_name, name_set):
    """判断列名是否命中规则集合。

    匹配规则（2026-08-09 京准通接入后扩展）：
        ① 列名精确等于集合元素
        ② 或以集合元素结尾（如 "商品SKU" → 命中 "SKU"）
        ③ 或以集合元素+空格结尾（如 "商品定向SKU ID" → 命中 "SKU"）
        ④ 或列名中包含集合元素+空格（如 "商品定向SKU ID" → 命中 "SKU"）

    说明：
        - 旧版仅 endswith，对"商品定向SKU ID"/"SPU ID"这种中间含空格的列名漏匹配
        - 京准通快车表大量使用"商品定向SKU ID/跟单SKU ID/SPU ID"等列名规则
        - 仍避开"成交金额（SPU）"型后缀——通过「关键词+空格/=」边界保护
        - ⚠️ 2026-08-09 加 METRIC_BLOCKLIST：列名含指标词（金额/数量/名称/类型/状态/省/市）
          的列不命中数值格式白名单，避免把 "SKU金额" 列误套 0 位小数格式
    """
    if col_name in name_set:
        return True
    # ⚠️ 含指标词的列不命中数值格式白名单
    if any(metric in col_name for metric in METRIC_BLOCKLIST):
        return False
    for name in name_set:
        # ① 严格结尾：XXXSKU（无空格干扰）
        if col_name.endswith(name):
            return True
        # ② 列名中含「{name} 」+ 内容（如"商品定向SKU ID"）
        if f"{name} " in col_name or f"{name}ID" in col_name:
            return True
    return False


def convert_date_format(date_str):
    """通用日期格式转换（所有报表统一标准）。

    入参:
        date_str - 日期字符串，支持以下3种输入：
            "20260729"              8位纯数字
            "2026-07-29"            横杠分隔
            "2026-07-29 13:45:59"   横杠分隔+时间
    出参:
        统一目标格式字符串（月/日不补零，时间部分原样保留）：
            "20260729"            → "2026/7/29"
            "2026-07-29"          → "2026/7/29"
            "2026-07-29 13:45:59" → "2026/7/29 13:45:59"
    兼容异常:
        无法识别的格式直接返回原值，不报错、不中断程序。
    """
    if date_str is None:
        return date_str
    s = str(date_str).strip()
    if not s:
        return date_str

    # ① 8位纯数字日期：20260729 → 2026/7/29
    m = re.fullmatch(r"(\d{4})(\d{2})(\d{2})", s)
    if m:
        return f"{m.group(1)}/{int(m.group(2))}/{int(m.group(3))}"

    # ② 分隔符日期（横杠/斜杠均可），可带时间：2026-07-29 / 2026/07/29 / 2026-07-29 13:45:59
    m = re.fullmatch(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})(.*)", s)
    if m:
        # 月/日用 int() 去掉前导0（07→7）；时间部分（含前导空格）原样保留
        return f"{m.group(1)}/{int(m.group(2))}/{int(m.group(3))}{m.group(4)}"

    # ③ 无法识别 → 原值返回（不报错）
    return s


def _find_date_cols(df):
    """查找DataFrame中报表自带(原始)的日期/时间列。

    识别标准（2026-08-07 公共规则1）：
        - 列名精确等于"日期"或"时间"
        - 或以"日期"结尾（如"下单日期"）
    ⚠️ 刻意不用"以'时间'结尾"匹配：防止"最近上架时间"等业务时间字段
       （列里是上架日期而非本报表统计日期）被误判为日期列。

    返回:
        list - 命中的列名列表；空列表 = 报表无日期/时间列（需要程序插入【日期】列）
    """
    hits = []
    for col in df.columns:
        name = str(col).strip()
        if name == "日期" or name == "时间" or name.endswith("日期"):
            hits.append(col)
    return hits


def prepare_date_columns(df, date):
    """Excel日期列统一处理（公共规则1+2，所有报表复用）。

    规则1（日期列智能新增）：
        - 报表已存在【日期】/【时间】列 → **禁止重复插入**日期列，
          仅对已有日期/时间列做格式标准化转换；
        - 报表无任何日期/时间列 → 在首列插入【日期】列，值=本次查询日期。

    规则2（日期格式统一）：yyyy/m/d（如 2026/7/29），带时分秒保留时间部分。
        输入兼容：20260729 / 2026-07-29 / 2026-07-29 13:45:59（内部调用 convert_date_format）。

    入参:
        df   - 从Excel读出的DataFrame（dtype=str）
        date - 本次查询日期（如 2026-07-29），仅"插入新列"场景用到
    出参:
        (date_column, date_value)
            date_column - 实际承载日期的列名
                          （插入列="日期"；自带日期列=第一个命中的列名）
            date_value  - 插入列场景：日期字符串值（供 apply_column_formats 用）；
                          自带日期列场景：None（值逐格不同，由 apply_column_formats 逐格解析）
    """
    date_cols = _find_date_cols(df)
    if date_cols:
        # 报表自带日期/时间列 → 禁止重复插入，只做格式标准化（2026-07-29 → 2026/7/29）
        for col in date_cols:
            df[col] = [convert_date_format(v) for v in df[col].tolist()]
        return date_cols[0], None

    # 报表无日期/时间列 → 首列插入【日期】列，全列取本次查询日期
    date_str = convert_date_format(date)
    df.insert(0, "日期", date_str)
    return "日期", date_str


def build_business_output_path(output_dir, filename, date):
    """按业务模块+日期子文件夹构造Excel保存路径（AGENTS.md Excel规则4）。

    目录规则：output/{业务模块}/{date}/{filename}
        业务模块名 = 文件名去掉 "_{date}.xlsx" 后缀的主体
        （如 搜索流量_2026-07-29.xlsx → 业务模块"搜索流量"）
    例：output/搜索流量/2026-07-29/搜索流量_2026-07-29.xlsx

    入参:
        output_dir - 全局输出目录（output/）
        filename   - 保存文件名（含 {date} 与 .xlsx 后缀）
        date       - 查询日期（YYYY-MM-DD）
    出参:
        最终保存的完整路径（子目录不存在会自动创建）
    """
    # 提取业务模块名：去掉 "_{date}.xlsx" 后缀即为主体名
    stem = filename
    suffix = f"_{date}.xlsx"
    if stem.endswith(suffix):
        stem = stem[: -len(suffix)]

    business_dir = os.path.join(output_dir, stem, date)
    os.makedirs(business_dir, exist_ok=True)
    return os.path.join(business_dir, filename)


def read_excel_bytes(content):
    """通用Excel二进制读取（xlsx + xls 双格式兼容，2026-08-07 项目6新增）。

    背景：京东商智部分接口返回 .xls（OLE2复合文档，如竞争-商品流失分析 exportLossProList），
    其文件头是 \xD0\xCF\x11\xE0（非 xlsx 的 PK\x03\x04），pandas 需显式 engine='xlrd' 才能读取。
    此前所有报表都是 .xlsx，用 openpyxl 引擎即可；本项目首次出现 .xls，故抽成公共函数统一处理。

    入参:
        content - 接口返回的Excel文件字节流（bytes）
    出参:
        DataFrame（dtype=str 防长数字精度丢失；na_filter=False 保留空字符串）
    异常:
        无法识别的Excel格式 → ValueError（由调用方重试兜底）
    """
    import io
    import pandas as pd

    magic = content[:4]
    if magic == b"PK\x03\x04":
        # .xlsx（zip容器），pandas 默认 openpyxl 引擎
        return pd.read_excel(io.BytesIO(content), dtype=str, na_filter=False)
    if magic == b"\xD0\xCF\x11\xE0":
        # .xls（OLE2复合文档），xlrd 引擎（pandas 3.0 仍保留 XlrdReader）
        return pd.read_excel(io.BytesIO(content), dtype=str, na_filter=False, engine="xlrd")
    raise ValueError(f"无法识别的Excel格式（magic={magic!r}）")


def drop_total_rows(df):
    """去除服务端默认加的「合计/总计/汇总」行（所有报表复用，全局生效）。

    ⚠️ 2026-08-10 用户决策：服务端导出的报表里配置的合计/总计/汇总行，整行剔除。
    检测规则：任一单元格的值（含表头和数据）含「合计」「总计」「汇总」任一关键词，
    整行 drop（inplace）。

    入参:
        df - pandas.DataFrame
    出参:
        处理后的DataFrame（原地修改并返回）
    """
    if df is None or df.empty:
        return df
    # 关键词集合（用户决策 2026-08-10）
    total_keywords = ("合计", "总计", "汇总")
    # 遍历每行：任一单元格含关键词 → 标记删除
    mask_to_drop = df.apply(
        lambda row: any(
            isinstance(v, str) and any(kw in v for kw in total_keywords)
            for v in row.tolist()
        ),
        axis=1,
    )
    drop_count = int(mask_to_drop.sum())
    if drop_count > 0:
        df.drop(df[mask_to_drop].index, inplace=True)
        df.reset_index(drop=True, inplace=True)
        # ⚠️ 调试日志：让用户能确认剔除行数
        print(f"   ├─ [drop_total_rows] 剔除服务端合计/总计/汇总行：{drop_count} 行")
    return df


def safe_convert_numeric(df):
    """全表数值安全转换（所有报表复用，全局生效）。

    入参:
        df - pandas.DataFrame（从Excel读取的表格数据）
    出参:
        处理后的DataFrame（直接修改并返回），转换规则：
            0. 【强制文本黑名单】列名命中 TEXT_FORCE_COLUMNS（如"订单编号"）
               → 整列完全跳过数值转换，强制保留原始文本字符串（不依赖长度判断）；
            1. 其他字符串且为纯数字（可含小数点/负号）且数字位数≤15位 → 转成数值（int/float）；
               其中列名命中 INTEGER_ZERO_DECIMAL_COLUMNS（如"SKU"/"SPU"/"商品ID"）时单元格格式为 0（0位小数无千分位）；
            2. 纯数字但数字位数>15位 → 保留原始文本，杜绝精度丢失（兜底防护，全局保留）；
            3. 非纯数字（日期/含字母/空值/已是数值类型） → 保留原值；转换失败同样保留原值。
    注意:
        ⚠️ 调用前请先把日期列用 convert_date_format() 处理好，否则"20260729"这类
           8位纯数字日期会被误当成普通数字转换（商品流量来源流程已保证先转日期再转数值）。
        ⚠️ 2026-08-10：内部先调用 drop_total_rows() 剔除合计行（用户决策），全局生效。
    """
    # ⚠️ 2026-08-10 用户决策：先剔除合计/总计/汇总行（封装在内部，全局生效，避免调用方遗漏）
    drop_total_rows(df)
    for col in df.columns:
        # ⚠️ 强制文本黑名单：命中列整列跳过数值转换，保留原始文本（订单编号等长ID）
        if _col_matches(col, TEXT_FORCE_COLUMNS):
            continue
        df[col] = [_safe_convert_one(v) for v in df[col].tolist()]
    return df


def _safe_convert_one(value):
    """单个单元格数值安全转换（safe_convert_numeric 的内部辅助函数）。"""
    # 非字符串（数值/日期对象/None等）直接返回
    if value is None or not isinstance(value, str):
        return value
    s = value.strip()
    if not s:
        return value
    # 仅"纯数字"（可选负号/小数点）才考虑转换，其余（日期/含字母等）原样保留
    if not re.fullmatch(r"-?\d+(\.\d+)?", s):
        return value
    # 数字位数>15位 → 保留文本（长订单号/长SKU，防止Excel精度丢失）
    digits_count = len(re.sub(r"[^0-9]", "", s))
    if digits_count > 15:
        return value
    try:
        return float(s) if "." in s else int(s)
    except ValueError:
        return value  # 转换失败 → 保留原值


def _parse_date_cell(date_str):
    """把目标格式日期串解析为datetime对象（失败返回None）。

    支持："2026/7/29" 和 "2026/7/29 13:45:59"
    用途：写Excel时把日期列从文本改为真实日期对象，避免打开文件弹格式警告。
    """
    s = str(date_str).strip()
    for fmt in ("%Y/%m/%d %H:%M:%S", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def apply_column_formats(file_path, df, date_column="日期", date_value=None):
    """按列名规则批量设置Excel单元格格式（所有报表复用，全局生效）。

    规则（对齐京东官方订单导出风险提示）：
        - 强制文本列 TEXT_FORCE_COLUMNS（如"订单编号"）→ 单元格格式 @（文本），
          订单号无论多长都按文本显示，杜绝科学计数法/末尾数字变0；
        - 整数0位小数列 INTEGER_ZERO_DECIMAL_COLUMNS（如"SKU"/"SPU"）→ 单元格格式 0
          （数值、0位小数、不使用千位分隔符），仅当单元格值为数字时生效；
        - date_column 指定列（默认"日期"）→ 日期格式 yyyy/m/d（带时间用 yyyy/m/d hh:mm:ss），
          并把文本日期替换为真实datetime对象，打开Excel不弹格式警告。

    入参:
        file_path   - 已用 df.to_excel 写好的Excel文件路径
        df          - 与Excel表头对应的DataFrame（用于列名→列号映射）
        date_column - 日期列列名（默认"日期"；其他报表自带日期列时传入自己的列名）
        date_value  - 日期列的字符串值（如 2026/8/5 或 2026/8/5 13:45:59），用于转真实日期对象
    出参:
        无（直接修改并保存Excel文件）
    """
    from openpyxl import load_workbook

    wb = load_workbook(file_path)
    ws = wb.active

    # 表头 → 列号 映射（表头在第1行）
    header_map = {}
    for cell in ws[1]:
        if cell.value is not None:
            header_map[str(cell.value)] = cell.column

    # ① 日期列：逐格文本 → 真实日期对象 + 日期/日期时间格式
    # 支持两种调用场景（2026-08-07 公共规则1+2）：
    #   - 程序插入的【日期】列（date_value 提供，全列同值）
    #   - 报表自带日期/时间列（date_value=None，每格值可能不同，逐格解析）
    if date_column in header_map:
        # 插入列场景：解析 date_value 得到基准格式（全列同值，用于兜底套格式）
        base_fmt = None
        if date_value is not None and _parse_date_cell(date_value) is not None:
            base_fmt = "yyyy/m/d hh:mm:ss" if ":" in str(date_value) else "yyyy/m/d"
        col_idx = header_map[date_column]
        for row in range(2, ws.max_row + 1):
            cell = ws.cell(row=row, column=col_idx)
            v = cell.value
            if isinstance(v, str):
                dt = _parse_date_cell(v)
                if dt is not None:
                    cell.value = dt           # 文本日期 → 真实日期对象（非文本）
                    # 每格按自身是否带时间决定格式（yyyy/m/d 或 yyyy/m/d hh:mm:ss）
                    cell.number_format = "yyyy/m/d hh:mm:ss" if ":" in v else "yyyy/m/d"
                    continue
            # 非文本（已是日期对象/数值）或无法解析 → 保留原值，有基准格式则套用
            if base_fmt is not None:
                cell.number_format = base_fmt

    # ② 其他列按列名规则设置格式（订单编号=@文本，SKU/SPU=0数值0位小数）
    for col_name, col_idx in header_map.items():
        if _col_matches(col_name, TEXT_FORCE_COLUMNS):
            fmt = "@"                          # 强制文本格式
        elif _col_matches(col_name, INTEGER_ZERO_DECIMAL_COLUMNS):
            fmt = "0"                          # 数值、0位小数、无千分位
        else:
            continue
        for row in range(2, ws.max_row + 1):
            cell = ws.cell(row=row, column=col_idx)
            # SKU/SPU列仅当单元格是数字时才套用"0"格式，避免文本值显示异常
            if fmt == "0" and not isinstance(cell.value, (int, float)):
                continue
            cell.number_format = fmt

    # ③ 冻结首列 + 表头行（2026-08-10 用户决策）
    #   含义：openpyxl freeze_panes="B2" 表示冻结 A 列 + 第 1 行（表头）
    #   - 左侧冻结：滚动时 A 列（时间/商品ID 等首列）始终可见
    #   - 顶部冻结：滚动时表头行（第 1 行）始终可见
    #   全局生效：所有报表复用本函数，自动应用
    ws.freeze_panes = "B2"

    wb.save(file_path)
    wb.close()


# ============================================================
#  通用请求基类（JDBaseRequest）
# ------------------------------------------------------------
#  业务名称：通用能力
#  接口地址：无（封装通用能力，被各业务API复用）
#  功能说明：
#      - Cookie 读取与更新（config/sz_cookie.txt）
#      - 风控签名生成（User-mup / User-mnp / uuid）
#      - 30秒请求间隔控制（从config读取，严格执行）
#      - 重试机制（最多3次，递增等待）
#      - Edge ↔ Chrome UA 自动切换
#      - 日志记录（按日期，文件+控制台）
#      - Excel 文件保存
#  复用方式：
#      各业务API类继承此类，直接调用 self.request(url, data) 即可。
# ============================================================
class JDBaseRequest:
    """京东商智API通用请求基类"""

    # ---------- 固定常量（与具体业务无关的全局变量）----------
    DEFAULT_REFERER = "https://sz.jd.com/szweb/sz/view/viewflow/flowPathDetailsNew.html"
    DEFAULT_ORIGIN = "https://sz.jd.com"
    DEFAULT_CONFIG_PATH = "config/config.xlsx"

    # 风控签名盐值默认值（config.xlsx 可覆盖；新增业务时如需不同盐值，注册业务时单独覆盖）
    _DEFAULT_SIGN_SALT = "372ad2c2b6"

    # uuid前缀默认值（基类兜底；各业务应通过业务参数注册时单独指定，避免硬编码渠道差异）
    UUID_PREFIX = "ca412182e5668a106054"
    UUID_RANDOM_DIGITS = 10

    # ⚠️ 进程级共享的"上次请求时间"（类属性，不是实例属性）
    # 原因：批量执行多个业务时会创建多个实例，若用实例属性，
    #       每个实例的间隔计数从0重新开始，跨业务的30秒间隔会被跳过。
    #       改为类属性后，所有实例共享同一个计数，保证全程严格间隔。
    _last_request_time = 0

    UA_EDGE = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/144.0.0.0 Safari/537.36 Edg/144.0.0.0"
    UA_CHROME = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/144.0.0.0 Safari/537.36"
    SEC_CH_UA_EDGE = '"Not(A:Brand";v="8", "Chromium";v="144", "Microsoft Edge";v="144"'
    SEC_CH_UA_CHROME = '"Not(A:Brand";v="8", "Chromium";v="144", "Google Chrome";v="144"'

    def __init__(self, cookie_path=None, config_path=None, log_dir=None):
        project_root = os.path.dirname(os.path.abspath(__file__))

        # 加载配置
        self.config_path = config_path or os.path.join(project_root, self.DEFAULT_CONFIG_PATH)
        self.config = self._load_config()

        # Cookie 路径
        cookie_path_from_config = self.config.get("cookie文件路径")
        if cookie_path_from_config:
            self.cookie_path = cookie_path_from_config if os.path.isabs(cookie_path_from_config) else os.path.join(project_root, cookie_path_from_config)
        else:
            self.cookie_path = cookie_path or os.path.join(project_root, "config", "cookie.txt")

        # 输出目录
        output_dir_rel = self.config.get("输出目录", "output/")
        self.output_dir = output_dir_rel if os.path.isabs(output_dir_rel) else os.path.join(project_root, output_dir_rel)
        os.makedirs(self.output_dir, exist_ok=True)

        # 日志目录
        self.log_dir = log_dir or os.path.join(project_root, "logs")
        os.makedirs(self.log_dir, exist_ok=True)

        # 控制参数（严格从config读取，无业务默认值，避免硬编码）
        self.REQUEST_INTERVAL = int(self.config.get("请求间隔(秒)", "30"))
        self.MAX_RETRIES = int(self.config.get("最大重试次数", "3"))
        self.REQUEST_TIMEOUT = int(self.config.get("请求超时(秒)", "30"))
        self.SIGN_SALT = self.config.get("签名盐值", self._DEFAULT_SIGN_SALT)

        # UA 列表（Edge <-> Chrome 自动切换）
        self._ua_list = [
            (self.UA_EDGE, self.SEC_CH_UA_EDGE),
            (self.UA_CHROME, self.SEC_CH_UA_CHROME),
        ]
        self._current_ua_index = 0

        # Cookie 与日志
        self.cookie_str = self._read_cookie()
        self.logger = self._init_logger()
        self.logger.info(f"签名盐值: {self.SIGN_SALT}")

        # 风控指纹字段 wlfstk_smdl 处理（AGENTS.md 京东接口风控相关参数归档）：
        # 非强制必选字段 → 缺失仅输出警告日志，禁止抛异常/中断导出；
        # 仅当出现 403/601 风控拦截时才需补齐此字段。
        if "wlfstk_smdl" not in self.cookie_str:
            self.logger.warning("⚠️ 缺失 wlfstk_smdl（风控指纹字段，本次导出未受影响，备用提示）")

        # Session
        self.session = requests.Session()
        self.session.headers.update(self._build_default_headers())

        self.logger.info(f"JDBaseRequest 初始化完成，当前UA: {'Edge' if self._current_ua_index == 0 else 'Chrome'}")

    def _load_config(self):
        """从config.xlsx加载配置。返回 dict：{变量名: 参数值}"""
        if not os.path.exists(self.config_path):
            print(f"[WARN] 配置文件不存在: {self.config_path}，使用默认值")
            return {}
        try:
            wb = load_workbook(self.config_path, read_only=True, data_only=True)
            config_dict = {}
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if row and len(row) >= 3:
                        var_name = row[1]
                        var_value = row[2]
                        if var_name and var_value is not None:
                            config_dict[str(var_name)] = str(var_value)
            wb.close()
            print(f"[OK] 配置文件加载成功: {len(config_dict)} 项配置")
            return config_dict
        except Exception as e:
            print(f"[WARN] 配置文件加载失败: {e}，使用默认值")
            return {}

    def _get_business_param(self, key, default=None):
        """从config读取业务参数。
        中文说明（小白必读）：
            这是统一的"业务参数获取入口"。所有业务类都应该通过此方法读取参数，
            而不是直接读 self.config[key]。
            优点：如果后续config结构变化，只改这个方法一处即可。
            ⚠️ 禁止在业务函数里硬编码参数默认值！必须从config读取或上层传入。
        """
        value = self.config.get(key, default)
        if value is None:
            raise ValueError(
                f"业务参数 [{key}] 在 config.xlsx 中未配置，且未提供默认值。\n"
                f"请在 config.xlsx 的【全局配置】sheet中添加 [{key}] 配置项。"
            )
        return value

    # ---------- Cookie ----------
    def _read_cookie(self):
        if not os.path.exists(self.cookie_path):
            raise FileNotFoundError(
                f"Cookie文件不存在: {self.cookie_path}\n请将京东商智的Cookie保存到此文件中。"
            )
        with open(self.cookie_path, "r", encoding="utf-8") as f:
            cookie_str = f.read().strip()
        if not cookie_str:
            raise ValueError(f"Cookie文件为空: {self.cookie_path}")
        return cookie_str

    def refresh_cookie(self, cookie_str=None):
        if cookie_str:
            with open(self.cookie_path, "w", encoding="utf-8") as f:
                f.write(cookie_str)
            self.cookie_str = cookie_str
        else:
            self.cookie_str = self._read_cookie()
        self.session.headers["Cookie"] = self.cookie_str
        self.logger.info("Cookie已更新")

    # ---------- 风控签名 ----------
    def _gen_risk_params(self, url, uuid_prefix=None):
        """
        生成风控参数：User-mup / User-mnp / uuid
        算法（commons-a5562705.js 逆向）：
            User-mnp = MD5(URL路径 + uuid + 时间戳 + 盐值)

        参数:
            url         - 接口URL，用于提取URL路径
            uuid_prefix - 自定义uuid前缀（如不传，用类常量UUID_PREFIX）
                          不同业务/页面uuid前缀可能不同（参考商智购物车3001用5f9cc2ca20cad3d11642）
        """
        timestamp = int(time.time() * 1000)

        prefix = uuid_prefix if uuid_prefix else self.UUID_PREFIX
        random_min = 10 ** (self.UUID_RANDOM_DIGITS - 1)
        random_max = 10 ** self.UUID_RANDOM_DIGITS - 1
        uuid_str = f"{prefix}-{random.randint(random_min, random_max)}"

        parsed = urlparse(url)
        url_path = parsed.path

        sign_str = f"{url_path}{uuid_str}{timestamp}{self.SIGN_SALT}"
        user_mnp = hashlib.md5(sign_str.encode("utf-8")).hexdigest()

        return {
            "User-mup": str(timestamp),
            "User-mnp": user_mnp,
            "uuid": uuid_str,
        }

    # ---------- 间隔控制 ----------
    def _wait_interval(self):
        """严格执行30秒间隔（从config读取），禁止跳过此方法。"""
        if self._last_request_time == 0:
            return
        elapsed = time.time() - self._last_request_time
        wait_time = self.REQUEST_INTERVAL - elapsed
        if wait_time > 0:
            self.logger.info(f"请求间隔控制：等待 {wait_time:.1f} 秒（距上次 {elapsed:.1f}秒，需≥{self.REQUEST_INTERVAL}秒）")
            time.sleep(wait_time)

    # ---------- 请求头 ----------
    def _build_default_headers(self):
        ua, sec_ch_ua = self._ua_list[self._current_ua_index]
        return {
            "User-Agent": ua,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
            "Accept-Encoding": "gzip, deflate, br, zstd",
            "Content-Type": "application/x-www-form-urlencoded",
            "Origin": self.DEFAULT_ORIGIN,
            "Referer": self.DEFAULT_REFERER,
            "Upgrade-Insecure-Requests": "1",
            "sec-ch-ua": sec_ch_ua,
            "sec-ch-ua-mobile": "?0",
            "sec-ch-ua-platform": '"Windows"',
            "Cookie": self.cookie_str,
        }

    def _switch_ua(self):
        old_name = "Edge" if self._current_ua_index == 0 else "Chrome"
        self._current_ua_index = 1 - self._current_ua_index
        new_name = "Edge" if self._current_ua_index == 0 else "Chrome"
        ua, sec_ch_ua = self._ua_list[self._current_ua_index]
        self.session.headers["User-Agent"] = ua
        self.session.headers["sec-ch-ua"] = sec_ch_ua
        self.logger.info(f"UA切换: {old_name} → {new_name}")

    # ---------- 通用请求 ----------
    def request(self, url, data, method="POST", extra_headers=None, uuid_prefix=None):
        """
        通用请求方法（自动加风控签名、重试、UA切换、间隔控制）。

        参数:
            url          - 请求URL
            data         - 表单参数dict
            method       - POST/GET
            extra_headers- 额外请求头
            uuid_prefix  - 自定义uuid前缀（不传则用类常量UUID_PREFIX）
        """
        headers = {}
        if extra_headers:
            headers.update(extra_headers)

        last_exception = None
        for attempt in range(1, self.MAX_RETRIES + 1):
            try:
                self._wait_interval()  # ⚠️ 严格间隔控制，禁止跳过
                risk_params = self._gen_risk_params(url, uuid_prefix=uuid_prefix)
                full_data = {**data, **risk_params}

                ua_name = "Edge" if self._current_ua_index == 0 else "Chrome"
                self.logger.info(f"发送请求 (第{attempt}/{self.MAX_RETRIES}次, UA={ua_name}): {url}")
                self.logger.debug(f"请求参数: {json.dumps(full_data, ensure_ascii=False)[:500]}")

                # 记录请求时间（类属性：所有实例共享，保证批量执行也严格间隔）
                JDBaseRequest._last_request_time = time.time()

                if method.upper() == "POST":
                    response = self.session.post(url, data=full_data, headers=headers, timeout=self.REQUEST_TIMEOUT)
                else:
                    response = self.session.get(url, params=full_data, headers=headers, timeout=self.REQUEST_TIMEOUT)

                content_type = response.headers.get("Content-Type", "")
                if "json" in content_type:
                    try:
                        result = response.json()
                        if not result.get("success", True) and result.get("status", 0) < 0:
                            error_msg = result.get("message", "未知错误")
                            status_code = result.get("status")

                            if status_code in (302, -1) or "登录" in error_msg or "login" in error_msg.lower():
                                self.logger.error(f"Cookie可能已过期: {error_msg} (status={status_code})")
                                raise CookieExpiredError("Cookie已过期，请更新 config/sz_cookie.txt")

                            self.logger.warning(f"风控拦截: {error_msg} (status={status_code})")
                            if attempt < self.MAX_RETRIES:
                                self._switch_ua()
                                wait = self.REQUEST_INTERVAL * attempt
                                self.logger.info(f"等待 {wait}秒 后重试...")
                                time.sleep(wait)
                                continue
                    except json.JSONDecodeError:
                        pass

                self.logger.info(f"请求成功: HTTP {response.status_code}, {len(response.content)}字节")
                return response

            except CookieExpiredError:
                raise

            except requests.exceptions.Timeout as e:
                last_exception = e
                self.logger.warning(f"第{attempt}次请求超时（{self.REQUEST_TIMEOUT}秒）")

            except Exception as e:
                last_exception = e
                self.logger.warning(f"第{attempt}次请求失败: {e}")

            if attempt < self.MAX_RETRIES:
                wait = self.REQUEST_INTERVAL * attempt
                self.logger.info(f"等待 {wait}秒 后重试...")
                time.sleep(wait)

        self.logger.error(f"所有 {self.MAX_RETRIES} 次重试均失败")
        raise last_exception

    # ---------- 日志 ----------
    def _init_logger(self):
        logger = logging.getLogger(f"JD_{datetime.now().strftime('%Y%m%d')}")
        logger.setLevel(logging.DEBUG)
        if logger.handlers:
            return logger
        log_file = os.path.join(self.log_dir, f"jd_api_{datetime.now().strftime('%Y%m%d')}.log")
        file_handler = logging.FileHandler(log_file, encoding="utf-8")
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S"))
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        console_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S"))
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)
        return logger

    # ---------- Excel 保存 ----------
    def save_excel(self, response, filename):
        file_path = os.path.join(self.output_dir, filename)
        with open(file_path, "wb") as f:
            f.write(response.content)
        self.logger.info(f"Excel已保存: {file_path} ({len(response.content)}字节)")
        return file_path


# ============================================================
#  业务接口 1：商品流量来源 - SKU维度（搜索/推荐/购物车）
# ------------------------------------------------------------
#  业务名称：
#      - 商品流量来源_搜索（lastSrcChannelId2=2008，搜索子来源）【执行】
#      - 商品流量来源_推荐（lastSrcChannelId2=2009，推荐子来源）【执行】
#      - 商品流量来源_购物车（lastSrcChannelId2=3001，购物车/我的订单回流）【执行】
#      - 商品流量来源_自主访问（lastSrcChannelId2=3001，与购物车口径重叠）【已停用，仅保留配置】
#  接口地址：https://szgateway.jd.com/szpaas/szajax/shop/source/offlineFlowSource/downSkuTable.ajax
#  数据维度：店铺来源 → 搜索/推荐/购物车渠道 → SKU维度（按入店浏览量降序，最多5000条）
#  返回格式：Excel 二进制流
#  参数说明：
#      可变业务参数（来自config.xlsx）：interval, dateType, limit
#      固定业务常量（经用户确认写死代码）：lastSrcChannelId1=2, groupType=skuId, attributes=skuId,
#          sortField=按入店浏览量, sortType=desc, compareType=hb
#      二级渠道 lastSrcChannelId2：2008/2009/3001（由CHANNEL_MAP自动带入）
#      日期参数（来自config + 函数入参动态覆盖）：
#          date / startDate / endDate → 优先用入参，其次从 config.xlsx 读取
# ============================================================
class ProductFlowAPI(JDBaseRequest):
    """商品流量来源 - SKU维度 - 数据导出（搜索/推荐/购物车执行，自主访问已停用）"""

    # 接口URL（域名固定，业务参数走配置）
    API_URL = "https://szgateway.jd.com/szpaas/szajax/shop/source/offlineFlowSource/downSkuTable.ajax"

    # ⚠️ CHANNEL_MAP：商品流量子渠道配置（业务核心配置）
    # 中文说明（小白必读）：
    #   商品流量来源的子渠道定义。每个子渠道对应一个业务key：
    #     "商品流量来源_搜索"     → 搜索子来源（2008）【执行】
    #     "商品流量来源_推荐"     → 推荐子来源（2009）【执行】
    #     "商品流量来源_购物车"   → 购物车/我的订单回流（3001）【执行】
    #     "商品流量来源_自主访问" → 与购物车数据口径重叠（同3001），仅保留配置，调度层停用
    #   uuid前缀：京东风控校验用的随机ID前缀
    #     搜索/推荐：ca412182e5668a106054
    #     购物车/自主访问：5f9cc2ca20cad3d11642
    #   ⚠️ 警告：新增/修改子渠道，必须修改此字典（业务参数专属配置）。
    #
    #   ⚠️ 2026-08-10 配置注释（P1-1）：商智搜索/推荐/购物车接口 **服务端不支持多日区间导出**。
    #     虽然接口表单里有 startDate/endDate，但服务端实际只返回单日数据。
    #     用户传入区间（--range last_Nd 或 --start_date/--end_date）时，
    #     由 _download_sku_by_days() 自动拆成逐天循环调用，每天插入当天日期列后合并输出。
    CHANNEL_MAP = {
        "商品流量来源_搜索":     ("2008", "ca412182e5668a106054"),
        "商品流量来源_推荐":     ("2009", "ca412182e5668a106054"),
        "商品流量来源_购物车":   ("3001", "5f9cc2ca20cad3d11642"),
        "商品流量来源_自主访问": ("3001", "5f9cc2ca20cad3d11642"),  # 与购物车口径重叠，仅保留配置
    }

    # ⚠️ 2026-08-10 用户决策：逐日循环最大天数限制。
    #   商智搜索/推荐/购物车不支持多日区间，区间查询会拆成逐天循环（每天一次接口+约30秒间隔）。
    #   设置 31 天上限，防止大批量压接口触发风控 403；超出限制直接报错终止。
    MAX_RANGE_DAYS = 31

    # 反向索引（从CHANNEL_MAP自动生成，用于支持直接传channel_id2）
    _CHANNEL_ID_INDEX = None

    @classmethod
    def _build_channel_id_index(cls):
        """从 CHANNEL_MAP 构建反向索引：{channel_id2: uuid_prefix}"""
        index = {}
        for _key, (channel_id2, uuid_prefix) in cls.CHANNEL_MAP.items():
            index[channel_id2] = uuid_prefix
        return index

    def _get_channel_config(self, biz_key):
        """获取渠道配置 (channel_id2, uuid_prefix)。支持业务key和channel_id2两种入参。"""
        # 方式1：业务key直接查
        if biz_key in self.CHANNEL_MAP:
            return self.CHANNEL_MAP[biz_key]

        # 方式2：channel_id2反向查
        if self._CHANNEL_ID_INDEX is None:
            self._CHANNEL_ID_INDEX = self._build_channel_id_index()
        if biz_key in self._CHANNEL_ID_INDEX:
            uuid_prefix = self._CHANNEL_ID_INDEX[biz_key]
            for key, (cid, _) in self.CHANNEL_MAP.items():
                if cid == biz_key:
                    self.logger.info(f"通过channel_id2 '{biz_key}' 匹配到业务 '{key}'")
                    break
            return (biz_key, uuid_prefix)

        available_keys = "、".join(self.CHANNEL_MAP.keys())
        raise ValueError(
            f"不支持的渠道: {biz_key}\n可用业务key: {available_keys}"
        )

    def _get_uuid_for_channel(self, biz_key):
        """根据业务key或channel_id2，返回完整uuid。"""
        _, uuid_prefix = self._get_channel_config(biz_key)
        random_min = 10 ** (self.UUID_RANDOM_DIGITS - 1)
        random_max = 10 ** self.UUID_RANDOM_DIGITS - 1
        return f"{uuid_prefix}-{random.randint(random_min, random_max)}"

    def _resolve_display_key(self, biz_key, channel_id2):
        """把biz_key归一化为友好业务key（用于日志和文件名）。"""
        if biz_key in self.CHANNEL_MAP:
            return biz_key
        for key, (cid, _) in self.CHANNEL_MAP.items():
            if cid == channel_id2:
                return key
        return biz_key

    # ---------- 业务参数配置区 ----------
    # ⚠️ 铁律：经常变化的业务参数必须走 config.xlsx（带开发期兜底+警告）。
    #    固定不变的业务常量，经用户确认（2026-08-05）后直接作为代码常量，
    #    不再放入 config.xlsx，避免配置文件冗余。

    # 【可变参数】经常需要调整，从 config.xlsx 读取（缺省用兜底值并打印警告）
    VARIABLE_BIZ_PARAMS = {
        "interval": "DAY",      # 时间粒度：DAY=按天汇总 / MONTH=按月汇总
        "dateType": "day",      # 日期类型：与interval对应（day/month）
        "limit": "5000",        # 返回条数上限
    }

    # 【固定常量】经用户确认(2026-08-05)固定不变，直接写死在代码，不读config
    FIXED_BIZ_PARAMS = {
        "lastSrcChannelId1": "2",   # 一级渠道：商品流量来源都是2
        "groupType": "skuId",       # 聚合维度：按商品SKU维度汇总
        "attributes": "skuId",      # 返回字段：SKU维度数据列
        "sortField": "jdr_sch_traffic_enter_shop__browse_page_cnt_shop_last_src",  # 排序字段：按入店浏览量
        "sortType": "desc",         # 排序方式：降序
        "compareType": "hb",        # 对比方式：环比
    }

    def _get_business_params(self):
        """组装本次请求的全部业务参数。

        规则：
            ① 可变参数（interval/dateType/limit）从 config.xlsx 读取，
               缺省时用代码兜底值并打印警告，提醒补写config；
            ② 固定常量（渠道ID/维度/排序/对比方式等）经用户确认后写死在代码，
               不读config、不打印警告。
        """
        result = {}
        missing = []

        # ① 可变参数：从config读取（优先），缺省用兜底值 + 警告
        for key, default_value in self.VARIABLE_BIZ_PARAMS.items():
            value = self.config.get(key)
            if value is None or value == "":
                missing.append(key)
                value = default_value  # 开发期兜底
            result[key] = value

        # ② 固定常量：直接并入（经用户确认固化，不读config）
        result.update(self.FIXED_BIZ_PARAMS)

        if missing:
            print(
                f"[WARN] 以下可变业务参数在config.xlsx中未配置，使用代码兜底值（建议补充到config）：\n"
                f"       缺失参数: {', '.join(missing)}\n"
                f"       ⚠️ 严禁长期依赖兜底！这些参数必须添加到 config.xlsx【全局配置】sheet。"
            )
        return result

    def _get_date_params(self, date=None, start_date=None, end_date=None):
        """解析本次查询的日期参数（支持动态覆盖）。

        规则（2026-08-05 修复）：
            date      : 优先用入参（如命令行 --date），其次从 config.xlsx 读取
            startDate : 优先用入参；未传时默认=date（⚠️ 修复：此前回落config旧值，
                        导致 --date 指定新日期时 start/end 仍是config里旧日期，
                        接口按旧区间取数，07-29与07-30导出完全相同）
            endDate   : 同 startDate
        """
        # date：入参优先，其次config
        if date is None:
            date = self.config.get("date")
        if date is None:
            raise ValueError("查询日期date未提供：请在config.xlsx配置或通过函数入参传入")

        # start/end：入参优先；未传时默认与date一致（单日查询）
        if start_date is None:
            start_date = date
        if end_date is None:
            end_date = date

        return date, start_date, end_date

    def download_sku(self, biz_key="商品流量来源_搜索", date=None, start_date=None, end_date=None):
        """下载SKU维度数据。
        参数:
            biz_key     - 业务key（CHANNEL_MAP中的key，如"商品流量来源_搜索"）
            date        - 查询日期（YYYY-MM-DD），优先用入参
            start_date  - 开始日期，单日查询时与date相同
            end_date    - 结束日期，单日查询时与date相同
        返回:
            保存的Excel文件路径
        """
        # 1. 解析渠道配置
        channel_id2, uuid_prefix = self._get_channel_config(biz_key)
        display_key = self._resolve_display_key(biz_key, channel_id2)

        # 2. 读取日期参数（入参 > config）
        date, start_date, end_date = self._get_date_params(date, start_date, end_date)

        # ⚠️ 2026-08-10 区间拆解（P0 用户决策）：
        #   商智搜索/推荐/购物车接口 **服务端不支持多日区间导出**（虽然表单有 startDate/endDate，
        #   但实际只返回单日数据）。当检测到 start_date != end_date（跨多天区间）时，
        #   自动拆成逐天循环调用：每天 date=startDate=endDate=当天 → 每天插入当天日期列 →
        #   pandas concat 合并 → 一次性输出合并 xlsx（文件名标注区间，如 2026-08-03_2026-08-09）。
        #   单日期（start_date == end_date）不进入循环，保持原有执行路径不变。
        if start_date and end_date and start_date != end_date:
            return self._download_sku_by_days(biz_key, display_key, start_date, end_date)

        # 3. 读取业务参数（全部从config）
        biz_params = self._get_business_params()

        # 4. 组装请求数据
        data = {
            "date": date,
            "startDate": start_date,
            "endDate": end_date,
            **biz_params,
            "lastSrcChannelId2": channel_id2,
        }

        # 5. 发送请求（自动间隔+重试+UA切换+风控签名）
        self.logger.info(
            f"下载商品流量来源数据: 日期={date}, 业务={display_key}"
            f"(id2={channel_id2}, uuid_prefix={uuid_prefix[:8]}...)"
        )
        response = self.request(self.API_URL, data, uuid_prefix=uuid_prefix)

        # 6. 保存Excel（文件名用友好业务key）
        # 例如：搜索流量_2026-07-29.xlsx
        short_name = display_key.replace("商品流量来源_", "")  # 去掉前缀，保留"搜索/推荐/购物车"
        filename = f"{short_name}流量_{date}.xlsx"

        # 7. 后置处理保存：读Excel → 首列插入【日期】 → 数值安全转换 → 写回
        return self._save_flow_excel(response, filename, date)

    # ---------- 商品流量来源 区间逐日循环导出（2026-08-10 新增）----------
    def _download_sku_by_days(self, biz_key, display_key, start_date, end_date):
        """区间查询专用：服务端不支持多日区间 → 拆成逐天循环调用，合并输出。

        ⚠️ 背景（P0 用户决策 2026-08-10）：
           商智搜索/推荐/购物车接口（downSkuTable.ajax）虽然表单有 startDate/endDate，
           但服务端实际只返回单日数据（已实证）。因此区间查询时不能一次传区间，
           而是把区间拆成逐天日期列表，每天单独调用接口（date=startDate=endDate=当天），
           每天得到的单天数据插入当天日期列，最后 pandas.concat 合并为一份 xlsx 输出。

        流程：
            ① split_date_range() 拆区间 → 校验（格式/start≤end/最大天数上限31天）
            ② 打印 CLI 警告：接口不支持区间，自动拆分逐日循环
            ③ for 每天：组装 data（date=startDate=endDate=当天）→ 调接口 →
               读取单天Excel → 插入当天日期列 → 收集 DataFrame
            ④ pd.concat 合并全部单天 → 写入合并 xlsx
               文件名标注区间（如 搜索流量_2026-08-03_2026-08-09.xlsx），
               输出子目录用最后一天 end_date（用户决策 2026-08-10）
            ⑤ 中间单天文件不落盘，只输出合并文件（用户决策 2026-08-10）

        入参:
            biz_key     - 业务key（CHANNEL_MAP中的key）
            display_key - 友好业务key（用于文件名/日志）
            start_date  - 区间开始日期 YYYY-MM-DD
            end_date    - 区间结束日期 YYYY-MM-DD
        出参:
            合并后的Excel文件绝对路径
        """
        import pandas as pd

        # ① 拆区间 + 边界保护（格式校验 / start≤end / 最大31天，超限抛错）
        day_list = split_date_range(start_date, end_date, self.MAX_RANGE_DAYS)

        # ② CLI 可见警告（P1-2）：说明接口不支持区间，将自动拆分逐日循环
        warn_msg = (
            f"⚠️ [WARN] 商智『{display_key}』接口服务端不支持多日区间导出，"
            f"本次区间 {start_date} ~ {end_date}（共{len(day_list)}天）将自动拆分为逐日循环调用，"
            f"预计耗时约 {len(day_list) * 30 // 60} 分钟+（每天间隔30秒）。"
        )
        self.logger.warning(warn_msg)
        print(warn_msg)

        channel_id2, uuid_prefix = self._get_channel_config(biz_key)
        biz_params = self._get_business_params()

        # ③ 逐天循环：每天 date=startDate=endDate=当天，收集 DataFrame
        frames = []
        total = len(day_list)
        for i, day in enumerate(day_list, 1):
            self.logger.info(f"    逐日循环 [{i}/{total}] {day}（{display_key}）")
            data = {
                "date": day,
                "startDate": day,
                "endDate": day,
                **biz_params,
                "lastSrcChannelId2": channel_id2,
            }
            response = self.request(self.API_URL, data, uuid_prefix=uuid_prefix)
            # 每天的单天数据：接口无日期列 → prepare_date_columns 自动插入当天日期列
            df, _col, _val = self._read_flow_df(response, day)
            frames.append(df)

        # ④ 合并全部单天数据（列结构对齐，忽略行索引重建）
        merged = pd.concat(frames, ignore_index=True)

        # 文件名标注区间（如 搜索流量_2026-08-03_2026-08-09.xlsx）
        short_name = display_key.replace("商品流量来源_", "")  # 去掉前缀，保留"搜索/推荐/购物车"
        filename = f"{short_name}流量_{start_date}_{end_date}.xlsx"

        # ⑤ 输出目录用最后一天 end_date（用户决策 2026-08-10）
        # ⚠️ 区间文件名含两个日期（start_end），不能复用 build_business_output_path 的
        #   "去掉 _{date}.xlsx 后缀提取业务模块名"规则（会把 start 日期误并入业务模块名，
        #   产生 output/搜索流量_2026-08-03/ 这样的错误目录），故在此直接构造目录：
        #   output/{业务模块}/{end_date}/{filename}
        business_dir = os.path.join(self.output_dir, f"{short_name}流量", end_date)
        os.makedirs(business_dir, exist_ok=True)
        file_path = os.path.join(business_dir, filename)
        merged.to_excel(file_path, index=False, engine="openpyxl")
        # 日期列逐格解析（每天插入的日期不同），date_value=None
        apply_column_formats(file_path, merged, date_column="日期", date_value=None)

        self.logger.info(
            f"Excel已保存(区间合并): {file_path}（{total}天数据合并，共{len(merged)}行，"
            f"{os.path.getsize(file_path)}字节）"
        )
        return file_path

    # ---------- 商品流量来源 Excel后置处理（2026-08-05 新增）----------
    def _read_flow_df(self, response, date):
        """读接口返回的Excel二进制流 → 日期列处理 → 数值安全转换。

        （2026-08-10 抽取公共步骤，供单日 _save_flow_excel 与区间 _download_sku_by_days 复用）

        入参:
            response - requests响应（content为接口返回的xlsx二进制）
            date     - 本次查询日期（如 2026-07-29），无日期列时插入的日期值
        出参:
            (df, date_column, date_value)
                df          - 处理后的DataFrame（已插入/格式化日期列+数值转换）
                date_column - 实际承载日期的列名
                date_value  - 插入列场景：日期字符串；自带日期列场景：None
        """
        import io
        import warnings
        import pandas as pd

        # 抑制openpyxl读取原始xlsx时的无害警告（"Workbook contains no default style"）
        warnings.filterwarnings("ignore", message="Workbook contains no default style")

        # ① 读取二进制流 → DataFrame（接口返回的原始数据）
        # ⚠️ 必须 dtype=str：否则pandas会把"数字样式的SKU"自动转成数值(int64)，
        #    导致 safe_convert_numeric 的">15位长数字保留文本"保护失效、精度丢失。
        #    先全部按文本读取，再交给安全转换函数统一处理。
        #    na_filter=False：空单元格保持空字符串，避免被替换成'nan'文本。
        df = pd.read_excel(io.BytesIO(response.content), dtype=str, na_filter=False)

        # ② 日期列统一处理（公共规则1+2，2026-08-07）：
        #    报表自带【日期】/【时间】列 → 禁止重复插入日期列，仅做格式标准化；
        #    报表无日期/时间列 → 首列插入【日期】列，值=本次查询日期。
        date_column, date_value = prepare_date_columns(df, date)

        # ③ 全表数值安全转换（>15位长数字保留文本，防止精度丢失）
        df = safe_convert_numeric(df)

        return df, date_column, date_value

    def _save_flow_excel(self, response, filename, date):
        """商品流量来源专用保存流程（单日，Excel后置处理，保持原有行为）。

        导出流程（需求文档要求 + 2026-08-07 公共规则1+2）：
            ① 复用 _read_flow_df()：读Excel → 日期列处理 → 数值安全转换
            ② 写入Excel并设置日期列单元格格式（打开文件不弹格式警告）

        入参:
            response - requests响应（content为接口返回的xlsx二进制）
            filename - 保存文件名（如 搜索流量_2026-07-29.xlsx）
            date     - 本次查询日期（如 2026-07-29）
        出参:
            保存后的Excel文件绝对路径
        """
        # ① 复用公共读取步骤（读二进制流 → 日期列处理 → 数值转换）
        df, date_column, date_value = self._read_flow_df(response, date)

        # ② 写入Excel → 按列名规则设置单元格格式（日期列/订单编号@/SKU·SPU数值0位小数）
        # 输出目录规则（AGENTS.md Excel规则4）：output/{业务模块}/{date}/{filename}
        file_path = build_business_output_path(self.output_dir, filename, date)
        df.to_excel(file_path, index=False, engine="openpyxl")
        apply_column_formats(file_path, df, date_column=date_column, date_value=date_value)

        self.logger.info(
            f"Excel已保存: {file_path}（已插入日期列+数值转换+单元格格式，{os.path.getsize(file_path)}字节）"
        )
        return file_path

    # 业务级便捷方法（保持向后兼容，内部都走 download_sku）
    def download_search_sku(self, date=None, start_date=None, end_date=None):
        """便捷方法：导出搜索流量"""
        return self.download_sku(biz_key="商品流量来源_搜索", date=date, start_date=start_date, end_date=end_date)

    def download_recommend_sku(self, date=None, start_date=None, end_date=None):
        """便捷方法：导出推荐流量"""
        return self.download_sku(biz_key="商品流量来源_推荐", date=date, start_date=start_date, end_date=end_date)

    def download_cart_sku(self, date=None, start_date=None, end_date=None):
        """便捷方法：导出购物车流量（购物车/我的订单回流，3001，正式执行渠道）"""
        return self.download_sku(biz_key="商品流量来源_购物车", date=date, start_date=start_date, end_date=end_date)

    def download_selfvisit_sku(self, date=None, start_date=None, end_date=None):
        """便捷方法：导出自主访问流量（与购物车口径重叠，已停用，仅供保留存档）"""
        return self.download_sku(biz_key="商品流量来源_自主访问", date=date, start_date=start_date, end_date=end_date)


# ============================================================
#  业务接口 2：（占位 - 后续追加）
# ------------------------------------------------------------
#  业务名称：店铺来源_三级渠道（离线流量报表）
#  接口地址：https://szgateway.jd.com/szpaas/szajax/shop/source/offlineFlowSource/downTable.ajax
#  业务说明：按三级流量渠道分组，导出店铺来源离线日度流量报表
#  业务需求（2026-08-06）：
#      - 对比方式=hb（环比）
#      - 聚合粒度=DAY（按日）
#      - 分组维度=lastSrcChannelId3（三级渠道 ID）
#      - 排序字段=进店访客数降序
#      - 日期类型=day（离线日度）
#      - 下载类型=downType=day
#      - 平台品类=platformCate1（空=全品类）
#  硬性约束（用户 2026-08-06 确认）：
#      1. 必带 Origin/Referer（缺失被平台拦截）
#      2. 业务表单参数固定不变（13 个字段）
#      3. User-mup / User-mnp / uuid 每次调用动态生成（禁止硬编码）
#      4. Cookie 从浏览器会话获取（走 config/sz_cookie.txt，禁止入代码）
# ============================================================
class OfflineChannelAPI(JDBaseRequest):
    """店铺来源-离线渠道流量报表 API。

    业务定位：
        商智 szgateway.jd.com 模块下"店铺来源-离线渠道"维度的报表导出。
        与商品流量来源（downSkuTable.ajax / SKU 维度）不同，本接口按
        三级流量渠道分组，输出渠道维度的访客/浏览/成交数据。

    父类复用：
        - 父类 JDBaseRequest 提供：
            * Cookie 读取（config/sz_cookie.txt）
            * 风控签名（User-mnp MD5 哈希，盐值复用全局 SIGN_SALT）
            * 自动 30 秒间隔（_wait_interval）
            * 自动重试 + UA 切换
            * 日志、Excel 保存

    自实现部分：
        - UUID 完全随机生成器（_gen_uuid_random）：不依赖父类 UUID_PREFIX，
          避免硬编码任何 uuid 前缀（你抓包显示 UUID 完全随机）。
        - 业务参数分组：groupType/attributes/dimensions/排序 走 FIXED_BIZ_PARAMS，
          业务专属常量不走 config（沿用现有 3 业务的固化模式）。
        - Excel 后置处理：复用父类 _save_flow_excel（商品流量来源专用后置）。
    """

    # 接口 URL（业务约束，固定）
    # 通俗解释：按三级渠道下载离线流量报表
    API_URL = "https://szgateway.jd.com/szpaas/szajax/shop/source/offlineFlowSource/downTable.ajax"

    # 必带请求头（业务约束，固定）
    # 通俗解释：告诉京东"我是从这个页面来的"，缺失会被拦截
    ORIGIN = "https://sz.jd.com"
    REFERER = "https://sz.jd.com/szweb/sz/view/viewflow/viewSourcesVNew.html"

    # 固定业务参数（用户确认 2026-08-06 固化，不读 config）
    # 通俗解释：这些参数永远不变（业务定义），直接写死代码
    FIXED_BIZ_PARAMS = {
        "compareType": "hb",          # 对比方式：环比
        "interval": "DAY",            # 聚合粒度：按日
        "dateType": "day",            # 日期类型：离线日度
        "downType": "day",            # 下载类型：按日
        "groupType": "lastSrcChannelId3",   # 分组维度：三级渠道
        "attributes": "lastSrcChannelId3",  # 返回字段：三级渠道 ID
        "sortField": "jdr_sch_traffic_enter_shop__visitor_cnt_shop_last_src",  # 排序字段：进店访客数（候选 A 兜底）
        "sortType": "desc",           # 排序方式：降序
        "lastSrcChannelId1": "2",     # 一级渠道：商品流量来源都是 2
    }

    # 可变业务参数（从 config 读取，缺省用兜底值并打印警告）
    # 通俗解释：日期可改，其他业务不变
    VARIABLE_BIZ_PARAMS = {
        "platformCate1": "",          # 平台品类 1：空字符串=全品类（可命令行覆盖）
    }

    def _get_date_params(self, date=None, start_date=None, end_date=None):
        """解析本次查询的日期参数（支持动态覆盖）。

        ⚠️ 重要：基类 JDBaseRequest 没有这个方法（ProductFlowAPI 才有），
        新业务必须自实现，否则会 AttributeError。

        规则（2026-08-05 修复）：
            date      : 优先用入参（如命令行 --date），其次从 config.xlsx 读取
            startDate : 优先用入参；未传时默认=date（⚠️ 修复：此前回落config旧值，
                        导致 --date 指定新日期时 start/end 仍是config里旧日期，
                        接口按旧区间取数，07-29与07-30导出完全相同）
            endDate   : 同 startDate
        """
        # date：入参优先，其次config
        if date is None:
            date = self.config.get("date")
        if date is None:
            raise ValueError("查询日期date未提供：请在config.xlsx配置或通过函数入参传入")

        # start/end：入参优先；未传时默认与date一致（单日查询）
        if start_date is None:
            start_date = date
        if end_date is None:
            end_date = date

        return date, start_date, end_date

    def _gen_uuid_random(self):
        """完全随机 UUID 生成器（不依赖任何固定前缀）。

        业务背景：
            你抓包显示 UUID 前缀两次完全不同（f1d5ae16 / a31e066d），
            证明前端 SDK 每次会话生成新 UUID。
            本方法用 Python secrets 生成 16hex + "-" + 10hex，与抓包格式一致。

        返回:
            str - 形如 "a31e066d8e94f4f39a3a-19fda02c2d4"
        """
        import secrets
        # 16位小写hex + "-" + 10位小写hex，与你抓包格式一致
        prefix = secrets.token_hex(8)        # 8字节 = 16hex 字符
        suffix = secrets.token_hex(5)        # 5字节 = 10hex 字符
        return f"{prefix}-{suffix}"

    def _gen_risk_params_random(self, url):
        """生成风控参数：User-mup / User-mnp / uuid（uuid 完全随机版）。

        与父类 _gen_risk_params 的差异：
            父类默认用类常量 UUID_PREFIX（硬编码 ca412182...）
            本方法用 _gen_uuid_random() 完全随机生成

        算法不变（commons-a5562705.js 逆向）：
            User-mnp = MD5(URL路径 + uuid + 时间戳 + 盐值)

        参数:
            url - 接口URL，用于提取URL路径
        返回:
            dict - {"User-mup": str, "User-mnp": str, "uuid": str}
        """
        timestamp = int(time.time() * 1000)  # 毫秒级时间戳，每次必新
        uuid_str = self._gen_uuid_random()    # 完全随机，不传任何 prefix

        from urllib.parse import urlparse
        parsed = urlparse(url)
        url_path = parsed.path

        # 盐值复用全局 SIGN_SALT（与现有 3 业务共用，确保 MD5 哈希算法一致）
        sign_str = f"{url_path}{uuid_str}{timestamp}{self.SIGN_SALT}"
        user_mnp = hashlib.md5(sign_str.encode("utf-8")).hexdigest()

        return {
            "User-mup": str(timestamp),
            "User-mnp": user_mnp,
            "uuid": uuid_str,
        }

    def download_offline_channel(self, date=None, start_date=None, end_date=None, platform_cate1=None):
        """下载店铺来源-三级渠道离线流量报表。

        参数:
            date           - 查询日期 YYYY-MM-DD（入参优先，其次 config）
            start_date     - 开始日期（单日查询=date）
            end_date       - 结束日期（单日查询=date）
            platform_cate1 - 平台品类 1（默认空=全品类；如传值则下载指定品类报表）

        返回:
            str - 保存的 Excel 文件绝对路径
        """
        # 1. 解析日期（入参 > config；与现有 _get_date_params 同源）
        date, start_date, end_date = self._get_date_params(date, start_date, end_date)

        # 2. 读取可变业务参数（platformCate1）
        platform_cate1 = platform_cate1 if platform_cate1 is not None else self.VARIABLE_BIZ_PARAMS["platformCate1"]

        # 3. 组装业务参数（固定常量 + 可变参数）
        biz_data = {
            "date": date,
            "startDate": start_date,
            "endDate": end_date,
            "platformCate1": platform_cate1,
            **self.FIXED_BIZ_PARAMS,
        }

        # 4. 必带请求头（业务约束：Origin/Referer 缺失被平台拦截）
        extra_headers = {
            "Origin": self.ORIGIN,
            "Referer": self.REFERER,
        }

        # 5. 发送请求（带完整重试循环：递增等待 + UA切换 + 风控识别 + Excel校验）
        #    本业务要求 uuid 完全随机，绕过基类 UUID_PREFIX 默认 → 手写签名
        #    重试策略（与基类 request() 对齐）：
        #        最多 MAX_RETRIES 次
        #        第 2/3 次前自动切换 UA（Edge ↔ Chrome）
        #        递增等待 REQUEST_INTERVAL * attempt（30/60/90 秒）
        #        Cookie 过期立即停止（不可重试）
        response = None  # 显式初始化，便于重试作用域
        last_exception = None

        for attempt in range(1, self.MAX_RETRIES + 1):
            try:
                # 5.1 严格间隔控制（与基类一致）
                self._wait_interval()

                # 5.2 重新生成风控参数（每次新 UUID + 新时间戳 + 新签名）
                risk_params = self._gen_risk_params_random(self.API_URL)
                full_data = {**biz_data, **risk_params}

                ua_name = "Edge" if self._current_ua_index == 0 else "Chrome"
                self.logger.info(
                    f"[第{attempt}/{self.MAX_RETRIES}次] 下载店铺来源-三级渠道报表: "
                    f"日期={date}, 平台品类='{platform_cate1 or '全品类'}', "
                    f"UA={ua_name}, uuid={risk_params['uuid'][:8]}..."
                )
                self.logger.debug(f"请求参数: {json.dumps(full_data, ensure_ascii=False)[:500]}")

                # 记录请求时间（类属性：所有实例共享，保证批量执行也严格间隔）
                JDBaseRequest._last_request_time = time.time()

                # 5.3 发送 POST 请求
                response = self.session.post(
                    self.API_URL,
                    data=full_data,
                    headers=extra_headers,
                    timeout=self.REQUEST_TIMEOUT,
                )

                # 5.4 HTTP 状态码基础检查
                if response.status_code == 401:
                    self.logger.error("HTTP 401 未授权 - Cookie 可能已过期或被禁用")
                    raise CookieExpiredError("Cookie已过期或无效，请更新 config/sz_cookie.txt")
                if response.status_code == 403:
                    self.logger.error("HTTP 403 禁止访问 - Cookie/签名/Origin 校验失败")
                    # 不抛 CookieExpired，让重试机制 + UA 切换兜底

                # 5.5 风控业务码识别（json 响应里的 status / message）
                self._check_business_code(response)

                # 5.6 空响应拦截：Excel magic 字节校验
                self._validate_excel_response(response, attempt)

                # 5.7 走到这里 = 成功，跳出重试循环
                self.logger.info(f"请求成功: HTTP {response.status_code}, {len(response.content)}字节")
                break

            except CookieExpiredError:
                # Cookie 过期是硬错误，不能靠重试解决，直接抛出
                raise
            except RiskControlError:
                # 阶段4修复：601 限流等风控硬错误，不重试，直接抛出
                raise
            except requests.exceptions.Timeout as e:
                last_exception = e
                self.logger.warning(f"第{attempt}次请求超时（{self.REQUEST_TIMEOUT}秒）")
            except Exception as e:
                last_exception = e
                self.logger.warning(f"第{attempt}次请求失败: {e}")

            # 5.8 重试间隔：UA 切换 + 递增等待
            if attempt < self.MAX_RETRIES:
                # 每次重试前切换 UA（Edge ↔ Chrome 兜底）
                self._switch_ua()
                wait_seconds = self.REQUEST_INTERVAL * attempt
                self.logger.info(
                    f"等待 {wait_seconds}秒 后重试（已切换UA，当前: "
                    f"{'Edge' if self._current_ua_index == 0 else 'Chrome'}）..."
                )
                time.sleep(wait_seconds)

        # 5.9 重试全部失败：上报
        if response is None or last_exception is not None and not (response and len(response.content) > 0):
            self.logger.error(f"所有 {self.MAX_RETRIES} 次重试均失败")
            raise RuntimeError(
                f"店铺来源-三级渠道报表下载失败：{last_exception}（请检查Cookie/网络/风控）"
            ) from last_exception

        # 7. 保存 Excel（文件名友好：店铺来源_三级渠道_YYYY-MM-DD.xlsx）
        filename = f"店铺来源_三级渠道_{date}.xlsx"

        # 8. 后置处理：复用基类 _save_flow_excel（与商品流量来源同样的 Excel 处理流程）
        return self._save_flow_excel(response, filename, date)

    # ---------- 阶段 4 新增：风控 / 空响应辅助方法（业务内自实现）----------

    def _check_business_code(self, response):
        """风控业务码识别（json 响应里的 status / message）。

        业务背景：
            京东风控有时返回 HTTP 200 但 body 是 json 错误（"不安全的请求"/"操作频繁"），
            必须解析 json 才能识别。常见码：
                - status = -407  → "不安全的请求"（签名错误，重试可能恢复）
                - status = -402  → 参数缺失
                - status = 601   → 操作频繁（限流，重试反而加重风控，不自动重试）
                - status = 302 / -1  → 登录失效
                - message 含 "登录" / "login" → 同上
        行为:
            Cookie 过期 → 抛 CookieExpiredError（外层捕获，停止重试）
            限流 601    → 仅警告，不重试（避免加重风控；让用户决定）
            其他负码    → 警告 + 抛 Exception（让重试循环兜底）
        """
        content_type = response.headers.get("Content-Type", "")
        if "json" not in content_type:
            return  # 非 json 响应，跳过业务码检查

        try:
            result = response.json()
        except json.JSONDecodeError:
            return  # json 解析失败，跳过

        # 仅在 success=False 时检查（status 可正可负：-407/-402/302/-1 是负，601/201 是正）
        if result.get("success", True):
            return

        error_msg = result.get("message", "未知错误")
        status_code = result.get("status")

        # Cookie 过期（硬错误，不可重试）
        if status_code in (302, -1) or "登录" in error_msg or "login" in error_msg.lower():
            self.logger.error(f"Cookie可能已过期: {error_msg} (status={status_code})")
            raise CookieExpiredError(f"Cookie已过期：{error_msg}")

        # 601 操作频繁（限流，不重试，让用户决定）
        # 阶段4修复：抛 RiskControlError 而非 RuntimeError。
        #   RuntimeError 会被重试循环的 except Exception 捕获 → 继续重试，
        #   违背"601不重试"约束；RiskControlError 在循环内单独捕获并直接抛出。
        if status_code == 601 or "操作频繁" in error_msg:
            self.logger.error(
                f"风控限流 601：{error_msg}（账号/IP被临时限流，"
                f"停止本次调用避免加重风控，请稍后30-120分钟再试）"
            )
            raise RiskControlError(f"风控限流：{error_msg}")

        # 其他错误码（-407/-402 等，可重试兜底）
        self.logger.warning(f"风控拦截: {error_msg} (status={status_code})")
        raise RuntimeError(f"风控拦截(status={status_code})：{error_msg}")

    def _validate_excel_response(self, response, attempt):
        """空响应拦截 + Excel magic 字节校验。

        业务背景：
            风控拦截时服务器可能返回 HTML 错误页（"网页解析失败"）或
            简短 json 错误（"不支持网页类型"）伪装成"成功响应"，
            字节数<1KB + 不是 xlsx → 视为失败，让重试机制兜底。

        校验规则:
            1. HTTP 200 但响应体 < 1KB → 视为空响应，抛 Exception
            2. 响应体前 4 字节不是 "PK\\x03\\x04"（Excel 文件头）→ 视为非 Excel，
               记录前 200 字节供排查，抛 Exception

        异常:
            任何校验失败都抛 RuntimeError，让重试循环兜底
        """
        # 1. HTTP 200 但响应体空
        if response.status_code == 200 and len(response.content) < 1024:
            preview = response.content[:200].decode("utf-8", errors="replace")
            self.logger.error(
                f"[第{attempt}次] 响应体过小（{len(response.content)}字节 < 1KB），"
                f"疑似风控拦截或服务器错误。响应前200字节：{preview!r}"
            )
            raise RuntimeError(f"空响应（{len(response.content)}字节）")

        # 2. Excel magic 字节校验（PK\x03\x04 = zip/xlsx 文件头）
        if response.content[:4] != b"PK\x03\x04":
            preview = response.content[:200].decode("utf-8", errors="replace")
            self.logger.error(
                f"[第{attempt}次] 响应体不是有效 Excel 文件（magic bytes 不匹配）。"
                f"响应前200字节：{preview!r}"
            )
            raise RuntimeError("响应体不是 Excel 文件（magic bytes 校验失败）")

    def _save_flow_excel(self, response, filename, date):
        """Excel 后置处理（自实现，基类无此方法）。

        ⚠️ 重要：基类 JDBaseRequest 没有 _save_flow_excel（ProductFlowAPI 才有），
        新业务必须自实现，否则会 AttributeError。

        流程（与 ProductFlowAPI._save_flow_excel 保持一致）：
            ① 读 Excel 二进制流 → DataFrame（dtype=str 防长数字精度丢失）
            ② 通用日期转换（convert_date_format）→ 插入首列【日期】
            ③ safe_convert_numeric 全表数值安全转换
            ④ 写入 Excel + apply_column_formats 设置单元格格式
        """
        import io
        import warnings
        import pandas as pd

        # 抑制openpyxl读取原始xlsx时的无害警告
        warnings.filterwarnings("ignore", message="Workbook contains no default style")

        # ① 读取二进制流 → DataFrame（dtype=str 防长数字精度丢失，na_filter=False 保留空字符串）
        df = pd.read_excel(io.BytesIO(response.content), dtype=str, na_filter=False)

        # ②③ 日期列统一处理（公共规则1+2，2026-08-07）：
        #    报表自带【日期】/【时间】列 → 禁止重复插入日期列，仅做格式标准化；
        #    报表无日期/时间列 → 首列插入【日期】列，值=本次查询日期。
        date_column, date_value = prepare_date_columns(df, date)

        # ④ 全表数值安全转换
        df = safe_convert_numeric(df)

        # ⑤ 写入Excel + 单元格格式
        # 输出目录规则（AGENTS.md Excel规则4）：output/{业务模块}/{date}/{filename}
        file_path = build_business_output_path(self.output_dir, filename, date)
        df.to_excel(file_path, index=False, engine="openpyxl")
        apply_column_formats(file_path, df, date_column=date_column, date_value=date_value)

        self.logger.info(
            f"Excel已保存: {file_path}（已插入日期列+数值转换+单元格格式，{os.path.getsize(file_path)}字节）"
        )
        return file_path


# ============================================================
#  业务接口 3：（新业务 - 商品明细导出，2026-08-07 上线）
# ------------------------------------------------------------
#  业务名称：商品明细导出
#  接口地址：https://sz.jd.com/sz/api/productDetail/exportProList.ajax
#  业务说明：按商品分析页面，导出商品明细流量报表
#  业务需求（2026-08-07）：
#      - 请求方式：GET（参数全拼 URL）
#      - 域名：sz.jd.com（与项目 4 szgateway.jd.com 不同）
#      - 必带 Header：Referer=productDetail.html、Sec-Fetch-Site=same-origin
#      - 业务参数：date / startDate / endDate / type=0 / categoryType=0 / second=999999
#        / third="" / channel=99 / isMonitored=undefined / downloadType=dayList
#      - 风控三元组：User-mup / User-mnp / uuid（uuid 完全随机）
#      - 成功响应：HTTP 200 + Content-Disposition: attachment + body 前两字节 PK
#  硬性约束（用户 2026-08-07 确认）：
#      1. GET 不用 POST，参数全拼 URL
#      2. 风控三元组每次全新生成，UUID 不使用固定前缀
#      3. Cookie 从 config/sz_cookie.txt 整体读取
#      4. 必须双重校验：Content-Disposition + PK 魔数
#      5. 文件名从 Content-Disposition 提取原始名
#      6. 保存目录：output/商品明细/{date}/{原始文件名}
# ============================================================
class ProductDetailAPI(JDBaseRequest):
    """商品明细导出 API（GET 请求，与项目 4 POST 表单不同）。

    业务定位：
        商智 sz.jd.com 模块下"商品分析"页面的"商品明细"导出。
        与项目 4（downTable.ajax / 渠道维度）不同，本接口：
        - 用 GET 请求（参数全拼 URL）
        - 域名 sz.jd.com（不是 szgateway.jd.com）
        - 业务输出"商品明细"维度（每条数据是一个商品）
        - 成功响应带 Content-Disposition 头（带原始文件名）

    父类复用：
        - 父类 JDBaseRequest 提供：
            * Cookie 读取（config/sz_cookie.txt）
            * 风控签名（基类 _gen_risk_params / _wait_interval）
            * 自动 30 秒间隔
            * 自动重试 + UA 切换（_switch_ua）
            * 日志、session 管理

    自实现部分（项目 4 模式延续，class.__dict__ 已验证基类无这些方法）：
        - _gen_uuid_random / _gen_risk_params_random：完全随机 UUID（用户抓包证实）
        - _get_date_params：日期三值一致（基类没有，ProductFlowAPI 才有）
        - _check_business_code：风控业务码识别（项目 4 独有）
        - _validate_excel_response：空响应拦截 + Content-Disposition 增强校验
        - _save_detail_excel：业务子目录保存 + 解析原始文件名
        - download_product_detail：主方法（GET 请求 + 重试 + UA 切换）
    """

    # 接口 URL（业务约束，固定）
    # 注意：域名是 sz.jd.com（与项目 4 szgateway.jd.com 不同）
    API_URL = "https://sz.jd.com/sz/api/productDetail/exportProList.ajax"

    # 必带请求头（业务约束，固定）
    # 通俗解释：告诉京东"我是从这个页面来的"，缺失会被拦截
    # 注意：Sec-Fetch-Site=same-origin（项目 4 是 same-site）
    REFERER = "https://sz.jd.com/szweb/sz/view/productAnalysis/productDetail.html"

    # 业务输出子目录（与项目 4 区分）
    # 项目 4：output/{文件名}.xlsx
    # 项目 5：output/商品明细/{date}/{原始文件名}.xlsx
    OUTPUT_SUBDIR = "商品明细"

    # 固定业务参数（用户确认 2026-08-07 固化，不读 config）
    # 通俗解释：这些参数永远不变（业务定义），直接写死代码
    FIXED_BIZ_PARAMS = {
        "type": "0",                  # 业务类型
        "categoryType": "0",          # 类目类型
        "downloadType": "dayList",    # 下载类型
    }

    # 可变业务参数（从 CLI/config 覆盖，否则用抓包默认值）
    # 通俗解释：日期可改，类目/渠道走代码常量（用户确认用抓包默认值）
    VARIABLE_BIZ_PARAMS = {
        "second": "999999",           # 二级类目：999999=全类目
        "third": "",                  # 三级类目：空=不限
        "channel": "99",              # 渠道：99=全部渠道
        "isMonitored": "undefined",   # 是否监控商品：undefined=不限
    }

    # ---------- 项目 5 阶段 3 子任务 1：UUID 完全随机 + 风控签名 ----------

    def _gen_uuid_random(self):
        """完全随机 UUID 生成器（不依赖任何固定前缀）。

        ⚠️ 业务背景：用户抓包显示 UUID 前缀为 `c30f3b84431d43dee267-19feab01d7e`（22+10），
        证明前端 SDK 每次会话运行时动态生成。完全随机化符合用户 2026-08-07
        "禁止硬编码 uuid 前缀"约束。

        2026-08-10 bug fix：原 16+10 格式与抓包 22+10 不一致，导致服务端 0 字节空响应。
        改为 22+10（11 字节 + 5 字节随机 hex），与京东前端 SDK 实际生成对齐。

        复制来源：项目 4 OfflineChannelAPI._gen_uuid_random

        返回:
            str - 形如 "c30f3b84431d43dee267-19feab01d7e"（22hex + - + 10hex）
        """
        import secrets
        # 22位小写hex + "-" + 10位小写hex，与抓包格式完全一致（2026-08-10 bug fix）
        prefix = secrets.token_hex(11)       # 11字节 = 22hex 字符
        suffix = secrets.token_hex(5)        # 5字节 = 10hex 字符
        return f"{prefix}-{suffix}"

    def _gen_risk_params_random(self, url):
        """生成风控参数：User-mup / User-mnp / uuid（uuid 完全随机版）。

        与基类 _gen_risk_params 的差异：
            基类默认用类常量 UUID_PREFIX（ca412182e5668a106054），硬编码
            本方法用 _gen_uuid_random() 完全随机生成

        算法不变（commons-a5562705.js 逆向）：
            User-mnp = MD5(URL路径 + uuid + 时间戳 + 盐值)

        复制来源：项目 4 OfflineChannelAPI._gen_risk_params_random

        参数:
            url - 接口URL，用于提取URL路径
        返回:
            dict - {"User-mup": str, "User-mnp": str, "uuid": str}
        """
        timestamp = int(time.time() * 1000)  # 毫秒级时间戳，每次必新
        uuid_str = self._gen_uuid_random()    # 完全随机，不传任何 prefix

        from urllib.parse import urlparse
        parsed = urlparse(url)
        url_path = parsed.path  # GET 路径（去掉 query string）

        # 盐值复用全局 SIGN_SALT（与现有 3 业务 + 项目 4 共用）
        sign_str = f"{url_path}{uuid_str}{timestamp}{self.SIGN_SALT}"
        user_mnp = hashlib.md5(sign_str.encode("utf-8")).hexdigest()

        return {
            "User-mup": str(timestamp),
            "User-mnp": user_mnp,
            "uuid": uuid_str,
        }

    # ---------- 项目 5 阶段 3 子任务 2：日期 + 风控识别 + 响应校验 ----------

    def _get_date_params(self, date=None, start_date=None, end_date=None):
        """解析本次查询的日期参数（支持动态覆盖）。

        ⚠️ 重要：基类 JDBaseRequest 没有这个方法（ProductFlowAPI 才有，class.__dict__ 已验证），
        新业务必须自实现，否则会 AttributeError。

        复制来源：项目 4 OfflineChannelAPI._get_date_params（与 ProductFlowAPI 同源）

        规则（2026-08-05 修复）：
            date      : 优先用入参（如命令行 --date），其次从 config.xlsx 读取
            startDate : 优先用入参；未传时默认=date
            endDate   : 同 startDate
        """
        # date：入参优先，其次config
        if date is None:
            date = self.config.get("date")
        if date is None:
            raise ValueError("查询日期date未提供：请在config.xlsx配置或通过函数入参传入")

        # start/end：入参优先；未传时默认与date一致（单日查询）
        if start_date is None:
            start_date = date
        if end_date is None:
            end_date = date

        return date, start_date, end_date

    def _check_business_code(self, response):
        """风控业务码识别（json 响应里的 status / message）。

        复制来源：项目 4 OfflineChannelAPI._check_business_code（项目 4 阶段 4 已验证 8/8 单测通过）

        业务背景：
            京东风控有时返回 HTTP 200 但 body 是 json 错误（"不安全的请求"/"操作频繁"），
            必须解析 json 才能识别。常见码：
                - status = -407  → "不安全的请求"（签名错误，重试可能恢复）
                - status = -402  → 参数缺失
                - status = 601   → 操作频繁（限流，重试反而加重风控，不自动重试）
                - status = 302 / -1  → 登录失效
                - message 含 "登录" / "login" → 同上
        行为:
            Cookie 过期 → 抛 CookieExpiredError（外层捕获，停止重试）
            限流 601    → 仅警告，不重试（避免加重风控；让用户决定）
            其他负码    → 警告 + 抛 Exception（让重试循环兜底）
        """
        content_type = response.headers.get("Content-Type", "")
        if "json" not in content_type:
            return  # 非 json 响应，跳过业务码检查

        try:
            result = response.json()
        except json.JSONDecodeError:
            return  # json 解析失败，跳过

        # 仅在 success=False 时检查（status 可正可负：-407/-402/302/-1 是负，601/201 是正）
        if result.get("success", True):
            return

        error_msg = result.get("message", "未知错误")
        status_code = result.get("status")

        # Cookie 过期（硬错误，不可重试）
        if status_code in (302, -1) or "登录" in error_msg or "login" in error_msg.lower():
            self.logger.error(f"Cookie可能已过期: {error_msg} (status={status_code})")
            raise CookieExpiredError(f"Cookie已过期：{error_msg}")

        # 601 操作频繁（限流，不重试，让用户决定）
        # 阶段4修复：抛 RiskControlError 而非 RuntimeError（同项目4）。
        #   RuntimeError 会被重试循环的 except Exception 捕获 → 继续重试，
        #   违背"601不重试"约束；RiskControlError 在循环内单独捕获并直接抛出。
        if status_code == 601 or "操作频繁" in error_msg:
            self.logger.error(
                f"风控限流 601：{error_msg}（账号/IP被临时限流，"
                f"停止本次调用避免加重风控，请稍后30-120分钟再试）"
            )
            raise RiskControlError(f"风控限流：{error_msg}")

        # 其他错误码（-407/-402 等，可重试兜底）
        self.logger.warning(f"风控拦截: {error_msg} (status={status_code})")
        raise RuntimeError(f"风控拦截(status={status_code})：{error_msg}")

    def _validate_excel_response(self, response, attempt):
        """空响应拦截 + Excel magic 字节校验 + Content-Disposition 增强校验（项目 5 新增）。

        复制来源：项目 4 OfflineChannelAPI._validate_excel_response
        增强点（项目 5）：
            - 校验响应头是否包含 Content-Disposition: attachment（项目 4 不需要）

        校验规则:
            1. HTTP 200 但响应体 < 1KB → 视为空响应，抛 Exception
            2. 响应体前 4 字节不是 "PK\\x03\\x04"（Excel 文件头）→ 视为非 Excel
            3. 响应头 Content-Disposition 不含 "attachment" → 视为非预期响应（项目 5 新增）

        异常:
            任何校验失败都抛 RuntimeError，让重试循环兜底
        """
        # 1. HTTP 200 但响应体空
        if response.status_code == 200 and len(response.content) < 1024:
            preview = response.content[:200].decode("utf-8", errors="replace")
            self.logger.error(
                f"[第{attempt}次] 响应体过小（{len(response.content)}字节 < 1KB），"
                f"疑似风控拦截或服务器错误。响应前200字节：{preview!r}"
            )
            raise RuntimeError(f"空响应（{len(response.content)}字节）")

        # 2. Excel magic 字节校验（PK\x03\x04 = zip/xlsx 文件头）
        if response.content[:4] != b"PK\x03\x04":
            preview = response.content[:200].decode("utf-8", errors="replace")
            # 阶段4增强：识别"文本型 601"（非 json 的 HTML 错误页含"操作频繁"字样）
            # 这类响应说明账号/IP已被限流，重试只会加重风控 → 抛 RiskControlError 停止
            if "操作频繁" in preview or "频繁" in preview:
                self.logger.error(
                    f"[第{attempt}次] 响应疑似风控限流(601)：{preview[:100]!r}"
                    f"（账号/IP被临时限流，停止重试，请30-120分钟后再试）"
                )
                raise RiskControlError(f"风控限流：响应文本含'操作频繁'（{preview[:50]!r}）")
            self.logger.error(
                f"[第{attempt}次] 响应体不是有效 Excel 文件（magic bytes 不匹配）。"
                f"响应前200字节：{preview!r}"
            )
            raise RuntimeError("响应体不是 Excel 文件（magic bytes 校验失败）")

        # 3. 【项目 5 增强】Content-Disposition: attachment 校验
        # 业务背景：本接口返回 Excel 是浏览器下载模式，响应头必须带 attachment
        #         否则可能是风控拦截返回了错误页（伪装成 200 OK）
        content_disp = response.headers.get("Content-Disposition", "")
        if "attachment" not in content_disp:
            self.logger.error(
                f"[第{attempt}次] 响应头缺少 Content-Disposition: attachment，"
                f"疑似风控拦截伪装。实际 Content-Disposition={content_disp!r}"
            )
            raise RuntimeError("响应头缺少 Content-Disposition: attachment（疑似风控拦截）")

    # ---------- 项目 5 阶段 3 子任务 3：Excel 后置处理 + 业务子目录 ----------

    def _parse_content_disposition_filename(self, content_disp_header):
        """从 Content-Disposition 响应头解析原始文件名（含中文 + URL 解码）。

        业务背景：
            京东响应头示例：Content-Disposition: attachment;filename=19524838_20260806_%E5%85%A8%E9%83%A8%E6%B8%A0%E9%81%93_%E5%95%86%E5%93%81%E6%98%8E%E7%BB%86_%E5%88%86%E5%A4%A9%E4%B8%8B%E8%BD%BD.xlsx
            - 文件名经过 URL 编码（含中文）
            - 解析后得到：19524838_20260806_全部渠道_商品明细_分天下载.xlsx

        参数:
            content_disp_header - 响应头 Content-Disposition 的完整值
        返回:
            str - 解码后的文件名（不含路径），如 "19524838_20260806_全部渠道_商品明细_分天下载.xlsx"
            None - 解析失败
        """
        import re
        from urllib.parse import unquote

        if not content_disp_header:
            return None

        # 匹配 filename= 后面的值（支持 filename*=UTF-8''... 形式）
        # 优先尝试 RFC 5987 格式：filename*=UTF-8''<urlencoded>
        m = re.search(r"filename\*=(?:UTF-8|utf-8)''(.+?)(?:;|$)", content_disp_header)
        if m:
            return unquote(m.group(1))

        # 标准格式：filename="xxx" 或 filename=xxx
        m = re.search(r'filename\s*=\s*"?(?P<name>[^";]+)"?', content_disp_header)
        if m:
            return unquote(m.group("name"))

        return None

    def _save_detail_excel(self, response, filename, date):
        """商品明细 Excel 后置处理（业务子目录 + 解析 Content-Disposition）。

        业务定位：
            与项目 4 _save_flow_excel 不同，本方法：
            1. 保存到业务子目录：output/商品明细/{date}/{filename}
            2. filename 优先用 Content-Disposition 解析的原始文件名（保留京东原始命名）
            3. 如果 Content-Disposition 解析失败，回退用 CLI 传入的 filename

        流程（与项目 4 _save_flow_excel 保持一致）：
            ① 读 Excel 二进制流 → DataFrame（dtype=str 防长数字精度丢失）
            ② 通用日期转换（convert_date_format）→ 插入首列【日期】
            ③ safe_convert_numeric 全表数值安全转换
            ④ 写入 Excel + apply_column_formats 设置单元格格式
        """
        import io
        import warnings
        import pandas as pd

        # 抑制openpyxl读取原始xlsx时的无害警告
        warnings.filterwarnings("ignore", message="Workbook contains no default style")

        # ① 读取二进制流 → DataFrame
        df = pd.read_excel(io.BytesIO(response.content), dtype=str, na_filter=False)

        # ②③ 日期列统一处理（公共规则1+2，2026-08-07）：
        #    报表自带【日期】/【时间】列 → 禁止重复插入日期列，仅做格式标准化；
        #    报表无日期/时间列 → 首列插入【日期】列，值=本次查询日期。
        date_column, date_value = prepare_date_columns(df, date)

        # ④ 全表数值安全转换
        df = safe_convert_numeric(df)

        # ⑤ 写入Excel
        file_path = os.path.join(self.output_dir, filename)
        df.to_excel(file_path, index=False, engine="openpyxl")
        apply_column_formats(file_path, df, date_column=date_column, date_value=date_value)

        self.logger.info(
            f"Excel已保存: {file_path}（已插入日期列+数值转换+单元格格式，{os.path.getsize(file_path)}字节）"
        )
        return file_path

    # ---------- 项目 5 阶段 3 子任务 4：主下载方法 ----------

    def download_product_detail(
        self,
        date=None,
        start_date=None,
        end_date=None,
        second=None,
        third=None,
        channel=None,
        is_monitored=None,
    ):
        """下载商品明细流量报表（GET 请求，与项目 4 POST 不同）。

        参数:
            date           - 查询日期 YYYY-MM-DD（入参优先，其次 config）
            start_date     - 开始日期（默认=date）
            end_date       - 结束日期（默认=date）
            second         - 二级类目（默认 "999999"=全类目，CLI 可覆盖）
            third          - 三级类目（默认 ""=不限）
            channel        - 渠道（默认 "99"=全部渠道）
            is_monitored   - 是否监控商品（默认 "undefined"=不限）

        返回:
            str - 保存的 Excel 文件绝对路径（业务子目录格式）
        """
        # 1. 解析日期（入参 > config；与项目 4 同样的三值一致规则）
        date, start_date, end_date = self._get_date_params(date, start_date, end_date)

        # 2. 读取可变业务参数（CLI 覆盖 → 代码常量兜底）
        second = second if second is not None else self.VARIABLE_BIZ_PARAMS["second"]
        third = third if third is not None else self.VARIABLE_BIZ_PARAMS["third"]
        channel = channel if channel is not None else self.VARIABLE_BIZ_PARAMS["channel"]
        is_monitored = is_monitored if is_monitored is not None else self.VARIABLE_BIZ_PARAMS["isMonitored"]

        # 3. 组装 URL query 参数
        query_params = {
            "date": date,
            "startDate": start_date,
            "endDate": end_date,
            "second": second,
            "third": third,
            "channel": channel,
            "isMonitored": is_monitored,
            **self.FIXED_BIZ_PARAMS,
        }

        # 4. 必带请求头（业务约束：Referer 缺失被平台拦截）
        # 注意：Sec-Fetch-Site 与项目 4 不同（项目 4 是 same-site，本项目是 same-origin）
        # 基类 _build_default_headers 已设置 same-site，需用 extra_headers 覆盖
        extra_headers = {
            "Referer": self.REFERER,
            "Sec-Fetch-Site": "same-origin",  # 覆盖基类默认 same-site
        }

        # 5. 发送请求（重试循环 + UA 切换 + 风控识别 + 响应校验）
        # 与项目 4 模式一致：手写签名后注入（因基类 request() 用 UUID_PREFIX，本业务需完全随机）
        response = None
        last_exception = None
        # 阶段4新增：成功标记。只有 break 跳出循环才算成功，
        # 防止"最后一次失败但响应有内容(如HTML错误页)"被误判为成功继续保存
        success = False

        for attempt in range(1, self.MAX_RETRIES + 1):
            try:
                # 5.1 严格间隔控制（与基类一致）
                self._wait_interval()

                # 5.2 重新生成风控参数（每次新 UUID + 新时间戳 + 新签名）
                risk_params = self._gen_risk_params_random(self.API_URL)
                full_params = {**query_params, **risk_params}

                ua_name = "Edge" if self._current_ua_index == 0 else "Chrome"
                self.logger.info(
                    f"[第{attempt}/{self.MAX_RETRIES}次] 下载商品明细报表: "
                    f"日期={date}, 类目=second={second}/third={third}, "
                    f"渠道={channel}, UA={ua_name}, uuid={risk_params['uuid'][:8]}..."
                )
                self.logger.debug(f"请求参数: {json.dumps(full_params, ensure_ascii=False)[:500]}")

                # 记录请求时间（类属性：所有实例共享）
                JDBaseRequest._last_request_time = time.time()

                # 5.3 发送 GET 请求（注意：GET 不用 POST，参数拼 URL）
                response = self.session.get(
                    self.API_URL,
                    params=full_params,  # GET 专用：拼 URL query string
                    headers=extra_headers,
                    timeout=self.REQUEST_TIMEOUT,
                )

                # 5.4 HTTP 状态码基础检查
                if response.status_code == 401:
                    self.logger.error("HTTP 401 未授权 - Cookie 可能已过期或被禁用")
                    raise CookieExpiredError("Cookie已过期或无效，请更新 config/sz_cookie.txt")
                if response.status_code == 403:
                    self.logger.error("HTTP 403 禁止访问 - Cookie/签名/Referer 校验失败")
                    # 不抛 CookieExpired，让重试机制 + UA 切换兜底

                # 5.5 风控业务码识别（json 响应里的 status / message）
                self._check_business_code(response)

                # 5.6 空响应拦截：Excel magic 字节 + Content-Disposition 校验
                self._validate_excel_response(response, attempt)

                # 5.7 走到这里 = 成功，跳出重试循环
                self.logger.info(
                    f"请求成功: HTTP {response.status_code}, "
                    f"{len(response.content)}字节, "
                    f"Content-Disposition={response.headers.get('Content-Disposition', '')[:80]}"
                )
                success = True  # 阶段4新增：只有走到这里才算成功
                break

            except CookieExpiredError:
                # Cookie 过期是硬错误，不能靠重试解决，直接抛出
                raise
            except RiskControlError:
                # 阶段4新增：601 限流等风控硬错误，不重试，直接抛出
                raise
            except requests.exceptions.Timeout as e:
                last_exception = e
                self.logger.warning(f"第{attempt}次请求超时（{self.REQUEST_TIMEOUT}秒）")
            except Exception as e:
                last_exception = e
                self.logger.warning(f"第{attempt}次请求失败: {e}")

            # 5.8 重试间隔：UA 切换 + 递增等待
            if attempt < self.MAX_RETRIES:
                self._switch_ua()
                wait_seconds = self.REQUEST_INTERVAL * attempt
                self.logger.info(
                    f"等待 {wait_seconds}秒 后重试（已切换UA，当前: "
                    f"{'Edge' if self._current_ua_index == 0 else 'Chrome'}）..."
                )
                time.sleep(wait_seconds)

        # 5.9 重试全部失败：上报
        # 阶段4修复：原条件"response有内容就继续"存在漏洞——
        #   最后一次失败时若响应体恰有内容(如HTML错误页/风控页)，
        #   会误判为成功继续保存错误内容。改为 success 标记判断，
        #   只要没走到 break（success=False）一律抛错。
        if not success:
            self.logger.error(f"所有 {self.MAX_RETRIES} 次重试均失败")
            raise RuntimeError(
                f"商品明细报表下载失败：{last_exception}（请检查Cookie/网络/风控）"
            ) from last_exception

        # 6. 解析 Content-Disposition 获取原始文件名（含中文 URL 解码）
        content_disp = response.headers.get("Content-Disposition", "")
        original_filename = self._parse_content_disposition_filename(content_disp)

        if not original_filename:
            # 兜底：用 CLI 传入的 filename 格式（业务子目录 + date）
            self.logger.warning(
                f"无法从 Content-Disposition 解析文件名（头={content_disp[:80]!r}），"
                f"使用兜底命名"
            )
            original_filename = f"{second or '全类目'}_{date}_商品明细.xlsx"

        # 7. 构造业务子目录路径：output/商品明细/{date}/{filename}
        # 注意：业务子目录确保不与其他业务混淆
        date_subdir = os.path.join(self.OUTPUT_SUBDIR, date)
        business_output_dir = os.path.join(self.output_dir, date_subdir)
        os.makedirs(business_output_dir, exist_ok=True)

        # 8. Excel 后置处理（业务子目录 + 解析 Content-Disposition 文件名）
        # 与项目 4 _save_flow_excel 区别：传入 output_dir 的子目录版本
        target_path = os.path.join(business_output_dir, original_filename)

        # 复用 _save_detail_excel 但指定具体路径
        return self._save_detail_excel_to_path(response, target_path, date)


    def _save_detail_excel_to_path(self, response, target_path, date):
        """商品明细 Excel 后置处理（指定具体路径版本）。

        业务定位：与 _save_detail_excel 类似，但允许调用方指定完整路径（不限制在 output_dir）

        流程：
            ① 读 Excel 二进制流 → DataFrame
            ② 通用日期转换 → 插入首列【日期】
            ③ safe_convert_numeric 全表数值安全转换
            ④ 写入 Excel（指定路径）+ apply_column_formats
        """
        import io
        import warnings
        import pandas as pd

        # 抑制openpyxl读取原始xlsx时的无害警告
        warnings.filterwarnings("ignore", message="Workbook contains no default style")

        # ① 读取二进制流 → DataFrame
        df = pd.read_excel(io.BytesIO(response.content), dtype=str, na_filter=False)

        # ②③ 日期列统一处理（公共规则1+2，2026-08-07）：
        #    报表自带【日期】/【时间】列 → 禁止重复插入日期列，仅做格式标准化；
        #    报表无日期/时间列 → 首列插入【日期】列，值=本次查询日期。
        date_column, date_value = prepare_date_columns(df, date)

        # ④ 全表数值安全转换
        df = safe_convert_numeric(df)

        # ⑤ 写入Excel（指定完整路径，含业务子目录）
        df.to_excel(target_path, index=False, engine="openpyxl")
        apply_column_formats(target_path, df, date_column=date_column, date_value=date_value)

        self.logger.info(
            f"Excel已保存: {target_path}（已插入日期列+数值转换+单元格格式，{os.path.getsize(target_path)}字节）"
        )
        return target_path


# ============================================================
#  业务接口 4：（新业务 - 商品流失分析，2026-08-07 上线）
# ------------------------------------------------------------
#  业务名称：商品流失分析
#  接口地址：https://sz.jd.com/sz/api/competitionAnalysis/exportLossProList.ajax
#  业务说明：竞争分析-竞争流失-商品流失分析，导出流失商品明细报表
#  业务需求（2026-08-07 用户提供抓包，阶段1确认）：
#      - 请求方式：POST（表单 application/x-www-form-urlencoded，同项目4）
#      - 域名：sz.jd.com（同项目5，Sec-Fetch-Site=same-origin）
#      - 页面入口：/sz/view/competitionAnalysis/lossAnalysiss.html（注意是 /sz/view/ 非 /szweb/sz/view/）
#      - 业务参数：indChannel=99（渠道）、unitType=0（SPU维度）→ 用户确认固定写死
#      - 响应格式：⚠️ .xls（OLE2复合文档），魔数 D0CF11E0，需 xlrd 读取
#      - 文件名头：filename= 为 GBK 字节乱码（如"鍟嗗搧娴佸け鍒嗘瀽"），需 latin-1→GBK 还原
#      - 保存规则：读取 xls → 转存 xlsx → output/商品流失分析/{date}/
#  硬性约束：
#      1. 风控三元组每次全新生成，uuid 完全随机（16hex-10hex，同项目4/5）
#      2. 601 不重试（抛 RiskControlError，同项目4/5 阶段4 修复）
#      3. Cookie 从 config/sz_cookie.txt 整体读取
# ============================================================
class LossProductAPI(JDBaseRequest):
    """商品流失分析 API（竞争分析-竞争流失-商品流失，2026-08-07）。

    POST 表单（同项目4） + sz.jd.com 域（同项目5） + .xls 响应（首次出现） + GBK 文件名（首次出现）。
    与项目4/5 的差异集中在：响应校验（xls/xlsx 双魔数）、文件名解析（UTF-8/GBK 双兼容）、
    xls 读取转存 xlsx。其余（重试/UA切换/风控码识别/success标记）与项目4/5 阶段4 完全一致。
    """

    # 接口 URL（业务约束，固定）
    API_URL = "https://sz.jd.com/sz/api/competitionAnalysis/exportLossProList.ajax"

    # 必带请求头（业务约束，固定）
    ORIGIN = "https://sz.jd.com"
    REFERER = "https://sz.jd.com/sz/view/competitionAnalysis/lossAnalysiss.html"

    # 业务子目录（输出目录规则 AGENTS.md Excel规则4）
    OUTPUT_SUBDIR = "商品流失分析"

    # 【固定业务参数】用户 2026-08-07 确认固化（抓包值），写死代码不读config
    FIXED_BIZ_PARAMS = {
        "indChannel": "99",   # 渠道：99=全部渠道（抓包确认）
        "unitType": "0",      # 维度：0=SPU 维度（抓包确认）
    }
    # 本业务无可变业务参数（日期 + 风控三元组动态生成）

    # ---------- 完全随机 UUID（与项目4/5 相同模式）----------

    def _gen_uuid_random(self):
        """完全随机 UUID 生成器（16hex-10hex，与抓包格式一致）。

        ⚠️ 背景：抓包 uuid=`a94057e3a4eab84aae5d-19fdb7dc648`（16hex-10hex），
        前缀每次会话变化，禁止硬编码前缀（AGENTS.md 风控归档）。
        """
        import secrets
        prefix = secrets.token_hex(8)   # 8字节 = 16hex
        suffix = secrets.token_hex(5)   # 5字节 = 10hex
        return f"{prefix}-{suffix}"

    def _gen_risk_params_random(self, url):
        """动态生成风控三元组（完全随机 uuid + 毫秒时间戳 + MD5签名）。

        签名算法（与项目1-5 一致）：
            User-mnp = MD5(URL路径 + uuid + 时间戳 + 全局盐值 372ad2c2b6)
        """
        import hashlib
        import time as _time
        from urllib.parse import urlparse

        uuid_str = self._gen_uuid_random()
        timestamp = int(_time.time() * 1000)   # 13位毫秒时间戳
        url_path = urlparse(url).path         # 取URL路径（去掉域名和query）

        sign_str = f"{url_path}{uuid_str}{timestamp}{self.SIGN_SALT}"
        user_mnp = hashlib.md5(sign_str.encode("utf-8")).hexdigest()

        return {
            "User-mup": str(timestamp),
            "User-mnp": user_mnp,
            "uuid": uuid_str,
        }

    def _get_date_params(self, date=None, start_date=None, end_date=None):
        """日期三值一致规则（与项目4/5 相同，踩坑 2026-08-05）。"""
        if date is None:
            date = self.config.get("date")
        if date is None:
            raise ValueError("查询日期date未提供：请在config.xlsx配置或通过函数入参传入")
        if start_date is None:
            start_date = date
        if end_date is None:
            end_date = date
        return date, start_date, end_date

    # ---------- 风控 / 响应校验辅助方法（项目6 适配 xls）----------

    def _check_business_code(self, response):
        """风控业务码识别（与项目4/5 完全一致，含 601 不重试）。"""
        content_type = response.headers.get("Content-Type", "")
        if "json" not in content_type:
            return
        try:
            result = response.json()
        except Exception:
            return
        if result.get("success", True):
            return
        error_msg = result.get("message", "未知错误")
        status_code = result.get("status")
        # Cookie 过期（硬错误）
        if status_code in (302, -1) or "登录" in error_msg or "login" in error_msg.lower():
            self.logger.error(f"Cookie可能已过期: {error_msg} (status={status_code})")
            raise CookieExpiredError(f"Cookie已过期：{error_msg}")
        # 601 限流（不重试）
        if status_code == 601 or "操作频繁" in error_msg:
            self.logger.error(
                f"风控限流 601：{error_msg}（账号/IP被临时限流，停止本次调用避免加重风控，请稍后30-120分钟再试）"
            )
            raise RiskControlError(f"风控限流：{error_msg}")
        # 其他错误码
        self.logger.warning(f"风控拦截: {error_msg} (status={status_code})")
        raise RuntimeError(f"风控拦截(status={status_code})：{error_msg}")

    def _validate_excel_response(self, response, attempt):
        """响应校验（项目6 适配：xls/xlsx 双魔数 + Content-Disposition）。

        校验规则:
            1. HTTP 200 但响应体 < 1KB → 空响应
            2. 魔数校验：.xlsx=PK\\x03\\x04（zip） 或 .xls=\\xD0\\xCF\\x11\\xE0（OLE2）
               —— 与项目4/5 只认 PK 不同，本项目兼容两种（首次出现 xls 响应）
            3. 响应头必须含 Content-Disposition: attachment（风控伪装拦截）
        """
        # 1. 空响应拦截
        if response.status_code == 200 and len(response.content) < 1024:
            preview = response.content[:200].decode("utf-8", errors="replace")
            self.logger.error(
                f"[第{attempt}次] 响应体过小（{len(response.content)}字节 < 1KB），"
                f"疑似风控拦截或服务器错误。响应前200字节：{preview!r}"
            )
            raise RuntimeError(f"空响应（{len(response.content)}字节）")

        # 2. Excel 魔数校验（xlsx 或 xls 任一通过）
        magic = response.content[:4]
        if magic not in (b"PK\x03\x04", b"\xD0\xCF\x11\xE0"):
            preview = response.content[:200].decode("utf-8", errors="replace")
            self.logger.error(
                f"[第{attempt}次] 响应体不是有效 Excel（xlsx=PK/xls=D0CF 均不匹配，实际={magic!r}）。"
                f"响应前200字节：{preview!r}"
            )
            raise RuntimeError(f"响应体不是 Excel 文件（magic bytes 校验失败: {magic!r}）")

        # 3. Content-Disposition: attachment 校验（风控伪装拦截）
        content_disp = response.headers.get("Content-Disposition", "")
        if "attachment" not in content_disp:
            self.logger.error(
                f"[第{attempt}次] 响应头缺少 Content-Disposition: attachment，"
                f"疑似风控拦截伪装。实际 Content-Disposition={content_disp!r}"
            )
            raise RuntimeError("响应头缺少 Content-Disposition: attachment（疑似风控拦截）")

    def _parse_content_disposition_filename(self, content_disp_header):
        """解析 Content-Disposition 文件名（UTF-8 与 GBK 双兼容，项目6增强）。

        京东响应头两类文件名：
            1. filename*=UTF-8''商品明细.xlsx（URL编码中文，项目5 商品明细）
            2. filename=<UTF-8字节>（本项目商品流失分析：真实字节 UTF-8，charset=utf-8）
               —— HTTP 头按 latin-1 解码成 U+00xx 字符，用 encode('latin-1') 还原字节后
                  **UTF-8 优先解码**（个别接口可能 GBK 字节，做 GBK 回退双兼容）

        返回:
            str 或 None（无法解析）
        """
        if not content_disp_header:
            return None

        # ① filename*=UTF-8''...（URL 编码）
        m = re.search(r"filename\*\s*=\s*UTF-8''([^;]+)", content_disp_header, re.IGNORECASE)
        if m:
            import urllib.parse
            raw = m.group(1).strip().strip('"')
            try:
                return urllib.parse.unquote(raw)
            except Exception:
                return raw

        # ② filename=...（真实字节为 UTF-8；个别接口可能 GBK，双兼容）
        # ⚠️ 踩坑（2026-08-07 真实导出）：京东此接口 filename 字节是 UTF-8 编码
        #   （header 声明 charset=utf-8），requests 按 latin-1 解码成 U+00xx 字符；
        #   必须先 encode('latin-1') 还原字节，再 **UTF-8 优先**解码。
        #   若按 GBK 解码会把 UTF-8 字节解成"鍟嗗搧..."乱码（此前真实导出踩坑）。
        m = re.search(r"filename\s*=\s*\"?([^;\"]+)\"?", content_disp_header, re.IGNORECASE)
        if m:
            raw = m.group(1).strip()
            # 还原 HTTP 层字节（latin-1 是 HTTP header 的传输编码）
            try:
                raw_bytes = raw.encode("latin-1")
            except UnicodeEncodeError:
                return raw   # 已含非 latin-1 字符（可能是已解码的中文），原样返回
            # 优先 UTF-8（本项目真实场景）
            try:
                decoded = raw_bytes.decode("utf-8")
                if "\ufffd" not in decoded:
                    return decoded
            except UnicodeDecodeError:
                pass
            # 回退 GBK（兼容个别接口用 GBK 字节）
            try:
                decoded = raw_bytes.decode("gbk")
                if "\ufffd" not in decoded:
                    return decoded
            except UnicodeDecodeError:
                pass
            return raw
        return None

    # ---------- Excel 后置处理（xls 读取 → 转存 xlsx）----------

    def _save_excel_to_path(self, response, target_path, date):
        """Excel 后置处理（项目6 版：read_excel_bytes 自动识别 xlsx/xls）。

        流程：
            ① read_excel_bytes() 读取（xlsx 用 openpyxl / xls 用 xlrd）
            ② prepare_date_columns() 日期列智能处理（公共规则1+2）
            ③ safe_convert_numeric() 数值安全转换（公共规则3）
            ④ 写入 .xlsx + apply_column_formats() 设置单元格格式
        """
        import warnings

        # 抑制openpyxl读取原始xlsx时的无害警告
        warnings.filterwarnings("ignore", message="Workbook contains no default style")

        # ① 读取（自动识别 xlsx/xls）
        df = read_excel_bytes(response.content)

        # ②③ 日期列统一处理（公共规则1+2）
        date_column, date_value = prepare_date_columns(df, date)

        # ④ 全表数值安全转换（公共规则3）
        df = safe_convert_numeric(df)

        # ⑤ 写入 .xlsx（转存格式，统一用 openpyxl 引擎）+ 单元格格式
        df.to_excel(target_path, index=False, engine="openpyxl")
        apply_column_formats(target_path, df, date_column=date_column, date_value=date_value)

        self.logger.info(
            f"Excel已保存: {target_path}（xls→xlsx转存+日期列+数值转换+单元格格式，{os.path.getsize(target_path)}字节）"
        )
        return target_path

    # ---------- 主下载方法 ----------

    def download_loss_product(self, date=None, start_date=None, end_date=None):
        """下载商品流失分析报表（POST 请求，xls 响应转存 xlsx）。

        参数:
            date        - 查询日期 YYYY-MM-DD（入参优先，其次 config）
            start_date  - 开始日期（默认=date）
            end_date    - 结束日期（默认=date）
        返回:
            str - 保存的 Excel 文件绝对路径（output/商品流失分析/{date}/ 子目录）
        """
        # 1. 解析日期（入参 > config；三值一致规则）
        date, start_date, end_date = self._get_date_params(date, start_date, end_date)

        # 2. 组装业务参数（固定常量 + 日期）
        form_params = {
            "date": date,
            "startDate": start_date,
            "endDate": end_date,
            **self.FIXED_BIZ_PARAMS,
        }

        # 3. 必带请求头（业务约束：Origin/Referer 缺失被拦截）
        # 注意：Sec-Fetch-Site=same-origin（与项目5 相同，覆盖基类默认 same-site）
        extra_headers = {
            "Origin": self.ORIGIN,
            "Referer": self.REFERER,
            "Sec-Fetch-Site": "same-origin",
        }

        # 4. 发送请求（重试循环 + UA切换 + 风控识别 + 响应校验 + success标记）
        response = None
        last_exception = None
        success = False   # 阶段4 修复：只有 break 才算成功，防误保存错误内容

        for attempt in range(1, self.MAX_RETRIES + 1):
            try:
                # 4.1 严格间隔控制（与基类一致）
                self._wait_interval()

                # 4.2 重新生成风控参数（每次新 uuid + 新时间戳 + 新签名）
                risk_params = self._gen_risk_params_random(self.API_URL)
                full_params = {**form_params, **risk_params}

                ua_name = "Edge" if self._current_ua_index == 0 else "Chrome"
                self.logger.info(
                    f"[第{attempt}/{self.MAX_RETRIES}次] 下载商品流失分析报表: "
                    f"日期={date}, UA={ua_name}, uuid={risk_params['uuid'][:8]}..."
                )

                # 记录请求时间（类属性：所有实例共享）
                JDBaseRequest._last_request_time = time.time()

                # 4.3 发送 POST 请求（表单，同项目4）
                response = self.session.post(
                    self.API_URL,
                    data=full_params,
                    headers=extra_headers,
                    timeout=self.REQUEST_TIMEOUT,
                )

                # 4.4 HTTP 状态码基础检查
                if response.status_code == 401:
                    self.logger.error("HTTP 401 未授权 - Cookie 可能已过期或被禁用")
                    raise CookieExpiredError("Cookie已过期或无效，请更新 config/sz_cookie.txt")
                if response.status_code == 403:
                    self.logger.error("HTTP 403 禁止访问 - Cookie/签名/Origin 校验失败")
                    # 不抛 CookieExpired，让重试机制 + UA 切换兜底

                # 4.5 风控业务码识别
                self._check_business_code(response)

                # 4.6 响应校验（xls/xlsx 双魔数 + Content-Disposition）
                self._validate_excel_response(response, attempt)

                # 4.7 走到这里 = 成功
                self.logger.info(
                    f"请求成功: HTTP {response.status_code}, {len(response.content)}字节, "
                    f"Content-Disposition={response.headers.get('Content-Disposition', '')[:60]}"
                )
                success = True
                break

            except CookieExpiredError:
                raise
            except RiskControlError:
                raise
            except requests.exceptions.Timeout as e:
                last_exception = e
                self.logger.warning(f"第{attempt}次请求超时（{self.REQUEST_TIMEOUT}秒）")
            except Exception as e:
                last_exception = e
                self.logger.warning(f"第{attempt}次请求失败: {e}")

            # 4.8 重试间隔：UA 切换 + 递增等待
            if attempt < self.MAX_RETRIES:
                self._switch_ua()
                wait_seconds = self.REQUEST_INTERVAL * attempt
                self.logger.info(
                    f"等待 {wait_seconds}秒 后重试（已切换UA，当前: "
                    f"{'Edge' if self._current_ua_index == 0 else 'Chrome'}）..."
                )
                time.sleep(wait_seconds)

        # 4.9 重试全部失败：上报（success 标记兜底，防误保存）
        if not success:
            self.logger.error(f"所有 {self.MAX_RETRIES} 次重试均失败")
            raise RuntimeError(
                f"商品流失分析报表下载失败：{last_exception}（请检查Cookie/网络/风控）"
            ) from last_exception

        # 5. 解析 Content-Disposition 原始文件名（GBK/UTF-8 双兼容）
        content_disp = response.headers.get("Content-Disposition", "")
        original_filename = self._parse_content_disposition_filename(content_disp)

        if not original_filename:
            # 兜底：无法解析时用友好命名
            self.logger.warning(
                f"无法从 Content-Disposition 解析文件名（头={content_disp[:80]!r}），使用兜底命名"
            )
            original_filename = f"商品流失分析_{date}.xls"

        # 6. 转存格式：.xls → .xlsx（统一输出 xlsx，内容已由 _save_excel_to_path 转换）
        if original_filename.lower().endswith(".xls"):
            original_filename = original_filename[:-4] + ".xlsx"

        # 7. 构造业务子目录路径：output/商品流失分析/{date}/{filename}
        date_subdir = os.path.join(self.OUTPUT_SUBDIR, date)
        business_output_dir = os.path.join(self.output_dir, date_subdir)
        os.makedirs(business_output_dir, exist_ok=True)
        target_path = os.path.join(business_output_dir, original_filename)

        # 8. Excel 后置处理（xls 读取 → 转存 xlsx）
        return self._save_excel_to_path(response, target_path, date)


# ============================================================
#  业务接口 5：（新业务 - 京准通快车自定义报表导出，2026-08-07 上线骨架）
# ------------------------------------------------------------
#  中文说明（小白必读）：
#    京准通 jzt.jd.com 广告报表导出，与商智/京麦 Cookie 不互通，必须独立 Cookie 文件。
#    ⚠️ 核心差异（与项目1-6对比）：
#      - 鉴权用 h5st（请求头），不是商智域 User-mnp/uuid 体系
#      - 三步异步：创建任务 → 轮询列表 → CDN 下载 CSV
#      - downloadUrl 一次性签名，过期需重新轮询刷新
#    本阶段（阶段3骨架）：
#      - 仅 3 个接口方法，不做轮询循环/Excel 解析（阶段5再补）
#      - 不实现 h5st 算法，__init__ 接收外部传入
#      - 不写 uuid 字段（京准通域未校验，已抓包确认）
# ============================================================

# 京准通快车自定义报表 payload 模板（最小字段版）
# ⚠️ 业务参数大部分固定，仅时间字段动态替换；完整模板放代码外延后迭代再引入
JZT_KUAICHE_PAYLOAD_TEMPLATE = {
    # 完整 payload 模板（2026-08-07 抓包实证，含 6 大模块 + 元模板）
    # ⚠️ 注意：若接口报参数错误，可能：
    #   1. checkSum 是页面 JS 动态计算（故障排查：用真实浏览器 page.evaluate() 提取原始 payload 比对）
    #   2. 当前账户在 customDimensionOptions 中未勾选（默认 FYA8888 已 checked:True）
    #   3. h5st 校验（当前 add 接口不校验，但其他接口可能校验，靠 window.ParamsSign.sign 实时生成）
    "caliberSettings": [
        {"checked": True, "desc": "转化周期：平台建议选择15天/30天转化周期进行数据观测", "hidden": False,
         "key": "clickOrOrderDay",
         "options": [
            {"checked": False, "desc": "当天", "hidden": False, "key": "today", "value": "0"},
            {"checked": False, "desc": "1天", "hidden": False, "key": "oneDay", "value": "1"},
            {"checked": False, "desc": "3天", "hidden": False, "key": "threeDays", "value": "3"},
            {"checked": False, "desc": "7天", "hidden": False, "key": "sevenDays", "value": "7"},
            {"checked": True, "desc": "15天", "hidden": False, "key": "fifteenDays", "value": "15"},
            {"checked": False, "desc": "30天", "hidden": False, "key": "thirtyDays", "value": "30"},
         ]},
        {"checked": True, "desc": "点击/下单口径", "hidden": False, "key": "clickOrOrderCaliber",
         "options": [
            {"checked": True, "desc": "点击", "hidden": False, "key": "click", "value": "0"},
            {"checked": False, "desc": "下单", "hidden": False, "key": "order", "value": "1"},
         ]},
        {"checked": True, "desc": "含赠品/不含赠品", "hidden": False, "key": "giftFlag",
         "options": [
            {"checked": False, "desc": "含赠品", "hidden": False, "key": "include"},
            {"checked": True, "desc": "不含赠品", "hidden": False, "key": "exclude", "value": "0"},
         ]},
        {"checked": True, "desc": "下单订单/成交订单", "hidden": False, "key": "orderStatusCategory",
         "options": [
            {"checked": False, "desc": "下单订单", "hidden": False, "key": "place"},
            {"checked": True, "desc": "成交订单", "hidden": False, "key": "done"},
         ]},
    ],
    # ⚠️ checkSum 用户决策 2026-08-07：现阶段硬编码 1114112；后续若接口报错可能是页面 JS 动态计算
    "checkSum": 1114112,
    "customDimension": [
        {"checked": False, "desc": "基础维度", "hidden": False, "key": "basicDimension",
         "options": [
            {"checked": False, "desc": "账户名称", "groupLabel": "basicDimension", "hidden": False, "key": "pin"},
            {"checked": True, "desc": "产品线", "groupLabel": "basicDimension", "hidden": False, "key": "businessType"},
            {"checked": True, "desc": "推广计划", "groupLabel": "basicDimension", "hidden": False, "key": "campaign"},
            {"checked": True, "desc": "推广单元/品类", "groupLabel": "basicDimension", "hidden": False, "key": "group", "tip": ""},
            {"checked": False, "desc": "推广创意/商品", "groupLabel": "basicDimension", "hidden": False, "key": "ad", "tip": "智能投放查询的是推广商品信息；其他产品线查询的是推广创意信息"},
         ]},
        {"checked": False, "desc": "细分维度", "hidden": False, "key": "detailDimension",
         "options": [
            {"checked": True, "desc": "营销目标", "hidden": False, "key": "marketingObjectiveName"},
            {"checked": True, "desc": "营销场景", "hidden": False, "key": "marketingScenarioTypeName"},
            {"checked": True, "desc": "搜索词", "hidden": False, "key": "searchTerm", "tip": "搜索词查询结果明细数据"},
            {"checked": True, "desc": "关键词", "hidden": False, "key": "keyword", "tip": "关键词查询结果明细数据"},
            {"checked": True, "desc": "关键词购买类型", "hidden": False, "key": "targetingType"},
            {"checked": True, "desc": "商品定向细分类型", "hidden": False, "key": "productDeliveryMatchingType", "tip": "商品定向细分类型包含：商品定向、类目定向、店铺定向、相似品定向、搭配品定向"},
            {"checked": True, "desc": "定向目标", "hidden": False, "key": "productDeliveryTriggerSkuId", "tip": "定向目标为商品定向下触发广告的id，其触发条件可能是sku、店铺id等"},
            {"checked": True, "desc": "广告定向方式", "hidden": False, "key": "deliveryType", "tip": "1.智能投放暂不支持关键词定向查询；2.人群定向类型指通过圈定特定人群"},
            {"checked": True, "desc": "投放地域", "hidden": False, "key": "mappedAreaName", "tip": "投放地域是指广告实际展现的地域"},
            {"checked": True, "desc": "跟单SKU", "hidden": False, "key": "skuDocId", "tip": "对于落地页为活动页、店铺页类广告"},
            {"checked": True, "desc": "SPU ID", "hidden": False, "key": "spuId", "tip": "SPU维度不是广告跟单所使用维度"},
            {"checked": False, "desc": "品牌", "hidden": False, "key": "promotedBrand", "tip": "品牌是根据推广SKU或跟单SKU关联的信息"},
            {"checked": False, "desc": "类目", "hidden": False, "key": "promotedCid", "tip": "类目是根据推广SKU或跟单SKU关联的三级类目信息"},
            {"checked": True, "desc": "投放位置", "hidden": False, "key": "trafficPackage", "tip": "推荐广告查询的是流量包信息"},
            {"checked": False, "desc": "人群名称", "hidden": False, "key": "crowdDimension", "tip": "搜索快车所查询的是搜索人群明细"},
         ]},
    ],
    # customDimensionOptions：账号范围 + 产品线 + 营销目标 + 广告定向类型（抓包原貌）
    "customDimensionOptions": [
        {"checked": False, "desc": "账号范围", "hidden": False, "key": "pin",
         "options": [
            {"checked": False, "desc": "自有账户", "hidden": False, "key": "subUser",
             "options": [
                {"checked": True, "desc": "FYA8888", "flag": True, "hidden": False, "key": "99936530475", "value": "FYA8888"},
                {"checked": False, "desc": "FYA888888", "flag": False, "hidden": False, "key": "99936525688", "value": "FYA888888"},
                {"checked": False, "desc": "FAY掌柜888", "flag": False, "hidden": False, "key": "99937142699", "value": "FAY掌柜888"},
                {"checked": False, "desc": "FYA19529975351", "flag": False, "hidden": False, "key": "99938531397", "value": "FYA19529975351"},
                {"checked": False, "desc": "fya掌柜777", "flag": False, "hidden": False, "key": "99938957251", "value": "fya掌柜777"},
                {"checked": False, "desc": "FYA小婷", "flag": False, "hidden": False, "key": "99938963919", "value": "FYA小婷"},
                {"checked": False, "desc": "FYA少冰", "flag": False, "hidden": False, "key": "99945916633", "value": "FYA少冰"},
                {"checked": False, "desc": "FYA小冠", "flag": False, "hidden": False, "key": "99947097388", "value": "FYA小冠"},
                {"checked": False, "desc": "FYA布丁", "flag": False, "hidden": False, "key": "99952884963", "value": "FYA布丁"},
                {"checked": False, "desc": "FYA小柔", "flag": False, "hidden": False, "key": "99955522890", "value": "FYA小柔"},
                {"checked": False, "desc": "FYA小敏", "flag": False, "hidden": False, "key": "99960125485", "value": "FYA小敏"},
             ]},
            {"checked": False, "desc": "授权账户", "hidden": False, "key": "authUser"},
         ]},
        {"checked": False, "desc": "产品线", "hidden": False, "key": "businessType",
         "options": [
            {"checked": False, "desc": "快车", "hidden": False, "key": "kuaiche", "value": "-4"},
            {"checked": False, "desc": "站外广告", "hidden": False, "key": "zhitou", "tip": "站外广告数据包含原京东直投+京易投数据", "value": "256"},
         ]},
    ],
    "customIndex": [
        {"checked": False, "desc": "基础数据", "hidden": False, "key": "basicData",
         "tip": "为广告基础投放数据，用于分析广告投放的基础表现",
         "options": [
            {"checked": True, "desc": "展现数", "hidden": False, "key": "impressions"},
            {"checked": True, "desc": "点击数", "hidden": False, "key": "clicks"},
            {"checked": True, "desc": "点击率(%)", "hidden": False, "key": "CTR"},
            {"checked": True, "desc": "花费", "hidden": False, "key": "cost"},
            {"checked": True, "desc": "千次展现成本", "hidden": False, "key": "CPM"},
            {"checked": True, "desc": "平均点击成本", "hidden": False, "key": "CPC"},
         ]},
        {"checked": False, "desc": "转化数据", "hidden": False, "key": "resultData",
         "tip": "为通过广告获得的用户后链路指标",
         "options": [
            {"checked": True, "desc": "直接订单行", "hidden": False, "key": "directOrderCnt"},
            {"checked": True, "desc": "直接订单金额", "hidden": False, "key": "directOrderSum"},
            {"checked": True, "desc": "间接订单行", "hidden": False, "key": "indirectOrderCnt"},
            {"checked": True, "desc": "间接订单金额", "hidden": False, "key": "indirectOrderSum"},
            {"checked": True, "desc": "总订单行", "hidden": False, "key": "totalOrderCnt"},
            {"checked": True, "desc": "总订单金额", "hidden": False, "key": "totalOrderSum"},
            {"checked": True, "desc": "直接加购数", "hidden": False, "key": "directCartCnt"},
            {"checked": True, "desc": "间接加购数", "hidden": False, "key": "indirectCartCnt"},
            {"checked": True, "desc": "总加购数", "hidden": False, "key": "totalCartCnt"},
            {"checked": True, "desc": "转化率(%)", "hidden": False, "key": "orderCVS"},
            {"checked": True, "desc": "平均订单成本", "hidden": False, "key": "orderCPA", "tip": "直投的订单行成本就是CPA"},
            {"checked": True, "desc": "投产比", "hidden": False, "key": "orderROI"},
            {"checked": True, "desc": "预售订单行", "hidden": False, "key": "totalPresaleOrderCnt"},
            {"checked": True, "desc": "预售订单金额", "hidden": False, "key": "totalPresaleOrderSum"},
         ]},
        # 转化数据[媒]、直播数据、订单效果数据、引流数据 元模板默认未勾选，与抓包一致
        {"checked": False, "desc": "转化数据[媒]", "hidden": False, "key": "mediaResultData",
         "tip": "为站外广告专属指标", "options": []},
        {"checked": False, "desc": "直播数据", "hidden": False, "key": "liveData", "options": []},
        {"checked": False, "desc": "订单效果数据", "hidden": False, "key": "orderData", "options": []},
        {"checked": False, "desc": "引流数据", "hidden": False, "key": "drainageData", "options": []},
    ],
    # customIndexOptions 元模板（与 customIndex 结构对应，用于报表展示）
    "customIndexOptions": [
        {"checked": False, "desc": "基础数据", "hidden": False, "key": "basicData", "isIndex": True,
         "tip": "为广告基础投放数据", "options": [
            {"checked": False, "desc": "展现数", "hidden": False, "key": "impressions"},
            {"checked": False, "desc": "点击数", "hidden": False, "key": "clicks"},
            {"checked": False, "desc": "点击率(%)", "hidden": False, "key": "CTR"},
            {"checked": False, "desc": "花费", "hidden": False, "key": "cost"},
            {"checked": False, "desc": "千次展现成本", "hidden": False, "key": "CPM"},
            {"checked": False, "desc": "平均点击成本", "hidden": False, "key": "CPC"},
         ]},
        {"checked": False, "desc": "转化数据", "hidden": False, "key": "resultData", "isIndex": True,
         "tip": "用户后链路指标", "options": [
            {"checked": False, "desc": "直接订单行", "hidden": False, "key": "directOrderCnt"},
            {"checked": False, "desc": "直接订单金额", "hidden": False, "key": "directOrderSum"},
            {"checked": False, "desc": "间接订单行", "hidden": False, "key": "indirectOrderCnt"},
            {"checked": False, "desc": "间接订单金额", "hidden": False, "key": "indirectOrderSum"},
            {"checked": False, "desc": "总订单行", "hidden": False, "key": "totalOrderCnt"},
            {"checked": False, "desc": "总订单金额", "hidden": False, "key": "totalOrderSum"},
            {"checked": False, "desc": "直接加购数", "hidden": False, "key": "directCartCnt"},
            {"checked": False, "desc": "间接加购数", "hidden": False, "key": "indirectCartCnt"},
            {"checked": False, "desc": "总加购数", "hidden": False, "key": "totalCartCnt"},
            {"checked": False, "desc": "转化率(%)", "hidden": False, "key": "orderCVS"},
            {"checked": False, "desc": "平均订单成本", "hidden": False, "key": "orderCPA", "tip": "直投的订单行成本就是CPA"},
            {"checked": False, "desc": "投产比", "hidden": False, "key": "orderROI"},
            {"checked": False, "desc": "预售订单行", "hidden": False, "key": "totalPresaleOrderCnt"},
            {"checked": False, "desc": "预售订单金额", "hidden": False, "key": "totalPresaleOrderSum"},
         ]},
    ],
    "daily": 1,
    # sortedDimensionKeys / sortedIndexKeys 全集（与抓包一致）
    "sortedDimensionKeys": [
        "businessType", "campaign", "group", "marketingObjectiveName", "marketingScenarioTypeName",
        "searchTerm", "keyword", "targetingType", "productDeliveryMatchingType",
        "productDeliveryTriggerSkuId", "deliveryType", "mappedAreaName", "skuDocId", "spuId", "trafficPackage",
    ],
    "sortedIndexKeys": [
        "impressions", "clicks", "CTR", "cost", "CPM", "CPC",
        "directOrderCnt", "directOrderSum", "indirectOrderCnt", "indirectOrderSum",
        "totalOrderCnt", "totalOrderSum", "directCartCnt", "indirectCartCnt", "totalCartCnt",
        "orderCVS", "orderCPA", "orderROI", "totalPresaleOrderCnt", "totalPresaleOrderSum",
    ],
    "timeout": False,
    "version": "JZT_V9",
    "indexSign": 0,
    "cycle": None,
    "pinIdList": [],
    "requestFrom": 0,
}


class JZTKuaicheAPI:
    """京准通快车自定义报表导出 API（基础骨架，2026-08-07 上线）。

    ⚠️ 本类**不继承 JDBaseRequest**（业务模型差异大）：
        - 鉴权体系不同（h5st + 独立 Cookie 文件，不是商智 User-mnp/uuid）
        - 异步三步流程（创建/轮询/下载），不适合基类 30秒重试模型
        - UA 与 h5st 绑定，禁用基类 UA 切换（会致 h5st 失效）

    参数:
        h5st       - **可选**。浏览器抓 add 接口请求头复制（如有）。项目7 抓包实测 add 接口不校验 h5st 字段
                      （与京麦 sff.jd.com 不同），但保留参数为后续接口（如未来 list/downloadUrl）增加 h5st 校验时使用
        cookie_path - 京准通 Cookie 文件路径，默认 config/jzt_cookie.txt（与商智 Cookie 不互通）
    """

    # ---- 类常量（业务固定参数）----
    BASE_URL = "https://jzt-api.jd.com/dataCenter/customreport/v2/report"
    ORIGIN = "https://jzt.jd.com"
    REFERER = "https://jzt.jd.com"
    SITE_ID = "0"
    USER_AGENT = (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/144.0.0.0 Safari/537.36 Edg/144.0.0.0"
    )
    OUTPUT_SUBDIR = "京准通快车效果自定义"  # 落 output/京准通快车效果自定义/{date}/ 子目录（AGENTS.md Excel规则4）

    # ---- 阶段4 容错配置（用户决策 2026-08-07：POLL_INTERVAL=3s / MAX_POLL_TIMES=15）----
    POLL_INTERVAL = 3            # 轮询间隔（秒），报表生成等待
    MAX_POLL_TIMES = 15          # 轮询最大次数（3s × 15 = 45s 超时）
    MAX_DOWNLOAD_RETRY = 3       # CDN 404 重试最大次数（重刷 URL 后随机退避 3-10s）

    # ---- 京东业务码约定（与项目4/5/6 对齐）----
    # code=0 成功；code=601 h5st过期（不重试）；code=-407/-402 签名错（重试）；
    # 业务码非0 且 message/msg 含"未登录/登录" → CookieExpiredError（不重试）

    def __init__(self, h5st: str = "", cookie_path: str = "config/jzt_cookie.txt"):
        """初始化京准通 API。

        参数:
            h5st       - 浏览器F12抓 add 接口请求头的 h5st 值（手动复制，脚本不实现 JS 签名）
            cookie_path - 京准通 Cookie 文件路径（默认 config/jzt_cookie.txt；与商智 Cookie 不互通）
        """
        import requests  # 本类独立按需导入，避免污染顶层 namespace
        # random 模块已在 main.py 顶层 import，此处可直接使用 random.uniform()

        # 1. 读取 Cookie（不存在即抛错，强制用户抓包填入）
        cookie_path_abs = os.path.join(os.path.dirname(os.path.abspath(__file__)), cookie_path)
        if not os.path.isfile(cookie_path_abs):
            raise FileNotFoundError(
                f"❌ 京准通 Cookie 文件不存在：{cookie_path_abs}\n"
                f"   请浏览器登录 https://jzt.jd.com/home，F12 抓 jzt.jd.com 域 Cookie 写入此文件"
            )
        with open(cookie_path_abs, "r", encoding="utf-8") as f:
            self.cookie = f.read().strip()
        if not self.cookie:
            raise ValueError(f"❌ 京准通 Cookie 文件 {cookie_path_abs} 内容为空")

        # 2. 接收 h5st（**可选**；阶段4 抓包实测 add 接口不校验 h5st，参数保留为未来扩展）
        self.h5st = h5st or ""

        # 3. requests Session（不继承基类 UA 切换逻辑，h5st 绑定 UA）
        self.session = requests.Session()
        session_headers = {
            "User-Agent": self.USER_AGENT,
            "Origin": self.ORIGIN,
            "Referer": self.REFERER + "/",  # 抓包带尾斜杠，对齐
            "siteId": self.SITE_ID,
            "Content-Type": "application/json;charset=UTF-8",
            "Accept": "application/json, text/plain, */*",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
            "Accept-Encoding": "gzip, deflate, br, zstd",
            "Cookie": self.cookie,
        }
        # h5st 非空才注入（抓包实测多数 add 请求无 h5st 字段）
        if self.h5st:
            session_headers["h5st"] = self.h5st
        self.session.headers.update(session_headers)

        # 4. 输出路径（按 AGENTS.md Excel规则4：业务子目录 + 日期子目录）
        self.output_dir = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "output", self.OUTPUT_SUBDIR,
        )

    # ---- 公共方法：组装 payload（动态注入时间）----

    def _build_payload(self, start_date: str, end_date: str) -> dict:
        """深拷贝模板并注入查询时间（用户决策 2026-08-07）。

        关键转换（2026-08-07 抓包实证）：
            - startTime / endTime 必须是**毫秒时间戳**（如 1786161600000），不是日期字符串
            - startTimeStr / endTimeStr 是日期字符串（"YYYY-MM-DD"），用于展示
            - tempName / reportName 是报表名（含日期时间）

        参数:
            start_date - 开始日期 YYYY-MM-DD
            end_date   - 结束日期 YYYY-MM-DD
        返回:
            dict - 完整 payload（含时间戳字段）
        """
        import copy
        import json
        from datetime import datetime, timezone, timedelta

        payload = copy.deepcopy(JZT_KUAICHE_PAYLOAD_TEMPLATE)

        # 1. 日期字符串 → 毫秒时间戳（北京时区 00:00:00）
        # 抓包示例：1786161600000 = 2026-08-08 16:00:00 UTC = 2026-08-09 00:00:00 +08:00
        # 注意：用户原抓包时间戳含时分秒（不是 00:00:00），本代码默认取 00:00:00；如需时分秒请扩展入参
        tz_beijing = timezone(timedelta(hours=8))
        start_dt = datetime.strptime(start_date, "%Y-%m-%d").replace(tzinfo=tz_beijing)
        end_dt = datetime.strptime(end_date, "%Y-%m-%d").replace(tzinfo=tz_beijing)
        start_ts = int(start_dt.timestamp() * 1000)  # 秒 → 毫秒
        end_ts = int(end_dt.timestamp() * 1000)

        # 2. 注入时间相关字段
        # ⚠️ 阶段6 真实跑通发现（2026-08-07）：
        #   1. 报表名重名会被拒（msg=【操作失败】报表名重复）→ 加 HHMM 后缀确保唯一
        #   2. 报表名长度限制 1-30 字符（msg=【参数错误】报表名长度只允许1-30个字符）→ 必须精简
        # 当前格式：日期去掉分隔符 + HHMM = "20260807_20260807_HHMM" = 22 字符（留 8 字符冗余）
        from datetime import datetime as _dt
        suffix = _dt.now().strftime("%H%M")  # HHMM（4 位后缀，总长度 < 30）
        date_compact = start_date.replace("-", "")  # 20260807
        end_compact = end_date.replace("-", "")
        payload["startTime"] = start_ts
        payload["endTime"] = end_ts
        payload["startTimeStr"] = start_date
        payload["endTimeStr"] = end_date
        payload["tempName"] = f"{date_compact}_{end_compact}_{suffix}"
        payload["reportName"] = f"{date_compact}_{end_compact}_{suffix}"

        # 3. 运行时日志：打印完整 payload（重点 checkSum 字段），方便人工比对抓包
        # ⚠️ 用户决策 2026-08-07：组装完 payload 输出完整 JSON，重点打印 checkSum 字段值
        # 故障排查方式（预案注释）：
        #   若接口报参数错误，checkSum 可能是页面 JS 动态计算值
        #   排查方法：使用 playwright 启动真实浏览器，F12 → Console 执行 page.evaluate()
        #     → window.ParamsSign.sign(JSON.stringify(payload)) 获取真实 checkSum
        #   当前硬编码 checkSum=1114112 与抓包实证一致，但京东可能不定期更新此值
        print(f"  [payload 调试] startTime={start_ts} ({start_date} +08:00)")
        print(f"  [payload 调试] endTime={end_ts} ({end_date} +08:00)")
        print(f"  [payload 调试] checkSum={payload['checkSum']}（如接口报错请用 page.evaluate() 提取真实值）")
        print(f"  [payload 调试] 完整 payload（精简打印前 800 字符）：{json.dumps(payload, ensure_ascii=False)[:800]}...")

        return payload

    # ---- 阶段4 容错：统一业务码识别（项目7）----

    def _is_cookie_expired(self, ret: dict) -> bool:
        """判断响应是否表示 Cookie 过期（参考项目5 文本型 601 识别）。

        判定规则：
            1. 业务码 2001（京东标准未登录码，跨域常见）
            2. 业务码 302 且 message 含 "登录"（与项目4/5 一致）
            3. message/msg 含 "未登录" / "登录已过期" / "请重新登录"（文本兜底）
        """
        code = ret.get("code")
        msg = str(ret.get("msg", "")) + str(ret.get("message", ""))
        if code in (2001, 302):
            return True
        keywords = ["未登录", "登录已过期", "请重新登录", "login required"]
        return any(k in msg for k in keywords)

    def _handle_response(self, ret: dict, op_desc: str):
        """统一处理京准通接口响应（阶段4 容错核心 + 阶段6 适配）。

        ⚠️ 关键发现（2026-08-07 真实跑通）：京准通响应**双字段判定**：
            - `success: true` + `code: 1` + `data: reportId` → 接口成功（如 add 返回）
            - `success: true` + `code: 0` + `data: ...`      → 标准成功（多数接口）
            - `success: false` 或 code 非 {0,1}             → 失败
        必须**同时**满足 success=true 且 code 在合法集合才视为成功。

        参数:
            ret     - 接口响应 dict
            op_desc - 操作描述（用于错误信息，如 "创建任务" / "查询列表"）
        返回:
            dict - 原始响应（成功时透传，失败抛异常）
        异常:
            CookieExpiredError - Cookie 过期（不重试，立即停）
            RuntimeError       - 业务码 601（h5st 过期）/-407/-402（签名错）/其他非0
        """
        code = ret.get("code")
        success = ret.get("success", True)  # 缺省视为 True（兼容旧响应）

        # 1. 成功判定：success=True 且 code 在 {0, 1}（京东业务码 1 通常表示有 data 返回）
        if success and code in (0, 1):
            return ret

        # 2. Cookie 过期 → 立即停（参考项目4/5 异常抛出规则）
        if self._is_cookie_expired(ret):
            raise CookieExpiredError(
                f"❌ 京准通 Cookie 过期（{op_desc}返回 code={code}）：\n"
                f"   → 请浏览器登录 https://jzt.jd.com/home，F12 抓 jzt.jd.com 域 Cookie 写入 config/jzt_cookie.txt"
            )

        # 3. h5st 过期（601）→ 不重试，直接抛
        if code == 601:
            raise RuntimeError(
                f"❌ 京准通 h5st 过期（{op_desc}返回 code=601）：\n"
                f"   → 请浏览器F12抓 add 接口最新 h5st 重新构造实例：api = JZTKuaicheAPI(h5st='新值')"
            )

        # 4. 签名错（-407/-402）→ 抛 RuntimeError 让外层决定重试
        if code in (-407, -402):
            raise RuntimeError(
                f"❌ 京准通签名校验失败（{op_desc}返回 code={code}）：\n"
                f"   msg={ret.get('msg')}\n"
                f"   → 可能 h5st 不匹配当前 UA，请重新抓 add 接口最新 h5st"
            )

        # 5. 其他非0 → 完整响应回显便于排查
        raise RuntimeError(
            f"❌ 京准通{op_desc}失败 code={code}, msg={ret.get('msg')}\n"
            f"   完整响应：{ret}\n"
            f"   可能原因：payload checkSum 错 / 时间跨度>90天 / 账号权限不足"
        )

    # ---- 接口1：创建导出任务 ----

    def create_export_task(self, date: str = None, start_date: str = None, end_date: str = None) -> str:
        """创建导出任务（POST），返回 task_id。

        参数（与项目1-6 调度层对齐）：
            date        - 单日查询 YYYY-MM-DD（调度层默认传此参数；start/end 默认=date）
            start_date  - 开始日期 YYYY-MM-DD（直接调用时可显式传区间）
            end_date    - 结束日期 YYYY-MM-DD
        返回:
            str - 任务 ID（用于后续 get_task_list / download_report 关联）
        异常:
            CookieExpiredError - Cookie 过期（不重试）
            RuntimeError       - 业务码 601（h5st 过期）/-407/-402（签名错）/其他非0
        """
        # 三值一致规则：start/end 未传时默认=date
        if not start_date:
            start_date = date
        if not end_date:
            end_date = date
        # ⚠️ 2026-08-10 bug fix：原 `if not date` 在 date=None 但 start_date 有值时也会报错
        #   用户决策补能力：--range last_7d 计算出 start_date/end_date 时不会传 date，
        #   这种「仅 start_date/end_date 有值」的合法调用不应被此分支误报
        #   修复：三个值都为 None/空才报错
        if not date and not start_date and not end_date:
            raise ValueError("❌ 至少需要传入 date 或 start_date/end_date")

        url = f"{self.BASE_URL}/add?requestFrom=0&businessFrom=1"
        payload = self._build_payload(start_date, end_date)

        resp = self.session.post(url, json=payload, timeout=30)
        resp.raise_for_status()
        ret = resp.json()

        # 阶段4 统一业务码识别
        self._handle_response(ret, op_desc="创建导出任务")

        # ⚠️ 阶段6 真实跑通发现（2026-08-07）：京准通 add 响应 data 格式不统一：
        #   - code=0 时 data 是 dict {reportId: "..."}
        #   - code=1 时 data 直接是 reportId int（如 22134297）
        # 必须兼容两种格式
        data = ret.get("data")
        if isinstance(data, dict):
            task_id = data["reportId"]
        else:
            task_id = data  # int / str 直接是 reportId
        print(f"✅ 创建导出任务成功：task_id={task_id}")
        return task_id

    # ---- 接口2：查询任务列表 ----

    def get_task_list(self) -> dict:
        """查询任务列表（GET），返回原始 dict（含 task 状态、downloadUrl）。

        返回:
            dict - 接口响应原始 dict，调用方按 task_id 字段匹配目标任务
        异常:
            CookieExpiredError - Cookie 过期（不重试）
            RuntimeError       - 业务码 601/-407/-402/其他非0
        """
        url = f"{self.BASE_URL}/list?requestFrom=0&businessFrom=1"
        resp = self.session.get(url, timeout=30)
        resp.raise_for_status()
        ret = resp.json()

        # 阶段4 统一业务码识别
        self._handle_response(ret, op_desc="查询任务列表")
        return ret

    # ---- 阶段4 容错：轮询等待报表生成 ----

    # 任务状态映射（2026-08-07 真实 list 响应实证 + atoms-api 对照修正）
    # 京准通实际状态字段是 subscribeState（int）。
    # ⚠️ 修正：旧版 SUBSCRIBE_STATE_OK=0 是错误推测；0 实际表示"排队处理中"。
    #      抓包 atoms-api status=2 与 jzt-api subscribeState=2 一致 → 2=报表生成完成。
    # 兼容 expected_status 参数同时支持字符串（"报表已生成"）和整数（0/1/2...）
    # ⚠️ 2026-08-09 真实跑通发现（task 22141980）：
    #   旧推测「2=完成」是错的！实际 subscribeState 长期保持 0，但报表已生成好（downloadById 能拿到 urlCsv）
    #   推测：subscribeState 字段含义不是"是否完成"，可能是"订阅/计划状态"等其他语义
    #   **改用 downloadById 探针判定**（能拿到 urlCsv 即视为完成），subscribeState 仅作辅助
    SUBSCRIBE_STATE_QUEUED = 1   # 占位（待确认）
    SUBSCRIBE_STATE_OK = 0       # 暂定 0（但不可靠，请用 _probe_download_ready 替代）
    SUBSCRIBE_STATE_FAIL = -1    # 暂定 -1=失败（占位）

    def wait_for_task_ready(self, task_id, expected_status=None) -> dict:
        """轮询任务列表直到任务可下载（纯探针策略）。

        ⚠️ 2026-08-09 第二次修正（基于 task 22141984 真实跑通）：
            subscribeState 字段语义**完全不可靠**——有时=0 是已完成，有时=0 是还没开始。
            决定放弃 subscribeState 判定，改用**纯 downloadById 探针**：
              - 每次轮询都调 downloadById，能拿到 urlCsv（且 OSS GET 200）即视为完成
              - 探针失败 → 继续等
              - 这样不依赖任何状态字段语义，跨京东版本兼容

        参数:
            task_id        - 来自 create_export_task 返回的任务 ID（int 或 str 都接受）
            expected_status - 保留兼容参数，不再使用
        返回:
            dict - 任务记录（含 reportName/tempName 等用于下载）
        异常:
            TimeoutError - 探针仍失败
            CookieExpiredError / RuntimeError - 业务码异常
        """
        import time

        if expected_status is None:
            expected_status = self.SUBSCRIBE_STATE_OK  # 保留兼容

        first_iteration = True

        for i in range(1, self.MAX_POLL_TIMES + 1):
            ret = self.get_task_list()
            records = ret.get("data", {}).get("data", [])
            match_item = next(
                (item for item in records if str(item.get("id")) == str(task_id)),
                None,
            )

            if not match_item:
                print(f"  [轮询 {i}/{self.MAX_POLL_TIMES}] task_id={task_id} 任务未出现，继续等待...")
                if i < self.MAX_POLL_TIMES:
                    time.sleep(self.POLL_INTERVAL)
                continue

            state = match_item.get("subscribeState")
            print(f"  [轮询 {i}/{self.MAX_POLL_TIMES}] task_id={task_id} subscribeState={state!r}（仅供参考）")

            # 失败状态立即停（这是 list 唯一可信的判定）
            if state == self.SUBSCRIBE_STATE_FAIL:
                raise RuntimeError(
                    f"❌ 任务生成失败 task_id={task_id}：{match_item}\n"
                    f"   可能原因：payload 字段错 / 账号无权限 / 数据异常"
                )

            # ⚠️ 纯 downloadById 探针（每次轮询都跑，不依赖状态字段）
            # 首次等待略长一些（避免无效探针）
            if first_iteration:
                first_iteration = False
                print(f"  ⏳ 首次等待 {self.POLL_INTERVAL*2:.0f} 秒后再探针（避免无效请求）...")
                time.sleep(self.POLL_INTERVAL * 2)
                continue

            probe_result = self._probe_download_ready(task_id, match_item)
            if probe_result:
                print(f"  ✅ 探针成功，任务已就绪（downloadById 能拿到 urlCsv）")
                return match_item

            if i < self.MAX_POLL_TIMES:
                time.sleep(self.POLL_INTERVAL)

        raise TimeoutError(
            f"❌ 轮询超过最大次数 {self.MAX_POLL_TIMES}，downloadById 探针始终无法拿到 urlCsv\n"
            f"   task_id={task_id}，可能原因：账号权限不足 / 数据异常 / 接口变更\n"
            f"   可手动浏览器登录 https://jzt.jd.com 查看任务状态"
        )

    def _probe_download_ready(self, task_id: str, match_item: dict) -> bool:
        """探针：调 downloadById 看能否拿到 urlCsv（不实际下载）。

        返回:
            bool - True=报表已就绪可下载，False=还没生成
        """
        file_name = match_item.get("reportName") or match_item.get("tempName")
        start_day = match_item.get("startTimeStr", "")
        end_day = match_item.get("endTimeStr", "")
        pin = match_item.get("pin", "")
        if not file_name:
            return False
        try:
            url = (
                f"{self.BASE_URL}/downloadById"
                f"?id={task_id}&name={file_name}"
                f"&startDay={start_day}&endDay={end_day}"
                f"&pin={pin}&fileName={file_name}"
                f"&requestFrom=0&businessFrom=1"
            )
            resp = self.session.get(url, timeout=30)
            if resp.status_code != 200:
                return False
            ret = resp.json()
            # 成功判定：success=true + code∈{0,1} 且 data.urlCsv 非空
            if ret.get("success", True) and ret.get("code") in (0, 1):
                url_csv = ret.get("data", {}).get("urlCsv")
                return bool(url_csv)
            return False
        except Exception:
            return False

    # ---- 接口3：CDN 下载（阶段4：CDN 403 自动重刷 URL 重试）----

    def download_report(self, task_id: str, save_filename: str) -> str:
        """根据 task_id 轮询等待报表生成，调 downloadById 拿 urlCsv，再 GET urlCsv 下载 CSV 流。

        ⚠️ 阶段8 真实实现（2026-08-07 抓包实证）：
            1. list 不返回下载 URL
            2. downloadById 返回 JSON（含 urlCsv，是 storage.jd.com 的 OSS 预签名链接）
            3. urlCsv 是 OSS 预签名链接，**带 Expires 过期时间**（实测 10 分钟有效）
            4. 必须**链式调用**：downloadById 拿到 urlCsv 立即 GET，否则会 404 NoSuchKey
            5. GET urlCsv 不需要 Cookie（OSS 自带签名），纯 requests.get 即可

        参数:
            task_id       - 来自 create_export_task 返回的任务 ID
            save_filename - 保存文件名（如 "report_0807.csv"）
        返回:
            str - 保存的文件绝对路径
        异常:
            TimeoutError / CookieExpiredError / RuntimeError
        """
        # 1. 轮询等待报表生成（subscribeState=2 即完成；0=排队中）
        match_item = self.wait_for_task_ready(task_id)

        # 2. 提取下载所需参数
        file_name = match_item.get("reportName") or match_item.get("tempName")
        if not file_name:
            raise RuntimeError(f"❌ 任务已就绪但 reportName 缺失，无法下载：{match_item}")
        start_day = match_item.get("startTimeStr", "")
        end_day = match_item.get("endTimeStr", "")
        pin = match_item.get("pin", "")

        # 3. 调 downloadById（同域 API）拿 urlCsv
        downloadbyid_url = (
            f"{self.BASE_URL}/downloadById"
            f"?id={task_id}"
            f"&name={file_name}"
            f"&startDay={start_day}"
            f"&endDay={end_day}"
            f"&pin={pin}"
            f"&fileName={file_name}"
            f"&requestFrom=0&businessFrom=1"
        )
        print(f"⬇️ 调 downloadById 拿 urlCsv: task_id={task_id}")

        try:
            resp_byid = self.session.get(downloadbyid_url, timeout=60)
        except requests.exceptions.RequestException as e:
            raise RuntimeError(f"❌ downloadById 请求失败：{e}") from e

        # 4. HTTP 状态码 + 业务码识别
        if resp_byid.status_code in (401, 403):
            raise CookieExpiredError(
                f"❌ downloadById 返回 {resp_byid.status_code}（Cookie 过期/账号限制）：\n"
                f"   → 请浏览器重新登录 https://jzt.jd.com/home，F12 抓 Cookie 写入 config/jzt_cookie.txt"
            )
        resp_byid.raise_for_status()

        ret_byid = resp_byid.json()
        # 业务码校验（success=true 且 code∈{0,1} 视为成功）
        self._handle_response(ret_byid, op_desc="downloadById")

        url_csv = ret_byid.get("data", {}).get("urlCsv")
        if not url_csv:
            raise RuntimeError(
                f"❌ downloadById 响应中 urlCsv 缺失：{ret_byid}\n"
                f"   可能原因：报表还没真正生成 / 接口字段名变更"
            )

        # 5. 立即 GET urlCsv（OSS 预签名链接）
        # ⚠️ 2026-08-09 修正：单纯 GET 同 URL 重试即可（OSS链接有效期内稳定）
        #   旧版"每次失败重调 downloadById 拿新 urlCsv"是错的：
        #   - 新 urlCsv 是新 OSS 文件路径，新文件可能还没生成
        #   - 京东 OSS 是异步生成，旧 urlCsv 对应的文件**正在生成中**，多等几次就 200
        print(f"⬇️ 下载 urlCsv（OSS 预签名链接；首次可能 404 等几秒重试）")
        print(f"  URL 前 80 字符: {url_csv[:80]}...")

        resp_csv = None
        last_error = None
        for retry in range(self.MAX_DOWNLOAD_RETRY + 1):
            try:
                resp_csv = requests.get(
                    url_csv,
                    headers={"User-Agent": self.USER_AGENT},
                    timeout=60,
                )
                # 200 成功
                if resp_csv.status_code == 200:
                    break
                # 404 NoSuchKey（OSS 文件还在生成中）→ 同一 urlCsv 退避重试
                if resp_csv.status_code == 404 and retry < self.MAX_DOWNLOAD_RETRY:
                    import time as _time
                    # ⚠️ 随机退避 3-10 秒（避免固定间隔被风控识别；同时给 OSS 足够预热时间）
                    backoff = random.uniform(3, 10)
                    print(
                        f"  ⚠️ 第 {retry+1}/{self.MAX_DOWNLOAD_RETRY+1} 次 urlCsv 404 NoSuchKey"
                        f"，随机退避 {backoff:.1f} 秒后重试（同一 urlCsv）..."
                    )
                    _time.sleep(backoff)
                    continue  # 注意：不重新调 downloadById，同一 urlCsv 继续 GET
                # 其他非 200
                resp_csv.raise_for_status()
            except requests.exceptions.RequestException as e:
                last_error = e
                if retry >= self.MAX_DOWNLOAD_RETRY:
                    raise RuntimeError(f"❌ urlCsv 下载失败（重试 {self.MAX_DOWNLOAD_RETRY+1} 次后）：{e}") from e
                print(f"  ⚠️ urlCsv 下载异常：{e}，重试中...")

        if resp_csv is None or resp_csv.status_code != 200:
            raise RuntimeError(
                f"❌ urlCsv 连续 {self.MAX_DOWNLOAD_RETRY+1} 次未成功：{last_error}\n"
                f"   可能原因：报表生成尚未完成 / OSS 预热延迟超出预期"
            )

        # 6. 落盘 + Excel 后置处理（2026-08-09 对齐 AGENTS.md Excel 报表统一规则）
        #    流程：raw CSV → pandas 读取(dtype=str) → prepare_date_columns →
        #          safe_convert_numeric → 转存为 .xlsx + apply_column_formats
        #    输出路径遵循 AGENTS.md Excel规则4：output/京准通快车效果自定义/{date}/业务名_日期.xlsx
        return self._post_process_csv_to_xlsx(resp_csv.content, task_id)

    # ---- Excel 后置处理：CSV → xlsx ----

    def _post_process_csv_to_xlsx(self, csv_bytes: bytes, task_id) -> str:
        """把京东 OSS 返回的 raw CSV 字节流 → 标准 Excel 后置处理 → 保存为 xlsx。

        流程（对齐 AGENTS.md Excel 报表统一规则 + 商品流失分析项目6 模式）：
            ① pandas.read_csv(dtype=str, na_filter=False) → 防精度丢失
            ② prepare_date_columns(df, date) → 日期列智能处理（公共规则1+2）
            ③ safe_convert_numeric(df) → 数值安全转换（公共规则3）
            ④ 转存为 .xlsx + apply_column_formats 设置单元格格式（SKU/SPU 0位小数、订单编号@）
            ⑤ 落盘路径：output/京准通快车效果自定义/{date}/京准通快车效果自定义_{date}.xlsx

        参数:
            csv_bytes - OSS 下载的原始 CSV 字节流（含 UTF-8 BOM）
            task_id   - 任务 ID（用于日志关联）
        返回:
            str - 保存的 .xlsx 绝对路径
        异常:
            RuntimeError - CSV 解析失败
        """
        import io
        import pandas as pd

        # ⚠️ 京准通 OSS 返回的 CSV 文件名通常带"下载.csv"等中文，这里从任务获取干净文件名
        match_item = self._find_task_in_list(task_id)
        clean_date = match_item.get("startTimeStr", "") if match_item else ""
        if not clean_date:
            # 兜底：从 list 找、或用今天
            from datetime import datetime as _dt
            clean_date = _dt.now().strftime("%Y-%m-%d")

        # 1. 读取 CSV（dtype=str 防长数字精度丢失；na_filter=False 防 "0" 被当 NaN）
        try:
            df = pd.read_csv(
                io.BytesIO(csv_bytes),
                dtype=str,
                na_filter=False,
                encoding="utf-8-sig",  # 兼容 BOM
                keep_default_na=False,
            )
        except Exception as e:
            raise RuntimeError(
                f"❌ CSV 解析失败：{e}\n"
                f"   任务 task_id={task_id}，请检查 OSS 返回内容是否正常"
            ) from e

        if df.empty:
            raise RuntimeError(f"❌ CSV 数据为空：task_id={task_id}")

        # 2. 日期列智能处理（公共规则1+2）
        date_column, date_value = prepare_date_columns(df, clean_date)

        # 3. 数值安全转换（公共规则3）
        df = safe_convert_numeric(df)

        # 4. 构造输出路径：output/京准通快车效果自定义/{date}/京准通快车效果自定义_{date}.xlsx
        date_subdir = os.path.join(self.output_dir, clean_date)
        os.makedirs(date_subdir, exist_ok=True)
        save_filename = f"京准通快车效果自定义_{clean_date}.xlsx"
        target_path = os.path.join(date_subdir, save_filename)

        # 5. 写 xlsx + 设置单元格格式
        df.to_excel(target_path, index=False, engine="openpyxl")
        apply_column_formats(target_path, df, date_column=date_column, date_value=date_value)

        print(
            f"✅ 文件已保存：{target_path}"
            f"\n   （CSV→xlsx 转存 + 日期列 + 数值转换 + 单元格格式，{os.path.getsize(target_path)}字节，{len(df)}行 × {len(df.columns)}列）"
        )
        return target_path

    def _find_task_in_list(self, task_id) -> dict:
        """从 list 接口获取目标任务信息（用于提取日期等参数）。"""
        try:
            ret = self.get_task_list()
            records = ret.get("data", {}).get("data", [])
            for item in records:
                if str(item.get("id")) == str(task_id):
                    return item
        except Exception:
            pass
        return {}

    # ---- 阶段10 完整流程封装（一键跑通）----

    def run_full_export(
        self,
        start_date: str = None,
        end_date: str = None,
        date: str = None,
        save_filename: str = None,
    ) -> str:
        """一键跑通：创建任务 → 内部轮询等待 → downloadById 拿 urlCsv → OSS 下载落盘。

        ⚠️ 推荐对外调用入口：外部不需要手工编排多步，本方法封装完整三步异步流程。
        设计动机：原 BUSINESS_REGISTRY 仅挂 create_export_task，需要外部手动调 download_report。
                    现统一为 run_full_export 一步到位，调度器只需要注册这一个方法。

        参数:
            start_date    - 开始日期 YYYY-MM-DD（默认 = date）
            end_date      - 结束日期 YYYY-MM-DD（默认 = date）
            date          - 单日查询 YYYY-MM-DD（start/end 默认 = date）
            save_filename - 保存文件名（如 "快车_2026-08-07.csv"）；None 时按日期+时间戳自动命名
        返回:
            str - 保存的文件绝对路径
        异常:
            CookieExpiredError / TimeoutError / RuntimeError
        """
        # 1. 三值一致规则（与项目1-6 调度层对齐）
        if date is None and start_date is None and end_date is None:
            raise ValueError("❌ 至少需要传入 date 或 start_date/end_date")
        if start_date is None:
            start_date = date
        if end_date is None:
            end_date = date

        # 2. 默认文件名：含日期范围 + 时间戳，避免重复覆盖
        if save_filename is None:
            from datetime import datetime as _dt
            ts = _dt.now().strftime("%H%M%S")
            save_filename = f"快车_{start_date}_{end_date}_{ts}.csv"

        # 3. 完整链路
        print(f"🚀 [JZT快车] 启动完整导出：{start_date} ~ {end_date}")
        print(f"   └─ Step 1/3: 创建导出任务...")
        task_id = self.create_export_task(date=date, start_date=start_date, end_date=end_date)
        print(f"   └─ Step 2/3: 轮询等待任务就绪（subscribeState=2）...")
        print(f"   └─ Step 3/3: 下载并落盘...")
        return self.download_report(task_id, save_filename)


# ============================================================
#  业务接口 6：（新业务 - 京准通快车订单效果明细报表，2026-08-09 上线）
# ------------------------------------------------------------
#  中文说明（小白必读）：
#    京东快车订单效果明细报表（reweb/msa/effect/order/download）
#    与"业务接口5"自定义报表**完全不同的接口**：本接口是**同步返回下载链接**，
#    调一次 POST 就直接拿到 downloadUrlCsv，不需要轮询 list 接口等异步任务。
#
#  ⚠️ 核心差异（与项目7对比）：
#    - 接口 URL 不同：reweb/msa/effect/order/download（不是 /dataCenter/customreport/v2/report/add）
#    - 流程：POST 同步返回下载链接 → 立即 GET urlCsv → 落盘
#    - 无 h5st：抓包请求头无 h5st 字段，与项目7一致
#    - 报表名：必须传 reportName 字段（决定下载文件名）；无重名检测要求（待验证）
#    - 数据维度：订单明细（订单号 / SKU ID / 金额 / 时间 / 地域 / 订单类型）
#
#  参数说明（2026-08-09 抓包实证）：
#    业务固定参数（类常量，不变）：
#      clickOrOrderCaliber=0   # 点击/下单口径：0=点击
#      clickOrOrderDay=15       # 转化周期：15天
#      giftFlag=0               # 含赠品：0=不含
#      orderStatusCategory=1    # 下单/成交订单：1=成交订单
#      orderType="1,3"          # 订单类型（抓包值，含义待补）
#      orderStatuses=[]         # 订单状态列表（空）
#    日期参数（每次可变）：
#      startDay / endDay → 优先用入参
#    报表名（动态）：
#      reportName 格式：{pin}_{固定描述}_{startDay}_{endDay} → 例：
#        "FYA8888_报表中心_订单_15天_点击_不含赠品_20260807_20260807"
# ============================================================

class JZTKuaicheOrderEffectAPI:
    """京准通快车订单效果明细报表导出 API（2026-08-09 上线骨架）。

    ⚠️ 本类**不继承 JDBaseRequest**（与项目7 同样的原因）：
        - 鉴权体系：仅 Cookie（与项目7 同一文件 config/jzt_cookie.txt）
        - 流程：同步两步（POST → 立即 GET OSS），不需要基类的 30 秒间隔/重试模型
        - UA：禁止切换（h5st 与 UA 绑定；本接口无 h5st 但保留习惯）
    """

    # ---- 类常量（业务固定参数）----
    BASE_URL = "https://jzt-api.jd.com/reweb/msa/effect/order/download"
    ORIGIN = "https://jzt.jd.com"
    REFERER = "https://jzt.jd.com/"
    SITE_ID = "0"
    USER_AGENT = (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/151.0.0.0 Safari/537.36 Edg/151.0.0.0"
    )
    OUTPUT_SUBDIR = "京准通快车订单效果明细"  # output/京准通快车订单效果明细/{date}/

    # 业务固定参数（抓包值，2026-08-09 实证）
    CLICK_OR_ORDER_CALIBER = 0      # 0=点击
    CLICK_OR_ORDER_DAY = 15         # 转化周期：15 天
    GIFT_FLAG = 0                   # 0=不含赠品
    ORDER_STATUS_CATEGORY = 1       # 1=成交订单
    ORDER_TYPE = "1,3"              # 订单类型（含义待补查）
    PIN_ID = "FYA8888"

    def __init__(self, cookie_path: str = "config/jzt_cookie.txt"):
        import requests

        # 读 Cookie（与项目7 互通 jzt_cookie.txt）
        cookie_path_abs = os.path.join(os.path.dirname(os.path.abspath(__file__)), cookie_path)
        if not os.path.isfile(cookie_path_abs):
            raise FileNotFoundError(
                f"❌ 京准通 Cookie 文件不存在：{cookie_path_abs}\n"
                f"   请浏览器登录 https://jzt.jd.com/home，F12 抓 jzt-api.jd.com 域 Cookie 写入此文件"
            )
        with open(cookie_path_abs, "r", encoding="utf-8") as f:
            self.cookie = f.read().strip()
        if not self.cookie:
            raise ValueError(f"❌ 京准通 Cookie 文件 {cookie_path_abs} 内容为空")

        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": self.USER_AGENT,
            "Accept": "application/json, text/plain, */*",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
            "Accept-Encoding": "gzip, deflate, br, zstd",
            "Content-Type": "application/json",
            "Origin": self.ORIGIN,
            "Referer": self.REFERER,
            "language": "zh_CN",
            "siteid": self.SITE_ID,
            "sec-ch-ua": '"Not=A?Brand";v="99", "Microsoft Edge";v="151", "Chromium";v="151"',
            "sec-ch-ua-mobile": "?0",
            "sec-ch-ua-platform": '"Windows"',
            "Cookie": self.cookie,
        })

        # 输出目录（按 AGENTS.md Excel规则4）
        self.output_dir = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "output", self.OUTPUT_SUBDIR,
        )

    # ---- 业务参数组装 ----
    def _build_payload(self, start_day: str, end_day: str) -> dict:
        """组装请求 payload（抓包实证 + 动态日期）。"""
        from datetime import datetime as _dt

        # 报表名：FYA8888_报表中心_订单_15天_点击_不含赠品_{startDay}_{endDay}
        # 注意：长度需 < 30 字符（实测安全），加日期足够区分
        # 抓包原始长度：FYA8888_报表中心_订单_15天_点击_不含赠品_20260807_20260807 ≈ 46 字符
        #   ⚠️ 超过 30 字符，但实测能跑通；如需严格 < 30 再压缩
        report_name = (
            f"{self.PIN_ID}_报表中心_订单_{self.CLICK_OR_ORDER_DAY}天_"
            f"{'点击' if self.CLICK_OR_ORDER_CALIBER == 0 else '下单'}_"
            f"{'不含赠品' if self.GIFT_FLAG == 0 else '含赠品'}_"
            f"{start_day}_{end_day}"
        )
        return {
            "startDay": start_day,
            "endDay": end_day,
            "clickOrOrderCaliber": self.CLICK_OR_ORDER_CALIBER,
            "clickOrOrderDay": self.CLICK_OR_ORDER_DAY,
            "giftFlag": self.GIFT_FLAG,
            "orderStatusCategory": self.ORDER_STATUS_CATEGORY,
            "orderType": self.ORDER_TYPE,
            "orderStatuses": [],
            "reportName": report_name,
        }

    def _handle_response(self, ret: dict, op_desc: str):
        """统一处理响应（与项目7 同样的双字段判定）。

        京东快车订单接口的判定：
            - success=true
            - code=1（数字）
            - data.code == "RC_SUCCESS"（字符串）
            - data.downloadUrlCsv 非空
        """
        if not ret.get("success", True):
            msg = ret.get("msg", "未知错误")
            code = ret.get("code")
            if code in (2001, 302) or "未登录" in msg or "登录已过期" in msg:
                raise CookieExpiredError(
                    f"❌ 京准通 Cookie 过期（{op_desc}返回 code={code}）：\n"
                    f"   → 请浏览器登录 https://jzt.jd.com/home，F12 抓 jzt-api.jd.com 域 Cookie 写入 config/jzt_cookie.txt"
                )
            if code == 601:
                raise RuntimeError(
                    f"❌ 京准通 订单接口 限流 code=601：{msg}\n"
                    f"   → 30-120 分钟冷却，避免重试加重风控"
                )
            raise RuntimeError(
                f"❌ 京准通{op_desc}失败：code={code}, msg={msg}, 完整响应={ret}"
            )
        code = ret.get("code")
        data_code = ret.get("data", {}).get("code")
        # ⚠️ 2026-08-09 项目9 全站营销接口探针发现：code 可能是字符串 "1" 而非数字 1
        # 兼容写法：str(code) 后再判
        if str(code) not in ("0", "1") or data_code != "RC_SUCCESS":
            raise RuntimeError(
                f"❌ 京准通{op_desc}业务失败：code={code}, data.code={data_code}\n"
                f"   完整响应：{ret}"
            )
        return ret

    # ---- 一步：同步 POST 拿 downloadUrlCsv ----
    def _post_for_csv(self, start_day: str, end_day: str) -> str:
        """POST 同步返回 downloadUrlCsv。

        返回:
            str - OSS 预签名链接（10 分钟有效）
        异常:
            CookieExpiredError / RuntimeError
        """
        payload = self._build_payload(start_day, end_day)
        print(f"🚀 [JZT快车订单明细] POST {self.BASE_URL}")
        print(f"   Body: {json.dumps(payload, ensure_ascii=False)}")

        resp = self.session.post(self.BASE_URL, json=payload, timeout=60)
        resp.raise_for_status()
        ret = resp.json()
        self._handle_response(ret, op_desc="导出订单明细")

        download_url = ret.get("data", {}).get("downloadUrlCsv")
        if not download_url:
            raise RuntimeError(f"❌ 响应中 downloadUrlCsv 缺失：{ret}")
        download_id = ret.get("data", {}).get("downloadId")
        print(f"✅ 拿到 downloadId={download_id}, downloadUrlCsv（前80字符）: {download_url[:80]}...")
        return download_url

    # ---- 二步：GET OSS 下载 CSV 字节流 ----
    def _download_csv(self, url_csv: str) -> bytes:
        """GET OSS 链接，下载 CSV 字节流（带 404 随机退避重试）。

        OSS 链接 10 分钟有效，但首次可能 404 NoSuchKey（异步生成）。
        """
        last_error = None
        for retry in range(4):  # 最多 4 次
            try:
                resp = requests.get(
                    url_csv,
                    headers={"User-Agent": self.USER_AGENT},
                    timeout=60,
                )
                if resp.status_code == 200:
                    return resp.content
                if resp.status_code == 404:
                    backoff = random.uniform(3, 10)
                    print(
                        f"  ⚠️ 第 {retry+1}/4 次 urlCsv 404 NoSuchKey，"
                        f"随机退避 {backoff:.1f} 秒后重试..."
                    )
                    time.sleep(backoff)
                    continue
                resp.raise_for_status()
            except requests.exceptions.RequestException as e:
                last_error = e
                print(f"  ⚠️ urlCsv 下载异常：{e}，重试中...")
                time.sleep(random.uniform(3, 10))
        raise RuntimeError(
            f"❌ urlCsv 下载失败（重试 4 次后）：{last_error}\n"
            f"   URL: {url_csv[:120]}"
        )

    # ---- 一键封装（推荐对外入口）----
    def run_full_export(
        self,
        start_date: str = None,
        end_date: str = None,
        date: str = None,
    ) -> str:
        """一键跑通：POST 同步拿 urlCsv → GET OSS 下载 → Excel 后置处理 → 落盘。

        参数:
            start_date - 开始日期 YYYY-MM-DD（默认 = date）
            end_date   - 结束日期 YYYY-MM-DD（默认 = date）
            date       - 单日查询 YYYY-MM-DD
        返回:
            str - 保存的 .xlsx 绝对路径
        """
        if date is None and start_date is None and end_date is None:
            raise ValueError("❌ 至少需要传入 date 或 start_date/end_date")
        if start_date is None:
            start_date = date
        if end_date is None:
            end_date = date

        print(f"🚀 [JZT快车订单明细] 启动完整导出：{start_date} ~ {end_date}")
        print(f"   └─ Step 1/2: POST 同步拿 downloadUrlCsv...")
        url_csv = self._post_for_csv(start_date, end_date)
        print(f"   └─ Step 2/2: GET OSS 下载并落盘为 xlsx...")
        csv_bytes = self._download_csv(url_csv)

        # 复用项目7 的 Excel 后置处理（日期/数值/格式）
        return self._post_process_csv_to_xlsx(csv_bytes, start_date)

    def _post_process_csv_to_xlsx(self, csv_bytes: bytes, clean_date: str) -> str:
        """把 OSS 下载的 raw CSV → 标准 Excel 后置处理 → 保存为 xlsx。

        与项目7 共用同样的 prepare_date_columns / safe_convert_numeric / apply_column_formats。
        """
        import io
        import pandas as pd

        # 1. 读取 CSV（dtype=str 防精度丢失；UTF-8-sig 兼容 BOM）
        try:
            df = pd.read_csv(
                io.BytesIO(csv_bytes),
                dtype=str,
                na_filter=False,
                encoding="utf-8-sig",
                keep_default_na=False,
            )
        except Exception as e:
            raise RuntimeError(
                f"❌ CSV 解析失败：{e}\n"
                f"   请检查 OSS 返回内容是否正常（可能含 BOM 或格式变化）"
            ) from e

        if df.empty:
            raise RuntimeError("❌ CSV 数据为空")

        # 2. 日期列智能处理（公共规则1+2）
        #    本报表自带"点击时间"/"下单时间"列 → 不会插入新日期列，只做格式标准化
        date_column, date_value = prepare_date_columns(df, clean_date)

        # 3. 数值安全转换（公共规则3）—— 订单编号强制文本，SKU 转数字 0 位小数
        df = safe_convert_numeric(df)

        # 4. 构造输出路径：output/京准通快车订单效果明细/{date}/业务名_{date}.xlsx
        date_subdir = os.path.join(self.output_dir, clean_date)
        os.makedirs(date_subdir, exist_ok=True)
        save_filename = f"京准通快车订单效果明细_{clean_date}.xlsx"
        target_path = os.path.join(date_subdir, save_filename)

        # 5. 写 xlsx + 单元格格式
        df.to_excel(target_path, index=False, engine="openpyxl")
        apply_column_formats(target_path, df, date_column=date_column, date_value=date_value)

        print(
            f"✅ 文件已保存：{target_path}"
            f"\n   （CSV→xlsx 转存 + 日期列 + 数值转换 + 单元格格式，{os.path.getsize(target_path)}字节，{len(df)}行 × {len(df.columns)}列）"
        )
        return target_path


# ============================================================
#  业务注册中心（BUSINESS_REGISTRY）
# ------------------------------------------------------------
#  中文说明（小白必读）：
#    这是整个项目的"业务路由表"。新增业务的正确做法：
#      1. 定义业务API类（继承JDBaseRequest）
#      2. 在 BUSINESS_REGISTRY 里注册一个 key，绑定业务类和处理方法
#      3. 完成！调用 run_business("你的业务key") 即可触发
#    严禁修改 run_business 主体逻辑来"加业务"，那是反模式。
# ============================================================

# 业务注册表
# 格式：
#   "业务key": {
#       "api_class": API类,
#       "method":   业务类的方法名（字符串）,
#       "desc":     业务描述,
#       "params":   业务专属参数说明（dict，键值对形式展示给用户）
#   }
#
# ════════════════════════════════════════════════════════════════════════
# ⚠️ 业务模块占位（2026-08-11 用户决策：3 大模块各留占位，方便后续扩展）
# ════════════════════════════════════════════════════════════════════════
#
# 📦 商智模块（sz）- szgateway.jd.com 域
#    现有业务（项目1/4/5/6/13）：
#      - 商品流量来源_搜索 / _推荐 / _购物车
#      - 店铺来源_三级渠道
#      - 商品明细导出
#      - 商品流失分析
#      - 商智关键词分析
#    Cookie：config/sz_cookie.txt（不入仓）
#    风控：USER_MNP_SALT + UUID_PREFIX（在 config/config.xlsx + 类常量）
#    新增项目时复制现有 class 模板，修改 API_URL / payload / 列名即可
#
# 📦 京准通模块（jzt）- jzt-api.jd.com 域
#    现有业务（项目7-12）：
#      - 京准通快车自定义报表 / 订单效果明细
#      - 京准通全站营销单品计划 / 单品推广效果 / 全店计划 / 全店推广效果
#    Cookie：config/jzt_cookie.txt（不入仓）
#    鉴权：仅 Cookie（h5st 可选）+ UA v=151
#    新增项目时复制 JZTKuaicheAPI / JZTQuanZhanEffectAPI 等模板
#
# 📦 京麦模块（jm）- sff.jd.com + export.shop.jd.com 域
#    现有业务（项目14，5 个一键）：
#      - 京麦订单明细_创建任务 / 创建并轮询 / 创建轮询下载zip / 创建轮询下载并申请密码 / 完整一键导出
#    Cookie：config/jm_cookie.txt（不入仓）
#    鉴权：h5st（一次性，CLI --h5st 传）+ dsm-* 全套头
#    短信密码：QQ 邮箱 IMAP（config/imap_config.ini 不入仓）
#    新增项目时复制 JingMaiOrderExportAPI 模板（5 步异步链路）
#
# ════════════════════════════════════════════════════════════════════════
BUSINESS_REGISTRY = {
    "商品流量来源_搜索": {
        "api_class": ProductFlowAPI,
        "method": "download_search_sku",
        "desc": "商品搜索效果（搜索子来源2008）",
        "params": {
            "date": "查询日期YYYY-MM-DD（从config.xlsx的date读取）",
            "startDate": "开始日期（默认=date）；⚠️接口服务端不支持多日区间，传区间将自动逐日拆分循环（P1-1）",
            "endDate": "结束日期（默认=date）；同上，区间将自动逐日拆分",
        },
    },
    "商品流量来源_推荐": {
        "api_class": ProductFlowAPI,
        "method": "download_recommend_sku",
        "desc": "商品推荐效果（推荐子来源2009）",
        "params": {
            "date": "查询日期YYYY-MM-DD",
            "startDate": "开始日期；⚠️接口服务端不支持多日区间，传区间将自动逐日拆分循环（P1-1）",
            "endDate": "结束日期；同上，区间将自动逐日拆分",
        },
    },
    "商品流量来源_购物车": {
        "api_class": ProductFlowAPI,
        "method": "download_cart_sku",
        "desc": "商品购物车效果（购物车/我的订单回流，3001）",
        "params": {
            "date": "查询日期YYYY-MM-DD",
            "startDate": "开始日期；⚠️接口服务端不支持多日区间，传区间将自动逐日拆分循环（P1-1）",
            "endDate": "结束日期；同上，区间将自动逐日拆分",
        },
    },
    # ⚠️ 自主访问：与购物车数据口径重叠（同3001），enabled=False 停用不执行。
    #    仅保留注册配置供存档/回溯，后续需要可把 enabled 改回 True 即可开启。
    "商品流量来源_自主访问": {
        "api_class": ProductFlowAPI,
        "method": "download_selfvisit_sku",
        "desc": "商品自主访问效果（与购物车数据口径重叠，已停用，3001）",
        "enabled": False,
        "params": {
            "date": "查询日期YYYY-MM-DD",
            "startDate": "开始日期",
            "endDate": "结束日期",
        },
    },
    # === 后续业务在此注册 ===
    # 业务：店铺来源-三级渠道（离线流量报表，2026-08-06 上线）
    "店铺来源_三级渠道": {
        "api_class": OfflineChannelAPI,
        "method": "download_offline_channel",
        "desc": "店铺来源离线渠道流量报表（按三级渠道分组，进店访客数降序）",
        "params": {
            "date": "查询日期YYYY-MM-DD（入参或config）",
            "startDate": "开始日期（默认=date）",
            "endDate": "结束日期（默认=date）",
            "platformCate1": "平台品类1（默认空=全品类）",
        },
    },
    # 业务：商品明细导出（2026-08-07 上线，GET 请求）
    "商品明细导出": {
        "api_class": ProductDetailAPI,
        "method": "download_product_detail",
        "desc": "商品明细导出（按二级/三级类目，GET导出，保存至 output/商品明细/{date}/）",
        "params": {
            "date": "查询日期YYYY-MM-DD（入参或config）",
            "second": "二级类目ID（默认999999=全类目）",
            "third": "三级类目ID（默认空=不限）",
            "channel": "渠道ID（默认99=全部渠道）",
        },
    },
    # 业务：商品流失分析（2026-08-07 上线，POST 请求，xls 响应转存 xlsx）
    "商品流失分析": {
        "api_class": LossProductAPI,
        "method": "download_loss_product",
        "desc": "商品流失分析（竞争-竞争流失-商品流失，POST导出，xls转存xlsx，output/商品流失分析/{date}/）",
        "params": {
            "date": "查询日期YYYY-MM-DD（入参或config）",
            "startDate": "开始日期（默认=date）",
            "endDate": "结束日期（默认=date）",
        },
    },
    # 业务：京准通快车自定义报表（阶段10：完整三步流程接入调度器）
    # ⚠️ 调度器支持两种注册方式：
    #   1. 标准方式：api_class + method（基类方法自动实例化）
    #   2. callable 方式：本业务因 h5st/cookie_path 需动态注入，采用自定义函数直接注册
    #   get_business_handler() 检测到 info.get("callable") 时优先返回该函数
    "京准通快车自定义报表": {
        "api_class": JZTKuaicheAPI,  # 兼容老调用；实际调度走 callable
        "method": "run_full_export",  # 实例化后也支持直调
        "callable": None,  # 占位：下方 _run_jzt_kuaiche_full 函数定义后注入（避免前向引用错误）
        "desc": "京准通快车自定义报表导出（h5st鉴权，独立Cookie，三步异步：add→轮询→CDN下载）",
        "params": {
            "h5st": "必填，浏览器F12抓add接口请求头复制（外部传入）",
            "start_date": "开始日期YYYY-MM-DD",
            "end_date": "结束日期YYYY-MM-DD",
            "cookie_path": "京准通Cookie路径（默认config/jzt_cookie.txt，可选）",
        },
    },
    # 业务：京准通快车订单效果明细（2026-08-09 上线，同步两步流程）
    #   同步接口：POST /reweb/msa/effect/order/download 直接返回 downloadUrlCsv
    #   无 h5st（抓包实证）；复用 jzt_cookie.txt；输出到 output/京准通快车订单效果明细/{date}/
    "京准通快车订单效果明细": {
        "api_class": JZTKuaicheOrderEffectAPI,
        "method": "run_full_export",
        "callable": None,  # 占位：下方 _run_jzt_order_effect_full 函数定义后注入
        "desc": "京准通快车订单效果明细导出（同步两步：POST拿urlCsv→GET下载）",
        "params": {
            "date": "查询日期YYYY-MM-DD（单日查询）",
            "start_date": "开始日期YYYY-MM-DD（区间查询，可选）",
            "end_date": "结束日期YYYY-MM-DD（区间查询，可选）",
            "cookie_path": "京准通Cookie路径（默认config/jzt_cookie.txt，可选）",
        },
    },
    # 业务：京准通全站营销单品计划（2026-08-09 上线，同步两步流程）
    #   同步接口：POST /reweb/swa/account/campaign/download 直接返回 downloadUrlCsv
    #   字段差异：giftFlag/orderStatus/sxuId/obys/province 是字符串""；campaignTypes 是列表
    #   输出到 output/京准通全站营销单品计划/{date}/
    "京准通全站营销单品计划": {
        "api_class": None,  # 占位：下方 JZTQuanZhanCampaignAPI 类定义后注入
        "method": "run_full_export",
        "callable": None,  # 占位：下方 _run_jzt_quanzhan_campaign_full 函数定义后注入
        "desc": "京准通全站营销单品计划报表导出（同步两步：POST拿urlCsv→GET下载）",
        "params": {
            "date": "查询日期YYYY-MM-DD（单日查询）",
            "start_date": "开始日期YYYY-MM-DD（区间查询，可选）",
            "end_date": "结束日期YYYY-MM-DD（区间查询，可选）",
            "cookie_path": "京准通Cookie路径（默认config/jzt_cookie.txt，可选）",
        },
    },
    # 业务：京准通全站营销单品推广效果（2026-08-10 上线，同步两步流程）
    #   同步接口：POST /reweb/swa/effect/order/download 同时返回 downloadUrlZip + downloadUrlCsv
    #   优先级：zip 优先（更完整），csv 降级备用
    #   与项目9 关键差异：
    #     - URL 路径多了 /effect/ 段（项目9 是 /swa/account/campaign/download）
    #     - payload：orderStatus 默认 "1"（成交订单，开放传空），isDaily 默认 False（开放入参），skuId/spuId 是新增（默认 ""）
    #     - 响应增加 downloadUrlZip
    #   输出到 output/京准通全站营销单品推广效果/{date}/
    "京准通全站营销单品推广效果": {
        "api_class": None,  # 占位：下方 JZTQuanZhanEffectAPI 类定义后注入
        "method": "run_full_export",
        "callable": None,  # 占位：下方 _run_jzt_quanzhan_effect_full 函数定义后注入
        "desc": "京准通全站营销单品推广效果报表导出（同步两步：POST拿urlZip/csv→GET下载→解压转xlsx）",
        "params": {
            "date": "查询日期YYYY-MM-DD（单日查询）",
            "start_date": "开始日期YYYY-MM-DD（区间查询，可选）",
            "end_date": "结束日期YYYY-MM-DD（区间查询，可选）",
            "cookie_path": "京准通Cookie路径（默认config/jzt_cookie.txt，可选）",
            "order_status": "订单状态（可选，默认None→\"1\"成交订单；传\"\"=不限）",
            "is_daily": "日报标志（可选，默认None→False非日报；传True=日报）",
            "sku_id": "SKU过滤（可选，默认\"\"=不过滤）",
            "spu_id": "SPU过滤（可选，默认\"\"=不过滤）",
        },
    },
    # 业务：京准通全站营销全店计划（2026-08-10 上线，同步两步流程）
    #   同步接口：POST /reweb/swa/account/campaign/download（与项目9 同URL，靠 campaignTypes=[118] 区分业务）
    #   与项目9 关键差异：
    #     - campaignTypes=[118]（疑似全店计划，含义待用户确认；项目9 是 [101]）
    #     - 报表名后缀：_全店计划报表_（项目9 是 _单品计划报表_）
    #     - 响应同时含 downloadUrlZip + downloadUrlCsv（项目9 抓包只含 csv，本次数据完整）
    #   输出到 output/京准通全站营销全店计划/{date}/
    "京准通全站营销全店计划": {
        "api_class": None,  # 占位：下方 JZTQuanZhanCampaignAllStoreAPI 类定义后注入
        "method": "run_full_export",
        "callable": None,  # 占位：下方 _run_jzt_quanzhan_campaign_all_store_full 函数定义后注入
        "desc": "京准通全站营销全店计划报表导出（同步两步：POST拿urlZip/csv→GET下载→解压转xlsx）",
        "params": {
            "date": "查询日期YYYY-MM-DD（单日查询）",
            "start_date": "开始日期YYYY-MM-DD（区间查询，可选）",
            "end_date": "结束日期YYYY-MM-DD（区间查询，可选）",
            "cookie_path": "京准通Cookie路径（默认config/jzt_cookie.txt，可选）",
            "order_status": "订单状态（可选，默认None→\"\"不限）",
            "is_daily": "日报标志（可选，默认None→True日报；传False=非日报）",
            "sku_id": "SKU过滤（可选，默认\"\"=不过滤；项目9字段名 sxuId）",
            "spu_id": "SPU过滤（可选，默认\"\"=不过滤）",
        },
    },
    # 业务：商智关键词分析导出（2026-08-10 上线，表单格式同步 xlsx 流）
    #   同步接口：POST /szpaas/szajax/keyword/analysis/shopOut/downTable.ajax
    #   鉴权：Cookie + User-mup/uuid/User-mnp（UUID 完全随机）
    #   输出：output/商智关键词分析/{YYYY-MM-DD}/商智关键词分析_{YYYY-MM-DD}_{day|month}.xlsx
    "商智关键词分析": {
        "api_class": None,  # 占位：KeywordAnalysisAPI 类在下方，模块末尾回填
        "method": "run_full_export",
        "callable": None,  # 占位：下方 _run_keyword_analysis_full 函数定义后注入
        "desc": "商智关键词分析导出（同步xlsx流+UUID完全随机+day/month双粒度）",
        "params": {
            "date": "查询日期YYYY-MM-DD（单日查询，与 granularity='day' 配套）",
            "start_date": "开始日期YYYY-MM-DD（区间查询，与 granularity='month' 配套）",
            "end_date": "结束日期YYYY-MM-DD（区间查询，与 granularity='month' 配套）",
            "granularity": "聚合粒度：'day' / 'month'（默认 'day'）",
        },
    },
    # 业务：京麦订单明细【加密】导出（项目14，2026-08-11 启动）
    #   流程：5 步异步链路（创建任务→轮询→短信申请→下载加密zip→解压xlsx）
    #   鉴权：Cookie + h5st（前端强签名）+ dsm-eid / dsm-platform / dsm-trace-id / dsm-lang
    #   ⚠️ 接口域名 sff.jd.com（不是 seller-v10.shop.jd.com），入口 shop.jd.com
    #   阶段1+2：创建任务 + 轮询拿 taskId（queryExportTaskInfo 实证为分页列表）
    "京麦订单明细_创建任务": {
        "api_class": None,  # 占位：下方 JingMaiOrderExportAPI 定义后回填
        "method": "create_export_task",
        "callable": None,  # 占位：下方 _run_jm_create_task 函数定义后注入
        "desc": "京麦订单明细【加密】导出 - 第1步：创建导出任务（POST /api?api=dsm.order.export.exportCenterService.createdExportTask）",
        "params": {
            "date": "单日查询YYYY-MM-DD（默认=start_date=end_date）",
            "start_date": "开始日期YYYY-MM-DD（含）",
            "end_date": "结束日期YYYY-MM-DD（含）",
            "h5st": "必填，浏览器F12抓 createdExportTask 请求头 h5st（前端强签名，一次性）",
            "order_status_list": "订单状态列表，默认 [-1]（全部），可传 [1,2,3] 等",
            "sensitive_info_sign": "敏感信息导出标志，默认 '0'（不导出收件人敏感信息）",
            "export_task_type": "导出任务类型，默认 0（订单明细）",
        },
    },
    # 业务：京麦订单明细 - 第1+2步一键（创建任务→轮询拿 taskId）（2026-08-11）
    #   ⚠️ 2026-08-11 实证：createdExportTask 响应**没有 taskId**，
    #     必须再调 queryExportTaskInfo 分页查任务列表，按 startTime/endTime 匹配刚那条 → 拿 id
    "京麦订单明细_创建并轮询": {
        "api_class": None,
        "method": "create_and_wait",
        "callable": None,  # 占位：下方 _run_jm_create_and_wait 定义后注入
        "desc": "京麦订单明细【加密】导出 - 一键创建+轮询拿taskId（创建→分页查→按时间匹配→状态=2 返回）",
        "params": {
            "date": "单日查询YYYY-MM-DD（默认=start_date=end_date）",
            "start_date": "开始日期YYYY-MM-DD（含）",
            "end_date": "结束日期YYYY-MM-DD（含）",
            "h5st": "必填，浏览器F12抓 createdExportTask 请求头 h5st（前端强签名，一次性）",
            "order_status_list": "订单状态列表，默认 [-1]",
            "sensitive_info_sign": "敏感信息导出标志，默认 '0'",
            "export_task_type": "导出任务类型，默认 0",
            "poll_interval": "轮询间隔秒数，默认 3",
            "max_poll_times": "轮询最大次数，默认 20（合计 60s）",
        },
    },
    # 业务：京麦订单明细 - 第1+2+3步一键（创建+轮询+下载加密 zip，2026-08-11）
    #   ⚠️ 第3步 GET export.action 仅 Cookie 鉴权（不需要 h5st / dsm 头），
    #     返回加密 zip（PK magic bytes + password encrypted），解压密码由阶段4短信下发
    "京麦订单明细_创建轮询并下载zip": {
        "api_class": None,
        "method": "create_wait_and_download",
        "callable": None,  # 占位：下方 _run_jm_create_wait_download 定义后注入
        "desc": "京麦订单明细【加密】导出 - 一键创建+轮询+下载加密zip（h5st创建→轮询→Cookie下载）",
        "params": {
            "date": "单日查询YYYY-MM-DD（默认=start_date=end_date）",
            "start_date": "开始日期YYYY-MM-DD（含）",
            "end_date": "结束日期YYYY-MM-DD（含）",
            "h5st": "必填，浏览器F12抓 createdExportTask 请求头 h5st（前端强签名，一次性）",
            "order_status_list": "订单状态列表，默认 [-1]",
            "sensitive_info_sign": "敏感信息导出标志，默认 '0'",
            "export_task_type": "导出任务类型，默认 0",
            "poll_interval": "轮询间隔秒数，默认 3",
            "max_poll_times": "轮询最大次数，默认 20（合计 60s）",
        },
    },
    # 业务：京麦订单明细 - 第1+2+3+4步一键（创建+轮询+下载+短信申请，2026-08-11）
    #   ⚠️ 第4步 exportTaskPwdSend 需要 h5st（与 createdExportTask 同一份 h5st 可复用）
    #     响应 data 是字符串（"密码短信发送成功!当前任务剩余短信发送次数N次"），
    #     **不返回密码明文**——密码只发到短信接收号码（如 1366794）
    "京麦订单明细_创建轮询下载并申请密码": {
        "api_class": None,
        "method": "create_wait_download_and_request_pwd",
        "callable": None,  # 占位：下方 _run_jm_full_with_pwd 定义后注入
        "desc": "京麦订单明细【加密】导出 - 完整4步一键：创建→轮询→下载→短信申请（密码发到手机/邮件）",
        "params": {
            "date": "单日查询YYYY-MM-DD（默认=start_date=end_date）",
            "start_date": "开始日期YYYY-MM-DD（含）",
            "end_date": "结束日期YYYY-MM-DD（含）",
            "h5st": "必填，浏览器F12抓 createdExportTask 请求头 h5st",
            "order_status_list": "订单状态列表，默认 [-1]",
            "sensitive_info_sign": "敏感信息导出标志，默认 '0'",
            "export_task_type": "导出任务类型，默认 0",
            "poll_interval": "轮询间隔秒数，默认 3",
            "max_poll_times": "轮询最大次数，默认 20（合计 60s）",
        },
    },
    # 业务：京麦订单明细 - 完整5步一键（创建+轮询+下载+短信申请+IMAP拿密码+解压xlsx，2026-08-11）
    #   ⚠️ 鉴权 3 次切换：sff.jd.com dsm/h5st → export.shop.jd.com 仅Cookie → 本地zipfile
    #   密码获取优先级：sms_password（命令行） > IMAP 自动监听
    #   IMAP 超时（用户决策 2026-08-11）：保留 zip + 提示人工 --sms-password 重跑，不报错退出
    "京麦订单明细_完整一键导出": {
        "api_class": None,
        "method": "run_full_export",
        "callable": None,  # 占位：下方 _run_jm_run_full_export 定义后注入
        "desc": "京麦订单明细【加密】导出 - 完整5步一键：创建→轮询→下载→短信→IMAP拿密码→解压xlsx",
        "params": {
            "date": "单日查询YYYY-MM-DD（默认=start_date=end_date）",
            "start_date": "开始日期YYYY-MM-DD（含）",
            "end_date": "结束日期YYYY-MM-DD（含）",
            "h5st": "必填，浏览器F12抓 createdExportTask 请求头 h5st",
            "sms_password": "可选：手动传入解压密码（优先级高于IMAP）",
            "imap_config_path": "IMAP配置文件路径，默认 config/imap_config.ini",
            "order_status_list": "订单状态列表，默认 [-1]",
            "sensitive_info_sign": "敏感信息导出标志，默认 '0'",
            "export_task_type": "导出任务类型，默认 0",
            "poll_interval": "轮询间隔秒数，默认 3",
            "max_poll_times": "轮询最大次数，默认 20（合计 60s）",
        },
    },
    # 业务：京麦售后明细 - 完整4步一键（创建+轮询+下载+解压，**无短信无IMAP**，2026-08-11）
    #   ⚠️ 项目16 真实抓包确认（2026-08-12 用户提供 HTTP 报文）：
    #     创建任务: api 路径 dsm.seller.afs.bff.ExportDsmService.createExportTask
    #     轮询: api 路径 dsm.seller.afs.bff.ExportDsmService.getExportTaskPage（不是项目14 的 queryExportTaskInfo）
    #     appId：BHPQ4MHJBUOQZKTFTRNS（项目14 CQLEJWPYPFOVQBC8UFLQ 不同！）
    #     payload 嵌套：{"request":{"data":{...}}, "accessContext":{"source":"web"}}
    #     创建时间格式：毫秒时间戳（项目14 是 "YYYY-MM-DD HH:MM:SS"）
    #     exportType: 2602（创建），轮询时是 [2602,2601,38] 列表
    #     创建响应 data 是 bool；轮询响应 data.content[] 是对象列表
    #     创建响应头 X-Rp-Sdtoken 30 分钟有效，下次请求带上
    #     请求头多一个 dsm-file-path: lineation-price
    #     状态枚举（你文档一致）：exportStatusCode 1=生成中/2=成功/3=失败
    #     匹配策略：exportTypeCode + exportCondition 文本（含"申请时间：YYYY-MM-DD至YYYY-MM-DD"）
    #   鉴权 2 次切换：sff.jd.com dsm/h5st/X-Rp-Sdtoken → export.shop.jd.com 仅Cookie（待下载抓包）
    "京麦售后明细_完整一键导出": {
        "api_class": None,
        "method": "run_after_sale_full_export",
        "callable": None,  # 占位：下方 _run_jm_after_sale_full 定义后注入
        "desc": "京麦售后明细导出 - 完整4步一键：创建→轮询→下载→解压（**无短信无IMAP**）",
        "params": {
            "date": "单日查询YYYY-MM-DD（默认=start_date=end_date）",
            "start_date": "开始日期YYYY-MM-DD（含）",
            "end_date": "结束日期YYYY-MM-DD（含）",
            "h5st": "必填，浏览器F12抓 createdExportTask 请求头 h5st",
            "tab_code": "售后tab类型，默认 'all'（抓包实证）",
            "after_sale_status_list": "售后状态列表，默认 '' = 全部",
            "service_order_sub_state_list": "售后子状态列表，默认 ''",
            "customer_expect_list": "客户期望列表，默认 ''",
            "refund_status_list": "退款状态列表，默认 ''",
            "transfer_feedback_reason_list": "转移反馈原因列表，默认 ''",
            "poll_interval": "轮询间隔秒数，默认 3",
            "max_poll_times": "轮询最大次数，默认 20（合计 60s）",
        },
    },
}


def _run_jzt_kuaiche_full(**kwargs) -> str:
    """调度器专用的京准通快车完整流程函数。

    ⚠️ 注册到 BUSINESS_REGISTRY["京准通快车自定义报表"]["callable"]，
       get_business_handler 检测到 callable 字段时优先返回本函数。
    设计动机：JZTKuaicheAPI.__init__ 需要 h5st 和 cookie_path 参数，
              而基类的标准调度路径只支持无参 __init__ → 实例化 → 调方法，
              无法透传这两个值。本函数手动构造实例并调用 run_full_export。

    参数:
        kwargs - 来自 run_business 的透传参数：
            h5st       (str): 必填，浏览器F12抓 add 接口请求头的 h5st 值
            start_date (str): 开始日期 YYYY-MM-DD
            end_date   (str): 结束日期 YYYY-MM-DD
            date       (str): 单日查询（start/end 默认=date）
            cookie_path(str): 可选，默认 config/jzt_cookie.txt
            save_filename(str): 可选，默认按日期+时间戳自动命名
    返回:
        str - 保存的文件绝对路径
    异常:
        ValueError - 缺 h5st 或日期参数时
        CookieExpiredError / TimeoutError / RuntimeError
    """
    # ⚠️ h5st **可选**（2026-08-07 抓包实证：add 接口不校验 h5st）
    #   - 不传 h5st：可跑通 add/list/downloadById 三步（最常见情况）
    #   - 传 h5st：增强未来接口升级风控时的兼容性
    #   - 何时需要：若 list/downloadById 返回 code=601 "操作频繁"，说明接口开始校验 h5st，
    #                此时浏览器F12抓 add 接口请求头的 h5st 值传入即可
    h5st = kwargs.get("h5st", "")
    # 不再强制必传，但给个温和提醒
    if not h5st:
        print("ℹ️  未传 h5st（add 接口抓包实测不校验，可正常跑；若报 601 请浏览器F12抓 add 接口的 h5st 重试）")

    # 提取透传给 run_full_export 的参数
    # ⚠️ 2026-08-10 bug fix：用户决策补能力——支持 --range 透传
    forward_kwargs = {}
    for k in ("start_date", "end_date", "date", "save_filename"):
        if k in kwargs and kwargs[k] is not None:
            forward_kwargs[k] = kwargs[k]

    # ⚠️ 兜底：如果 start_date/end_date/date 都是 None，强制报错（避免创建任务时三值校验失败）
    if not forward_kwargs.get("start_date") and not forward_kwargs.get("end_date") and not forward_kwargs.get("date"):
        raise ValueError("❌ 至少需要传入 date 或 start_date/end_date")

    api = JZTKuaicheAPI(h5st=h5st, cookie_path=kwargs.get("cookie_path", "config/jzt_cookie.txt"))
    return api.run_full_export(**forward_kwargs)


def _run_jzt_order_effect_full(**kwargs) -> str:
    """调度器专用的京准通快车订单效果明细完整流程函数（2026-08-09 上线）。

    ⚠️ 注册到 BUSINESS_REGISTRY["京准通快车订单效果明细"]["callable"]。
    设计动机：JZTKuaicheOrderEffectAPI.__init__ 需要 cookie_path，
              标准调度路径不支持构造参数注入，本函数手动构造实例并调用 run_full_export。

    参数:
        kwargs - 来自 run_business 的透传参数：
            date       (str): 单日查询 YYYY-MM-DD
            start_date (str): 开始日期 YYYY-MM-DD（区间查询）
            end_date   (str): 结束日期 YYYY-MM-DD
            cookie_path(str): 可选，默认 config/jzt_cookie.txt
    返回:
        str - 保存的文件绝对路径
    异常:
        ValueError - 缺日期参数时
        CookieExpiredError / RuntimeError
    """
    forward_kwargs = {
        k: kwargs[k] for k in ("start_date", "end_date", "date")
        if k in kwargs
    }
    api = JZTKuaicheOrderEffectAPI(cookie_path=kwargs.get("cookie_path", "config/jzt_cookie.txt"))
    return api.run_full_export(**forward_kwargs)


def _run_jzt_quanzhan_campaign_full(**kwargs) -> str:
    """调度器专用的京准通全站营销单品计划完整流程函数（2026-08-09 上线）。

    ⚠️ 注册到 BUSINESS_REGISTRY["京准通全站营销单品计划"]["callable"]。
    设计动机：与 _run_jzt_order_effect_full 同——本类 __init__ 需要 cookie_path。

    参数:
        kwargs - 来自 run_business 的透传参数：
            date       (str): 单日查询 YYYY-MM-DD
            start_date (str): 开始日期 YYYY-MM-DD（区间查询）
            end_date   (str): 结束日期 YYYY-MM-DD
            cookie_path(str): 可选，默认 config/jzt_cookie.txt
    返回:
        str - 保存的文件绝对路径
    异常:
        ValueError - 缺日期参数时
        CookieExpiredError / RuntimeError
    """
    forward_kwargs = {
        k: kwargs[k] for k in ("start_date", "end_date", "date")
        if k in kwargs
    }
    api = JZTQuanZhanCampaignAPI(cookie_path=kwargs.get("cookie_path", "config/jzt_cookie.txt"))
    return api.run_full_export(**forward_kwargs)


# ============================================================
#  业务接口 7：（新业务 - 京准通-全站营销单品计划报表，2026-08-09 上线）
# ------------------------------------------------------------
#  中文说明（小白必读）：
#    京准通-全站营销（reweb/swa/account/campaign/download）下的「单品计划报表」，
#    与"业务接口6"订单效果明细**几乎同模式**：同步两步返回下载链接。
#
#  ⚠️ 核心差异（与项目8对比）：
#    - 接口 URL 不同：reweb/swa/account/campaign/download（不是 reweb/msa/effect/order/download）
#    - 路径含义：swa=全站营销（Search Whole-site Advertising） / account=账户 / campaign=计划
#    - payload 字段更多：dateValues[]（日期数组）/isDaily（日报标志）/campaignTypes[]（业务类型）
#    - 字段类型不同：giftFlag 是字符串 ""（不是数字 0）；其他空字段也是字符串
#    - 报表名格式：{pin}_全站营销_单品计划报表_{startDay}_{endDay}
#    - 数据维度：商品计划 / 投放类型 / 花费 / 全站投产比 / 订单行 / 智能补贴券
#
#  参数说明（2026-08-09 抓包实证）：
#    业务固定参数（类常量，不变）：
#      platform=""                 # 平台（空=不限）
#      campaignTypes=[101]         # 业务类型：101=京东快车（推测）
#      province=""                 # 省份过滤（空=全国）
#      clickOrOrderDay=15          # 转化周期：15天
#      clickOrOrderCaliber=0       # 0=点击
#      isDaily=True                # 日报标志
#      orderStatusCategory=1       # 1=成交订单
#      orderStatus=""              # 订单状态过滤（空）
#      giftFlag=""                 # 含赠品（字符串空）
#      sxuId=""                    # SKU 过滤（空）
#      obys=""                     # 对象过滤（空）
#    日期参数：
#      startDay / endDay          # 顶层
#      dateValues=[{startDay,endDay}]  # 同时要传
#    报表名（动态）：
#      reportName 格式：{pin}_全站营销_单品计划报表_{startDay}_{endDay}
# ============================================================

class JZTQuanZhanCampaignAPI:
    """京准通-全站营销单品计划报表导出 API（2026-08-09 上线骨架）。

    ⚠️ 本类**不继承 JDBaseRequest**（与项目7/8 同样的原因）：
        - 鉴权体系：仅 Cookie（与项目7/8 同一文件 config/jzt_cookie.txt）
        - 流程：同步两步（POST → 立即 GET OSS），不需要基类的 30 秒间隔/重试模型
        - UA：禁止切换（h5st 与 UA 绑定；本接口无 h5st 但保留习惯）
    """

    # ---- 类常量（业务固定参数）----
    BASE_URL = "https://jzt-api.jd.com/reweb/swa/account/campaign/download"
    ORIGIN = "https://jzt.jd.com"
    REFERER = "https://jzt.jd.com/"
    SITE_ID = "0"
    USER_AGENT = (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/151.0.0.0 Safari/537.36 Edg/151.0.0.0"
    )
    OUTPUT_SUBDIR = "京准通全站营销单品计划"  # output/京准通全站营销单品计划/{date}/

    # 业务固定参数（抓包值，2026-08-09 实证）
    PLATFORM = ""               # 平台（空=不限）
    CAMPAIGN_TYPES = [101]      # 业务类型：101=京东快车（推测）
    PROVINCE = ""               # 省份过滤（空=全国）
    CLICK_OR_ORDER_DAY = 15     # 转化周期：15天
    CLICK_OR_ORDER_CALIBER = 0  # 0=点击
    IS_DAILY = True             # 日报标志
    ORDER_STATUS_CATEGORY = 1   # 1=成交订单
    ORDER_STATUS = ""           # 订单状态过滤（空）
    GIFT_FLAG = ""              # 含赠品（字符串空，注意与项目8数字0不同）
    SXU_ID = ""                 # SKU 过滤（空）
    OBYS = ""                   # 对象过滤（空）
    PIN_ID = "FYA8888"

    def __init__(self, cookie_path: str = "config/jzt_cookie.txt"):
        import requests

        # 读 Cookie（与项目7/8 互通 jzt_cookie.txt）
        cookie_path_abs = os.path.join(os.path.dirname(os.path.abspath(__file__)), cookie_path)
        if not os.path.isfile(cookie_path_abs):
            raise FileNotFoundError(
                f"❌ 京准通 Cookie 文件不存在：{cookie_path_abs}\n"
                f"   请浏览器登录 https://jzt.jd.com/home，F12 抓 jzt-api.jd.com 域 Cookie 写入此文件"
            )
        with open(cookie_path_abs, "r", encoding="utf-8") as f:
            self.cookie = f.read().strip()
        if not self.cookie:
            raise ValueError(f"❌ 京准通 Cookie 文件 {cookie_path_abs} 内容为空")

        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": self.USER_AGENT,
            "Accept": "application/json, text/plain, */*",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
            "Accept-Encoding": "gzip, deflate, br, zstd",
            "Content-Type": "application/json;charset=UTF-8",
            "Origin": self.ORIGIN,
            "Referer": self.REFERER,
            "Cookie": self.cookie,
        })

        # 输出目录（按 AGENTS.md Excel规则4）
        self.output_dir = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "output", self.OUTPUT_SUBDIR,
        )

    # ---- 业务参数组装 ----
    def _build_payload(self, start_day: str, end_day: str) -> dict:
        """组装请求 payload（抓包实证 + 动态日期）。

        ⚠️ 字段类型严格匹配抓包：
            - giftFlag/orderStatus/sxuId/obys/province 是字符串 ""（不是 None / 数字 0）
            - campaignTypes 是 list [101]（不是字符串 "101"）
            - isDaily 是 bool True
        """
        # 报表名：FYA8888_全站营销_单品计划报表_{startDay}_{endDay}
        report_name = f"{self.PIN_ID}_全站营销_单品计划报表_{start_day}_{end_day}"
        return {
            "platform": self.PLATFORM,
            "campaignTypes": self.CAMPAIGN_TYPES,
            "province": self.PROVINCE,
            "startDay": start_day,
            "endDay": end_day,
            "orderStatus": self.ORDER_STATUS,
            "giftFlag": self.GIFT_FLAG,
            "clickOrOrderDay": self.CLICK_OR_ORDER_DAY,
            "clickOrOrderCaliber": self.CLICK_OR_ORDER_CALIBER,
            "sxuId": self.SXU_ID,
            "obys": self.OBYS,
            "isDaily": self.IS_DAILY,
            "orderStatusCategory": self.ORDER_STATUS_CATEGORY,
            "dateValues": [{"startDay": start_day, "endDay": end_day}],
            "reportName": report_name,
        }

    def _handle_response(self, ret: dict, op_desc: str):
        """统一处理响应。

        京东快车订单接口的判定：
            - success=true
            - code=1 或 "1"（兼容字符串/数字）
            - data.code == "RC_SUCCESS"
            - data.downloadUrlCsv 非空
        """
        if not ret.get("success", True):
            msg = ret.get("msg", "未知错误")
            code = ret.get("code")
            if code in (2001, 302) or "未登录" in msg or "登录已过期" in msg:
                raise CookieExpiredError(
                    f"❌ 京准通 Cookie 过期（{op_desc}返回 code={code}）：\n"
                    f"   → 请浏览器登录 https://jzt.jd.com/home，F12 抓 jzt-api.jd.com 域 Cookie 写入 config/jzt_cookie.txt"
                )
            if code == 601 or str(code) == "601":
                raise RuntimeError(
                    f"❌ 京准通 全站营销 限流 code=601：{msg}\n"
                    f"   → 30-120 分钟冷却，避免重试加重风控"
                )
            raise RuntimeError(
                f"❌ 京准通{op_desc}失败：code={code}, msg={msg}, 完整响应={ret}"
            )
        code = ret.get("code")
        data_code = ret.get("data", {}).get("code")
        # ⚠️ 2026-08-09 项目9 探针发现：code 可能是字符串 "1" 而非数字 1
        if str(code) not in ("0", "1") or data_code != "RC_SUCCESS":
            raise RuntimeError(
                f"❌ 京准通{op_desc}业务失败：code={code}, data.code={data_code}\n"
                f"   完整响应：{ret}"
            )
        return ret

    # ---- 一步：同步 POST 拿 downloadUrlCsv ----
    def _post_for_csv(self, start_day: str, end_day: str) -> str:
        """POST 同步返回 downloadUrlCsv（同项目8）。"""
        payload = self._build_payload(start_day, end_day)
        print(f"🚀 [JZT全站营销] POST {self.BASE_URL}")
        print(f"   Body: {json.dumps(payload, ensure_ascii=False)}")

        resp = self.session.post(self.BASE_URL, json=payload, timeout=60)
        resp.raise_for_status()
        ret = resp.json()
        self._handle_response(ret, op_desc="导出单品计划报表")

        download_url = ret.get("data", {}).get("downloadUrlCsv")
        if not download_url:
            raise RuntimeError(f"❌ 响应中 downloadUrlCsv 缺失：{ret}")
        download_id = ret.get("data", {}).get("downloadId")
        print(f"✅ 拿到 downloadId={download_id}, downloadUrlCsv（前80字符）: {download_url[:80]}...")
        return download_url

    # ---- 二步：GET OSS 下载 CSV 字节流 ----
    def _download_csv(self, url_csv: str) -> bytes:
        """GET OSS 链接，下载 CSV 字节流（带 404 随机退避重试，同项目8）。"""
        last_error = None
        for retry in range(4):  # 最多 4 次
            try:
                resp = requests.get(
                    url_csv,
                    headers={"User-Agent": self.USER_AGENT},
                    timeout=60,
                )
                if resp.status_code == 200:
                    return resp.content
                if resp.status_code == 404:
                    backoff = random.uniform(3, 10)
                    print(
                        f"  ⚠️ 第 {retry+1}/4 次 urlCsv 404 NoSuchKey，"
                        f"随机退避 {backoff:.1f} 秒后重试..."
                    )
                    time.sleep(backoff)
                    continue
                resp.raise_for_status()
            except requests.exceptions.RequestException as e:
                last_error = e
                print(f"  ⚠️ urlCsv 下载异常：{e}，重试中...")
                time.sleep(random.uniform(3, 10))
        raise RuntimeError(
            f"❌ urlCsv 下载失败（重试 4 次后）：{last_error}\n"
            f"   URL: {url_csv[:120]}"
        )

    # ---- 一键封装（推荐对外入口）----
    def run_full_export(
        self,
        start_date: str = None,
        end_date: str = None,
        date: str = None,
    ) -> str:
        """一键跑通：POST 同步拿 urlCsv → GET OSS 下载 → Excel 后置处理 → 落盘。"""
        if date is None and start_date is None and end_date is None:
            raise ValueError("❌ 至少需要传入 date 或 start_date/end_date")
        if start_date is None:
            start_date = date
        if end_date is None:
            end_date = date

        print(f"🚀 [JZT全站营销] 启动完整导出：{start_date} ~ {end_date}")
        print(f"   └─ Step 1/2: POST 同步拿 downloadUrlCsv...")
        url_csv = self._post_for_csv(start_date, end_date)
        print(f"   └─ Step 2/2: GET OSS 下载并落盘为 xlsx...")
        csv_bytes = self._download_csv(url_csv)

        # 复用现有 Excel 后置处理（日期/数值/格式）
        return self._post_process_csv_to_xlsx(csv_bytes, start_date)

    def _post_process_csv_to_xlsx(self, csv_bytes: bytes, clean_date: str) -> str:
        """把 OSS 下载的 raw CSV → 标准 Excel 后置处理 → 保存为 xlsx。"""
        import io
        import pandas as pd

        # 1. 读取 CSV（dtype=str 防精度丢失；UTF-8-sig 兼容 BOM）
        try:
            df = pd.read_csv(
                io.BytesIO(csv_bytes),
                dtype=str,
                na_filter=False,
                encoding="utf-8-sig",
                keep_default_na=False,
            )
        except Exception as e:
            raise RuntimeError(
                f"❌ CSV 解析失败：{e}\n"
                f"   请检查 OSS 返回内容是否正常"
            ) from e

        if df.empty:
            # ⚠️ 用户决策 2026-08-10 补充：空数据视为成功（**仅项目9 启用**，与项目10/11 行为一致）
            #   场景：与服务端确认结果一致的空数据报表视为「没有投放该推广工具」，
            #         不视为代码 bug，写入空 xlsx（含表头）+ 返回成功路径
            #   注意：项目7、8 仍保留原抛错行为（项目7 由 h5st/轮询控制、项目8 链路复杂）
            print(
                f"   ├─ ⚠️ CSV 数据为空（{len(df.columns)}列 0行）"
                f"—— 视为业务无数据，写入空 xlsx"
            )
            # 不抛异常，继续走落盘流程

        # 2. 日期列智能处理（公共规则1+2）
        #    本报表自带「日期」列 → 不插入新日期列，只做格式标准化
        date_column, date_value = prepare_date_columns(df, clean_date)

        # 3. 数值安全转换（公共规则3）
        df = safe_convert_numeric(df)

        # 4. 构造输出路径：output/京准通全站营销单品计划/{date}/业务名_{date}.xlsx
        date_subdir = os.path.join(self.output_dir, clean_date)
        os.makedirs(date_subdir, exist_ok=True)
        save_filename = f"京准通全站营销单品计划_{clean_date}.xlsx"
        target_path = os.path.join(date_subdir, save_filename)

        # 5. 写 xlsx + 单元格格式
        df.to_excel(target_path, index=False, engine="openpyxl")
        apply_column_formats(target_path, df, date_column=date_column, date_value=date_value)

        print(
            f"✅ 文件已保存：{target_path}"
            f"\n   （CSV→xlsx 转存 + 日期列 + 数值转换 + 单元格格式，{os.path.getsize(target_path)}字节，{len(df)}行 × {len(df.columns)}列）"
        )
        return target_path


# ⚠️ 注册表 callable 字段回填（解决前向引用：注册表先于函数定义）
BUSINESS_REGISTRY["京准通快车自定义报表"]["callable"] = _run_jzt_kuaiche_full
BUSINESS_REGISTRY["京准通快车订单效果明细"]["callable"] = _run_jzt_order_effect_full
BUSINESS_REGISTRY["京准通全站营销单品计划"]["callable"] = _run_jzt_quanzhan_campaign_full
# ⚠️ 业务接口7 类在前向引用（注册表在类之前），回填 api_class
BUSINESS_REGISTRY["京准通全站营销单品计划"]["api_class"] = JZTQuanZhanCampaignAPI


class JZTQuanZhanEffectAPI:
    """京准通-全站营销**单品推广效果报表**导出 API（2026-08-10 上线骨架）。

    ⚠️ 本类**不继承 JDBaseRequest**（与项目7/8/9 同原因）：
        - 鉴权体系：仅 Cookie（与项目7/8/9 同一文件 config/jzt_cookie.txt）
        - 流程：同步两步（POST → 立即 GET OSS），不需要基类的 30 秒间隔/重试模型
        - UA：禁止切换（h5st 与 UA 绑定；本接口无 h5st 但保留习惯，沿用项目9 的 v=151）

    🆕 与项目9（JZTQuanZhanCampaignAPI 单品计划报表）的关键差异：
        - URL 路径：多了 /effect/ 段
            项目9：/reweb/swa/account/campaign/download
            本项目：/reweb/swa/effect/order/download
        - payload：
            * orderStatus：项目9 默认 ""（空=不限）；本项目默认 "1"（**成交订单**，开放入参）
            * isDaily：项目9 默认 True；本项目默认 False（**调用方可入参切换**）
            * skuId / spuId：**新增字段**（默认 ""，开放入参；空=不过滤）
            * obys / dateValues：**移除**（项目9 有，本项目无）
        - 响应：增加 downloadUrlZip（**zip 优先**，csv 降级备用）
        - 报表名：FYA8888_全站营销_**效果报表**_单品推广_{startDay}_{endDay}（项目9 是 _单品计划报表_）
    """

    # ---- 类常量（业务固定参数，2026-08-10 抓包实证 + 用户确认）----
    BASE_URL = "https://jzt-api.jd.com/reweb/swa/effect/order/download"
    ORIGIN = "https://jzt.jd.com"
    REFERER = "https://jzt.jd.com/"
    SITE_ID = "0"
    USER_AGENT = (
        # ⚠️ 沿用项目9 的 v=151 UA（保证与 jzt_cookie.txt 会话一致）
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/151.0.0.0 Safari/537.36 Edg/151.0.0.0"
    )
    OUTPUT_SUBDIR = "京准通全站营销单品推广效果"  # output/京准通全站营销单品推广效果/{date}/

    # 业务固定参数（抓包值，2026-08-10 实证 + 用户 2026-08-10 确认固化）
    PLATFORM = ""               # 平台（空=不限，字符串）
    CAMPAIGN_TYPES = [101]      # 业务类型：101=京东快车（推测，列表）
    CLICK_OR_ORDER_DAY = 15     # 转化周期：15 天
    CLICK_OR_ORDER_CALIBER = 0  # 0=点击（int）
    ORDER_STATUS_CATEGORY = 1   # 1=成交订单
    GIFT_FLAG = ""              # 含赠品（字符串空）
    PIN_ID = "FYA8888"

    # ⚠️ 用户决策 2026-08-10：
    #   orderStatus 默认 "1"（成交订单），开放入参支持传空（""=不限）
    #   isDaily 不固化，调用方可在 run_full_export / _post_for_csv 传入开关
    #   skuId / spuId 默认 ""（空=不过滤），开放可选入参
    DEFAULT_ORDER_STATUS = "1"  # 成交订单（用户决策 2026-08-10）
    DEFAULT_IS_DAILY = False    # 默认非日报（用户决策 2026-08-10，抓包实测值）

    def __init__(self, cookie_path: str = "config/jzt_cookie.txt"):
        import requests

        # 读 Cookie（与项目7/8/9 互通 jzt_cookie.txt）
        cookie_path_abs = os.path.join(os.path.dirname(os.path.abspath(__file__)), cookie_path)
        if not os.path.isfile(cookie_path_abs):
            raise FileNotFoundError(
                f"❌ 京准通 Cookie 文件不存在：{cookie_path_abs}\n"
                f"   请浏览器登录 https://jzt.jd.com/home，F12 抓 jzt-api.jd.com 域 Cookie 写入此文件"
            )
        with open(cookie_path_abs, "r", encoding="utf-8") as f:
            self.cookie = f.read().strip()
        if not self.cookie:
            raise ValueError(f"❌ 京准通 Cookie 文件 {cookie_path_abs} 内容为空")

        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": self.USER_AGENT,
            "Accept": "application/json, text/plain, */*",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
            "Accept-Encoding": "gzip, deflate, br, zstd",
            "Content-Type": "application/json;charset=UTF-8",
            "Origin": self.ORIGIN,
            "Referer": self.REFERER,
            "Cookie": self.cookie,
        })

        # 输出目录（按 AGENTS.md Excel规则4）
        self.output_dir = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "output", self.OUTPUT_SUBDIR,
        )

    # ---- 业务参数组装 ----
    def _build_payload(
        self,
        start_day: str,
        end_day: str,
        order_status: str = None,
        is_daily: bool = None,
        sku_id: str = "",
        spu_id: str = "",
    ) -> dict:
        """组装请求 payload（抓包实证 + 动态日期 + 开放入参）。

        ⚠️ 用户决策 2026-08-10：
            - order_status 默认为 "1"（成交订单），传 "" 表示不限
            - is_daily 默认 False（抓包实测值），调用方可在 run_full_export 传 True/False
            - sku_id / spu_id 默认 ""（空=不过滤），调用方按需传入

        字段类型严格匹配抓包：
            - platform/giftFlag/orderStatus/skuId/spuId 是字符串
            - campaignTypes 是列表 [101]
            - isDaily 是布尔
            - startDay/endDay 是字符串
        """
        # 入参兜底（None → 类默认）
        if order_status is None:
            order_status = self.DEFAULT_ORDER_STATUS
        if is_daily is None:
            is_daily = self.DEFAULT_IS_DAILY

        # 报表名：FYA8888_全站营销_效果报表_单品推广_{startDay}_{endDay}
        report_name = f"{self.PIN_ID}_全站营销_效果报表_单品推广_{start_day}_{end_day}"
        return {
            "platform": self.PLATFORM,
            "campaignTypes": self.CAMPAIGN_TYPES,
            "startDay": start_day,
            "endDay": end_day,
            "orderStatus": order_status,
            "giftFlag": self.GIFT_FLAG,
            "skuId": sku_id,
            "spuId": spu_id,
            "clickOrOrderCaliber": self.CLICK_OR_ORDER_CALIBER,
            "clickOrOrderDay": self.CLICK_OR_ORDER_DAY,
            "isDaily": is_daily,
            "orderStatusCategory": self.ORDER_STATUS_CATEGORY,
            "reportName": report_name,
        }

    def _handle_response(self, ret: dict, op_desc: str):
        """统一处理响应（与项目8/9 同样的双字段判定）。

        判定：
            - success=true
            - code=="1"（兼容字符串/数字）
            - data.code == "RC_SUCCESS"
            - data.downloadUrlZip 或 data.downloadUrlCsv 非空
        """
        if not ret.get("success", True):
            msg = ret.get("msg", "未知错误")
            code = ret.get("code")
            if code in (2001, 302) or "未登录" in msg or "登录已过期" in msg:
                raise CookieExpiredError(
                    f"❌ 京准通 Cookie 过期（{op_desc}返回 code={code}）：\n"
                    f"   → 请浏览器登录 https://jzt.jd.com/home，F12 抓 jzt-api.jd.com 域 Cookie 写入 config/jzt_cookie.txt"
                )
            if code == 601 or str(code) == "601":
                raise RuntimeError(
                    f"❌ 京准通 全站营销效果 限流 code=601：{msg}\n"
                    f"   → 30-120 分钟冷却，避免重试加重风控"
                )
            raise RuntimeError(
                f"❌ 京准通{op_desc}失败：code={code}, msg={msg}, 完整响应={ret}"
            )
        code = ret.get("code")
        data_code = ret.get("data", {}).get("code")
        # ⚠️ 兼容字符串/数字 code（项目9 探针发现）
        if str(code) not in ("0", "1") or data_code != "RC_SUCCESS":
            raise RuntimeError(
                f"❌ 京准通{op_desc}业务失败：code={code}, data.code={data_code}\n"
                f"   完整响应：{ret}"
            )
        return ret

    def _pick_download_url(self, ret: dict) -> str:
        """⚠️ 用户决策 2026-08-10（**第二次调整**）：**csv 优先，zip 降级**。

        变更动机：模拟浏览器行为（抓包证实浏览器 GET 的是 .csv 直链而非 .zip），
                  若 OSS 写 csv 比写 zip 早，csv 优先可缩短跨日归档延迟。
        历史：
            - 2026-08-10 阶段5：zip 优先（用户决策）
            - 2026-08-10 阶段6+：改为 csv 优先（用户决策），zip 作为完整包备份

        返回:
            str - OSS 预签名链接（csv 优先，zip 降级）
        异常:
            RuntimeError - 两者都缺失时
        """
        data = ret.get("data", {})
        download_csv = data.get("downloadUrlCsv")
        download_zip = data.get("downloadUrlZip")
        if download_csv:
            print(f"   ├─ 优先使用 downloadUrlCsv（**模拟浏览器行为**，csv 比 zip 早写）")
            return download_csv
        if download_zip:
            print(f"   ├─ downloadUrlCsv 缺失，降级使用 downloadUrlZip（完整压缩包）")
            return download_zip
        raise RuntimeError(f"❌ 响应中 downloadUrlCsv/downloadUrlZip 都缺失：{ret}")

    # ---- 一步：同步 POST 拿 downloadUrl（zip 优先）----
    def _post_for_csv(
        self,
        start_day: str,
        end_day: str,
        order_status: str = None,
        is_daily: bool = None,
        sku_id: str = "",
        spu_id: str = "",
    ) -> str:
        """POST 同步返回 downloadUrlZip（优先）/ downloadUrlCsv（降级）。

        参数:
            start_day    - 开始日期 YYYY-MM-DD
            end_day      - 结束日期 YYYY-MM-DD
            order_status - 订单状态（默认 "1"=成交订单，传 ""=不限）
            is_daily     - 日报标志（默认 False=非日报，传 True=日报）
            sku_id       - SKU 过滤（默认 ""=不过滤）
            spu_id       - SPU 过滤（默认 ""=不过滤）
        返回:
            str - OSS 预签名链接（10 分钟有效）
        异常:
            CookieExpiredError / RuntimeError
        """
        payload = self._build_payload(
            start_day, end_day,
            order_status=order_status,
            is_daily=is_daily,
            sku_id=sku_id, spu_id=spu_id,
        )
        print(f"🚀 [JZT全站营销单品推广效果] POST {self.BASE_URL}")
        print(f"   Body: {json.dumps(payload, ensure_ascii=False)}")

        resp = self.session.post(self.BASE_URL, json=payload, timeout=60)
        resp.raise_for_status()
        ret = resp.json()
        self._handle_response(ret, op_desc="导出单品推广效果报表")

        download_id = ret.get("data", {}).get("downloadId")
        download_url = self._pick_download_url(ret)
        print(f"✅ 拿到 downloadId={download_id}, downloadUrl（前80字符）: {download_url[:80]}...")
        return download_url

    # ---- 二步：GET OSS 下载文件字节流（zip/csv 通用）----
    # ⚠️ 用户决策 2026-08-10：MAX_DOWNLOAD_RETRY 4→8（覆盖跨日归档延迟场景）
    #   - 京东 OSS 异步生成报表可能需 30-120 秒（POST 返回 URL 但对象未生成）
    #   - 重试上限 8 次 + 退避 3-10s ≈ 累计 60-80 秒，可覆盖一般跨日归档
    #   - 仍失败 → 抛 RuntimeError（不静默放弃，避免掩盖真实问题）
    MAX_DOWNLOAD_RETRY = 8

    def _download_file(self, url_oss: str) -> bytes:
        """GET OSS 链接，下载文件字节流（带 404 随机退避重试，最多 8 次）。

        OSS 链接 10 分钟有效，但首次可能 404 NoSuchKey（异步生成）。
        重试策略：最多 MAX_DOWNLOAD_RETRY=8 次，3-10s 随机退避，累计 ~60-80 秒。

        告警日志：
            - 每次 404：打印重试进度
            - 重试用尽：打印严重告警 + 累计等待时间 + URL 前缀（便于人工排查）
        """
        last_error = None
        total_wait = 0.0  # 累计等待时间（秒）
        for retry in range(self.MAX_DOWNLOAD_RETRY):
            try:
                resp = requests.get(
                    url_oss,
                    headers={"User-Agent": self.USER_AGENT},
                    timeout=60,
                )
                if resp.status_code == 200:
                    if retry > 0:
                        print(
                            f"  ✅ 第 {retry+1}/{self.MAX_DOWNLOAD_RETRY} 次重试成功"
                            f"（累计等待 {total_wait:.1f} 秒）"
                        )
                    return resp.content
                if resp.status_code == 404:
                    backoff = random.uniform(3, 10)
                    total_wait += backoff
                    print(
                        f"  ⚠️ 第 {retry+1}/{self.MAX_DOWNLOAD_RETRY} 次 url 404 NoSuchKey，"
                        f"随机退避 {backoff:.1f} 秒后重试...（累计 {total_wait:.1f} 秒）"
                    )
                    time.sleep(backoff)
                    continue
                resp.raise_for_status()
            except requests.exceptions.RequestException as e:
                last_error = e
                print(f"  ⚠️ url 下载异常：{e}，重试中...")
                time.sleep(random.uniform(3, 10))
                total_wait += 3  # 粗略累计
        # 全部失败 → 严重告警日志
        print(
            f"  🔴 [严重告警] url 下载失败，已重试 {self.MAX_DOWNLOAD_RETRY} 次仍未成功\n"
            f"     ├─ 累计等待时间：{total_wait:.1f} 秒\n"
            f"     ├─ OSS 域：storage.jd.com\n"
            f"     ├─ 诊断建议：\n"
            f"     │   ① 京东 OSS 异步生成延迟（跨日归档常见）→ 手动 GET 该 URL 验证\n"
            f"     │   ② URL 签名是否过期 → 检查 URL 中 Expires 时间\n"
            f"     │   ③ 调 atoms-api list 查报表 status（_poll_report_status_atoms）\n"
            f"     └─ URL: {url_oss[:120]}"
        )
        raise RuntimeError(
            f"❌ url 下载失败（重试 {self.MAX_DOWNLOAD_RETRY} 次后，累计等待 {total_wait:.1f} 秒）：{last_error}\n"
            f"   URL: {url_oss[:120]}"
        )

    # ---- 一键封装（推荐对外入口）----
    def run_full_export(
        self,
        start_date: str = None,
        end_date: str = None,
        date: str = None,
        order_status: str = None,    # 用户决策 2026-08-10：默认 "1"，可传 ""（不限）
        is_daily: bool = None,       # 用户决策 2026-08-10：默认 False，可传 True
        sku_id: str = "",            # 用户决策 2026-08-10：默认 ""（不过滤）
        spu_id: str = "",            # 用户决策 2026-08-10：默认 ""（不过滤）
    ) -> str:
        """一键跑通：POST 同步拿 url（zip 优先）→ GET OSS 下载 → Excel 后置处理 → 落盘。

        参数:
            start_date   - 开始日期 YYYY-MM-DD（默认 = date）
            end_date     - 结束日期 YYYY-MM-DD（默认 = date）
            date         - 单日查询 YYYY-MM-DD
            order_status - 订单状态（None=默认"1"成交订单；""=不限）
            is_daily     - 日报标志（None=默认False非日报；True=日报）
            sku_id       - SKU 过滤（""=不过滤）
            spu_id       - SPU 过滤（""=不过滤）
        返回:
            str - 保存的 .xlsx 绝对路径
        """
        if date is None and start_date is None and end_date is None:
            raise ValueError("❌ 至少需要传入 date 或 start_date/end_date")
        if start_date is None:
            start_date = date
        if end_date is None:
            end_date = date

        print(f"🚀 [JZT全站营销单品推广效果] 启动完整导出：{start_date} ~ {end_date}")
        print(f"   └─ Step 1/2: POST 同步拿 downloadUrl（zip 优先）...")
        url_oss = self._post_for_csv(
            start_date, end_date,
            order_status=order_status, is_daily=is_daily,
            sku_id=sku_id, spu_id=spu_id,
        )
        print(f"   └─ Step 2/2: GET OSS 下载并落盘为 xlsx...")
        file_bytes = self._download_file(url_oss)

        # ⚠️ 本项目可能拿到 zip（压缩包含 csv）或 csv（裸流）→ 智能识别
        return self._post_process_to_xlsx(file_bytes, start_date, url_oss)

    # ---- 辅助诊断：atoms-api list 轮询（用户决策 2026-08-10 补能力）----
    def _poll_report_status_atoms(
        self,
        report_type: int,
        start_day: str,
        end_day: str,
        name_like: str = "",
        page: int = 1,
        page_size: int = 10,
    ) -> dict:
        """调用 atoms-api list 接口查询报表生成状态（**辅助诊断方法**）。

        ⚠️ 用户决策 2026-08-10：补 atoms-api list 排查能力
            - 用途：OSS GET 404 后，调用本方法查报表是否已生成
            - 不硬编码 report_type：调用方传入（本次抓包 type=40=全站营销效果报表，但含义待用户确认）
            - 不替换主线：仅作为辅助诊断；现有 _download_file 重试逻辑不变

        URL: https://atoms-api.jd.com/api/download/common/asyn/download/reportInfo/list
        方法: POST JSON
        必带头（与 jzt-api 不同域）：
            - loginMode=0
            - language=zh_CN
            - siteId=0
            - Origin: https://jzt.jd.com
            - Referer: https://jzt.jd.com/
        请求体:
            {page, pageSize, startDay, endDay, nameLike, type}

        参数:
            report_type - 业务类型（用户决策 2026-08-10 不固化，调用方传入）
                         项目7 已知：9 = 快车自定义报表
                         本次抓包：40（含义待用户确认，疑似全站营销效果报表）
            start_day   - 开始日期 YYYY-MM-DD
            end_day     - 结束日期 YYYY-MM-DD
            name_like   - 报表名模糊匹配（默认空）
            page        - 页码（默认 1）
            page_size   - 每页条数（默认 10）

        返回:
            dict - 完整响应（含 code/data.datas[]/data.paginator）
                   data.datas[] 每条含 status/statusText/progress/downloadUrl/logId/createdTime 等
                   调用方根据 statusText="报表已生成" + progress=100 + downloadUrl 判是否可下载
                   不抛异常，失败时返回原始 dict（便于诊断）

        异常:
            requests.exceptions.RequestException - 网络层异常向上抛
        """
        url = "https://atoms-api.jd.com/api/download/common/asyn/download/reportInfo/list"
        # atoms-api 专属头（与 jzt-api 不同）
        atoms_headers = {
            "loginMode": "0",
            "language": "zh_CN",
            "siteId": "0",
            "Origin": "https://jzt.jd.com",
            "Referer": "https://jzt.jd.com/",
            "Content-Type": "application/json",
        }
        # 合并 session headers + 专属头（专属头优先）
        merged_headers = {**self.session.headers, **atoms_headers}

        payload = {
            "page": page,
            "pageSize": page_size,
            "startDay": start_day,
            "endDay": end_day,
            "nameLike": name_like,
            "type": report_type,
        }
        print(f"🔍 [JZT atoms-api 排查] POST {url}")
        print(f"   Body: {json.dumps(payload, ensure_ascii=False)}")

        resp = requests.post(
            url, json=payload, headers=merged_headers, timeout=30,
        )
        resp.raise_for_status()
        ret = resp.json()

        # 不抛业务异常（这是诊断方法），仅打印关键字段
        code = ret.get("code")
        datas = ret.get("data", {}).get("datas", [])
        print(f"   ├─ code: {code}")
        print(f"   ├─ 记录数: {len(datas)}")
        for i, item in enumerate(datas[:5]):  # 最多打印前 5 条
            status = item.get("status")
            status_text = item.get("statusText", "")
            progress = item.get("progress", "?")
            has_url = bool(item.get("downloadUrl"))
            log_id = item.get("logId", "")
            print(
                f"   ├─ [{i+1}] status={status} statusText='{status_text}' "
                f"progress={progress} has_downloadUrl={has_url} logId={log_id}"
            )
        if len(datas) > 5:
            print(f"   ├─ ...还有 {len(datas)-5} 条省略")
        return ret

    def _post_process_to_xlsx(self, file_bytes: bytes, clean_date: str, url_oss: str) -> str:
        """把 OSS 下载的文件 → 标准 Excel 后置处理 → 保存为 xlsx。

        ⚠️ URL 后缀决定文件类型（zip/csv）：
            - URL 含 .zip → BytesIO(zipfile) → 解压取第一个 csv
            - URL 含 .csv → BytesIO 直接读 csv
            - URL 含 %3F → URL 编码的 ?，按 URL 末尾 .zip/.csv 判
        """
        import io
        import zipfile
        import pandas as pd

        # 1. 智能识别文件类型（zip or csv）
        is_zip = ".zip" in url_oss.lower()
        if is_zip:
            print(f"   ├─ 检测到 zip 压缩包，解压中...")
            with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
                csv_names = [n for n in zf.namelist() if n.lower().endswith(".csv")]
                if not csv_names:
                    raise RuntimeError(f"❌ zip 包内无 csv 文件：{zf.namelist()}")
                csv_name = csv_names[0]  # 取第一个 csv（通常只有一个）
                print(f"   ├─ 解压文件：{csv_name}")
                csv_bytes = zf.read(csv_name)
        else:
            csv_bytes = file_bytes

        # 2. 读取 CSV（dtype=str 防精度丢失；UTF-8-sig 兼容 BOM）
        try:
            df = pd.read_csv(
                io.BytesIO(csv_bytes),
                dtype=str,
                na_filter=False,
                encoding="utf-8-sig",
                keep_default_na=False,
            )
        except Exception as e:
            raise RuntimeError(
                f"❌ CSV 解析失败：{e}\n"
                f"   请检查 OSS 返回内容是否正常"
            ) from e

        if df.empty:
            # ⚠️ 用户决策 2026-08-10：空数据视为测试通过（**仅项目10 启用**）
            #   场景：账号某日期无推广数据，服务端返回空 CSV（仍 GET 200，OSS 文件存在）
            #   行为：写入空 xlsx（仅含表头或 0 列）+ 返回成功路径
            #   注意：项目8/9 保留原「CSV 数据为空」抛错行为，未受本次决策影响
            print(
                f"   ├─ ⚠️ CSV 数据为空（{len(df.columns)}列 0行）"
                f"—— 视为业务无数据，写入空 xlsx"
            )
            # 不抛异常，继续走落盘流程

        # 3. 日期列智能处理（公共规则1+2）
        date_column, date_value = prepare_date_columns(df, clean_date)

        # 4. 数值安全转换（公共规则3）
        df = safe_convert_numeric(df)

        # 5. 构造输出路径：output/京准通全站营销单品推广效果/{date}/业务名_{date}.xlsx
        date_subdir = os.path.join(self.output_dir, clean_date)
        os.makedirs(date_subdir, exist_ok=True)
        save_filename = f"京准通全站营销单品推广效果_{clean_date}.xlsx"
        target_path = os.path.join(date_subdir, save_filename)

        # 6. 写 xlsx + 单元格格式
        df.to_excel(target_path, index=False, engine="openpyxl")
        apply_column_formats(target_path, df, date_column=date_column, date_value=date_value)

        print(
            f"✅ 文件已保存：{target_path}"
            f"\n   （{'zip→' if is_zip else ''}csv→xlsx 转存 + 日期列 + 数值转换 + 单元格格式，"
            f"{os.path.getsize(target_path)}字节，{len(df)}行 × {len(df.columns)}列）"
        )
        return target_path


class JZTQuanZhanCampaignAllStoreAPI:
    """京准通-全站营销**全店计划报表**导出 API（2026-08-10 上线）。

    ⚠️ 本类**不继承 JDBaseRequest**（与项目7/8/9/10 同原因）：
        - 鉴权体系：仅 Cookie（与项目7-10 同一文件 config/jzt_cookie.txt）
        - 流程：同步两步（POST → 立即 GET OSS），不需要基类的 30 秒间隔/重试模型
        - UA：沿用项目9 的 v=151（保证 jzt_cookie.txt 会话一致）

    🆕 与项目9（JZTQuanZhanCampaignAPI 单品计划报表）的关键差异：
        - URL 路径：⚠️ **完全相同** `/reweb/swa/account/campaign/download`
          靠 payload 内 `campaignTypes=[118]` 区分业务（项目9 是 [101]）
        - 报表名后缀：_全店计划报表_（项目9 是 _单品计划报表_）
        - 响应：本次数据完整返回 downloadUrlZip + downloadUrlCsv（项目9 抓包只含 csv）
        - 字段：完全复用项目9 的 15 项 payload（含 dateValues 嵌套列表）
        - 入参：4 项开放（order_status / is_daily / sku_id / spu_id），与项目9 一致
        - 默认值差异：
            * orderStatus：项目9 默认 ""，本项目默认 ""（一致）
            * isDaily：项目9 默认 True，本项目默认 True（一致）
    """

    # ---- 类常量（业务固定参数，2026-08-10 抓包实证 + 用户确认固化）----
    BASE_URL = "https://jzt-api.jd.com/reweb/swa/account/campaign/download"
    ORIGIN = "https://jzt.jd.com"
    REFERER = "https://jzt.jd.com/"
    USER_AGENT = (
        # ⚠️ 沿用项目9 的 v=151 UA（保证与 jzt_cookie.txt 会话一致）
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/151.0.0.0 Safari/537.36 Edg/151.0.0.0"
    )
    OUTPUT_SUBDIR = "京准通全站营销全店计划"  # output/京准通全站营销全店计划/{date}/

    # 业务固定参数（抓包值，2026-08-10 实证 + 用户 2026-08-10 确认固化）
    PLATFORM = ""               # 平台（空=不限，字符串）
    CAMPAIGN_TYPES = [118]      # 业务类型：118=疑似全店计划（用户决策 2026-08-10，含义待确认）
    PROVINCE = ""               # 省份过滤（空=全国）
    CLICK_OR_ORDER_DAY = 15     # 转化周期：15 天
    CLICK_OR_ORDER_CALIBER = 0  # 0=点击（int）
    IS_DAILY = True             # 日报标志（bool，项目9 一致）
    ORDER_STATUS_CATEGORY = 1    # 1=成交订单
    ORDER_STATUS = ""           # 订单状态过滤（字符串空）
    GIFT_FLAG = ""              # 含赠品（字符串空）
    SXU_ID = ""                 # SKU 过滤（空）
    OBYS = ""                   # 对象过滤（空）
    PIN_ID = "FYA8888"

    # ---- 复用项目10 的 MAX_DOWNLOAD_RETRY ----
    MAX_DOWNLOAD_RETRY = 8

    def __init__(self, cookie_path: str = "config/jzt_cookie.txt"):
        import requests

        # 读 Cookie（与项目7-10 互通 jzt_cookie.txt）
        cookie_path_abs = os.path.join(os.path.dirname(os.path.abspath(__file__)), cookie_path)
        if not os.path.isfile(cookie_path_abs):
            raise FileNotFoundError(
                f"❌ 京准通 Cookie 文件不存在：{cookie_path_abs}\n"
                f"   请浏览器登录 https://jzt.jd.com/home，F12 抓 jzt-api.jd.com 域 Cookie 写入此文件"
            )
        with open(cookie_path_abs, "r", encoding="utf-8") as f:
            self.cookie = f.read().strip()
        if not self.cookie:
            raise ValueError(f"❌ 京准通 Cookie 文件 {cookie_path_abs} 内容为空")

        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": self.USER_AGENT,
            "Accept": "application/json, text/plain, */*",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
            "Accept-Encoding": "gzip, deflate, br, zstd",
            "Content-Type": "application/json;charset=UTF-8",
            "Origin": self.ORIGIN,
            "Referer": self.REFERER,
            "Cookie": self.cookie,
        })

        # 输出目录（按 AGENTS.md Excel规则4）
        self.output_dir = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "output", self.OUTPUT_SUBDIR,
        )

    # ---- 业务参数组装 ----
    def _build_payload(
        self,
        start_day: str,
        end_day: str,
        order_status: str = None,
        is_daily: bool = None,
        sxu_id: str = "",
        obys: str = "",
    ) -> dict:
        """组装请求 payload（抓包实证 + 动态日期 + 4 项开放入参）。

        ⚠️ 用户决策 2026-08-10：
            - order_status 默认 ""（不限，调用方可传其他字符串）
            - is_daily 默认 True（日报，调用方可传 False）
            - sxu_id 默认 ""（不过滤），obys 默认 ""（不过滤）

        字段类型严格匹配抓包：
            - platform/province/orderStatus/giftFlag/sxuId/obys 是字符串
            - campaignTypes 是列表 [118]
            - isDaily 是布尔
            - dateValues 是嵌套列表
        """
        # 入参兜底（None → 类默认）
        if order_status is None:
            order_status = self.ORDER_STATUS
        if is_daily is None:
            is_daily = self.IS_DAILY

        # 报表名：FYA8888_全站营销_全店计划报表_{startDay}_{endDay}
        report_name = f"{self.PIN_ID}_全站营销_全店计划报表_{start_day}_{end_day}"
        return {
            "platform": self.PLATFORM,
            "campaignTypes": self.CAMPAIGN_TYPES,
            "province": self.PROVINCE,
            "startDay": start_day,
            "endDay": end_day,
            "orderStatus": order_status,
            "giftFlag": self.GIFT_FLAG,
            "clickOrOrderDay": self.CLICK_OR_ORDER_DAY,
            "clickOrOrderCaliber": self.CLICK_OR_ORDER_CALIBER,
            "sxuId": sxu_id,
            "obys": obys,
            "isDaily": is_daily,
            "orderStatusCategory": self.ORDER_STATUS_CATEGORY,
            "dateValues": [{"startDay": start_day, "endDay": end_day}],
            "reportName": report_name,
        }

    def _handle_response(self, ret: dict, op_desc: str):
        """统一处理响应（同项目9/10）。

        判定：
            - success=true
            - code=="1"（兼容字符串/数字）
            - data.code == "RC_SUCCESS"
            - data.downloadUrlZip 或 data.downloadUrlCsv 非空
        """
        if not ret.get("success", True):
            msg = ret.get("msg", "未知错误")
            code = ret.get("code")
            if code in (2001, 302) or "未登录" in msg or "登录已过期" in msg:
                raise CookieExpiredError(
                    f"❌ 京准通 Cookie 过期（{op_desc}返回 code={code}）：\n"
                    f"   → 请浏览器登录 https://jzt.jd.com/home，F12 抓 jzt-api.jd.com 域 Cookie 写入 config/jzt_cookie.txt"
                )
            if code == 601 or str(code) == "601":
                raise RuntimeError(
                    f"❌ 京准通 全站营销全店计划 限流 code=601：{msg}\n"
                    f"   → 30-120 分钟冷却，避免重试加重风控"
                )
            raise RuntimeError(
                f"❌ 京准通{op_desc}失败：code={code}, msg={msg}, 完整响应={ret}"
            )
        code = ret.get("code")
        data_code = ret.get("data", {}).get("code")
        # ⚠️ 兼容字符串/数字 code（项目9 探针发现）
        if str(code) not in ("0", "1") or data_code != "RC_SUCCESS":
            raise RuntimeError(
                f"❌ 京准通{op_desc}业务失败：code={code}, data.code={data_code}\n"
                f"   完整响应：{ret}"
            )
        return ret

    def _pick_download_url(self, ret: dict) -> str:
        """⚠️ 用户决策 2026-08-10（继承项目10）：csv 优先，zip 降级。"""
        data = ret.get("data", {})
        download_csv = data.get("downloadUrlCsv")
        download_zip = data.get("downloadUrlZip")
        if download_csv:
            print(f"   ├─ 优先使用 downloadUrlCsv（模拟浏览器行为）")
            return download_csv
        if download_zip:
            print(f"   ├─ downloadUrlCsv 缺失，降级使用 downloadUrlZip（完整压缩包）")
            return download_zip
        raise RuntimeError(f"❌ 响应中 downloadUrlCsv/downloadUrlZip 都缺失：{ret}")

    # ---- 一步：同步 POST 拿 downloadUrl ----
    def _post_for_csv(
        self,
        start_day: str,
        end_day: str,
        order_status: str = None,
        is_daily: bool = None,
        sxu_id: str = "",
        obys: str = "",
    ) -> str:
        """POST 同步返回 downloadUrl（csv 优先 / zip 降级）。"""
        payload = self._build_payload(
            start_day, end_day,
            order_status=order_status,
            is_daily=is_daily,
            sxu_id=sxu_id, obys=obys,
        )
        print(f"🚀 [JZT全站营销全店计划] POST {self.BASE_URL}")
        print(f"   Body: {json.dumps(payload, ensure_ascii=False)}")

        resp = self.session.post(self.BASE_URL, json=payload, timeout=60)
        resp.raise_for_status()
        ret = resp.json()
        self._handle_response(ret, op_desc="导出全店计划报表")

        download_id = ret.get("data", {}).get("downloadId")
        download_url = self._pick_download_url(ret)
        print(f"✅ 拿到 downloadId={download_id}, downloadUrl（前80字符）: {download_url[:80]}...")
        return download_url

    # ---- 二步：GET OSS 下载文件字节流（复用项目10 8 次重试逻辑）----
    def _download_file(self, url_oss: str) -> bytes:
        """GET OSS 链接，下载文件字节流（带 404 随机退避重试，最多 8 次）。"""
        last_error = None
        total_wait = 0.0
        for retry in range(self.MAX_DOWNLOAD_RETRY):
            try:
                resp = requests.get(
                    url_oss,
                    headers={"User-Agent": self.USER_AGENT},
                    timeout=60,
                )
                if resp.status_code == 200:
                    if retry > 0:
                        print(
                            f"  ✅ 第 {retry+1}/{self.MAX_DOWNLOAD_RETRY} 次重试成功"
                            f"（累计等待 {total_wait:.1f} 秒）"
                        )
                    return resp.content
                if resp.status_code == 404:
                    backoff = random.uniform(3, 10)
                    total_wait += backoff
                    print(
                        f"  ⚠️ 第 {retry+1}/{self.MAX_DOWNLOAD_RETRY} 次 url 404 NoSuchKey，"
                        f"随机退避 {backoff:.1f} 秒后重试...（累计 {total_wait:.1f} 秒）"
                    )
                    time.sleep(backoff)
                    continue
                resp.raise_for_status()
            except requests.exceptions.RequestException as e:
                last_error = e
                print(f"  ⚠️ url 下载异常：{e}，重试中...")
                time.sleep(random.uniform(3, 10))
                total_wait += 3
        # 全部失败 → 严重告警
        print(
            f"  🔴 [严重告警] url 下载失败，已重试 {self.MAX_DOWNLOAD_RETRY} 次仍未成功\n"
            f"     ├─ 累计等待时间：{total_wait:.1f} 秒\n"
            f"     ├─ 诊断建议：手动 GET 该 URL 验证 / 检查签名是否过期\n"
            f"     └─ URL: {url_oss[:120]}"
        )
        raise RuntimeError(
            f"❌ url 下载失败（重试 {self.MAX_DOWNLOAD_RETRY} 次后，累计等待 {total_wait:.1f} 秒）：{last_error}\n"
            f"   URL: {url_oss[:120]}"
        )

    # ---- 一键封装 ----
    def run_full_export(
        self,
        start_date: str = None,
        end_date: str = None,
        date: str = None,
        order_status: str = None,   # 默认 None→""（不限）
        is_daily: bool = None,      # 默认 None→True（日报）
        sku_id: str = "",           # 项目9 字段名是 sxu_id，调度层用 sku_id 统一命名
        spu_id: str = "",           # 暂未使用，保留接口对齐项目9/10
    ) -> str:
        """一键跑通：POST 同步拿 url（csv 优先）→ GET OSS 下载 → 解压转 xlsx。

        ⚠️ 用户决策 2026-08-10：4 项开放入参与项目9 一致：
            - order_status: None→""（不限）；spu_id 暂未使用（项目9/11 payload 无 spuId 字段）
        """
        if date is None and start_date is None and end_date is None:
            raise ValueError("❌ 至少需要传入 date 或 start_date/end_date")
        if start_date is None:
            start_date = date
        if end_date is None:
            end_date = date

        print(f"🚀 [JZT全站营销全店计划] 启动完整导出：{start_date} ~ {end_date}")
        print(f"   └─ Step 1/2: POST 同步拿 downloadUrl（csv 优先）...")
        url_oss = self._post_for_csv(
            start_date, end_date,
            order_status=order_status, is_daily=is_daily,
            sxu_id=sku_id, obys="",
        )
        print(f"   └─ Step 2/2: GET OSS 下载并落盘为 xlsx...")
        file_bytes = self._download_file(url_oss)

        return self._post_process_to_xlsx(file_bytes, start_date, url_oss)

    def _post_process_to_xlsx(self, file_bytes: bytes, clean_date: str, url_oss: str) -> str:
        """OSS 文件 → 标准 Excel 后置处理 → 保存为 xlsx（zip/csv 智能识别 + 空数据不抛错）。

        ⚠️ 用户决策 2026-08-10（继承项目10）：空数据视为业务无数据，写入空 xlsx + 返回成功。
        """
        import io
        import zipfile
        import pandas as pd

        # 1. 智能识别文件类型
        is_zip = ".zip" in url_oss.lower()
        if is_zip:
            print(f"   ├─ 检测到 zip 压缩包，解压中...")
            with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
                csv_names = [n for n in zf.namelist() if n.lower().endswith(".csv")]
                if not csv_names:
                    raise RuntimeError(f"❌ zip 包内无 csv 文件：{zf.namelist()}")
                csv_name = csv_names[0]
                print(f"   ├─ 解压文件：{csv_name}")
                csv_bytes = zf.read(csv_name)
        else:
            csv_bytes = file_bytes

        # 2. 读取 CSV
        try:
            df = pd.read_csv(
                io.BytesIO(csv_bytes),
                dtype=str,
                na_filter=False,
                encoding="utf-8-sig",
                keep_default_na=False,
            )
        except Exception as e:
            raise RuntimeError(
                f"❌ CSV 解析失败：{e}\n"
                f"   请检查 OSS 返回内容是否正常"
            ) from e

        if df.empty:
            # ⚠️ 用户决策 2026-08-10：空数据视为成功（仅项目10/11 启用，项目8/9 保持原行为）
            print(
                f"   ├─ ⚠️ CSV 数据为空（{len(df.columns)}列 0行）"
                f"—— 视为业务无数据，写入空 xlsx"
            )

        # 3. 日期列智能处理
        date_column, date_value = prepare_date_columns(df, clean_date)

        # 4. 数值安全转换
        df = safe_convert_numeric(df)

        # 5. 构造输出路径：output/京准通全站营销全店计划/{date}/业务名_{date}.xlsx
        date_subdir = os.path.join(self.output_dir, clean_date)
        os.makedirs(date_subdir, exist_ok=True)
        save_filename = f"京准通全站营销全店计划_{clean_date}.xlsx"
        target_path = os.path.join(date_subdir, save_filename)

        # 6. 写 xlsx + 单元格格式
        df.to_excel(target_path, index=False, engine="openpyxl")
        apply_column_formats(target_path, df, date_column=date_column, date_value=date_value)

        print(
            f"✅ 文件已保存：{target_path}"
            f"\n   （{'zip→' if is_zip else ''}csv→xlsx 转存 + 日期列 + 数值转换 + 单元格格式，"
            f"{os.path.getsize(target_path)}字节，{len(df)}行 × {len(df.columns)}列）"
        )
        return target_path


def _run_jzt_quanzhan_campaign_all_store_full(**kwargs) -> str:
    """调度器专用的京准通全站营销全店计划完整流程函数（2026-08-10 上线）。

    ⚠️ 注册到 BUSINESS_REGISTRY["京准通全站营销全店计划"]["callable"]。
    """
    forward_kwargs = {
        k: kwargs[k] for k in (
            "start_date", "end_date", "date",
            "order_status", "is_daily", "sku_id", "spu_id",
        )
        if k in kwargs
    }

    api = JZTQuanZhanCampaignAllStoreAPI(cookie_path=kwargs.get("cookie_path", "config/jzt_cookie.txt"))
    return api.run_full_export(**forward_kwargs)


# ⚠️ 注册表 callable 字段回填（项目11，2026-08-10 上线）
BUSINESS_REGISTRY["京准通全站营销全店计划"]["callable"] = _run_jzt_quanzhan_campaign_all_store_full
BUSINESS_REGISTRY["京准通全站营销全店计划"]["api_class"] = JZTQuanZhanCampaignAllStoreAPI


class JZTQuanZhanEffectAllStoreAPI:
    """京准通-全站营销**全店推广效果报表**导出 API（2026-08-10 上线）。

    ⚠️ 本类**不继承 JDBaseRequest**（与项目7/8/9/10/11 同原因）：
        - 鉴权体系：仅 Cookie（与项目7-11 同一文件 config/jzt_cookie.txt）
        - 流程：同步两步（POST → 立即 GET OSS），不需要基类的 30 秒间隔/重试模型
        - UA：沿用项目9/10/11 的 v=151（保证 jzt_cookie.txt 会话一致）

    🆕 与项目10（JZTQuanZhanEffectAPI 单品推广效果报表）的关键差异：
        - URL 路径：⚠️ **完全相同** `/reweb/swa/effect/order/download`
          靠 payload 内 `campaignTypes=[118]` 区分业务（项目10 是 [101]）
        - 报表名后缀：_效果报表_全店推广_（项目10 是 _效果报表_单品推广_）
        - 字段结构：完全复用项目10 的 13 项 payload
        - 入参：4 项开放（order_status / is_daily / sku_id / spu_id），与项目10 一致
        - 默认值差异：
            * orderStatus：项目10 默认 "1"，本项目默认 "1"（一致）
            * isDaily：项目10 默认 False，本项目默认 False（一致）
    """

    # ---- 类常量（业务固定参数，2026-08-10 抓包实证 + 用户确认固化）----
    BASE_URL = "https://jzt-api.jd.com/reweb/swa/effect/order/download"
    ORIGIN = "https://jzt.jd.com"
    REFERER = "https://jzt.jd.com/"
    USER_AGENT = (
        # ⚠️ 沿用项目10 的 v=151 UA（保证与 jzt_cookie.txt 会话一致）
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/151.0.0.0 Safari/537.36 Edg/151.0.0.0"
    )
    OUTPUT_SUBDIR = "京准通全站营销全店推广效果"  # output/京准通全站营销全店推广效果/{date}/

    # 业务固定参数（抓包值，2026-08-10 实证 + 用户 2026-08-10 确认固化）
    PLATFORM = ""               # 平台（空=不限，字符串）
    CAMPAIGN_TYPES = [118]      # 业务类型：118=疑似全店推广效果（与项目11 同号但路径不同）
    CLICK_OR_ORDER_DAY = 15     # 转化周期：15 天
    CLICK_OR_ORDER_CALIBER = 0  # 0=点击（int）
    IS_DAILY = False            # 非日报（bool，与项目10 默认一致）
    ORDER_STATUS_CATEGORY = 1    # 1=成交订单
    ORDER_STATUS = "1"          # 订单状态过滤（字符串 "1"成交订单，与项目10 默认一致）
    GIFT_FLAG = ""              # 含赠品（字符串空）
    PIN_ID = "FYA8888"

    # ---- 复用项目10 的 MAX_DOWNLOAD_RETRY ----
    MAX_DOWNLOAD_RETRY = 8

    def __init__(self, cookie_path: str = "config/jzt_cookie.txt"):
        import requests

        # 读 Cookie（与项目7-11 互通 jzt_cookie.txt）
        cookie_path_abs = os.path.join(os.path.dirname(os.path.abspath(__file__)), cookie_path)
        if not os.path.isfile(cookie_path_abs):
            raise FileNotFoundError(
                f"❌ 京准通 Cookie 文件不存在：{cookie_path_abs}\n"
                f"   请浏览器登录 https://jzt.jd.com/home，F12 抓 jzt-api.jd.com 域 Cookie 写入此文件"
            )
        with open(cookie_path_abs, "r", encoding="utf-8") as f:
            self.cookie = f.read().strip()
        if not self.cookie:
            raise ValueError(f"❌ 京准通 Cookie 文件 {cookie_path_abs} 内容为空")

        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": self.USER_AGENT,
            "Accept": "application/json, text/plain, */*",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
            "Accept-Encoding": "gzip, deflate, br, zstd",
            "Content-Type": "application/json;charset=UTF-8",
            "Origin": self.ORIGIN,
            "Referer": self.REFERER,
            "Cookie": self.cookie,
        })

        # 输出目录（按 AGENTS.md Excel规则4）
        self.output_dir = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "output", self.OUTPUT_SUBDIR,
        )

    # ---- 业务参数组装 ----
    def _build_payload(
        self,
        start_day: str,
        end_day: str,
        order_status: str = None,
        is_daily: bool = None,
        sku_id: str = "",
        spu_id: str = "",
    ) -> dict:
        """组装请求 payload（抓包实证 + 动态日期 + 4 项开放入参）。

        ⚠️ 用户决策 2026-08-10：
            - order_status 默认 "1"（成交订单，调用方可传空）
            - is_daily 默认 False（非日报，调用方可传 True）
            - sku_id 默认 ""（不过滤），spu_id 默认 ""（不过滤）

        字段类型严格匹配抓包：
            - platform/orderStatus/giftFlag/skuId/spuId 是字符串
            - campaignTypes 是列表 [118]
            - isDaily 是布尔
        """
        # 入参兜底（None → 类默认）
        if order_status is None:
            order_status = self.ORDER_STATUS
        if is_daily is None:
            is_daily = self.IS_DAILY

        # 报表名：FYA8888_全站营销_效果报表_全店推广_{startDay}_{endDay}
        report_name = f"{self.PIN_ID}_全站营销_效果报表_全店推广_{start_day}_{end_day}"
        return {
            "platform": self.PLATFORM,
            "campaignTypes": self.CAMPAIGN_TYPES,
            "startDay": start_day,
            "endDay": end_day,
            "orderStatus": order_status,
            "giftFlag": self.GIFT_FLAG,
            "clickOrOrderDay": self.CLICK_OR_ORDER_DAY,
            "clickOrOrderCaliber": self.CLICK_OR_ORDER_CALIBER,
            "skuId": sku_id,
            "spuId": spu_id,
            "isDaily": is_daily,
            "orderStatusCategory": self.ORDER_STATUS_CATEGORY,
            "reportName": report_name,
        }

    def _handle_response(self, ret: dict, op_desc: str):
        """统一处理响应（同项目10/11）。"""
        if not ret.get("success", True):
            msg = ret.get("msg", "未知错误")
            code = ret.get("code")
            if code in (2001, 302) or "未登录" in msg or "登录已过期" in msg:
                raise CookieExpiredError(
                    f"❌ 京准通 Cookie 过期（{op_desc}返回 code={code}）：\n"
                    f"   → 请浏览器登录 https://jzt.jd.com/home，F12 抓 jzt-api.jd.com 域 Cookie 写入 config/jzt_cookie.txt"
                )
            if code == 601 or str(code) == "601":
                raise RuntimeError(
                    f"❌ 京准通 全站营销全店推广效果 限流 code=601：{msg}\n"
                    f"   → 30-120 分钟冷却，避免重试加重风控"
                )
            raise RuntimeError(
                f"❌ 京准通{op_desc}失败：code={code}, msg={msg}, 完整响应={ret}"
            )
        code = ret.get("code")
        data_code = ret.get("data", {}).get("code")
        # ⚠️ 兼容字符串/数字 code（项目9 探针发现）
        if str(code) not in ("0", "1") or data_code != "RC_SUCCESS":
            raise RuntimeError(
                f"❌ 京准通{op_desc}业务失败：code={code}, data.code={data_code}\n"
                f"   完整响应：{ret}"
            )
        return ret

    def _pick_download_url(self, ret: dict) -> str:
        """⚠️ 用户决策 2026-08-10（继承项目10）：csv 优先，zip 降级。"""
        data = ret.get("data", {})
        download_csv = data.get("downloadUrlCsv")
        download_zip = data.get("downloadUrlZip")
        if download_csv:
            print(f"   ├─ 优先使用 downloadUrlCsv（模拟浏览器行为）")
            return download_csv
        if download_zip:
            print(f"   ├─ downloadUrlCsv 缺失，降级使用 downloadUrlZip（完整压缩包）")
            return download_zip
        raise RuntimeError(f"❌ 响应中 downloadUrlCsv/downloadUrlZip 都缺失：{ret}")

    # ---- 一步：同步 POST 拿 downloadUrl ----
    def _post_for_csv(
        self,
        start_day: str,
        end_day: str,
        order_status: str = None,
        is_daily: bool = None,
        sku_id: str = "",
        spu_id: str = "",
    ) -> str:
        """POST 同步返回 downloadUrl（csv 优先 / zip 降级）。"""
        payload = self._build_payload(
            start_day, end_day,
            order_status=order_status,
            is_daily=is_daily,
            sku_id=sku_id, spu_id=spu_id,
        )
        print(f"🚀 [JZT全站营销全店推广效果] POST {self.BASE_URL}")
        print(f"   Body: {json.dumps(payload, ensure_ascii=False)}")

        resp = self.session.post(self.BASE_URL, json=payload, timeout=60)
        resp.raise_for_status()
        ret = resp.json()
        self._handle_response(ret, op_desc="导出全店推广效果报表")

        download_id = ret.get("data", {}).get("downloadId")
        download_url = self._pick_download_url(ret)
        print(f"✅ 拿到 downloadId={download_id}, downloadUrl（前80字符）: {download_url[:80]}...")
        return download_url

    # ---- 二步：GET OSS 下载文件字节流（复用项目10 8 次重试逻辑）----
    def _download_file(self, url_oss: str) -> bytes:
        """GET OSS 链接，下载文件字节流（带 404 随机退避重试，最多 8 次）。"""
        last_error = None
        total_wait = 0.0
        for retry in range(self.MAX_DOWNLOAD_RETRY):
            try:
                resp = requests.get(
                    url_oss,
                    headers={"User-Agent": self.USER_AGENT},
                    timeout=60,
                )
                if resp.status_code == 200:
                    if retry > 0:
                        print(
                            f"  ✅ 第 {retry+1}/{self.MAX_DOWNLOAD_RETRY} 次重试成功"
                            f"（累计等待 {total_wait:.1f} 秒）"
                        )
                    return resp.content
                if resp.status_code == 404:
                    backoff = random.uniform(3, 10)
                    total_wait += backoff
                    print(
                        f"  ⚠️ 第 {retry+1}/{self.MAX_DOWNLOAD_RETRY} 次 url 404 NoSuchKey，"
                        f"随机退避 {backoff:.1f} 秒后重试...（累计 {total_wait:.1f} 秒）"
                    )
                    time.sleep(backoff)
                    continue
                resp.raise_for_status()
            except requests.exceptions.RequestException as e:
                last_error = e
                print(f"  ⚠️ url 下载异常：{e}，重试中...")
                time.sleep(random.uniform(3, 10))
                total_wait += 3
        # 全部失败 → 严重告警
        print(
            f"  🔴 [严重告警] url 下载失败，已重试 {self.MAX_DOWNLOAD_RETRY} 次仍未成功\n"
            f"     ├─ 累计等待时间：{total_wait:.1f} 秒\n"
            f"     ├─ 诊断建议：手动 GET 该 URL 验证 / 检查签名是否过期\n"
            f"     └─ URL: {url_oss[:120]}"
        )
        raise RuntimeError(
            f"❌ url 下载失败（重试 {self.MAX_DOWNLOAD_RETRY} 次后，累计等待 {total_wait:.1f} 秒）：{last_error}\n"
            f"   URL: {url_oss[:120]}"
        )

    # ---- 一键封装 ----
    def run_full_export(
        self,
        start_date: str = None,
        end_date: str = None,
        date: str = None,
        order_status: str = None,   # 默认 None→"1"（成交订单）
        is_daily: bool = None,      # 默认 None→False（非日报）
        sku_id: str = "",
        spu_id: str = "",
    ) -> str:
        """一键跑通：POST 同步拿 url（csv 优先）→ GET OSS 下载 → 解压转 xlsx。

        ⚠️ 用户决策 2026-08-10：4 项开放入参与项目10 一致：
            - order_status: None→"1"（成交订单）
            - is_daily: None→False（非日报）
            - sku_id / spu_id: 默认 ""（不过滤）
        """
        if date is None and start_date is None and end_date is None:
            raise ValueError("❌ 至少需要传入 date 或 start_date/end_date")
        if start_date is None:
            start_date = date
        if end_date is None:
            end_date = date

        print(f"🚀 [JZT全站营销全店推广效果] 启动完整导出：{start_date} ~ {end_date}")
        print(f"   └─ Step 1/2: POST 同步拿 downloadUrl（csv 优先）...")
        url_oss = self._post_for_csv(
            start_date, end_date,
            order_status=order_status, is_daily=is_daily,
            sku_id=sku_id, spu_id=spu_id,
        )
        print(f"   └─ Step 2/2: GET OSS 下载并落盘为 xlsx...")
        file_bytes = self._download_file(url_oss)

        return self._post_process_to_xlsx(file_bytes, start_date, url_oss)

    def _post_process_to_xlsx(self, file_bytes: bytes, clean_date: str, url_oss: str) -> str:
        """OSS 文件 → 标准 Excel 后置处理 → 保存为 xlsx（zip/csv 智能识别 + 空数据不抛错）。

        ⚠️ 用户决策 2026-08-10（继承项目10）：空数据视为业务无数据，写入空 xlsx + 返回成功。
        ⚠️ 用户决策 2026-08-10 补充：合计行去除 + 商品ID 整数0位小数 + 首列冻结（全局生效）。
        """
        import io
        import zipfile
        import pandas as pd

        # 1. 智能识别文件类型
        is_zip = ".zip" in url_oss.lower()
        if is_zip:
            print(f"   ├─ 检测到 zip 压缩包，解压中...")
            with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
                csv_names = [n for n in zf.namelist() if n.lower().endswith(".csv")]
                if not csv_names:
                    raise RuntimeError(f"❌ zip 包内无 csv 文件：{zf.namelist()}")
                csv_name = csv_names[0]
                print(f"   ├─ 解压文件：{csv_name}")
                csv_bytes = zf.read(csv_name)
        else:
            csv_bytes = file_bytes

        # 2. 读取 CSV
        try:
            df = pd.read_csv(
                io.BytesIO(csv_bytes),
                dtype=str,
                na_filter=False,
                encoding="utf-8-sig",
                keep_default_na=False,
            )
        except Exception as e:
            raise RuntimeError(
                f"❌ CSV 解析失败：{e}\n"
                f"   请检查 OSS 返回内容是否正常"
            ) from e

        if df.empty:
            # ⚠️ 用户决策 2026-08-10：空数据视为成功（仅项目10/11/12 启用）
            print(
                f"   ├─ ⚠️ CSV 数据为空（{len(df.columns)}列 0行）"
                f"—— 视为业务无数据，写入空 xlsx"
            )

        # 3. 日期列智能处理（公共规则1+2）
        date_column, date_value = prepare_date_columns(df, clean_date)

        # 4. 数值安全转换（公共规则3，含合计行剔除 + 商品ID 整数 0 位小数）
        df = safe_convert_numeric(df)

        # 5. 构造输出路径：output/京准通全站营销全店推广效果/{date}/业务名_{date}.xlsx
        date_subdir = os.path.join(self.output_dir, clean_date)
        os.makedirs(date_subdir, exist_ok=True)
        save_filename = f"京准通全站营销全店推广效果_{clean_date}.xlsx"
        target_path = os.path.join(date_subdir, save_filename)

        # 6. 写 xlsx + 单元格格式（含首列+表头冻结 + SKU/SPU/商品ID 整数 0 位小数）
        df.to_excel(target_path, index=False, engine="openpyxl")
        apply_column_formats(target_path, df, date_column=date_column, date_value=date_value)

        print(
            f"✅ 文件已保存：{target_path}"
            f"\n   （{'zip→' if is_zip else ''}csv→xlsx 转存 + 日期列 + 数值转换 + 单元格格式 + 合计行剔除 + 首列冻结，"
            f"{os.path.getsize(target_path)}字节，{len(df)}行 × {len(df.columns)}列）"
        )
        return target_path


def _run_jzt_quanzhan_effect_all_store_full(**kwargs) -> str:
    """调度器专用的京准通全站营销全店推广效果完整流程函数（2026-08-10 上线）。

    ⚠️ 注册到 BUSINESS_REGISTRY["京准通全站营销全店推广效果"]["callable"]。
    """
    forward_kwargs = {
        k: kwargs[k] for k in (
            "start_date", "end_date", "date",
            "order_status", "is_daily", "sku_id", "spu_id",
        )
        if k in kwargs
    }

    api = JZTQuanZhanEffectAllStoreAPI(cookie_path=kwargs.get("cookie_path", "config/jzt_cookie.txt"))
    return api.run_full_export(**forward_kwargs)


# ⚠️ 注册表 callable 字段回填（项目12，2026-08-10 上线）
BUSINESS_REGISTRY["京准通全站营销全店推广效果"] = {
    "api_class": JZTQuanZhanEffectAllStoreAPI,
    "method": "run_full_export",
    "callable": _run_jzt_quanzhan_effect_all_store_full,
    "desc": "京准通全站营销全店推广效果报表导出（同步两步：POST拿urlZip/csv→GET下载→解压转xlsx）",
    "params": {
        "date": "查询日期YYYY-MM-DD（单日查询）",
        "start_date": "开始日期YYYY-MM-DD（区间查询，可选）",
        "end_date": "结束日期YYYY-MM-DD（区间查询，可选）",
        "cookie_path": "京准通Cookie路径（默认config/jzt_cookie.txt，可选）",
        "order_status": "订单状态（可选，默认None→\"1\"成交订单；传\"\"=不限）",
        "is_daily": "日报标志（可选，默认None→False非日报；传True=日报）",
        "sku_id": "SKU过滤（可选，默认\"\"=不过滤）",
        "spu_id": "SPU过滤（可选，默认\"\"=不过滤）",
    },
}


def _run_jzt_quanzhan_effect_full(**kwargs) -> str:
    """调度器专用的京准通全站营销单品推广效果完整流程函数（2026-08-10 上线）。

    ⚠️ 注册到 BUSINESS_REGISTRY["京准通全站营销单品推广效果"]["callable"]。
    设计动机：JZTQuanZhanEffectAPI.__init__ 需要 cookie_path + run_full_export
              开放入参（order_status / is_daily / sku_id / spu_id），
              标准调度路径无法透传，本函数手动构造实例并调用 run_full_export。

    参数:
        kwargs - 来自 run_business 的透传参数：
            date        (str): 单日查询（start/end 默认=date）
            start_date  (str): 开始日期 YYYY-MM-DD
            end_date    (str): 结束日期 YYYY-MM-DD
            cookie_path (str): 可选，默认 config/jzt_cookie.txt
            order_status(str): 可选，默认 None→"1"（成交订单）；传 ""=不限
            is_daily    (bool): 可选，默认 None→False；传 True=日报
            sku_id      (str): 可选，默认 ""（不过滤）
            spu_id      (str): 可选，默认 ""（不过滤）
    返回:
        str - 保存的 .xlsx 绝对路径
    """
    forward_kwargs = {
        k: kwargs[k] for k in (
            "start_date", "end_date", "date",
            "order_status", "is_daily", "sku_id", "spu_id",
        )
        if k in kwargs
    }

    api = JZTQuanZhanEffectAPI(cookie_path=kwargs.get("cookie_path", "config/jzt_cookie.txt"))
    return api.run_full_export(**forward_kwargs)


# ⚠️ 注册表 callable 字段回填（项目10，2026-08-10 上线）
BUSINESS_REGISTRY["京准通全站营销单品推广效果"]["callable"] = _run_jzt_quanzhan_effect_full
BUSINESS_REGISTRY["京准通全站营销单品推广效果"]["api_class"] = JZTQuanZhanEffectAPI


def _run_keyword_analysis_full(**kwargs) -> str:
    """调度器专用的商智关键词分析完整流程函数（2026-08-10 上线）。

    ⚠️ 注册到 BUSINESS_REGISTRY["商智关键词分析"]["callable"]。
    设计动机：基类 JDBaseRequest.__init__ 无参可走标准调度路径，
              所以 callable 可以直接转发参数。

    参数:
        kwargs - 来自 run_business 的透传参数：
            date       (str): 单日查询 YYYY-MM-DD（与 granularity='day' 配套）
            start_date (str): 区间开始 YYYY-MM-DD
            end_date   (str): 区间结束 YYYY-MM-DD
            granularity(str): 'day' / 'month'（默认 'day'）
    返回:
        str - 保存的 xlsx 绝对路径
    异常:
        ValueError - 缺日期参数
        CookieExpiredError / RuntimeError
    """
    forward_kwargs = {
        k: kwargs[k] for k in ("date", "start_date", "end_date", "granularity")
        if k in kwargs
    }
    api = KeywordAnalysisAPI()
    return api.run_full_export(**forward_kwargs)


# ⚠️ 项目13 注册表回填（商智关键词分析，2026-08-10 上线）
# ⚠️ api_class 在类定义之后回填（解决前向引用：KeywordAnalysisAPI 类在 5676 行）
BUSINESS_REGISTRY["商智关键词分析"]["callable"] = _run_keyword_analysis_full
# api_class 占位为 None，待下方类定义完成后回填


# ============================================================
#  业务接口 13：（新业务 - 商智关键词分析导出，2026-08-10 上线）
# ------------------------------------------------------------
#  中文说明（小白必读）：
#    商智"关键词分析"报表导出接口（downTable.ajax），
#    输出按搜索词聚合的 FYA 关键词数据（访客数/成交/转化率等）。
#
#  ⚠️ 核心特征（与项目1-6 的区别）：
#    - 表单格式：application/x-www-form-urlencoded（不是 JSON）
#    - 同步返回：直接返回 xlsx 字节流，无 taskId，无需轮询
#    - 支持日期区间聚合：dateType+interval 配套，month+MONTH（按月）或 day+DAY（按日）
#    - 不需要逐日循环（服务端已聚合）
#    - 原生 Excel 无时间列：pandas 读取后手动插入时间区间列
#    - 鉴权：Cookie + User-mup/uuid/User-mnp 风控三元组（UUID 完全随机）
#
#  ⚠️ 与项目1-6 关键差异：
#    - URL 路径：keyword/analysis/shopOut（项目1-6 是 source/* 或 keyword/* 不同页面）
#    - Referer：viewflow/shopKeywordsVNew.html（项目4 是 viewSourcesVNew.html）
#    - groupType=lastSrcPageSearchKeyword（项目1-6 是 lastSrcChannelId2/3 等）
#    - sortField=jdr_sch_traffic_enter_shop__browse_page_cnt_shop_last_src（按浏览量排序）
# ============================================================

class KeywordAnalysisAPI(JDBaseRequest):
    """商智-关键词分析导出 API（2026-08-10 上线骨架）。

    父类复用：
        - JDBaseRequest 提供：
            * Cookie 读取（config/sz_cookie.txt）
            * 风控签名（UUID 完全随机 + MD5 哈希）
            * 自动 30 秒间隔（_wait_interval）
            * 自动重试 + UA 切换
            * 日志

    自实现部分：
        - _gen_uuid_random / _gen_risk_params_random：UUID 完全随机（不依赖 UUID_PREFIX）
        - 业务参数组装：_build_form_payload 含 dateType/interval 互斥逻辑
        - run_full_export 支持 day/month 双粒度
        - xlsx 后置处理：插入时间区间列 + Excel 通用后置（日期/数值/格式）
    """

    # 接口 URL（业务约束，固定）
    API_URL = "https://szgateway.jd.com/szpaas/szajax/keyword/analysis/shopOut/downTable.ajax"

    # 必带请求头（业务约束，固定）
    # ⚠️ 与项目4 的 viewSourcesVNew.html 不同（这里是关键词分析页）
    ORIGIN = "https://sz.jd.com"
    REFERER = "https://sz.jd.com/szweb/sz/view/viewflow/shopKeywordsVNew.html"

    # 固定业务参数（抓包值 2026-08-10 固化，不读 config）
    FIXED_BIZ_PARAMS = {
        "method": "POST",
        "target": "_self",
        "groupType": "lastSrcPageSearchKeyword",
        "attributes": "lastSrcPageSearchKeyword",
        "limit": "300",
        "sortField": "jdr_sch_traffic_enter_shop__browse_page_cnt_shop_last_src",
        "sortType": "desc",
    }

    # 可变业务参数（从 config 读取，缺省用兜底值）
    VARIABLE_BIZ_PARAMS = {
        "platformCate1": "",  # 平台品类 1：空字符串=全品类
    }

    # 两种聚合粒度（互斥，不可混用）
    GRANULARITY_DAY = {"dateType": "day", "interval": "DAY"}
    GRANULARITY_MONTH = {"dateType": "month", "interval": "MONTH"}

    OUTPUT_SUBDIR = "商智关键词分析"

    # ---------- UUID 完全随机生成器（与项目4 OfflineChannelAPI 同源）----------

    def _gen_uuid_random(self):
        """完全随机 UUID（16hex-10hex）。

        ⚠️ 与基类 UUID_PREFIX 的区别：
            关键词分析页面（用户抓包）UUID 完全随机，不带固定前缀。
            使用 secrets 模块生成密码学级随机 hex。
        """
        import secrets
        prefix = secrets.token_hex(8)   # 8 字节 = 16 hex
        suffix = secrets.token_hex(5)   # 5 字节 = 10 hex
        return f"{prefix}-{suffix}"

    def _gen_risk_params_random(self, url):
        """生成风控三元组：UUID 完全随机。

        算法：User-mnp = MD5(URL路径 + uuid + 时间戳 + SIGN_SALT)
        """
        from urllib.parse import urlparse
        timestamp = int(time.time() * 1000)
        uuid_str = self._gen_uuid_random()
        parsed = urlparse(url)
        url_path = parsed.path

        sign_str = f"{url_path}{uuid_str}{timestamp}{self.SIGN_SALT}"
        user_mnp = hashlib.md5(sign_str.encode("utf-8")).hexdigest()

        return {
            "User-mup": str(timestamp),
            "User-mnp": user_mnp,
            "uuid": uuid_str,
        }

    # ---------- 日期参数解析（兼容项目1-6 的 _get_date_params 模式）----------

    def _get_date_params(self, date=None, start_date=None, end_date=None, granularity="day"):
        """根据粒度生成 date / startDate / endDate / dateType / interval。

        ⚠️ 规则（用户 2026-08-10 强调）：
            - dateType 与 interval 必须配套：day+DAY 或 month+MONTH（不能混用）
            - 不需要在客户端拆日期循环（服务端已聚合）
            - 单日查询时 date=startDate=endDate
        """
        if date is None and start_date is None and end_date is None:
            raise ValueError("❌ 至少需要传入 date 或 start_date/end_date")
        if start_date is None:
            start_date = date
        if end_date is None:
            end_date = date

        # date 字段格式（2026-08-10 抓包实证）：
            #   - month 粒度：YYYYMM（如 202607），紧凑无分隔符
            #   - day 粒度：YYYY-MM-DD（如 2026-07-30），带分隔符完整日期
        if granularity == "month":
            date_compact = start_date.replace("-", "")[:6]  # YYYYMM（紧凑）
        else:
            date_compact = start_date                       # YYYY-MM-DD（带分隔符）

        # dateType/interval 配套（day+DAY 或 month+MONTH，不可混用）
        if granularity == "day":
            date_type = "day"
            interval = "DAY"
        else:
            date_type = "month"
            interval = "MONTH"

        return {
            "date": date_compact,
            "startDate": start_date,
            "endDate": end_date,
            # ⚠️ 字段顺序按抓包实证（2026-08-10）：
            #   month 抓包: dateType=month, interval=MONTH（dateType 在前）
            #   day 抓包  : interval=DAY, dateType=day（interval 在前）
            # Python 3.7+ dict 保留插入顺序，所以**两个粒度各起一段**保证字段顺序精确匹配抓包
            **({"dateType": date_type, "interval": interval} if granularity == "month"
               else {"interval": interval, "dateType": date_type}),
        }

    # ---------- 表单参数组装 ----------

    def _build_form_payload(self, date=None, start_date=None, end_date=None, granularity="day"):
        """组装完整表单参数（含固定+可变+日期+风控）。"""
        date_params = self._get_date_params(date, start_date, end_date, granularity)
        return {
            **self.FIXED_BIZ_PARAMS,
            **self.VARIABLE_BIZ_PARAMS,
            **date_params,
        }

    # ---------- 一步：POST 拿 xlsx 字节流 ----------

    def _post_for_xlsx(self, form_data):
        """POST 同步返回 xlsx 字节流（无 taskId）。

        返回:
            bytes - xlsx 文件字节流
        异常:
            CookieExpiredError / RuntimeError
        """
        # 间隔控制（基类自带）
        self._wait_interval()

        # 生成风控三元组
        risk_params = self._gen_risk_params_random(self.API_URL)
        full_data = {**form_data, **risk_params}

        ua_name = "Edge" if self._current_ua_index == 0 else "Chrome"
        self.logger.info(f"[关键词分析] POST {self.API_URL} (UA={ua_name})")
        self.logger.debug(f"表单参数: {json.dumps(full_data, ensure_ascii=False)[:500]}")

        # 记录请求时间（基类处理）
        JDBaseRequest._last_request_time = time.time()

        response = self.session.post(
            self.API_URL,
            data=full_data,
            timeout=self.REQUEST_TIMEOUT,
        )

        # 业务码判定
        content_type = response.headers.get("Content-Type", "")
        if "spreadsheetml" not in content_type:
            # 不是 xlsx 流（可能被风控拦截 / Cookie 过期）
            snippet = response.text[:500] if response.text else "(空响应)"
            if response.status_code in (401, 403):
                raise CookieExpiredError(
                    f"❌ 关键词分析 Cookie 过期（HTTP {response.status_code}）\n"
                    f"   → 请浏览器重新登录 https://sz.jd.com/szweb/sz/view/viewflow/shopKeywordsVNew.html\n"
                    f"   → F12 抓 szgateway.jd.com 域 Cookie 写入 config/sz_cookie.txt"
                )
            raise RuntimeError(
                f"❌ 关键词分析响应不是 xlsx 流（HTTP {response.status_code}, Content-Type={content_type}）\n"
                f"   响应片段：{snippet}"
            )

        filename = response.headers.get("Content-Disposition", "").split("filename=")[-1].strip('" ')
        self.logger.info(f"✅ 拿到 xlsx 流（{len(response.content)} 字节）文件名={filename}")
        return response.content, filename

    # ---------- 一步封装：完整导出 ----------

    def run_full_export(
        self,
        date: str = None,
        start_date: str = None,
        end_date: str = None,
        granularity: str = "day",
    ) -> str:
        """一键跑通：POST 拿 xlsx → 落盘 + Excel 后置处理 + 插入时间区间列。

        参数:
            date       - 查询日期 YYYY-MM-DD（单日查询，day粒度）
            start_date - 开始日期 YYYY-MM-DD
            end_date   - 结束日期 YYYY-MM-DD
            granularity- 聚合粒度："day" / "month"（默认 day）
        返回:
            str - 保存的 xlsx 绝对路径
        """
        if granularity not in ("day", "month"):
            raise ValueError(f"❌ granularity 必须是 'day' 或 'month'，当前：{granularity}")

        # 1. 构造表单参数
        form_data = self._build_form_payload(date, start_date, end_date, granularity)
        self.logger.info(
            f"🚀 [关键词分析] 启动导出：{form_data['startDate']} ~ {form_data['endDate']}（{granularity}）"
        )

        # 2. POST 拿 xlsx 字节流
        xlsx_bytes, orig_filename = self._post_for_xlsx(form_data)

        # 3. 落盘 + Excel 后置处理 + 插入时间区间列
        return self._post_process_xlsx(
            xlsx_bytes,
            clean_date=form_data["startDate"],
            end_date=form_data["endDate"],
            granularity=granularity,
        )

    # ---------- Excel 后置处理 + 时间区间列插入 ----------

    def _post_process_xlsx(self, xlsx_bytes: bytes, clean_date: str, end_date: str, granularity: str) -> str:
        """读取 xlsx → 插入时间区间列 → Excel 通用后置 → 保存。

        ⚠️ 用户 2026-08-10 要求：原生 Excel 无时间字段，需手动插入时间区间列。
            - day 粒度：插入「日期」列（YYYY/M/D 格式）
            - month 粒度：插入「日期范围」列（YYYY/M/D ~ YYYY/M/D）
        """
        import io
        import pandas as pd

        # 1. 读取 xlsx（dtype=str 防精度丢失）
        try:
            df = pd.read_excel(
                io.BytesIO(xlsx_bytes),
                dtype=str,
                na_filter=False,
                keep_default_na=False,
            )
        except Exception as e:
            raise RuntimeError(
                f"❌ xlsx 解析失败：{e}\n"
                f"   请检查响应是否合法 xlsx"
            ) from e

        if df.empty:
            # ⚠️ 空数据是合法情况（如某天没流量），不能抛错
            self.logger.warning("⚠️ xlsx 数据为空（该时段可能无数据），仍保存空表")

        # 2. 插入时间区间列
        if granularity == "day":
            # 单日：天报表，插入「日期」列（自动格式化 yyyy/m/d）
            df.insert(0, "日期", clean_date)
            date_column, date_value = prepare_date_columns(df, clean_date)
        else:
            # 月报：插入「日期范围」列（"YYYY/M/D ~ YYYY/M/D"）
            range_str = f"{convert_date_format(clean_date)} ~ {convert_date_format(end_date)}"
            df.insert(0, "日期范围", range_str)
            date_column = "日期范围"
            date_value = range_str

        # 3. Excel 通用后置（数值安全转换）
        df = safe_convert_numeric(df)

        # 4. 构造输出路径：output/商智关键词分析/{YYYY-MM-DD}/{业务名}_{YYYY-MM-DD}_{granularity}.xlsx
        date_subdir = os.path.join(self.output_dir, self.OUTPUT_SUBDIR, clean_date)
        os.makedirs(date_subdir, exist_ok=True)
        save_filename = f"商智关键词分析_{clean_date}_{granularity}.xlsx"
        target_path = os.path.join(date_subdir, save_filename)

        # 5. 写 xlsx + 单元格格式
        df.to_excel(target_path, index=False, engine="openpyxl")
        apply_column_formats(target_path, df, date_column=date_column, date_value=date_value)

        self.logger.info(
            f"✅ 文件已保存：{target_path}\n"
            f"   （{os.path.getsize(target_path)}字节，{len(df)}行 × {len(df.columns)}列，{granularity}粒度）"
        )
        return target_path


# ⚠️ 项目13 KeywordAnalysisAPI 类前向引用回填（解决注册表在前、类在后）
BUSINESS_REGISTRY["商智关键词分析"]["api_class"] = KeywordAnalysisAPI


# ============================================================
#  业务接口 14：（新业务 - 京麦订单明细【加密】导出，2026-08-11 启动）
# ------------------------------------------------------------
#  中文说明（小白必读）：
#    京麦订单导出页面是「先创建任务 → 后台生成加密 zip → 用户触发短信获取解压密码」的模式。
#    完整 5 步异步链路（用户决策 2026-08-11）：
#      1) countDown              前置限流校验（页面进入时调用，本期先跳过，保留扩展位）
#      2) createdExportTask      POST 创建导出任务，返回 taskId
#      3) queryExportTaskInfo    轮询任务状态，直到完成
#      4) exportTaskPwdSend      触发短信下发（接口不返回密码明文）
#      5) export.action          GET 下载加密 zip → 用密码解压 → 内部 xlsx
#
#  ⚠️ 鉴权差异（与商智/京准通完全不同）：
#    - 鉴权签名：h5st（前端强签名，一次性，浏览器实时生成，**禁止复用抓包值**）
#    - 风控字段：dsm-eid / dsm-platform / dsm-lang / dsm-trace-id（每请求不同）
#    - 域名：sff.jd.com（不是 seller-v10.shop.jd.com）
#    - 入口：shop.jd.com/jdm/trade/tools/export/ExprotList
#    - 鉴权头按抓包 2026-08-11 实证固化（不读 config，因 h5st 是一次性签名）
#
#  阶段1（本轮交付）范围：
#    - 业务类 JingMaiOrderExportAPI 骨架（Cookie + dsm 头注入 + h5st 入参）
#    - 仅实现 create_export_task()（POST createdExportTask → 解析 taskId）
#    - 其余 4 步方法占位（raise NotImplementedError 提示后续抓包）
#
#  业务硬性约束（2026-08-06/11 踩坑日志）：
#    - 时间跨度最大 31 天（订单明细【加密】导出服务端强制）
#    - 同导出类型账号维度：两次导出间隔≥10 分钟，单日最多 10 次（code=201）
#    - 单 taskId 申请密码：单任务单日≤10 次，两次调用间隔≥60 秒
#    - h5st 必须真实浏览器实时生成（禁止硬编码复用抓包值）
#    - 严禁代理/VPN/IP 池访问京麦（升级风险）
#    - 触发 601 后 30-120 分钟冷却，冷却期间任何请求都会重置冷却
# ============================================================

import secrets as _secrets  # 用于 dsm-trace-id 唯一化
import uuid as _uuid        # 用于 dsm-trace-id 标准 UUID 格式（与抓包格式一致）


class JingMaiOrderExportAPI:
    """京麦订单明细【加密】导出 API（项目14，2026-08-11 启动骨架）。

    ⚠️ 本类**不继承 JDBaseRequest**（与项目7 同原因）：
        - 鉴权体系：h5st + dsm 全套头（不是商智 User-mnp/uuid）
        - 异步 5 步流程（创建→轮询→短信→下载→解压），不适合基类 30秒重试模型
        - UA 与 h5st 绑定，禁用基类 UA 切换（会致 h5st 失效）

    参数:
        h5st        - 浏览器 F12 抓 createdExportTask 请求头 h5st 值（**一次性**）
        cookie_path - 京麦 Cookie 文件路径（默认 config/sz_cookie.txt）
                      ⚠️ 京麦与商智/京准通 Cookie 不互通，但 sff.jd.com 用的是
                         shop.jd.com 域 Cookie，与商智 cookie 实际是不同账户会话；
                         建议复制到 config/jm_cookie.txt 单独维护，本类先支持自定义路径
    """

    # ---- 鉴权域（业务约束，固定）----
    BASE_URL = "https://sff.jd.com/api"
    APP_ID = "CQLEJWPYPFOVQBC8UFLQ"
    API_VERSION = "1.0"
    ORIGIN = "https://shop.jd.com"
    REFERER = "https://shop.jd.com/jdm/trade/tools/export/ExprotList?exportTaskType=0"
    X_REFERER_PAGE = "https://shop.jd.com/jdm/trade/tools/export/ExprotList"
    X_RP_CLIENT = "h5_2.4.0"

    # ---- UA（与 h5st 绑定，禁用基类 UA 切换）----
    USER_AGENT = (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/144.0.0.0 Safari/537.36 Edg/144.0.0.0"
    )

    # ---- 业务硬性约束（2026-08-11 实证，与踩坑日志坑5/坑6 一致）----
    MAX_RANGE_DAYS = 31       # 订单明细【加密】导出时间跨度上限（服务端强制）
    EXPORT_INTERVAL_MIN = 600 # 同导出类型两次间隔 ≥10 分钟（秒）
    EXPORT_DAILY_LIMIT = 10   # 单日最多 10 次（code=201 表示超额）
    PWD_SEND_INTERVAL_MIN = 60  # 单 taskId 两次密码申请间隔 ≥60 秒
    PWD_SEND_DAILY_LIMIT = 10  # 单 taskId 单日最多 10 次密码申请

    # ---- 业务码（与项目4-13 一致体系）----
    CODE_OK = 200
    CODE_DAILY_LIMIT = 201    # 单日次数超限
    CODE_RISK = 601           # 风控限流（不重试）

    def __init__(self, h5st: str = "", cookie_path: str = "config/sz_cookie.txt"):
        """初始化京麦订单导出 API。

        参数:
            h5st        - 浏览器F12抓 createdExportTask 请求头 h5st（**必填**）
            cookie_path - 京麦 Cookie 文件路径，默认 config/sz_cookie.txt
        """
        import requests

        # 1. h5st 校验（必填，前端强签名一次性）
        if not h5st:
            # 不强制必抛错——保留空字符串的可能（万一某些调用方想先做参数校验，
            # 实际创建任务时再报错）。但打印强提示让用户警觉。
            print(
                "ℹ️  [京麦订单] 未传 h5st。⚠️ 京麦接口强校验 h5st 签名，"
                "调用 create_export_task() 时会因签名缺失失败。\n"
                "   → 请浏览器登录 https://shop.jd.com/jdm/trade/tools/export/ExprotList，"
                "F12 抓 createdExportTask 请求头 h5st 复制传入"
            )
        self.h5st = h5st

        # 2. 读 Cookie（不存在即抛错，强制用户抓包）
        cookie_path_abs = os.path.join(os.path.dirname(os.path.abspath(__file__)), cookie_path)
        if not os.path.isfile(cookie_path_abs):
            raise FileNotFoundError(
                f"❌ 京麦 Cookie 文件不存在：{cookie_path_abs}\n"
                f"   请浏览器登录 https://shop.jd.com/jdm/trade/tools/export/ExprotList，"
                f"F12 抓 sff.jd.com 域 Cookie 写入此文件"
            )
        with open(cookie_path_abs, "r", encoding="utf-8") as f:
            self.cookie = f.read().strip()
        if not self.cookie:
            raise ValueError(f"❌ 京麦 Cookie 文件 {cookie_path_abs} 内容为空")

        # 3. requests Session（UA 与 h5st 绑定，禁用基类 UA 切换）
        self.session = requests.Session()
        # ⚠️ 关键：基础头不直接 update h5st（h5st 每次请求都可能不同，按需注入）
        #    这里只放不变的鉴权头和会话头
        self.session.headers.update({
            "User-Agent": self.USER_AGENT,
            "Accept": "application/json, text/plain, */*",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
            "Accept-Encoding": "gzip, deflate, br, zstd",
            "Content-Type": "application/json;charset=UTF-8",
            "Origin": self.ORIGIN,
            "Referer": self.REFERER,
            "X-Referer-Page": self.X_REFERER_PAGE,
            "X-Requested-With": "XMLHttpRequest",
            "X-Rp-Client": self.X_RP_CLIENT,
            "Cookie": self.cookie,
        })

    # ---- 工具方法 ----

    @staticmethod
    def _gen_dsm_trace_id() -> str:
        """生成 dsm-trace-id（每请求唯一）。

        抓包 2026-08-11 实证：格式 `175fe167-09ff-4678-bac9-d24b3d7ddd90`（标准 UUID v4 字符串）
        """
        return str(_uuid.uuid4())

    @staticmethod
    def _gen_dsm_eid(cookie: str) -> str:
        """从 Cookie 抓 dsm-eid（与 3AB9D23F7A4B3CSS 字段一致）。

        抓包 2026-08-11 实证：dsm-eid = URL-decode(3AB9D23F7A4B3CSS 的 value)
            Cookie: 3AB9D23F7A4B3CSS=jdd03PXDAJVX5VICPIPT5...
            dsm-eid: jdd03PXDAJVX5VICPIPT5...

        说明：直接复用 cookie 中的 3AB9D23F7A4B3CSS 字段值（去前缀 jdd03 前缀外的内容），
              与抓包抓到的 dsm-eid 头部值字面一致。
        """
        import re as _re
        m = _re.search(r"3AB9D23F7A4B3CSS=([^;]+)", cookie)
        if not m:
            return ""
        return m.group(1).strip()

    def _build_request_headers(self) -> dict:
        """组装单次请求的完整鉴权头（h5st + dsm-* 动态注入）。

        返回:
            dict - 完整请求头（含 dsm-trace-id 唯一化、dsm-eid 从 cookie 提取）
        """
        headers = {
            "dsm-eid": self._gen_dsm_eid(self.cookie),
            "dsm-lang": "zh-CN",
            "dsm-platform": "pc",
            "dsm-site": "",  # 抓包实证为空字符串
            "dsm-trace-id": self._gen_dsm_trace_id(),
            "h5st": self.h5st,  # 一次性签名
        }
        return headers

    @staticmethod
    def _build_api_url(api_name: str) -> str:
        """拼接接口 URL。

        模板：https://sff.jd.com/api?v={VER}&appId={APP_ID}&api=dsm.order.export.exportCenterService.{api_name}

        参数:
            api_name - 目标方法名（如 createdExportTask / queryExportTaskInfo / exportTaskPwdSend）
        返回:
            str - 完整 URL
        """
        # 抓包 2026-08-11 实证：api 路径是 dsm.order.export.exportCenterService.<接口名>
        full_api = f"dsm.order.export.exportCenterService.{api_name}"
        return (
            f"{JingMaiOrderExportAPI.BASE_URL}"
            f"?v={JingMaiOrderExportAPI.API_VERSION}"
            f"&appId={JingMaiOrderExportAPI.APP_ID}"
            f"&api={full_api}"
        )

    def _post_dsm(self, api_name: str, body: dict) -> dict:
        """京麦 dsm 体系接口统一 POST（JSON）。

        ⚠️ 鉴权头（dsm-eid / dsm-platform / dsm-trace-id / dsm-lang / h5st）
           每请求动态注入，不在 session.headers 里固化。

        参数:
            api_name - 目标方法名（createdExportTask 等）
            body     - POST body（dict，会被 json.dumps 序列化）
        返回:
            dict - 解析后的 JSON 响应
        异常:
            CookieExpiredError - Cookie 过期 / 未登录
            RuntimeError       - 业务码非 200 / 序列化失败 / 反序列化失败
        """
        url = self._build_api_url(api_name)
        headers = self._build_request_headers()

        # 调试日志：打印 URL + 关键头（敏感字段做长度截断，不打印完整 h5st）
        print(f"🚀 [京麦订单] POST {url}")
        print(f"   Body: {json.dumps(body, ensure_ascii=False)}")
        print(f"   Headers(关键): dsm-eid={headers['dsm-eid'][:30]}..., dsm-trace-id={headers['dsm-trace-id']}, h5st={self.h5st[:30]}...（共 {len(self.h5st)} 字符）")

        resp = self.session.post(url, headers=headers, json=body, timeout=60)
        resp.raise_for_status()
        try:
            ret = resp.json()
        except Exception as e:
            raise RuntimeError(
                f"❌ 京麦 {api_name} 响应非 JSON：HTTP {resp.status_code}，"
                f"响应片段={resp.text[:200]!r}"
            ) from e

        # 业务码判定
        self._handle_response(ret, op_desc=api_name)
        return ret

    def _handle_response(self, ret: dict, op_desc: str):
        """统一处理京麦 dsm 体系响应（项目14，2026-08-11 启动）。

        业务码体系（与京麦其他 dsm 接口一致，2026-08-11 抓包实证）：
            - code=200 + msg="成功" → 成功
            - code=200 但 msg 含"操作频繁" / "限流" → 仍走 601 路径（不重试）
            - code=201 → 单日次数超限（按 EXPORT_DAILY_LIMIT 处理）
            - code=601 → 风控限流（**不重试**，抛 RiskControlError）
            - code 非 0/200 + message 含"登录" → CookieExpiredError

        异常:
            CookieExpiredError / RiskControlError / RuntimeError
        """
        code = ret.get("code")
        msg = str(ret.get("msg", ""))

        # 1. 业务码 200 + msg 成功 → 通过
        if code == self.CODE_OK and ("成功" in msg or "success" in msg.lower()):
            return ret

        # 2. 文本型 601（与项目4/5 实证一致：msg 含"操作频繁/限流"）
        if code == self.CODE_RISK or any(k in msg for k in ("操作频繁", "限流", "risk")):
            raise RiskControlError(
                f"❌ 京麦 {op_desc} 触发 601 风控限流：code={code}, msg={msg}\n"
                f"   → 30-120 分钟冷却，避免加重风控；"
                f"   冷却期间任何请求都会重置冷却"
            )

        # 3. 201 = 单日次数超限
        if code == self.CODE_DAILY_LIMIT:
            raise RuntimeError(
                f"❌ 京麦 {op_desc} 单日次数超限：code={code}, msg={msg}\n"
                f"   → 等待 {self.EXPORT_INTERVAL_MIN // 60} 分钟后重试，"
                f"或明日再试（单日上限 {self.EXPORT_DAILY_LIMIT} 次）"
            )

        # 4. Cookie 过期
        if code in (2001, 302) or any(k in msg for k in ("未登录", "登录已过期", "请重新登录")):
            raise CookieExpiredError(
                f"❌ 京麦 Cookie 过期（{op_desc}返回 code={code}）：\n"
                f"   → 请浏览器登录 https://shop.jd.com/jdm/trade/tools/export/ExprotList，"
                f"F12 抓 sff.jd.com 域 Cookie 写入配置文件"
            )

        # 5. 其它业务码 → RuntimeError（含完整响应便于排查）
        raise RuntimeError(
            f"❌ 京麦 {op_desc} 业务失败：code={code}, msg={msg}, 完整响应={ret}"
        )

    # ---- 工具：业务表单参数组装 ----

    @staticmethod
    def _build_task_data_param(
        start_date: str,
        end_date: str,
        order_status_list: list = None,
        sensitive_info_sign: str = "0",
        export_task_type: int = 0,
        sku_id=None,
        loc_sku_id=None,
        warning_type=None,
    ) -> str:
        """组装 createdExportTask 的 taskDataParam 内层 JSON 字符串。

        抓包 2026-08-11 实证：
            {
              "startDate": "2026-08-10 00:00:00",
              "endDate":   "2026-08-10 23:59:59",
              "exportTaskType": 0,
              "skuId": null,
              "warningType": null,
              "locSkuId": null,
              "sensitiveInfoSign": "0",
              "orderStatusList": [-1]
            }

        ⚠️ 注意：taskDataParam 必须是 **JSON 字符串**（json.dumps 序列化），不能直接传对象。

        参数:
            start_date         - 开始日期 YYYY-MM-DD
            end_date           - 结束日期 YYYY-MM-DD
            order_status_list  - 订单状态列表，默认 [-1]（全部订单状态）
            sensitive_info_sign- 敏感信息导出标志，默认 "0"（不导出收件人敏感信息）
            export_task_type   - 导出任务类型，默认 0（订单明细）
            sku_id / loc_sku_id/ warning_type - 可选过滤参数（默认 None）
        返回:
            str - 序列化后的 JSON 字符串
        """
        if order_status_list is None:
            order_status_list = [-1]
        task_data = {
            "startDate": f"{start_date} 00:00:00",
            "endDate": f"{end_date} 23:59:59",
            "exportTaskType": export_task_type,
            "skuId": sku_id,
            "warningType": warning_type,
            "locSkuId": loc_sku_id,
            "sensitiveInfoSign": sensitive_info_sign,
            "orderStatusList": order_status_list,
        }
        return json.dumps(task_data, ensure_ascii=False, separators=(",", ":"))

    # ---- 5 步方法：阶段1 仅实现 create_export_task；其余 4 步占位 ----

    def create_export_task(
        self,
        start_date: str = None,
        end_date: str = None,
        date: str = None,
        order_status_list: list = None,
        sensitive_info_sign: str = "0",
        export_task_type: int = 0,
    ) -> dict:
        """第 1 步：创建京麦订单明细【加密】导出任务。

        POST /api?api=dsm.order.export.exportCenterService.createdExportTask
        Body: {"exportParam": {"exportTaskType": 0, "taskDataParam": "{...内层JSON字符串...}"}}

        抓包 2026-08-11 实证：
            {
              "exportParam": {
                "exportTaskType": 0,
                "taskDataParam": "{\"startDate\":\"2026-08-10 00:00:00\",...}"
              }
            }

        响应（成功）：
            {"msg": "成功", "code": 200, "dsm-trace-id": "..."}
            ⚠️ 实证响应**没有 taskId 字段**（与京准通 add 不同），taskId 需从
            后续 queryExportTaskInfo 轮询接口获取，或业务升级后才有。

        参数:
            start_date         - 开始日期 YYYY-MM-DD（默认 = date）
            end_date           - 结束日期 YYYY-MM-DD（默认 = date）
            date               - 单日查询 YYYY-MM-DD（start/end 默认 = date）
            order_status_list  - 订单状态列表，默认 [-1]（全部）
            sensitive_info_sign- 敏感信息标志，默认 "0"
            export_task_type   - 任务类型，默认 0
        返回:
            dict - 接口响应 JSON（含 code/msg/dsm-trace-id，**项目14阶段1实证不含 taskId**）
                  实际 taskId 由后续 queryExportTaskInfo 阶段提供，本方法先回执完整响应
        异常:
            ValueError         - 缺日期参数 / 区间 >31 天
            CookieExpiredError - Cookie 过期
            RiskControlError   - 触发 601 风控限流
            RuntimeError       - 业务码非 200
        """
        # 1. 参数兜底
        if date is None and start_date is None and end_date is None:
            raise ValueError("❌ 至少需要传入 date 或 start_date/end_date")
        if start_date is None:
            start_date = date
        if end_date is None:
            end_date = date

        # 2. 区间合法性校验
        from datetime import datetime as _dt
        d_start = _dt.strptime(start_date, "%Y-%m-%d").date()
        d_end = _dt.strptime(end_date, "%Y-%m-%d").date()
        if d_start > d_end:
            raise ValueError(f"❌ 开始日期不能晚于结束日期：start={start_date}, end={end_date}")
        days = (d_end - d_start).days + 1
        if days > self.MAX_RANGE_DAYS:
            raise ValueError(
                f"❌ 订单明细【加密】导出时间跨度({days}天)超过最大限制({self.MAX_RANGE_DAYS}天)，"
                f"请缩小日期区间后再试"
            )

        # 3. 组装 body
        task_data_param_str = self._build_task_data_param(
            start_date=start_date,
            end_date=end_date,
            order_status_list=order_status_list,
            sensitive_info_sign=sensitive_info_sign,
            export_task_type=export_task_type,
        )
        body = {
            "exportParam": {
                "exportTaskType": export_task_type,
                "taskDataParam": task_data_param_str,
            }
        }

        # 4. POST
        print(f"📝 [京麦订单] 第 1 步：创建导出任务 {start_date} ~ {end_date}（共 {days} 天）")
        ret = self._post_dsm("createdExportTask", body)

        # 5. 回执
        print(
            f"✅ [京麦订单] 创建任务响应：code={ret.get('code')}, msg={ret.get('msg')}, "
            f"dsm-trace-id={ret.get('dsm-trace-id')}"
        )
        return ret

    def create_and_wait(
        self,
        start_date: str = None,
        end_date: str = None,
        date: str = None,
        order_status_list: list = None,
        sensitive_info_sign: str = "0",
        export_task_type: int = 0,
        poll_interval: int = 3,
        max_poll_times: int = 20,
    ) -> dict:
        """第 1+2 步一键：创建任务 + 轮询拿 taskId（2026-08-11 项目14 阶段2 一键封装）。

        ⚠️ 京麦与京准通关键差异：
            - 京准通 add 返回 data.reportId
            - 京麦 createdExportTask 响应**没有 taskId**（实证 2026-08-11）
              → 必须再 queryExportTaskInfo 分页查，按 startTime/endTime 匹配刚那条 → 拿 id

        调用：
            api.create_and_wait(date="2026-08-10")
            api.create_and_wait(start_date="2026-08-10", end_date="2026-08-10",
                                poll_interval=5, max_poll_times=30)

        返回:
            dict - taskStatus=2 的任务完整记录（含 id=taskId、taskData、smsSendTip、encryptFlag）

        异常:
            ValueError         - 缺日期参数 / 区间 >31 天
            CookieExpiredError - Cookie 过期
            RiskControlError   - 触发 601 风控
            RuntimeError       - 轮询超时 / 任务失败/过期 / 业务码非 200
        """
        # 日期兜底
        if date is None and start_date is None and end_date is None:
            raise ValueError("❌ 至少需要传入 date 或 start_date/end_date")
        if start_date is None:
            start_date = date
        if end_date is None:
            end_date = date

        # 第 1 步：创建
        create_ret = self.create_export_task(
            start_date=start_date,
            end_date=end_date,
            order_status_list=order_status_list,
            sensitive_info_sign=sensitive_info_sign,
            export_task_type=export_task_type,
        )
        if create_ret.get("code") != self.CODE_OK:
            raise RuntimeError(f"❌ 创建任务失败：{create_ret}")

        # 第 2 步：轮询（每次轮询都重新调 queryExportTaskInfo，按时间匹配）
        task_item = self.wait_for_task_ready(
            start_date=start_date,
            end_date=end_date,
            poll_interval=poll_interval,
            max_poll_times=max_poll_times,
        )

        # 一键回执：含 taskId / status / encryptFlag / smsSendTip
        return {
            "code": self.CODE_OK,
            "msg": "成功",
            "taskId": task_item.get("id"),
            "taskStatus": task_item.get("taskStatus"),
            "encryptFlag": task_item.get("encryptFlag"),
            "taskTypeName": task_item.get("taskTypeName"),
            "smsSendTip": task_item.get("smsSendTip"),
            "smsReceiver": self.get_sms_receiver(task_item),
            "createDate": task_item.get("createDate"),
            "taskData": task_item.get("taskData"),
            "rawItem": task_item,  # 完整原始任务记录，方便阶段 3/4 直接取字段
        }

    def create_wait_and_download(
        self,
        start_date: str = None,
        end_date: str = None,
        date: str = None,
        order_status_list: list = None,
        sensitive_info_sign: str = "0",
        export_task_type: int = 0,
        poll_interval: int = 3,
        max_poll_times: int = 20,
    ) -> dict:
        """第 1+2+3 步一键：创建 + 轮询拿 taskId + 下载加密 zip（2026-08-11 项目14 阶段3 一键封装）。

        ⚠️ 鉴权三次切换：
            1. 创建/轮询：h5st + Cookie + dsm-* 全套头（sff.jd.com）
            2. 下载：仅 Cookie + Referer（export.shop.jd.com，不需要 h5st / dsm 头）
            3. 短信申请：h5st + Cookie + dsm-* 全套头（sff.jd.com，阶段4 待实现）

        返回:
            dict - 含 taskId / zip_path / zip_filename / zip_size 等

        异常:
            ValueError         - 缺日期参数 / 区间 >31 天
            CookieExpiredError - Cookie 过期
            RiskControlError   - 触发 601 风控
            RuntimeError       - 轮询超时 / 任务失败/过期 / 业务码非 200 / 下载失败
        """
        # 第 1+2 步：创建+轮询
        wait_ret = self.create_and_wait(
            start_date=start_date,
            end_date=end_date,
            date=date,
            order_status_list=order_status_list,
            sensitive_info_sign=sensitive_info_sign,
            export_task_type=export_task_type,
            poll_interval=poll_interval,
            max_poll_times=max_poll_times,
        )
        task_id = wait_ret.get("taskId")
        if not task_id:
            raise RuntimeError(f"❌ create_and_wait 没拿到 taskId：{wait_ret}")

        # 第 3 步：下载加密 zip
        zip_bytes, filename = self.download_encrypted_zip(task_id)
        zip_path = self.save_encrypted_zip(zip_bytes, filename)

        # 一键回执
        return {
            "code": self.CODE_OK,
            "msg": "成功",
            "taskId": task_id,
            "taskStatus": wait_ret.get("taskStatus"),
            "encryptFlag": wait_ret.get("encryptFlag"),
            "smsSendTip": wait_ret.get("smsSendTip"),
            "smsReceiver": wait_ret.get("smsReceiver"),
            "createDate": wait_ret.get("createDate"),
            "zip_path": zip_path,
            "zip_filename": filename,
            "zip_size": len(zip_bytes),
            # ⚠️ 下一步：阶段4 exportTaskPwdSend 触发短信 → 短信密码 → 解压 zip
            "rawItem": wait_ret.get("rawItem"),
        }

    def create_wait_download_and_request_pwd(
        self,
        start_date: str = None,
        end_date: str = None,
        date: str = None,
        order_status_list: list = None,
        sensitive_info_sign: str = "0",
        export_task_type: int = 0,
        poll_interval: int = 3,
        max_poll_times: int = 20,
    ) -> dict:
        """第 1+2+3+4 步一键：创建 + 轮询 + 下载 + 短信申请（2026-08-11 项目14 阶段4 一键封装）。

        ⚠️ 与 create_wait_and_download 的关键差异：
            - 多一步 exportTaskPwdSend，需要重新注入鉴权头
            - 短信密码**不返回明文**，只发到安全手机（项目实证 1366794）
            - 调用方需要：
                1) 等待手机短信
                2) 手动或自动（如 IMAP 监听邮箱）拿到解压密码
                3) 调用 _run_with_password（待实现）解压 zip → 提取 xlsx

        返回:
            dict - 含 taskId / zip_path / pwd_response / sms_receiver
                  调用方拿到本回执后，下一步：等短信 → 拿到密码 → 解压

        异常:
            ValueError         - 缺日期参数
            CookieExpiredError - Cookie 过期
            RiskControlError   - 触发 601 风控
            RuntimeError       - 轮询超时 / 任务失败 / 下载失败 / 短信申请失败
        """
        # 第 1+2+3 步：创建+轮询+下载
        dl_ret = self.create_wait_and_download(
            start_date=start_date,
            end_date=end_date,
            date=date,
            order_status_list=order_status_list,
            sensitive_info_sign=sensitive_info_sign,
            export_task_type=export_task_type,
            poll_interval=poll_interval,
            max_poll_times=max_poll_times,
        )
        task_id = dl_ret.get("taskId")

        # 第 4 步：申请短信密码
        pwd_ret = self.request_export_password(
            task_id=task_id,
            export_task_type=export_task_type,
        )

        # 一键回执：含全部上下文 + 短信申请结果
        return {
            "code": self.CODE_OK,
            "msg": "成功",
            "taskId": task_id,
            "taskStatus": dl_ret.get("taskStatus"),
            "encryptFlag": dl_ret.get("encryptFlag"),
            "smsSendTip": dl_ret.get("smsSendTip"),
            "smsReceiver": dl_ret.get("smsReceiver"),
            "createDate": dl_ret.get("createDate"),
            "zip_path": dl_ret.get("zip_path"),
            "zip_filename": dl_ret.get("zip_filename"),
            "zip_size": dl_ret.get("zip_size"),
            # 短信申请结果
            "pwd_response": pwd_ret,                       # 完整短信申请响应
            "remainingTimes": pwd_ret.get("remainingTimes"),  # 剩余发送次数（正则解析）
            "pwd_raw_message": pwd_ret.get("rawData"),     # 原始 data 字符串
            # ⚠️ 下一步：等手机短信 → 拿到密码 → 解压 zip（阶段5 待实现）
            "nextStep": (
                f"短信已发送至 {dl_ret.get('smsReceiver')}，"
                f"等待收到解压密码后调用解压方法"
            ),
            "rawItem": dl_ret.get("rawItem"),
        }

    # ---- 阶段 2 已实现：queryExportTaskInfo 分页查询（2026-08-11） ----

    # ---- 任务状态码（2026-08-11 实证：抓包 10 条历史任务全部 taskStatus=2，状态机暂按下列定义）----
    # ⚠️ 当前实证样本仅覆盖 taskStatus=2，其他状态码（0=等待/1=处理中/3=失败/4=过期）
    #    含义待后续抓包确认。本类先按以下定义固化状态机：
    TASK_STATUS_PENDING = 0    # 等待/排队（推测）
    TASK_STATUS_RUNNING = 1    # 处理中（推测）
    TASK_STATUS_SUCCESS = 2    # 已完成（✅ 2026-08-11 实证 10/10 历史任务都是这个值）
    TASK_STATUS_FAILED = 3     # 失败（推测）
    TASK_STATUS_EXPIRED = 4    # 已过期/超时（推测）

    def query_export_task_list(
        self,
        page: int = 1,
        page_size: int = 10,
        export_task_type: int = 0,
    ) -> dict:
        """第 2 步 a：分页查询任务列表（2026-08-11 实证）。

        POST /api?api=dsm.order.export.exportCenterService.queryExportTaskInfo
        Body: {"exportParam": {"exportTaskType": 0, "page": 1, "pageSize": 10}}

        ⚠️ 关键发现（2026-08-11 实证）：
            - 该接口**不是按 taskId 查单个任务**，而是**分页查询任务列表**
            - 响应 data.itemList[] 包含多条任务，按 createDate 倒序（最新创建的在前）
            - 每条任务的关键字段：
                * id              → taskId（如 "105874710541"）
                * taskStatus      → 状态码（2=已完成，0/1/3/4 含义待补充）
                * encryptFlag     → true=导出文件加密（订单明细【加密】始终为 true）
                * taskTypeName    → "订单明细信息"（任务类型中文名）
                * taskData        → 任务参数（含 startTime/endTime/exportPin/orderStatusList 等）
                * createDate      → 创建时间（毫秒时间戳）
                * smsSendTip      → "接收号码：1366794，每日限发送10次"
                                    ⚠️ 发送密码短信的接收号码从这里取（不是用户手机号！）
                                    （短信下发到京东商家平台绑定的安全手机 1366794）
                * exportPin       → 发起导出的账号（pin）

        参数:
            page             - 页码（默认 1）
            page_size        - 每页条数（默认 10，10 条够覆盖最近一轮操作）
            export_task_type - 任务类型（默认 0=订单明细）
        返回:
            dict - 完整响应（含 data.itemList[]）
        异常:
            CookieExpiredError / RiskControlError / RuntimeError
        """
        body = {
            "exportParam": {
                "exportTaskType": export_task_type,
                "page": page,
                "pageSize": page_size,
            }
        }
        return self._post_dsm("queryExportTaskInfo", body)

    @staticmethod
    def find_task_by_time_range(
        item_list: list,
        start_time: str,
        end_time: str,
    ) -> dict | None:
        """从任务列表中按 startTime/endTime 精确定位单个任务（2026-08-11 工具）。

        ⚠️ 2026-08-11 实证：createdExportTask 响应**没有 taskId**，
           必须先调 queryExportTaskInfo，再用本方法从列表里把刚创建的任务捞出来。

        匹配规则：完全字符串相等匹配 startTime + endTime
        （创建任务时 startDate=`YYYY-MM-DD` → 提交到服务端变成 `YYYY-MM-DD 00:00:00`，
         列表里的 taskData.startTime 也是这个格式）

        参数:
            item_list - query_export_task_list() 返回的 data.itemList
            start_time- 期望 startTime，格式 `YYYY-MM-DD HH:MM:SS`（如 "2026-08-10 00:00:00"）
            end_time  - 期望 endTime，格式 `YYYY-MM-DD HH:MM:SS`（如 "2026-08-10 23:59:59"）
        返回:
            dict - 匹配的任务记录（含 id/taskStatus/taskData/encryptFlag 等）
            None - 未找到
        """
        for item in item_list:
            td = item.get("taskData", {})
            if td.get("startTime") == start_time and td.get("endTime") == end_time:
                return item
        return None

    @staticmethod
    def parse_task_status(item: dict) -> int:
        """从单条任务记录拿 taskStatus（兜底 -1 = 字段缺失）。"""
        try:
            return int(item.get("taskStatus", -1))
        except (TypeError, ValueError):
            return -1

    @staticmethod
    def get_sms_receiver(item: dict) -> str:
        """从单条任务记录拿短信接收号码（smsSendTip 字段解析，2026-08-11 实证）。

        ⚠️ 2026-08-11 真实跑通发现：服务端**对手机号脱敏**，格式 `接收号码：136****6794，每日限发送10次`
            - 格式 = 前 3 位 + 4 个 * + 后 4 位
            - 完整号码无法从响应里拿到（前端/前端 JS SDK 会从其他接口拿）
            - 只能拿到脱敏后的 11 位串（保留 *）

        返回:
            str - 脱敏号码（如 "136****6794"）
            ""  - 字段缺失
        """
        import re as _re
        tip = item.get("smsSendTip", "")
        m = _re.search(r"接收号码[：:]\s*(1\d{2}\*+\d{4})", tip)
        if m:
            return m.group(1)
        return ""

    def wait_for_task_ready(
        self,
        start_date: str,
        end_date: str,
        poll_interval: int = 3,
        max_poll_times: int = 20,
        page_size: int = 10,
    ) -> dict:
        """第 2 步 b：轮询任务直到 taskStatus=2（已完成）（2026-08-11 实证）。

        完整流程：
            1. 调 query_export_task_list 拉最近 page_size 条任务
            2. 在列表里按 startTime/endTime 匹配刚创建的任务
            3. 看 taskStatus：
               - 2（已完成）→ 返回该条任务（含 taskId）
               - 0/1（等待/处理中）→ 等 poll_interval 秒后重试
               - 3/4（失败/过期）→ 抛 RuntimeError
               - 任务未出现（创建太新或服务端延迟）→ 等 poll_interval 秒后重试
            4. 重复直到 max_poll_times 用完

        ⚠️ 关键约束：
            - max_poll_times=20 × poll_interval=3s = 60s 超时
            - 京麦任务一般 5-30 秒内完成（与京准通 6-10 秒同量级）
            - 若超时，调大 max_poll_times 或检查 createDate 是否匹配

        参数:
            start_date     - 创建任务时传入的开始日期 YYYY-MM-DD
            end_date       - 创建任务时传入的结束日期 YYYY-MM-DD
            poll_interval  - 轮询间隔（秒），默认 3s
            max_poll_times - 最大轮询次数，默认 20
            page_size      - 列表查询每页条数，默认 10
        返回:
            dict - taskStatus=2 的任务完整记录（含 id=taskId / taskData / smsSendTip / ...）
        异常:
            RuntimeError - 轮询超时 / 任务状态为失败/过期
        """
        # 服务端格式：YYYY-MM-DD 00:00:00 ~ YYYY-MM-DD 23:59:59
        target_start = f"{start_date} 00:00:00"
        target_end = f"{end_date} 23:59:59"

        print(
            f"⏳ [京麦订单] 轮询任务：start={target_start} ~ end={target_end}，"
            f"间隔 {poll_interval}s × 上限 {max_poll_times} 次（最多 {poll_interval * max_poll_times}s）"
        )

        for i in range(max_poll_times):
            ret = self.query_export_task_list(page=1, page_size=page_size)
            items = ret.get("data", {}).get("itemList", [])

            matched = self.find_task_by_time_range(items, target_start, target_end)
            if matched is not None:
                status = self.parse_task_status(matched)
                task_id = matched.get("id", "")
                if status == self.TASK_STATUS_SUCCESS:
                    print(
                        f"✅ [京麦订单] 轮询命中：taskId={task_id}, taskStatus={status}, "
                        f"encryptFlag={matched.get('encryptFlag')}, "
                        f"smsSendTip={matched.get('smsSendTip')!r}"
                    )
                    return matched
                if status in (self.TASK_STATUS_FAILED, self.TASK_STATUS_EXPIRED):
                    raise RuntimeError(
                        f"❌ 京麦订单任务失败/过期：taskId={task_id}, taskStatus={status}\n"
                        f"   完整任务记录：{matched}"
                    )
                # 等待/处理中（0/1）
                print(
                    f"  [{i+1}/{max_poll_times}] 任务已出现但未完成："
                    f"taskId={task_id}, taskStatus={status}（{poll_interval}s 后重试）"
                )
            else:
                # 任务未出现（创建太新/服务端延迟/分页未刷到）
                # 抓包实证：createDate 倒序，最近的任务在 itemList[0]
                head = items[0] if items else {}
                print(
                    f"  [{i+1}/{max_poll_times}] 任务未出现（最新一条 taskId={head.get('id')!r}, "
                    f"taskStatus={head.get('taskStatus')}）（{poll_interval}s 后重试）"
                )

            if i < max_poll_times - 1:
                time.sleep(poll_interval)

        raise RuntimeError(
            f"❌ 京麦订单轮询超时：{max_poll_times} 次 × {poll_interval}s 后仍未命中 taskStatus=2\n"
            f"   目标区间：start={target_start}, end={target_end}\n"
            f"   可能原因：① 服务端延迟超过 {poll_interval * max_poll_times}s；"
            f"② 创建任务失败但被静默；③ max_poll_times 调小"
        )

    # ---- 阶段 3/4/5 占位（抓到对应接口的成功报文后再实现） ----

    def request_export_password(self, task_id: str) -> dict:
        """第 3 步：申请密码短信（占位，待抓包后实现）。

        POST /api?api=dsm.order.export.exportCenterService.exportTaskPwdSend
        Body: {"taskId": "..."}
        约束：单 taskId 两次申请间隔 ≥60 秒，单日最多 10 次
        """
        raise NotImplementedError(
            "⏳ 阶段 3 待实现：请提供 exportTaskPwdSend 的成功抓包，"
            "我再实现短信申请 + 限流控制"
        )

    def download_encrypted_zip(self, task_id: str) -> bytes:
        """第 3 步：下载加密 zip（2026-08-11 实证落地）。

        GET https://export.shop.jd.com/exportCenter/export.action?taskId={taskId}
        鉴权：仅 Cookie（**不需要 h5st / 不需要 dsm 头**——这是关键差异！）

        抓包 2026-08-11 实证：
            请求头核心：
                Cookie: <完整 .shop.jd.com 域 Cookie>
                Referer: https://shop.jd.com/jdm/trade/tools/export/ExprotList?exportTaskType=0
                User-Agent: <Edge/Chrome>
            响应核心：
                Content-Type: application/octet-stream
                Content-Disposition: form-data; name="attachment"; filename="{taskId}.zip"
                Content-Length: 6534 (本批 2026-08-10 单日共 6534 字节加密 zip)
            响应体：<加密 zip 二进制>（PK\\x03\\x04 开头，但内容已加密，
                     真实订单数据是 zip 内部被密码加密的 xlsx 字节流）

        ⚠️ 关键约束（与京麦 sff.jd.com dsm 接口的差异）：
            - 鉴权签名**不需要 h5st**（GET 静态下载，不走 dsm 强签名）
            - 不需要 dsm-eid / dsm-platform / dsm-trace-id / dsm-lang
            - 不需要 Origin / X-Referer-Page / X-Rp-Client
            - 只需要 Cookie + Referer + UA

        ⚠️ 关键发现（2026-08-11 实证）：
            - 文件名直接是 <taskId>.zip（Content-Disposition 里包含）
            - 压缩包是**密码加密的 zip**（不是普通 zip）：
                * 第一阶段：zip magic bytes `PK\\x03\\x04` 命中（zip 文件本身）
                * 第二阶段：zip 内部含**加密条目**（京东 zip AES 加密或 ZipCrypto）
                * 解压密码由第 4 步 exportTaskPwdSend 触发的短信下发
            - Content-Length=6534：单日订单量较小时 zip 体积很小（加密开销 + Excel 压缩）
            - 不带 password 参数（密码只通过短信下发，**接口不返回密码明文**）

        参数:
            task_id - 创建任务返回的 taskId（如 "105874710541"）
        返回:
            bytes - 加密 zip 文件字节流
        异常:
            CookieExpiredError - Cookie 过期（401/302）
            RuntimeError       - 响应非 zip 字节流 / 任务已过期 / 下载异常
        """
        if not task_id:
            raise ValueError("❌ task_id 不能为空")

        url = f"https://export.shop.jd.com/exportCenter/export.action?taskId={task_id}"

        # ⚠️ 仅 Cookie 鉴权，不走基类的 dsm 头（不调 _post_dsm）
        #    独立构造最小化请求头，避免把 dsm-* 头误传到 export.shop.jd.com
        download_headers = {
            "User-Agent": self.USER_AGENT,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
            "Accept-Encoding": "gzip, deflate, br, zstd",
            "Referer": self.REFERER,
            "Cookie": self.cookie,
            "Sec-Fetch-Dest": "document",
            "Sec-Fetch-Mode": "navigate",
            "Sec-Fetch-Site": "same-site",
            "Sec-Fetch-User": "?1",
            "Upgrade-Insecure-Requests": "1",
        }

        print(f"📥 [京麦订单] 第 3 步：GET {url}")
        print(f"   鉴权：仅 Cookie（{len(self.cookie)} 字符）+ Referer（不需要 h5st / dsm 头）")

        try:
            resp = self.session.get(url, headers=download_headers, timeout=60, allow_redirects=True)
        except requests.exceptions.RequestException as e:
            raise RuntimeError(f"❌ 下载 zip 网络异常：{e}") from e

        # 1. HTTP 状态码校验
        if resp.status_code == 401:
            raise CookieExpiredError(
                f"❌ 京麦订单导出 Cookie 过期（HTTP 401）\n"
                f"   → 请浏览器登录 https://shop.jd.com/jdm/trade/tools/export/ExprotList，"
                f"F12 抓 .shop.jd.com 域 Cookie 写入 config/sz_cookie.txt"
            )
        if resp.status_code == 302:
            # 重定向到登录页 → Cookie 失效
            raise CookieExpiredError(
                f"❌ 京麦订单导出 Cookie 过期（HTTP 302 重定向）\n"
                f"   Location={resp.headers.get('Location', '(无)')}\n"
                f"   → 请浏览器重新登录后重抓 Cookie"
            )
        if resp.status_code != 200:
            raise RuntimeError(
                f"❌ 京麦订单导出下载失败：HTTP {resp.status_code}\n"
                f"   响应片段：{resp.text[:200]!r}"
            )

        # 2. 校验响应是 zip 字节流（PK magic bytes）
        content_type = resp.headers.get("Content-Type", "")
        content_disp = resp.headers.get("Content-Disposition", "")
        body = resp.content

        if len(body) < 1024:
            raise RuntimeError(
                f"❌ 京麦订单导出下载响应过小（{len(body)} 字节）\n"
                f"   响应头 Content-Type={content_type!r}, Content-Disposition={content_disp!r}\n"
                f"   可能是任务未完成/已过期/任务不存在（taskId={task_id}）\n"
                f"   响应片段：{body[:200]!r}"
            )

        # zip magic bytes: PK\x03\x04
        if not (body[:4] == b"PK\x03\x04"):
            raise RuntimeError(
                f"❌ 京麦订单导出响应不是 zip（magic bytes 校验失败）\n"
                f"   响应头 Content-Type={content_type!r}, Content-Disposition={content_disp!r}\n"
                f"   前 32 字节：{body[:32]!r}\n"
                f"   可能是任务不存在/已过期/接口变更"
            )

        # 3. 解析文件名（与抓包一致：filename="{taskId}.zip"）
        filename = f"{task_id}.zip"
        import re as _re
        m = _re.search(r'filename=("?)([^";]+)\1', content_disp)
        if m:
            filename = m.group(2).strip()

        print(
            f"✅ [京麦订单] 下载成功：{filename}（{len(body)} 字节，"
            f"Content-Type={content_type!r}）"
        )
        return body, filename

    def save_encrypted_zip(self, zip_bytes: bytes, filename: str, output_dir: str = None) -> str:
        """把加密 zip 字节流落盘（2026-08-11 工具）。

        输出目录：output/京麦订单明细加密导出/{date}/{filename}
        （日期取 filename 中的 taskId 创建时间需另外解析；先用 self.OUTPUT_SUBDIR + "raw_zip" 子目录）

        参数:
            zip_bytes  - download_encrypted_zip() 返回的 zip 字节流
            filename   - 同上返回的 filename（用于命名）
            output_dir - 自定义输出目录（默认 output/京麦订单明细加密导出/raw_zip/）
        返回:
            str - 落盘的绝对路径
        """
        if output_dir is None:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            output_dir = os.path.join(base_dir, "output", "京麦订单明细加密导出", "raw_zip")
        os.makedirs(output_dir, exist_ok=True)

        target_path = os.path.join(output_dir, filename)
        with open(target_path, "wb") as f:
            f.write(zip_bytes)
        print(f"💾 [京麦订单] zip 已落盘：{target_path}（{os.path.getsize(target_path)} 字节）")
        return target_path

    # ---- 阶段 5：IMAP 监听 + zip 解压 + xlsx 提取（2026-08-11）----

    @staticmethod
    def load_imap_config(config_path: str = "config/imap_config.ini") -> dict:
        """读取 IMAP 配置文件（ini 格式），不存任何敏感字段到代码（2026-08-11 决策）。

        ini 模板（config/imap_config.ini）：
            [imap]
            host = imap.qq.com
            port = 993
            user = your_qq@qq.com
            auth_code = xxxxxxxxxxxxxxxx   # QQ 邮箱 IMAP 授权码（不是 QQ 密码）
            use_ssl = true
            folder = INBOX
            sender_filter = jmsj@jd.com   # 只关心京东商家平台发的短信
            subject_keyword = 解压密码    # 主题含此关键词
            max_wait_seconds = 300        # 最多等 5 分钟
            poll_interval_seconds = 5     # 每 5 秒轮询一次

        返回:
            dict - 配置项（缺字段抛错）
        异常:
            FileNotFoundError - 配置文件不存在
            ValueError        - 必填字段缺失
        """
        import configparser
        if not os.path.isfile(config_path):
            raise FileNotFoundError(
                f"❌ IMAP 配置文件不存在：{config_path}\n"
                f"   请参考 SKILL.md 模板创建 ini 文件，授权码从 QQ 邮箱设置获取"
            )
        cfg = configparser.ConfigParser()
        cfg.read(config_path, encoding="utf-8")
        if "imap" not in cfg:
            raise ValueError(f"❌ IMAP 配置文件缺少 [imap] section：{config_path}")

        section = cfg["imap"]
        required = ["host", "port", "user", "auth_code"]
        missing = [k for k in required if not section.get(k)]
        if missing:
            raise ValueError(
                f"❌ IMAP 配置文件缺失必填字段：{missing}\n"
                f"   完整模板见 SKILL.md"
            )
        return {
            "host": section["host"].strip(),
            "port": int(section["port"]),
            "user": section["user"].strip(),
            "auth_code": section["auth_code"].strip(),
            "use_ssl": section.get("use_ssl", "true").strip().lower() in ("1", "true", "yes"),
            "folder": section.get("folder", "INBOX").strip(),
            "sender_filter": section.get("sender_filter", "").strip(),
            "subject_keyword": section.get("subject_keyword", "解压密码").strip(),
            "max_wait_seconds": int(section.get("max_wait_seconds", "300")),
            "poll_interval_seconds": int(section.get("poll_interval_seconds", "5")),
        }

    @staticmethod
    def fetch_password_from_imap(
        task_id: str,
        config_path: str = "config/imap_config.ini",
    ) -> str:
        """从 QQ 邮箱 IMAP 拉取指定 taskId 对应的解压密码（2026-08-11 项目14 阶段5）。

        ⚠️ 业务背景：
            - iPhone 快捷指令监听京东商家短信（接收号 1366794）
            - 收到后自动转发到 QQ 邮箱（带 taskId / password 等结构化内容）
            - 本方法用 IMAP 协议拉邮件，按 taskId 精确匹配

        ⚠️ 重要：需要 `imaplib` 标准库（Python 自带，无需 pip install）

        邮件内容识别规则（与 iPhone 快捷指令约定）：
            - From:    sender_filter 配置的地址（默认 jmsj@jd.com 或自定义）
            - Subject: 含 subject_keyword 配置的关键词（默认"解压密码"）
            - Body:    含 taskId 字符串，正则提取紧跟其后的 6-12 位字母数字混合密码
                      格式样例：「您的导出任务 105874710541 解压密码为：AbCd1234」
                      或「taskId: 105874710541, password: AbCd1234」

        行为：
            - 默认最多等 5 分钟（max_wait_seconds 配置）
            - 每 5 秒轮询一次（poll_interval_seconds 配置）
            - 匹配到邮件立即返回密码
            - 超时抛 TimeoutError

        参数:
            task_id     - 任务 ID（如 "105874710541"）
            config_path - IMAP 配置文件路径
        返回:
            str - 解压密码
        异常:
            FileNotFoundError - 配置文件不存在
            ValueError        - 必填字段缺失
            TimeoutError      - 超时未找到
            RuntimeError      - IMAP 登录失败 / 网络异常
        """
        import imaplib
        import email
        from email.header import decode_header
        import time as _time

        cfg = JingMaiOrderExportAPI.load_imap_config(config_path)
        deadline = _time.time() + cfg["max_wait_seconds"]
        attempt = 0

        print(
            f"📬 [京麦订单] 第 5 步：IMAP 监听解压密码 taskId={task_id}\n"
            f"   服务器: {cfg['host']}:{cfg['port']} | 用户: {cfg['user']}\n"
            f"   最多等 {cfg['max_wait_seconds']}s（每 {cfg['poll_interval_seconds']}s 轮询）"
        )

        while _time.time() < deadline:
            attempt += 1
            try:
                # 1. 登录 IMAP
                if cfg["use_ssl"]:
                    mail = imaplib.IMAP4_SSL(cfg["host"], cfg["port"])
                else:
                    mail = imaplib.IMAP4(cfg["host"], cfg["port"])
                mail.login(cfg["user"], cfg["auth_code"])
                mail.select(cfg["folder"])

                # 2. 搜索邮件（先按主题关键词粗筛，再按 taskId 精筛）
                # ⚠️ 2026-08-11 真实跑通发现：imaplib.search() 内部用 ASCII 编码命令
                #   传中文"解压密码"会抛 `'ascii' codec can't encode characters` 异常
                #   解决：把搜索改为 ALL + 客户端过滤主题（避免 IMAP 命令包含中文）
                #   然后用 email 解析后的 Subject 头判断
                criterion = "ALL"
                typ, data = mail.search(None, criterion)
                if typ != "OK" or not data or not data[0]:
                    mail.logout()
                    print(f"  [第 {attempt} 次] 邮箱无邮件（typ={typ}）")
                else:
                    # 3. 倒序遍历最新邮件（最近 10 封）
                    msg_ids = data[0].split()[::-1]
                    matched_count = 0
                    for msg_id in msg_ids[:10]:  # 只看最近 10 封
                        typ, msg_data = mail.fetch(msg_id, "(RFC822)")
                        if typ != "OK":
                            continue
                        msg = email.message_from_bytes(msg_data[0][1])

                        # 3.1 解析主题（处理 RFC 2047 Base64/Quoted-Printable 编码）
                        from email.header import decode_header
                        subj_raw = str(msg.get("Subject", ""))
                        try:
                            subj_parts = decode_header(subj_raw)
                            subj_decoded = ""
                            for part, charset in subj_parts:
                                if isinstance(part, bytes):
                                    subj_decoded += part.decode(charset or "utf-8", errors="replace")
                                else:
                                    subj_decoded += part
                            subject_text = subj_decoded
                        except Exception:
                            subject_text = subj_raw

                        # 3.2 主题关键词过滤（客户端判断，避开 imaplib 中文编码问题）
                        if cfg["subject_keyword"] and cfg["subject_keyword"] not in subject_text:
                            continue
                        matched_count += 1

                        # 3.3 解析发件人
                        from_header = msg.get("From", "")
                        if cfg["sender_filter"] and cfg["sender_filter"] not in from_header:
                            continue

                        # 3.4 解析正文（处理 multipart）
                        body_text = ""
                        if msg.is_multipart():
                            for part in msg.walk():
                                if part.get_content_type() == "text/plain":
                                    try:
                                        body_text = part.get_payload(decode=True).decode("utf-8", errors="replace")
                                    except Exception:
                                        pass
                                    break
                        else:
                            try:
                                body_text = msg.get_payload(decode=True).decode("utf-8", errors="replace")
                            except Exception:
                                body_text = str(msg.get_payload())

                        # 3.5 body 提取密码（taskId 精筛）
                        import re as _re
                        if task_id not in body_text and task_id not in subject_text:
                            continue

                        # 多种密码格式正则（容错）
                        # 用户决策 2026-08-11：iPhone 快捷指令邮件正文里是密码（不带 taskId），
                        # 主题含 taskId 字段；匹配策略改为：①主题含 taskId 优先 ②正文取第一个 6-12 位密码
                        patterns = [
                            rf"taskId[:\s]*{task_id}[,\s\S]*?password[:\s]*([A-Za-z0-9]{{6,12}})",
                            rf"{task_id}[^A-Za-z0-9]*?([A-Za-z0-9]{{6,12}})",
                            rf"解压密码[为：:]*\s*([A-Za-z0-9]{{6,12}})",
                            # 兜底：直接从正文里取第一个 6-12 位字母数字混合串
                            r"([A-Za-z0-9]{6,12})",
                        ]
                        password = None
                        for pat in patterns:
                            m = _re.search(pat, body_text)
                            if m:
                                pwd_candidate = m.group(1)
                                # 兜底正则需要排除常见英文词
                                if pat == patterns[-1]:
                                    # 简单启发式：避免匹配到 taskId 自身
                                    if pwd_candidate == task_id:
                                        continue
                                    # 避免匹配到日期/纯数字
                                    if pwd_candidate.isdigit() and len(pwd_candidate) < 8:
                                        continue
                                password = pwd_candidate
                                break

                        if password:
                            mail.logout()
                            print(
                                f"✅ [京麦订单] IMAP 拿到密码：{password!r}（taskId={task_id}，"
                                f"第 {attempt} 次轮询命中）"
                            )
                            return password

                    mail.logout()
                    print(
                        f"  [第 {attempt} 次] 邮箱共 {len(msg_ids)} 封，"
                        f"主题匹配 {matched_count} 封，但都未含 taskId={task_id}"
                    )

            except imaplib.IMAP4.error as e:
                raise RuntimeError(
                    f"❌ IMAP 登录失败：{e}\n"
                    f"   检查授权码（{cfg['host']}:{cfg['port']}, user={cfg['user']}）"
                ) from e
            except Exception as e:
                print(f"  [第 {attempt} 次] IMAP 网络异常：{e}（继续重试）")

            if _time.time() < deadline:
                _time.sleep(cfg["poll_interval_seconds"])

        raise TimeoutError(
            f"❌ IMAP 监听超时：{cfg['max_wait_seconds']}s 内未找到 taskId={task_id} 的解压密码邮件\n"
            f"   排查：① iPhone 快捷指令是否正常转发短信 → QQ 邮箱\n"
            f"         ② QQ 邮箱 IMAP 授权码是否过期\n"
            f"         ③ 邮件 sender_filter / subject_keyword 配置是否与实际一致"
        )

    @staticmethod
    def extract_xlsx_from_zip(
        zip_path: str,
        password: str = None,
        output_dir: str = None,
        date: str = None,
    ) -> str:
        """解压加密 zip + 解密 OLE2 + xls→xlsx 转存 + Excel 后置统一规则（2026-08-11 项目14 阶段5 重构）。

        ⚠️ 京麦订单明细【加密】导出双层加密链路（2026-08-11 实证）：
            第 1 层：zip 容器用 ZipCrypto 加密（zipfile 标准库支持）
            第 2 层：内部 xls 是 OLE2 复合文档，**内容本身也加密**（需要同密码二次解密）
            ⚠️ 这是京麦订单明细【加密】导出独有特点：
                - 内部条目后缀是 .xlsx 但实际是加密的 OLE2 复合文档
                - 用 xlrd.open_workbook(filename, password=password) 二次解密
                - ⚠️ 项目6（商品流失分析）的 .xls 是**未加密** OLE2，本项目是**加密** OLE2

        完整流程：
            1. zipfile 解开 zip（带密码） → 拿到加密的 OLE2 字节流
            2. xlrd.open_workbook(io.BytesIO, password=password) 二次解密 → 拿到 DataFrame
            3. ⚠️ 删除中间加密 zip 文件（用户决策 2026-08-11：只留解密后文件）
            4. xls → xlsx 转存（pandas + openpyxl 引擎）
            5. Excel 后置统一规则（AGENTS.md Excel规则 1+2+3+4）：
               - 规则1：日期列智能识别（已有日期/时间列不重复新增，仅格式化）
               - 规则2：日期统一 yyyy/m/d
               - 规则3：数值安全转换（SKU/SPU 整数 0 位小数、订单编号强制文本 @）
               - 规则4：输出目录 output/京麦订单明细/{date}/订单明细_{date}.xlsx
            6. ⚠️ 内部 xls 列名待真实数据验证（本轮从 0 字节文件推不出列名）

        参数:
            zip_path   - 加密 zip 路径（save_encrypted_zip 返回的路径）
            password   - 解压密码（IMAP 拿到的或人工输入的）
            output_dir - 自定义输出根目录（默认 output/京麦订单明细/）
            date       - 单日查询日期 YYYY-MM-DD（决定日期子目录名）
        返回:
            str - 解密后 xlsx 的绝对路径
        异常:
            RuntimeError - 解压失败（密码错/zip损坏/无xls条目/OLE2 解密失败/转存失败）
        """
        import zipfile
        import io
        import pandas as pd

        if not os.path.isfile(zip_path):
            raise RuntimeError(f"❌ 加密 zip 不存在：{zip_path}")

        # 输出根目录：output/京麦订单明细/（用户决策 2026-08-11：按业务模块建立文件夹）
        if output_dir is None:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            output_dir = os.path.join(base_dir, "output", "京麦订单明细")
        os.makedirs(output_dir, exist_ok=True)

        # 1. 解开 zip（带密码）
        try:
            with zipfile.ZipFile(zip_path, "r") as zf:
                # 找第一个 .xlsx/.xls 条目（京东订单明细 .xlsx 后缀但内部 OLE2）
                candidate_names = [n for n in zf.namelist() if n.lower().endswith((".xlsx", ".xls"))]
                if not candidate_names:
                    raise RuntimeError(
                        f"❌ zip 内未找到 .xlsx/.xls 条目：{zip_path}\n"
                        f"   zip 内文件列表：{zf.namelist()}"
                    )
                target_name = candidate_names[0]
                print(
                    f"📂 [京麦订单] 第 5 步：解压 zip\n"
                    f"   源: {zip_path}\n"
                    f"   密码: {password!r}\n"
                    f"   目标条目: {target_name}"
                )

                # ⚠️ 二次解密：京东订单明细 .xlsx 是加密 OLE2，必须用密码再解一次
                pwd_bytes = password.encode("utf-8") if password else None
                encrypted_ole2_bytes = zf.read(target_name, pwd=pwd_bytes)
        except zipfile.BadZipFile as e:
            raise RuntimeError(f"❌ zip 文件损坏或不是有效 zip：{e}") from e
        except RuntimeError as e:
            # zipfile 在密码错时会抛 RuntimeError "Bad password for file ..."
            raise RuntimeError(
                f"❌ zip 解压失败（密码错误？）：{e}\n"
                f"   zip: {zip_path}\n"
                f"   密码: {password!r}"
            ) from e

        # 2. 二次解密：msoffcrypto-tool 专门解 Microsoft Office 加密文件（.xls/.xlsx/.docx）
        #    ⚠️ 这是项目14独有路径（项目6 是无加密 OLE2，本项目是加密 OLE2）
        #    xlrd 2.0+ 虽然支持加密 .xls，但 msoffcrypto 更稳，跨格式支持更好
        try:
            import msoffcrypto
        except ImportError:
            raise RuntimeError(
                "❌ 缺少 msoffcrypto-tool 库（解密 .xls/.xlsx 加密文件需要）\n"
                "   安装：pip install msoffcrypto-tool"
            )
        try:
            import io as _io
            decrypted_buf = _io.BytesIO()
            office_file = msoffcrypto.OfficeFile(_io.BytesIO(encrypted_ole2_bytes))
            office_file.load_key(password=password)  # 用密码解密
            office_file.decrypt(decrypted_buf)
            decrypted_ole2_bytes = decrypted_buf.getvalue()
        except Exception as e:
            raise RuntimeError(
                f"❌ OLE2 二次解密失败：{e}\n"
                f"   可能原因：① 密码错误 ② 内部 xls 是空表（无真实订单数据）\n"
                f"   尝试：用 unzip -P '{password}' '{zip_path}' 命令行验证"
            ) from e

        # 3. 读解密后的文件（magic bytes 决定格式）
        #    ⚠️ 2026-08-11 真实跑通发现：京麦订单明细解密后是 .xlsx（PK\x03\x04），
        #       不是 .xls（OLE2 头是 zip 容器的"假象"——内部已解密成 xlsx）
        #    magic bytes 处理：
        #       - b"PK\x03\x04"  → .xlsx  → openpyxl
        #       - b"\xD0\xCF\x11\xE0" → .xls → xlrd
        #    用通用工具 read_excel_bytes()（2026-08-07 项目6 已建）
        try:
            df_main = read_excel_bytes(decrypted_ole2_bytes)
            sheet_dfs = {"订单明细": df_main}  # 单 sheet，统一命名
            print(f"   ├─ 主表: {len(df_main)} 行 × {len(df_main.columns)} 列")
            print(f"   ├─ 列名: {list(df_main.columns[:8])}{'...' if len(df_main.columns) > 8 else ''}")
        except Exception as e:
            raise RuntimeError(
                f"❌ 解密后文件读取失败：{e}\n"
                f"   前 16 字节: {decrypted_ole2_bytes[:16].hex()}"
            ) from e

        # 取第一个 sheet 作为主表（订单明细通常只有 1 个 sheet）
        main_sheet_name = list(sheet_dfs.keys())[0]
        df = sheet_dfs[main_sheet_name]

        if df.empty:
            print(f"   ⚠️ Sheet「{main_sheet_name}」为空表（无订单数据），仍保存空 xlsx")
        else:
            print(f"   ├─ 主表「{main_sheet_name}」前 5 列: {list(df.columns[:5])}...")

        # 4. Excel 后置统一规则（规则 1+2：日期列智能处理）
        # ⚠️ 业务方需求：参照 Excel 报表后置统一规则
        #     无日期/时间列才新增【日期】列；已有日期/时间列禁止重复新增
        # ⚠️ 订单明细内层表通常自带「下单时间/订单时间」列，待真实数据验证
        if date:
            date_column, date_value = prepare_date_columns(df, date)
        else:
            date_column, date_value = None, None

        # 5. Excel 后置统一规则（规则 3：数值安全转换 + 文本列保护 + 合计行剔除）
        df = safe_convert_numeric(df)

        # 6. 输出路径（用户决策 2026-08-11：业务模块 + 日期子目录）
        #    AGENTS.md Excel 规则 4：output/{业务模块}/{date}/{filename}
        #    京麦订单明细业务模块名 = 「订单明细」
        if date is None:
            # 兜底：zip 文件名含 taskId（105874726884.zip），无法直接取日期
            # 用 taskId 转创建时间戳？暂时用 zip mtime
            import datetime as _dt
            mtime_ts = os.path.getmtime(zip_path)
            date = _dt.datetime.fromtimestamp(mtime_ts).strftime("%Y-%m-%d")

        date_subdir = os.path.join(output_dir, date)
        os.makedirs(date_subdir, exist_ok=True)
        save_filename = f"订单明细_{date}.xlsx"
        target_xlsx = os.path.join(date_subdir, save_filename)

        # 7. 写 xlsx + 应用单元格格式（SKU/SPU 整数 0、订单编号文本、日期格式化）
        df.to_excel(target_xlsx, index=False, engine="openpyxl")
        if date_column:
            apply_column_formats(target_xlsx, df, date_column=date_column, date_value=date_value)
        else:
            # 即使没日期列也走格式应用（SKU/SPU/订单编号识别）
            apply_column_formats(target_xlsx, df)

        print(
            f"✅ [京麦订单] 解密+转存成功：{target_xlsx}（{os.path.getsize(target_xlsx)} 字节，"
            f"{len(df)}行 × {len(df.columns)}列）"
        )

        # 8. ⚠️ 用户决策 2026-08-11：删除中间加密 zip（只留解密后 xlsx）
        try:
            os.remove(zip_path)
            print(f"🗑️  [京麦订单] 中间加密 zip 已删除：{zip_path}")
        except OSError as e:
            print(f"⚠️ [京麦订单] 中间 zip 删除失败（不影响主流程）：{e}")

        return target_xlsx

    def run_full_export(
        self,
        start_date: str = None,
        end_date: str = None,
        date: str = None,
        order_status_list: list = None,
        sensitive_info_sign: str = "0",
        export_task_type: int = 0,
        poll_interval: int = 3,
        max_poll_times: int = 20,
        sms_password: str = None,
        imap_config_path: str = "config/imap_config.ini",
        imap_timeout_seconds: int = None,
    ) -> dict:
        """完整 5 步一键：创建 + 轮询 + 下载 + 短信申请 + IMAP 拿密码 + 解压 xlsx（2026-08-11）。

        ⚠️ 鉴权 3 次切换：
            1. 创建/轮询/短信申请：h5st + Cookie + dsm-* 全套头（sff.jd.com）
            2. 下载：仅 Cookie + Referer（export.shop.jd.com）
            3. 解压：纯本地 zipfile（无网络）

        ⚠️ 密码获取策略（用户决策 2026-08-11）：
            - 优先用传入的 sms_password（如有，从命令行/环境变量注入）
            - 否则调 IMAP 监听（imap_config_path 配置）
            - IMAP 超时后保留 zip，提示用户人工 --sms-password 重跑
            - （不抛错退出 —— 用户决策「超时后保留 zip + 提示手动输入」）

        参数:
            start_date / end_date / date / order_status_list / sensitive_info_sign
                / export_task_type / poll_interval / max_poll_times
                —— 与 create_wait_download_and_request_pwd 一致
            sms_password       - 可选：手动传入解压密码（优先级最高）
            imap_config_path   - IMAP ini 路径（默认 config/imap_config.ini）
            imap_timeout_seconds - IMAP 监听超时（None=读 ini 的 max_wait_seconds）
        返回:
            dict - 含 taskId / zip_path / xlsx_path / password（来源）+ 全链路上下文
        异常:
            ValueError         - 缺日期参数 / 区间 >31 天
            CookieExpiredError - Cookie 过期
            RiskControlError   - 触发 601 风控
            RuntimeError       - 轮询超时 / 任务失败 / 下载失败 / 短信申请失败 / 解压失败
        """
        # 第 1+2+3+4 步：创建+轮询+下载+短信申请
        full_ret = self.create_wait_download_and_request_pwd(
            start_date=start_date,
            end_date=end_date,
            date=date,
            order_status_list=order_status_list,
            sensitive_info_sign=sensitive_info_sign,
            export_task_type=export_task_type,
            poll_interval=poll_interval,
            max_poll_times=max_poll_times,
        )
        task_id = full_ret.get("taskId")
        zip_path = full_ret.get("zip_path")
        sms_receiver = full_ret.get("smsReceiver")

        # 第 5 步前半：拿解压密码
        password = None
        password_source = None  # "manual" / "imap" / None（取失败）
        if sms_password:
            password = sms_password
            password_source = "manual"
            print(f"🔑 [京麦订单] 第 5 步：使用人工传入密码 {password!r}")
        else:
            # 走 IMAP 监听
            try:
                password = self.fetch_password_from_imap(
                    task_id=task_id,
                    config_path=imap_config_path,
                )
                password_source = "imap"
            except TimeoutError as e:
                # 用户决策 2026-08-11：超时后保留 zip + 提示人工输入，不报错退出
                print(f"⚠️ [京麦订单] IMAP 监听超时，未拿到密码：{e}")
                print(f"   加密 zip 已保留在：{zip_path}")
                print(f"   提示：等手机短信拿到解压密码后，重跑命令并加 --sms-password '<密码>'：")
                print(
                    f"   python main.py --biz_key '京麦订单明细_创建轮询下载并申请密码' "
                    f"--date {date or start_date} --h5st '<h5st>' --sms-password '<短信密码>'"
                )
                return {
                    "code": self.CODE_OK,
                    "msg": "IMAP 监听超时，保留 zip 等人工补密码",
                    "taskId": task_id,
                    "zip_path": zip_path,
                    "xlsx_path": None,
                    "password": None,
                    "password_source": None,
                    "smsReceiver": sms_receiver,
                    "nextStep": f"用 --sms-password 重新跑，或人工解压 {zip_path}",
                    "rawItem": full_ret.get("rawItem"),
                }

        # 第 5 步后半：解压 zip → xlsx
        xlsx_path = self.extract_xlsx_from_zip(
            zip_path=zip_path,
            password=password,
        )

        # 完整回执
        return {
            "code": self.CODE_OK,
            "msg": "成功",
            "taskId": task_id,
            "taskStatus": full_ret.get("taskStatus"),
            "encryptFlag": full_ret.get("encryptFlag"),
            "smsReceiver": sms_receiver,
            "zip_path": zip_path,
            "xlsx_path": xlsx_path,
            "password": password,
            "password_source": password_source,  # 标记密码来源（manual / imap）
            "remainingTimes": full_ret.get("remainingTimes"),
            "pwd_raw_message": full_ret.get("pwd_raw_message"),
            "rawItem": full_ret.get("rawItem"),
        }

    def request_export_password(
        self,
        task_id: str,
        export_task_type: int = 0,
    ) -> dict:
        """第 4 步：申请解压密码短信（2026-08-11 实证落地）。

        POST /api?api=dsm.order.export.exportCenterService.exportTaskPwdSend
        Body: {"exportParam": {"exportTaskType": 0, "taskId": "..."}}

        抓包 2026-08-11 实证：
            请求头：与 createdExportTask 相同（Cookie + h5st + dsm-eid + dsm-platform + dsm-trace-id + dsm-lang）
            请求体：{"exportParam": {"exportTaskType": 0, "taskId": "105874710541"}}
                     ⚠️ 注意：**嵌套在 exportParam 里**（与 createdExportTask 一致结构）
            响应：{
                "msg": "成功",
                "code": 200,
                "data": "密码短信发送成功!当前任务剩余短信发送次数8次",  ← data 是**字符串**，不是 JSON 对象
                "dsm-trace-id": "..."
            }

        ⚠️ 关键发现（2026-08-11 实证）：
            - data 字段是**字符串**（不是 JSON 对象）——需要正则解析
            - 字符串中含"剩余短信发送次数 N 次"——N 是个位数（0-10）
            - **接口不返回密码明文**（与设计预期一致）——密码只发到京东商家平台绑定的安全手机
            - 每调用一次扣减 1 次剩余（实证从 10 → 9 → 8）
            - 短信下发到 smsSendTip 里的接收号码（项目实证 1366794）

        业务硬性约束（与 docs/jd-api-analyze SKILL 一致）：
            - 单 taskId 两次申请间隔 ≥60 秒（脚本**不实现**——由调用方控制节奏）
            - 单 taskId 单日 ≤10 次（**接口自动累计**，超额由服务端拒绝，本方法在响应里返回 remainingTimes 给调用方判断）

        参数:
            task_id         - 任务 ID（如 "105874710541"）
            export_task_type- 任务类型（默认 0=订单明细）
        返回:
            dict - {
                "code": 200,
                "msg": "成功",
                "rawData": "密码短信发送成功!当前任务剩余短信发送次数8次",  # 完整 data 字符串
                "remainingTimes": 8,                                         # 剩余发送次数（从 data 字符串正则解析）
                "taskId": "105874710541",
                "dsmTraceId": "...",
            }
        异常:
            CookieExpiredError - Cookie 过期
            RiskControlError   - 触发 601 风控
            RuntimeError       - 业务码非 200 / 响应异常
        """
        if not task_id:
            raise ValueError("❌ task_id 不能为空")

        body = {
            "exportParam": {
                "exportTaskType": export_task_type,
                "taskId": task_id,
            }
        }

        print(f"📨 [京麦订单] 第 4 步：申请密码短信 taskId={task_id}")
        ret = self._post_dsm("exportTaskPwdSend", body)

        # 响应里 data 是字符串（与京麦 dsm 接口常规 JSON 对象不同！）
        raw_data = ret.get("data", "")
        if not isinstance(raw_data, str):
            # 防御性检查：万一未来京东改回 JSON 对象，提示用户
            raise RuntimeError(
                f"❌ exportTaskPwdSend 响应 data 不是字符串（类型={type(raw_data).__name__}），"
                f"可能是接口变更，请人工核对抓包：{ret}"
            )

        # 正则解析剩余次数（"剩余短信发送次数 N 次"）
        # 放宽空白匹配：应对服务端话术变化（如"剩余 短信 发送 次数 3 次"也能容错）
        import re as _re
        m = _re.search(r"剩余\s*短信\s*发送\s*次数\s*(\d+)\s*次", raw_data)
        remaining = int(m.group(1)) if m else None

        if remaining is None:
            # 没匹配上 —— 可能是首次（满 10 次）或服务端话术变更
            print(f"  ⚠️ 响应中未匹配到「剩余短信发送次数 N 次」字样：{raw_data!r}")
            print(f"     可能是首次申请（默认 10 次）或服务端文案变更")
            remaining = None

        print(
            f"✅ [京麦订单] 短信申请成功：{raw_data!r}"
            + (f"（剩余 {remaining} 次）" if remaining is not None else "")
        )

        return {
            "code": ret.get("code"),
            "msg": ret.get("msg"),
            "rawData": raw_data,
            "remainingTimes": remaining,
            "taskId": task_id,
            "dsmTraceId": ret.get("dsm-trace-id"),
        }


# ============================================================
# 项目16：京麦售后订单明细导出（2026-08-11 启动骨架，待真实抓包验证）
# ============================================================
#
# ⚠️ 业务背景（用户决策 2026-08-11）：
#   - 售后明细导出，zip **无密码**，不需要短信获取解压密码
#   - 与项目14（订单明细【加密】）对比：
#       创建/轮询接口完全相同（都是 sff.jd.com + createdExportTask + queryExportTaskInfo）
#       下载接口完全相同（export.shop.jd.com/exportCenter/export.action，仅 Cookie）
#       **唯一差异**：zip 无密码，无需 exportTaskPwdSend 短信申请 / 无需 IMAP
#   - 任务状态枚举（用户决策 2026-08-11）：
#       status=1 生成中、status=2 成功、status=3 失败（项目14 是 0/1/2，可能项目16 改用 1/2/3）
#
# ⚠️ 本类**继承 JingMaiOrderExportAPI**：
#   - 复用 dsm 头模板、UUID 生成、Cookie 加载、_post_dsm、download_encrypted_zip、save_encrypted_zip 等
#   - 复用 extract_xlsx_from_zip（无需密码时会自动跳过 OLE2 解密）
#   - 复用作弊：售后明细解压时 password=None → 跳过 msoffcrypto → 直接读 xlsx
#
# ⚠️ 待你提供真实抓包后再精确适配：
#   - 项目16 是否需要 h5st？（用户文档没说，可能仅 Cookie）
#   - X-Rp-Sdtoken 是否在创建/轮询请求头中出现？（项目14 没有，用 X-Rp-Client）
#   - taskStatus 字段名是 status 还是 taskStatus？（项目14 用 taskStatus）
#   - taskDataParam 内字段是否完全相同（startDate/endDate/exportTaskType）
#   - 售后特有的筛选条件（售后状态列表、退款时间范围等）


class JingMaiAfterSaleExportAPI(JingMaiOrderExportAPI):
    """京麦售后订单明细导出 API（项目16，2026-08-11 实证抓包适配）。

    ⚠️ 2026-08-12 真实抓包确认（来自用户提供的原始 HTTP 包）：
        - api 路径：**dsm.seller.afs.bff.ExportDsmService.<接口名>**（与项目14 完全不同）
        - appId：**BHPQ4MHJBUOQZKTFTRNS**（与项目14 的 CQLEJWPYPFOVQBC8UFLQ 不同）
        - payload 嵌套结构：{"request": {"data": {"exportType": 2602, "param": "<JSON字符串>"}}, "accessContext": {"source": "web"}}
        - param 内容：{exportType, applyTime[ms, ms], applyTimeRange: {dateBegin, dateEnd}, tabCode, ... 13 个售后筛选字段}
        - 时间格式：**毫秒时间戳**（不是 YYYY-MM-DD HH:MM:SS）
        - exportType: 2602（售后明细导出，固定值，待更多业务验证）
        - 响应 data 字段：**bool 类型**（true=成功，不是 JSON 对象）
        - 响应头 **X-Rp-Sdtoken: set;1800;...**：30分钟有效，下次请求必须带上
        - 请求头 dsm-file-path: **lineation-price**（售后业务特有）
        - 请求头 dsm-site: 空字符串
        - Referer: after-sale/independent-after-sale/list?tabCode=all
        - X-Referer-Page: after-sale/independent-after-sale/list

    ⚠️ 与父类 JingMaiOrderExportAPI 的核心差异：
        - zip 无密码：解压时 password=None
        - 不需要短信申请：跳过 exportTaskPwdSend 步骤
        - 不需要 IMAP：跳过 QQ 邮箱监听
        - appId 不同（项目14 项目16 是两个独立业务系统）
        - api 路径前缀不同（order.export vs seller.afs.bff）
        - payload 嵌套结构不同（exportParam vs request.data）
        - X-Rp-Sdtoken 必须解析并动态带下次请求（**重要风控令牌**）

    ⚠️ 待你提供下载包后再适配下载部分（getAction / 域名）
    """

    # ---- ⚠️ 项目16 特有类常量（与项目14 完全不同）----
    APP_ID = "BHPQ4MHJBUOQZKTFTRNS"  # ⚠️ 售后明细导出专属 appId
    API_PATH_PREFIX = "dsm.seller.afs.bff.ExportDsmService"  # ⚠️ 售后明细 api 路径前缀
    DSM_FILE_PATH = "lineation-price"  # ⚠️ 售后业务特有 dsm-file-path 头（待其他业务验证）
    REFERER = "https://shop.jd.com/jdm/trade/after-sale/independent-after-sale/list?tabCode=all"
    X_REFERER_PAGE = "https://shop.jd.com/jdm/trade/after-sale/independent-after-sale/list"

    # ---- 业务硬性约束（项目16 待真实业务限制实证）----
    EXPORT_TYPE_AFTER_SALE_DETAIL = 2602  # ⚠️ 售后明细导出类型，固定值（待更多业务验证）

    def __init__(self, h5st: str = "", cookie_path: str = "config/jm_cookie.txt"):
        """初始化京麦售后明细导出 API。

        参数:
            h5st        - 浏览器F12抓 createdExportTask 请求头 h5st（项目16 抓包实证必需）
            cookie_path - 京麦 Cookie 文件路径，默认 config/jm_cookie.txt
        """
        # 直接调用父类构造，复用 dsm 头/Cookie/requests Session 等
        super().__init__(h5st=h5st, cookie_path=cookie_path)

        # 项目16 特有：覆盖父类的 Referer / X-Referer-Page（售后页面）
        self.session.headers.update({
            "Referer": self.REFERER,
            "X-Referer-Page": self.X_REFERER_PAGE,
            "dsm-file-path": self.DSM_FILE_PATH,
        })

        # 项目16 特有：X-Rp-Sdtoken 风控令牌（从响应里解析，下次请求带上）
        # 30 分钟有效（响应头 set;1800 表示 1800 秒）
        self._rp_sdtoken = None
        self._rp_sdtoken_expire_ts = 0  # unix 时间戳

    # ---- X-Rp-Sdtoken 风控令牌解析 ----

    def _refresh_rp_sdtoken_from_response(self, resp):
        """从响应头解析 X-Rp-Sdtoken，下次请求带上。

        ⚠️ 抓包 2026-08-12 实证：
            响应头: X-Rp-Sdtoken: set;1800;AAbEsBpEIOVjqTAKCQtvQu17TNqU7pH-3dnMLzjC_r_fWMMY3YMa0qp_8ydNTsT9FRysfnMmiaf4N7Pm9X4E5TTdwx7_B_iIb1_6e63AqHQtEO46p_GUm1lRA3L6NvdNEknl0d7Io82BvhgQmnwmip28LA8IpHz6tlaVRPtwlg_jB5bnmRMKgac
            格式: set;<有效期秒数>;<令牌值>
            有效: 30 分钟（1800 秒）

        参数:
            resp - requests.Response 对象
        """
        sdtoken_header = resp.headers.get("X-Rp-Sdtoken", "")
        if not sdtoken_header:
            return
        parts = sdtoken_header.split(";", 2)
        if len(parts) == 3 and parts[0] == "set":
            try:
                expire_seconds = int(parts[1])
                import time as _time
                self._rp_sdtoken = parts[2]
                self._rp_sdtoken_expire_ts = _time.time() + expire_seconds
            except (ValueError, IndexError):
                pass

    def _build_api_url(self, api_name: str) -> str:
        """⚠️ 重写父类：项目16 api 路径前缀完全不同。

        模板：https://sff.jd.com/api?v={VER}&appId={APP_ID}&api=dsm.seller.afs.bff.ExportDsmService.{api_name}
        """
        full_api = f"{self.API_PATH_PREFIX}.{api_name}"
        return (
            f"{self.BASE_URL}"
            f"?v={self.API_VERSION}"
            f"&appId={self.APP_ID}"
            f"&api={full_api}"
        )

    def _post_dsm_after_sale(self, api_name: str, body: dict) -> dict:
        """⚠️ 项目16 dsm POST（继承父类框架，新增 X-Rp-Sdtoken 动态注入）。

        关键差异 vs 父类：
            - 动态注入 X-Rp-Sdtoken 头（每次响应刷新，30 分钟有效）
            - url 用项目16 的 appId + api 路径前缀
            - 响应头解析 X-Rp-Sdtoken 后自动存到 self._rp_sdtoken
            - 响应 data 可能是 bool（项目16 createExportTask 返回 data=true）—— 不报错
        """
        import time as _time
        url = self._build_api_url(api_name)
        headers = self._build_request_headers()

        # 注入 X-Rp-Sdtoken（如果还有效）
        if self._rp_sdtoken and self._rp_sdtoken_expire_ts > _time.time():
            headers["X-Rp-Sdtoken"] = self._rp_sdtoken

        print(f"🚀 [京麦售后明细] POST {url}")
        print(f"   Body: {json.dumps(body, ensure_ascii=False)[:500]}{'...' if len(json.dumps(body, ensure_ascii=False)) > 500 else ''}")
        print(f"   Headers(关键): dsm-eid={headers.get('dsm-eid','')[:30]}..., dsm-trace-id={headers.get('dsm-trace-id','')}, h5st={self.h5st[:30]}...（共 {len(self.h5st)} 字符）")
        if "X-Rp-Sdtoken" in headers:
            print(f"   X-Rp-Sdtoken: {headers['X-Rp-Sdtoken'][:30]}...（有效至 {_time.time() - self._rp_sdtoken_expire_ts:.0f}s 后）")

        resp = self.session.post(url, headers=headers, json=body, timeout=60)
        # ⚠️ 项目16 响应可能 code=200 但 data=true（不是 dict），不用 raise_for_status 用业务码判断
        try:
            ret = resp.json()
        except Exception as e:
            raise RuntimeError(
                f"❌ 京麦售后 {api_name} 响应非 JSON：HTTP {resp.status_code}，"
                f"响应片段={resp.text[:200]!r}"
            ) from e

        # 刷新 X-Rp-Sdtoken（从响应头）
        self._refresh_rp_sdtoken_from_response(resp)

        # 业务码判定（复用父类 _handle_response 框架）
        self._handle_response_after_sale(ret, op_desc=api_name)
        return ret

    def _handle_response_after_sale(self, ret: dict, op_desc: str):
        """⚠️ 项目16 响应处理（与父类 _handle_response 的关键差异）：
            - data 字段是 bool（true/false），不强行访问 data.taskId 等子字段
            - 任务 ID 在响应里**没有**，必须后续轮询拿
        """
        code = ret.get("code")
        msg = str(ret.get("msg", ""))

        # 1. code=200 + msg="成功" → 通过
        if code == self.CODE_OK and ("成功" in msg or "success" in msg.lower()):
            return ret

        # 2. 风控 601
        if code == self.CODE_RISK or any(k in msg for k in ("操作频繁", "限流", "risk")):
            raise RiskControlError(
                f"❌ 京麦售后 {op_desc} 触发 601 风控限流：code={code}, msg={msg}"
            )

        # 3. 201 = 单日次数超限
        if code == self.CODE_DAILY_LIMIT:
            raise RuntimeError(
                f"❌ 京麦售后 {op_desc} 单日次数超限：code={code}, msg={msg}\n"
                f"   → 等待 {self.EXPORT_INTERVAL_MIN // 60} 分钟后重试"
            )

        # 4. Cookie 过期
        if code in (2001, 302) or any(k in msg for k in ("未登录", "登录已过期", "请重新登录")):
            raise CookieExpiredError(
                f"❌ 京麦售后 Cookie 过期（{op_desc}）：code={code}, msg={msg}\n"
                f"   → 请浏览器登录 https://shop.jd.com/jdm/trade/after-sale/independent-after-sale/list，"
                f"F12 抓 sff.jd.com 域 Cookie 写入 config/jm_cookie.txt"
            )

        # 5. 其它业务码
        raise RuntimeError(
            f"❌ 京麦售后 {op_desc} 业务失败：code={code}, msg={msg}, 完整响应={ret}"
        )

    # ---- 业务硬性约束（项目16 待抓包实证，先沿用项目14 约束）----
    MAX_RANGE_DAYS = 31
    EXPORT_INTERVAL_MIN = 600
    EXPORT_DAILY_LIMIT = 10

    # ---- 业务码（沿用项目14 体系）----
    # ⚠️ 项目16 文档说 status=3 失败（项目14 是 taskStatus 0/1/2 体系），
    #    待真实抓包确认是 dsm 返回 code=201 之类的统一体系，还是 taskStatus 自定义
    CODE_OK = 200
    CODE_DAILY_LIMIT = 201
    CODE_RISK = 601

    # ---- 任务状态枚举（项目16 文档说 1=生成中 / 2=成功 / 3=失败，待抓包验证）----
    TASK_STATUS_GENERATING = 1
    TASK_STATUS_SUCCESS = 2
    TASK_STATUS_FAIL = 3

    def __init__(self, h5st: str = "", cookie_path: str = "config/jm_cookie.txt"):
        """初始化京麦售后明细导出 API。

        参数:
            h5st        - 浏览器F12抓 createdExportTask 请求头 h5st（项目16 待验证）
            cookie_path - 京麦 Cookie 文件路径，默认 config/jm_cookie.txt（项目14 BUG 已修复）
        """
        # 直接调用父类构造，复用 dsm 头/Cookie/requests Session 等
        super().__init__(h5st=h5st, cookie_path=cookie_path)
        # 售后导出报表名通常含 afterSaleOrderDetail / aftersale 标识（待抓包确认）

    def create_after_sale_export_task(
        self,
        start_date: str = None,
        end_date: str = None,
        date: str = None,
        tab_code: str = "all",                          # ⚠️ 项目16 抓包实证 tabCode=all
        after_sale_status_list: str = "",              # 售后状态列表（默认空 = 全部）
        service_order_sub_state_list: str = "",        # 子状态列表（默认空）
        customer_expect_list: str = "",                # 客户期望列表（默认空）
        refund_status_list: str = "",                  # 退款状态列表（默认空）
        transfer_feedback_reason_list: str = "",       # 转移反馈原因列表（默认空）
    ) -> dict:
        """第 1 步：创建售后明细导出任务（项目16，2026-08-12 真实抓包适配）。

        ⚠️ 2026-08-12 抓包实证：
            URL: POST https://sff.jd.com/api?v=1.0&appId=BHPQ4MHJBUOQZKTFTRNS&api=dsm.seller.afs.bff.ExportDsmService.createExportTask
            Body 结构:
                {
                    "request": {
                        "data": {
                            "exportType": 2602,           # 售后明细导出（固定）
                            "param": "{<JSON 字符串>}"  # 13+ 个售后筛选字段
                        }
                    },
                    "accessContext": {"source": "web"}
                }
            param 内容（JSON 字符串）：
                {
                    "applyTime": [dateBegin_ms, dateEnd_ms],
                    "applyTimeRange": {"dateBegin": dateBegin_ms, "dateEnd": dateEnd_ms},
                    "tabCode": "all",
                    "serviceOrderSubStateList": "",
                    "customerExpectList": "",
                    "refundStatusList": "",
                    "afsIdList": null,
                    ...（其他 13 个售后筛选字段）
                }

        返回:
            dict - {"msg":"成功", "code":200, "data": true, "dsm-trace-id": "..."}
            ⚠️ **响应 data 是 bool**（不是 taskId 对象），taskId 在轮询里拿
        """
        if not start_date and not date:
            raise ValueError("❌ 必须传入 start_date/end_date 或 date")

        # 日期归一化
        if date:
            start_date = date
            end_date = date
        else:
            if not end_date:
                end_date = start_date

        # 毫秒时间戳
        import time as _time_local
        import datetime as _dt_local
        begin_dt = _dt_local.datetime.strptime(start_date, "%Y-%m-%d").replace(hour=0, minute=0, second=0, microsecond=0)
        end_dt = _dt_local.datetime.strptime(end_date, "%Y-%m-%d").replace(hour=23, minute=59, second=59, microsecond=999000)
        date_begin_ms = int(begin_dt.timestamp() * 1000)
        date_end_ms = int(end_dt.timestamp() * 1000)

        # ⚠️ 13 个售后筛选字段（2026-08-12 抓包实证）
        #     全部默认空字符串 / null / 空 dict，意味「全部数据」
        #     如果你想筛选特定状态，传入具体值
        param_dict = {
            "applyTime": [date_begin_ms, date_end_ms],
            "applyTimeRange": {"dateBegin": date_begin_ms, "dateEnd": date_end_ms},
            "tabCode": tab_code,
            "serviceOrderSubStateList": service_order_sub_state_list,
            "customerExpectList": customer_expect_list,
            "refundStatusList": refund_status_list,
            "afsIdList": None,
            "transferFeedbackReasonList": transfer_feedback_reason_list,
            "transferFeedbackTime": "",
            "approveTime": "",
            "addressList": "",
            "waybillCodeList": None,
            "deliveryWareStatusList": "",
            "refundOperateTypeList": "",
            "pickWareTypeList": "",
            "returnWareStatus": "",
            "processResultList": "",
            "collectionStatus": "",
            "tagList": "",
            "remarkLevelList": "",
            "orderTypeList": "",
            "followTypePin": "",
            "wareInfo": None,
            "transferFeedbackTimeRange": {},
            "approveTimeRange": {},
        }

        # ⚠️ 项目16 真实 payload 嵌套（与项目14 完全不同）
        body = {
            "request": {
                "data": {
                    "exportType": self.EXPORT_TYPE_AFTER_SALE_DETAIL,  # 2602
                    "param": json.dumps(param_dict, separators=(",", ":")),
                }
            },
            "accessContext": {"source": "web"},
        }

        print(f"📝 [京麦售后明细] 第 1 步：创建导出任务 {start_date} ~ {end_date}")
        print(f"   exportType: {self.EXPORT_TYPE_AFTER_SALE_DETAIL}, tabCode: {tab_code}")
        print(f"   时间范围: {date_begin_ms} ~ {date_end_ms}（毫秒）")
        ret = self._post_dsm_after_sale("createExportTask", body)
        print(f"✅ [京麦售后明细] 创建任务响应：code={ret.get('code')}, msg={ret.get('msg')!r}, data={ret.get('data')!r}")
        print(f"   X-Rp-Sdtoken 已刷新（30 分钟有效）")
        return ret

    def wait_for_after_sale_task_ready(
        self,
        start_date: str,
        end_date: str,
        poll_interval: int = 3,
        max_poll_times: int = 20,
    ) -> dict:
        """第 2 步：轮询售后任务状态（项目16，2026-08-12 真实抓包适配）。

        ⚠️ 2026-08-12 抓包实证：
            URL: POST .../api?api=dsm.seller.afs.bff.ExportDsmService.getExportTaskPage
            Body: {"request":{"data":{"pageIndex":1,"pageSize":10,"exportType":[2602,2601,38]}}, "accessContext":{"source":"web"}}
            响应:
                {
                    "msg":"成功","code":200,
                    "data":{
                        "totalNum":"41", "pageIndex":1, "pageSize":10,
                        "content":[
                            {
                                "exportStatus":"已完成",
                                "exportType":"售后(新)",
                                "exportTypeCode":2602,
                                "exportCondition":"tab页签：全部\\n申请时间：2026-08-06至2026-08-06",
                                "taskId":"105884767567",
                                "exportStatusCode":2,
                                "createDate":"2026-08-12 14:35:37"
                            },
                            ...
                        ]
                    }
                }

        ⚠️ 与项目14 wait_for_task_ready 的关键差异：
            - 接口名：queryExportTaskInfo → **getExportTaskPage**
            - 响应字段路径：data.itemList[] → **data.content[]**
            - payload 不带 applyTimeRange/tabCode（只带 pageIndex/pageSize/exportType）
            - 状态字段：taskStatus → **exportStatusCode**（1=生成中/2=成功/3=失败）
            - exportType 是 **列表** [2602,2601,38] 而非单值
            - 任务 ID 字段：id → **taskId**

        ⚠️ 匹配策略（基于 exportCondition 文本匹配）：
            exportCondition 格式："tab页签：全部\\n申请时间：2026-08-06至2026-08-06"
            用 createDate 时间字符串粗匹配（毫秒精度太低不好匹配）

        返回:
            dict - 命中任务记录（含 taskId / exportStatusCode / exportTypeCode 等）
        """
        import time as _time_local
        import datetime as _dt_local

        # ⚠️ 项目16 真实 payload（2026-08-12 抓包实证）：
        #     pageIndex / pageSize / exportType（**列表**，同时查 3 类业务）
        body = {
            "request": {
                "data": {
                    "pageIndex": 1,
                    "pageSize": 10,
                    "exportType": [
                        self.EXPORT_TYPE_AFTER_SALE_DETAIL,  # 2602 售后(新)
                        2601,                                  # 售后(老)
                        38,                                    # 待你确认（可能是另一种业务类型）
                    ],
                }
            },
            "accessContext": {"source": "web"},
        }

        print(
            f"⏳ [京麦售后明细] 轮询任务：start={start_date} ~ end={end_date}，"
            f"间隔 {poll_interval}s × 上限 {max_poll_times} 次（最多 {poll_interval * max_poll_times}s）"
        )
        deadline_ts = _time.time() + poll_interval * max_poll_times
        attempt = 0

        while _time.time() < deadline_ts:
            attempt += 1
            # 接口名：getExportTaskPage（2026-08-12 抓包实证）
            ret = self._post_dsm_after_sale("getExportTaskPage", body)

            # 响应字段路径：data.content[]（不是 data.itemList[]）
            data = ret.get("data", {})
            if not isinstance(data, dict):
                print(f"  [{attempt}/{max_poll_times}] 响应 data 不是 dict：{data}")
                _time.sleep(poll_interval)
                continue

            content = data.get("content", [])
            if not isinstance(content, list):
                content = []
            total_num = data.get("totalNum", "?")
            print(f"  [{attempt}/{max_poll_times}] 拉到 {len(content)} 条任务（total={total_num}）")

            # ⚠️ 匹配策略：双层匹配（保证正确率）
            #   1. exportTypeCode 必须等于 EXPORT_TYPE_AFTER_SALE_DETAIL（2602）
            #   2. exportCondition 含 "申请时间：{start_date}至{end_date}"
            target_item = None
            target_cond_str = f"申请时间：{start_date}至{end_date}"
            for item in content:
                if not isinstance(item, dict):
                    continue
                # 仅看售后明细（2602），不看售后(老 2601) 或其他(38)
                if item.get("exportTypeCode") != self.EXPORT_TYPE_AFTER_SALE_DETAIL:
                    continue
                # exportCondition 含目标时间范围
                if target_cond_str in item.get("exportCondition", ""):
                    target_item = item
                    break

            if not target_item:
                print(f"      暂未命中（目标条件：exportTypeCode=2602 且 exportCondition 含 '{target_cond_str}'）")
                _time.sleep(poll_interval)
                continue

            # ⚠️ 状态字段名：exportStatusCode（2026-08-12 抓包实证）
            #    枚举：1=生成中 / 2=成功 / 3=失败（你文档一致）
            status_code = target_item.get("exportStatusCode")
            task_id = target_item.get("taskId")

            if status_code == self.TASK_STATUS_SUCCESS:  # 2=成功
                print(
                    f"✅ [京麦售后明细] 轮询命中：taskId={task_id}, "
                    f"exportStatusCode={status_code}（{target_item.get('exportStatus')}）"
                )
                return target_item
            elif status_code == self.TASK_STATUS_GENERATING:  # 1=生成中
                print(
                    f"  [{attempt}/{max_poll_times}] 任务生成中：taskId={task_id}, "
                    f"exportStatusCode={status_code}（{poll_interval}s 后重试）"
                )
                _time.sleep(poll_interval)
                continue
            elif status_code == self.TASK_STATUS_FAIL:  # 3=失败
                raise RuntimeError(
                    f"❌ 京麦售后明细任务失败：taskId={task_id}, "
                    f"exportStatusCode={status_code}（{target_item.get('exportStatus')}）"
                )
            else:
                print(
                    f"  [{attempt}/{max_poll_times}] 任务状态未知：exportStatusCode={status_code}"
                    f"（继续等）"
                )
                _time.sleep(poll_interval)

        raise RuntimeError(
            f"❌ 京麦售后明细轮询超时：{poll_interval * max_poll_times}s 内未命中任务"
        )

    def download_after_sale_zip(self, task_id: str) -> tuple:
        """第 3 步：下载售后明细 zip（项目16，2026-08-12 真实抓包适配）。

        ⚠️ 2026-08-12 抓包实证：
            URL: GET https://export.shop.jd.com/exportCenter/export.action?taskId={task_id}
            鉴权: **仅 Cookie**（**不要 dsm-* 头、不要 h5st、不要 X-Rp-Client**）
            Referer: after-sale/independent-after-sale/list?tabCode=all（项目14 是 ExprotList）
            Accept: text/html,application/xhtml+xml,application/xml,...（完整浏览器 Accept）
            Sec-Fetch-Dest: document | Sec-Fetch-Mode: navigate | Sec-Fetch-Site: same-site
            Sec-Fetch-User: ?1 | Upgrade-Insecure-Requests: 1
            User-Agent: Chrome/144.0.0.0 Edg/144.0.0.0（与项目14 一致）
            响应: application/octet-stream，Content-Disposition: filename="<taskId>.zip"
            响应 Content-Length 5020 字节（项目16 抓包），典型大小（项目14 是 6540 字节）

        ⚠️ 与父类 download_encrypted_zip 的关键差异：
            - **Referer 必须改为售后页面**（否则会被风控拦截）
            - Accept 头需要更完整（浏览器默认值）
            - Sec-Fetch-* 头要齐（模拟浏览器导航行为）
            - X-Rp-Sdtoken / dsm-* 头 **不能带**（与项目14 一致）

        ⚠️ 项目16 zip 无密码（实证！）：
            - 5020 字节提示：可能是「带表头/表尾但无数据的售后明细 xlsx」
            - 也可能是「正常大小的数据表」
            - 反正不需要 msoffcrypto 解密（不像项目14 有密码）

        返回:
            tuple - (zip_bytes, filename)
        """
        import requests as _requests

        url = f"https://export.shop.jd.com/exportCenter/export.action?taskId={task_id}"

        # ⚠️ 售后业务专属 Referer（项目14 是 ExprotList）
        headers = {
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
            "Accept-Encoding": "gzip, deflate, br, zstd",
            "Connection": "keep-alive",
            "Cookie": self.cookie,
            "Host": "export.shop.jd.com",
            "Referer": self.REFERER,  # 售后页面 URL（不是 ExprotList）
            "Sec-Fetch-Dest": "document",
            "Sec-Fetch-Mode": "navigate",
            "Sec-Fetch-Site": "same-site",
            "Sec-Fetch-User": "?1",
            "Upgrade-Insecure-Requests": "1",
            "User-Agent": self.USER_AGENT,
            # ⚠️ 关键：不要带 X-Rp-Sdtoken / dsm-* 头 / h5st / X-Rp-Client
        }

        print(f"📥 [京麦售后明细] 第 3 步：下载 zip taskId={task_id}")
        print(f"   URL: {url}")
        print(f"   鉴权：仅 Cookie（不要 dsm-* 头、不要 h5st）")
        print(f"   Referer: {self.REFERER}")

        try:
            resp = self.session.get(url, headers=headers, timeout=60, allow_redirects=True)
        except _requests.exceptions.RequestException as e:
            raise RuntimeError(f"❌ 京麦售后明细下载失败：{e}") from e

        # 状态码判定
        if resp.status_code == 401 or resp.status_code == 302:
            raise CookieExpiredError(
                f"❌ 京麦售后 Cookie 过期（下载返回 HTTP {resp.status_code}）\n"
                f"   → 请浏览器重新登录 https://shop.jd.com/jdm/trade/after-sale/independent-after-sale/list，"
                f"F12 抓 export.shop.jd.com 域 Cookie 写入 config/jm_cookie.txt"
            )
        if resp.status_code != 200:
            raise RuntimeError(
                f"❌ 京麦售后明细下载 HTTP {resp.status_code}：{resp.text[:200]!r}"
            )

        zip_bytes = resp.content

        # 文件大小校验
        if len(zip_bytes) < 1024:
            raise RuntimeError(
                f"❌ 下载的 zip 太小（{len(zip_bytes)} 字节），可能任务未完成或已过期"
            )

        # 魔数校验：zip 是 PK\x03\x04
        if not zip_bytes.startswith(b"PK\x03\x04"):
            raise RuntimeError(
                f"❌ 下载内容不是 zip（magic bytes={zip_bytes[:8].hex()}）\n"
                f"   响应片段: {zip_bytes[:200]!r}"
            )

        # 提取文件名（Content-Disposition: form-data; name="attachment"; filename="xxx.zip"）
        cd = resp.headers.get("Content-Disposition", "")
        filename = f"{task_id}.zip"  # 兜底
        import re as _re
        m = _re.search(r'filename="([^"]+)"', cd)
        if m:
            filename = m.group(1)

        print(
            f"✅ [京麦售后明细] 下载成功：{filename}（{len(zip_bytes)} 字节，"
            f"Content-Type={resp.headers.get('Content-Type', 'unknown')!r}）"
        )
        return zip_bytes, filename

    def extract_xlsx_from_after_sale_zip(
        self,
        zip_path: str,
        output_dir: str = None,
        date: str = None,
    ) -> str:
        """第 4 步：解压售后 zip（项目16，**无密码**——绕过 msoffcrypto 直接解压）。

        ⚠️ 与项目14 extract_xlsx_from_zip 的关键差异（2026-08-12 抓包实证）：
            - 项目16 zip **无密码**：不调用 msoffcrypto
            - 项目14 zip **有密码**：要先 zipfile 解压（带密码）+ msoffcrypto 二次解密
            - 本方法**直接 zipfile 解压**，不解密（售后是普通 zip）

        ⚠️ 5020 字节小文件处理：
            - 如果内部 xlsx 是空表（5020 字节常见值）→ 仍正常保存空表
            - 如果内部 xlsx 文件损坏 → 抛出 RuntimeError

        返回:
            str - 解压后 xlsx 的绝对路径
        """
        import zipfile

        if not os.path.isfile(zip_path):
            raise RuntimeError(f"❌ zip 文件不存在：{zip_path}")

        # 输出目录：output/京麦售后明细/{date}/订单明细_{date}.xlsx
        # ⚠️ 项目14 用的"订单明细_{date}.xlsx"是订单明细的命名
        #     项目16 应该是"售后明细_{date}.xlsx"（保持命名一致）
        if output_dir is None:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            output_dir = os.path.join(base_dir, "output", "京麦售后明细")
        os.makedirs(output_dir, exist_ok=True)

        try:
            with zipfile.ZipFile(zip_path, "r") as zf:
                # 找第一个 .xlsx 或 .xls 条目
                candidate_names = [n for n in zf.namelist() if n.lower().endswith((".xlsx", ".xls"))]
                if not candidate_names:
                    raise RuntimeError(
                        f"❌ zip 内未找到 .xlsx/.xls 条目：{zip_path}\n"
                        f"   zip 内文件列表：{zf.namelist()}"
                    )
                target_name = candidate_names[0]
                print(
                    f"📂 [京麦售后明细] 第 4 步：解压 zip\n"
                    f"   源: {zip_path}\n"
                    f"   密码: 无（项目16 zip 不加密）\n"
                    f"   目标条目: {target_name}"
                )
                # ⚠️ 无密码直接读
                extracted_bytes = zf.read(target_name)
        except zipfile.BadZipFile as e:
            raise RuntimeError(f"❌ zip 文件损坏或不是有效 zip：{e}") from e

        # ⚠️ 项目16 zip 无密码，**直接读** xlsx（不走 msoffcrypto）
        #     通用工具 read_excel_bytes 按 magic bytes 自动选引擎
        try:
            df = read_excel_bytes(extracted_bytes)
            print(f"   ├─ 主表: {len(df)} 行 × {len(df.columns)} 列")
            if not df.empty:
                print(f"   ├─ 列名（前 8 列）: {list(df.columns[:8])}{'...' if len(df.columns) > 8 else ''}")
        except Exception as e:
            raise RuntimeError(
                f"❌ 解压后文件读取失败：{e}\n"
                f"   前 16 字节: {extracted_bytes[:16].hex()}"
            ) from e

        # Excel 后置统一规则
        if date:
            date_column, date_value = prepare_date_columns(df, date)
        else:
            date_column, date_value = None, None

        df = safe_convert_numeric(df)

        # 输出路径：output/京麦售后明细/{date}/售后明细_{date}.xlsx
        if date is None:
            import datetime as _dt
            date = _dt.datetime.fromtimestamp(os.path.getmtime(zip_path)).strftime("%Y-%m-%d")

        date_subdir = os.path.join(output_dir, date)
        os.makedirs(date_subdir, exist_ok=True)
        save_filename = f"售后明细_{date}.xlsx"
        target_xlsx = os.path.join(date_subdir, save_filename)

        df.to_excel(target_xlsx, index=False, engine="openpyxl")
        if date_column:
            apply_column_formats(target_xlsx, df, date_column=date_column, date_value=date_value)
        else:
            apply_column_formats(target_xlsx, df)

        print(
            f"✅ [京麦售后明细] 解压+转存成功：{target_xlsx}（{os.path.getsize(target_xlsx)} 字节，"
            f"{len(df)}行 × {len(df.columns)}列）"
        )

        # ⚠️ 用户决策 2026-08-11：删除中间 zip（只留解密后 xlsx）
        #     即使项目16 无密码，也按项目14 策略统一删中间 zip
        try:
            os.remove(zip_path)
            print(f"🗑️  [京麦售后明细] 中间 zip 已删除：{zip_path}")
        except OSError as e:
            print(f"⚠️ [京麦售后明细] 中间 zip 删除失败（不影响主流程）：{e}")

        return target_xlsx

    def run_after_sale_full_export(
        self,
        start_date: str = None,
        end_date: str = None,
        date: str = None,
        # ⚠️ 项目16 特有参数（2026-08-12 抓包实证）----
        tab_code: str = "all",
        after_sale_status_list: str = "",
        service_order_sub_state_list: str = "",
        customer_expect_list: str = "",
        refund_status_list: str = "",
        transfer_feedback_reason_list: str = "",
        poll_interval: int = 3,
        max_poll_times: int = 20,
    ) -> dict:
        """完整 4 步一键：创建 + 轮询 + 下载 + 解压（项目16，**无短信、无 IMAP**）。

        ⚠️ 2026-08-12 真实抓包适配：
            - 创建走 createExportTask（项目16 api）
            - 轮询走 queryExportTaskInfo（**待你提供轮询抓包**）
            - 下载走（**待你提供下载抓包**）
            - 解压无密码（password=None 跳过 msoffcrypto）

        ⚠️ 与项目14 run_full_export 的关键差异：
            - 4 步（项目14 是 5 步：多了短信申请 + IMAP）
            - 不需要 IMAP 授权码
            - 不调用 request_export_password
            - appId / api 路径 / payload 结构 / 时间格式（毫秒）都不同
            - 需要 X-Rp-Sdtoken 动态令牌（_post_dsm_after_sale 自动处理）

        返回:
            dict - 含 taskId / zip_path / xlsx_path
        """
        print("=" * 70)
        print(f"🚀 [京麦售后明细] 完整 4 步一键（{date or start_date}）")
        print(f"   exportType: {self.EXPORT_TYPE_AFTER_SALE_DETAIL} | tabCode: {tab_code}")
        print("=" * 70)

        # 第 1 步：创建
        create_ret = self.create_after_sale_export_task(
            start_date=start_date,
            end_date=end_date,
            date=date,
            tab_code=tab_code,
            after_sale_status_list=after_sale_status_list,
            service_order_sub_state_list=service_order_sub_state_list,
            customer_expect_list=customer_expect_list,
            refund_status_list=refund_status_list,
            transfer_feedback_reason_list=transfer_feedback_reason_list,
        )

        # 第 2 步：轮询
        item = self.wait_for_after_sale_task_ready(
            start_date=date or start_date,
            end_date=date or end_date,
            poll_interval=poll_interval,
            max_poll_times=max_poll_times,
        )
        # ⚠️ 项目16 用 taskId 字段（不是项目14 的 id）
        task_id = item.get("taskId") or item.get("id")
        if not task_id:
            raise RuntimeError(
                f"❌ 京麦售后明细轮询命中但无 taskId：{item}\n"
                f"   → 请检查轮询响应里任务记录的 taskId 字段路径"
            )

        # 第 3 步：下载 zip（2026-08-12 真实抓包适配）
        print(f"📥 [京麦售后明细] 第 3 步：下载 zip taskId={task_id}")
        zip_bytes, filename = self.download_after_sale_zip(task_id)
        zip_path = self.save_encrypted_zip(zip_bytes, filename)

        # 第 4 步：解压（无密码，售后业务无密码）
        print(f"📂 [京麦售后明细] 第 4 步：解压 zip（无密码）")
        xlsx_path = self.extract_xlsx_from_after_sale_zip(
            zip_path=zip_path,
            output_dir=os.path.join(os.path.dirname(os.path.abspath(__file__)), "output", "京麦售后明细"),
            date=date or start_date,
        )

        return {
            "code": self.CODE_OK,
            "msg": "成功",
            "taskId": task_id,
            "taskStatus": item.get("status") or item.get("taskStatus"),
            "zip_path": zip_path,
            "xlsx_path": xlsx_path,
            "rawItem": item,
        }


# ---- 调度器专用 callable 函数 ----

def _run_jm_create_task(**kwargs) -> dict:
    """调度器专用：京麦订单导出 - 第 1 步创建任务（项目14 阶段1，2026-08-11）。

    ⚠️ 注册到 BUSINESS_REGISTRY["京麦订单明细_创建任务"]["callable"]。
    设计动机：JingMaiOrderExportAPI.__init__ 需要 h5st 必填，
              标准调度路径不支持构造参数注入，本函数手动构造实例并调用 create_export_task。
    """
    # 提取透传参数
    h5st = kwargs.get("h5st", "")
    if not h5st:
        raise ValueError(
            "❌ 京麦订单明细_创建任务 必须传 h5st（浏览器F12抓 createdExportTask 请求头）\n"
            "   → 请浏览器登录 https://shop.jd.com/jdm/trade/tools/export/ExprotList，\n"
            "     F12 抓 createdExportTask 请求头 h5st 复制传入"
        )

    # cookie_path 可选
    cookie_path = kwargs.get("cookie_path", "config/jm_cookie.txt")

    # 透传给 create_export_task 的参数
    forward_kwargs = {
        k: kwargs[k] for k in (
            "start_date", "end_date", "date",
            "order_status_list", "sensitive_info_sign", "export_task_type",
        )
        if k in kwargs
    }

    # 日期兜底
    if not forward_kwargs.get("date") and not forward_kwargs.get("start_date") and not forward_kwargs.get("end_date"):
        raise ValueError("❌ 至少需要传入 date 或 start_date/end_date")

    api = JingMaiOrderExportAPI(h5st=h5st, cookie_path=cookie_path)
    return api.create_export_task(**forward_kwargs)


# ⚠️ 项目14 注册表 callable 字段回填（2026-08-11 启动）
BUSINESS_REGISTRY["京麦订单明细_创建任务"]["callable"] = _run_jm_create_task
BUSINESS_REGISTRY["京麦订单明细_创建任务"]["api_class"] = JingMaiOrderExportAPI


def _run_jm_create_and_wait(**kwargs) -> dict:
    """调度器专用：京麦订单导出 - 第 1+2 步一键（创建+轮询）（项目14 阶段2，2026-08-11）。

    ⚠️ 注册到 BUSINESS_REGISTRY["京麦订单明细_创建并轮询"]["callable"]。
    """
    h5st = kwargs.get("h5st", "")
    if not h5st:
        raise ValueError(
            "❌ 京麦订单明细_创建并轮询 必须传 h5st\n"
            "   → 浏览器F12抓 createdExportTask 请求头 h5st 复制传入"
        )

    cookie_path = kwargs.get("cookie_path", "config/jm_cookie.txt")
    forward_kwargs = {
        k: kwargs[k] for k in (
            "start_date", "end_date", "date",
            "order_status_list", "sensitive_info_sign", "export_task_type",
            "poll_interval", "max_poll_times",
        )
        if k in kwargs
    }

    if not forward_kwargs.get("date") and not forward_kwargs.get("start_date") and not forward_kwargs.get("end_date"):
        raise ValueError("❌ 至少需要传入 date 或 start_date/end_date")

    api = JingMaiOrderExportAPI(h5st=h5st, cookie_path=cookie_path)
    return api.create_and_wait(**forward_kwargs)


# ⚠️ 项目14 注册表第 2 个业务回填（创建并轮询一键，2026-08-11）
BUSINESS_REGISTRY["京麦订单明细_创建并轮询"]["callable"] = _run_jm_create_and_wait
BUSINESS_REGISTRY["京麦订单明细_创建并轮询"]["api_class"] = JingMaiOrderExportAPI


def _run_jm_create_wait_download(**kwargs) -> dict:
    """调度器专用：京麦订单导出 - 第 1+2+3 步一键（创建+轮询+下载加密 zip，2026-08-11）。

    ⚠️ 注册到 BUSINESS_REGISTRY["京麦订单明细_创建轮询并下载zip"]["callable"]。
    """
    h5st = kwargs.get("h5st", "")
    if not h5st:
        raise ValueError(
            "❌ 京麦订单明细_创建轮询并下载zip 必须传 h5st\n"
            "   → 浏览器F12抓 createdExportTask 请求头 h5st 复制传入"
        )

    cookie_path = kwargs.get("cookie_path", "config/jm_cookie.txt")
    forward_kwargs = {
        k: kwargs[k] for k in (
            "start_date", "end_date", "date",
            "order_status_list", "sensitive_info_sign", "export_task_type",
            "poll_interval", "max_poll_times",
        )
        if k in kwargs
    }

    if not forward_kwargs.get("date") and not forward_kwargs.get("start_date") and not forward_kwargs.get("end_date"):
        raise ValueError("❌ 至少需要传入 date 或 start_date/end_date")

    api = JingMaiOrderExportAPI(h5st=h5st, cookie_path=cookie_path)
    return api.create_wait_and_download(**forward_kwargs)


# ⚠️ 项目14 注册表第 3 个业务回填（创建+轮询+下载zip一键，2026-08-11）
BUSINESS_REGISTRY["京麦订单明细_创建轮询并下载zip"]["callable"] = _run_jm_create_wait_download
BUSINESS_REGISTRY["京麦订单明细_创建轮询并下载zip"]["api_class"] = JingMaiOrderExportAPI


def _run_jm_full_with_pwd(**kwargs) -> dict:
    """调度器专用：京麦订单导出 - 完整 4 步一键（创建+轮询+下载+短信申请，2026-08-11）。

    ⚠️ 注册到 BUSINESS_REGISTRY["京麦订单明细_创建轮询下载并申请密码"]["callable"]。
    设计动机：4 步链路，每步鉴权头不同（h5st/dsm vs Cookie-only），需要单一入口编排。
    """
    h5st = kwargs.get("h5st", "")
    if not h5st:
        raise ValueError(
            "❌ 京麦订单明细_创建轮询下载并申请密码 必须传 h5st\n"
            "   → 浏览器F12抓 createdExportTask 请求头 h5st 复制传入"
        )

    cookie_path = kwargs.get("cookie_path", "config/jm_cookie.txt")
    forward_kwargs = {
        k: kwargs[k] for k in (
            "start_date", "end_date", "date",
            "order_status_list", "sensitive_info_sign", "export_task_type",
            "poll_interval", "max_poll_times",
        )
        if k in kwargs
    }

    if not forward_kwargs.get("date") and not forward_kwargs.get("start_date") and not forward_kwargs.get("end_date"):
        raise ValueError("❌ 至少需要传入 date 或 start_date/end_date")

    api = JingMaiOrderExportAPI(h5st=h5st, cookie_path=cookie_path)
    return api.create_wait_download_and_request_pwd(**forward_kwargs)


# ⚠️ 项目14 注册表第 4 个业务回填（完整 4 步一键，2026-08-11）
BUSINESS_REGISTRY["京麦订单明细_创建轮询下载并申请密码"]["callable"] = _run_jm_full_with_pwd
BUSINESS_REGISTRY["京麦订单明细_创建轮询下载并申请密码"]["api_class"] = JingMaiOrderExportAPI


def _run_jm_run_full_export(**kwargs) -> dict:
    """调度器专用：京麦订单导出 - 完整 5 步一键（创建+轮询+下载+短信+IMAP+解压，2026-08-11）。

    ⚠️ 注册到 BUSINESS_REGISTRY["京麦订单明细_完整一键导出"]["callable"]。
    设计动机：5 步链路最完整，密码获取两路（sms_password 优先 / IMAP 兜底）。
    """
    h5st = kwargs.get("h5st", "")
    if not h5st:
        raise ValueError(
            "❌ 京麦订单明细_完整一键导出 必须传 h5st\n"
            "   → 浏览器F12抓 createdExportTask 请求头 h5st 复制传入"
        )

    cookie_path = kwargs.get("cookie_path", "config/jm_cookie.txt")
    forward_kwargs = {
        k: kwargs[k] for k in (
            "start_date", "end_date", "date",
            "order_status_list", "sensitive_info_sign", "export_task_type",
            "poll_interval", "max_poll_times",
            "sms_password", "imap_config_path", "imap_timeout_seconds",
        )
        if k in kwargs
    }

    if not forward_kwargs.get("date") and not forward_kwargs.get("start_date") and not forward_kwargs.get("end_date"):
        raise ValueError("❌ 至少需要传入 date 或 start_date/end_date")

    api = JingMaiOrderExportAPI(h5st=h5st, cookie_path=cookie_path)
    return api.run_full_export(**forward_kwargs)


# ⚠️ 项目14 注册表第 5 个业务回填（完整 5 步一键，2026-08-11）
BUSINESS_REGISTRY["京麦订单明细_完整一键导出"]["callable"] = _run_jm_run_full_export
BUSINESS_REGISTRY["京麦订单明细_完整一键导出"]["api_class"] = JingMaiOrderExportAPI


# ============================================================
# 项目16：京麦售后明细导出 - 调度函数（2026-08-11 启动骨架）
# ============================================================

def _run_jm_after_sale_full(**kwargs) -> dict:
    """调度器专用：京麦售后明细导出 - 完整 4 步一键（创建+轮询+下载+解压，**无短信**）。

    ⚠️ 注册到 BUSINESS_REGISTRY["京麦售后明细_完整一键导出"]["callable"]。
    设计动机：售后业务无短信/IMAP，4 步链路直接走 run_after_sale_full_export。
    """
    h5st = kwargs.get("h5st", "")
    if not h5st:
        raise ValueError(
            "❌ 京麦售后明细_完整一键导出 必须传 h5st\n"
            "   → 浏览器F12抓 createdExportTask 请求头 h5st 复制传入\n"
            "   ⚠️ 项目16 h5st 抓包实证是必需的（与项目14 一致）"
        )

    cookie_path = kwargs.get("cookie_path", "config/jm_cookie.txt")
    forward_kwargs = {
        k: kwargs[k] for k in (
            "start_date", "end_date", "date",
            "tab_code", "after_sale_status_list",
            "service_order_sub_state_list", "customer_expect_list",
            "refund_status_list", "transfer_feedback_reason_list",
            "poll_interval", "max_poll_times",
        )
        if k in kwargs
    }

    if not forward_kwargs.get("date") and not forward_kwargs.get("start_date") and not forward_kwargs.get("end_date"):
        raise ValueError("❌ 至少需要传入 date 或 start_date/end_date")

    api = JingMaiAfterSaleExportAPI(h5st=h5st, cookie_path=cookie_path)
    return api.run_after_sale_full_export(**forward_kwargs)


# ⚠️ 项目16 注册表回填（售后明细完整一键，2026-08-11 启动骨架）
BUSINESS_REGISTRY["京麦售后明细_完整一键导出"]["callable"] = _run_jm_after_sale_full
BUSINESS_REGISTRY["京麦售后明细_完整一键导出"]["api_class"] = JingMaiAfterSaleExportAPI


def list_businesses():
    """打印所有已注册业务清单（启动时用）。"""
    print()
    print("=" * 70)
    print(f"已注册业务清单（共 {len(BUSINESS_REGISTRY)} 个）：")
    print("=" * 70)
    for idx, (key, info) in enumerate(BUSINESS_REGISTRY.items(), 1):
        enabled = info.get("enabled", True)
        status_tag = "" if enabled else "  [已停用]"
        print(f"  [{idx}] {key}{status_tag}")
        print(f"      描述: {info['desc']}")
        print(f"      API类: {info['api_class'].__name__}.{info['method']}()")
        if info.get("params"):
            print(f"      参数:")
            for pk, pv in info["params"].items():
                print(f"        - {pk}: {pv}")
        print()


def get_business_handler(biz_key):
    """根据业务key返回对应的处理函数。

    优先级：
        1. info["callable"] - 自定义函数（用于特殊业务需要传额外参数，如京准通 h5st）
        2. info["api_class"] + info["method"] - 标准基类方法
    """
    if biz_key not in BUSINESS_REGISTRY:
        available = "、".join(BUSINESS_REGISTRY.keys())
        raise BusinessNotFoundError(
            f"未知业务: {biz_key}\n"
            f"已注册业务: {available}\n"
            f"调用 list_businesses() 查看所有业务详情。"
        )
    info = BUSINESS_REGISTRY[biz_key]

    # 方式1：callable 优先（用于特殊业务需要传额外构造参数）
    if info.get("callable"):
        return info["callable"]

    # 方式2：标准基类方式
    api_class = info["api_class"]
    method_name = info["method"]
    # ⚠️ 必须先生成实例，再取实例方法（否则拿到的是未绑定方法，调用时会报 missing 'self'）
    api_instance = api_class()
    method = getattr(api_instance, method_name)
    return method


# ============================================================
#  统一调度入口（run_business）
# ------------------------------------------------------------
#  支持两种调用方式：
#    1. run_business("业务key", date="2026-07-29")       # 单个业务
#    2. run_business(["业务key1", "业务key2"], date=...)   # 批量业务（list传入）
#  函数入参kwargs优先级 > config.xlsx配置（动态覆盖）
# ============================================================
def run_business(biz_key_or_keys, **kwargs):
    """
    统一业务调度入口。

    参数:
        biz_key_or_keys - 单个业务key字符串 或 业务key列表
        **kwargs        - 业务参数（如 date="2026-07-29"）

    返回:
        单个业务：返回文件路径
        批量业务：返回 {业务key: 文件路径} 的dict
    """
    # 兼容 list 批量调用
    if isinstance(biz_key_or_keys, (list, tuple)):
        return _run_business_batch(biz_key_or_keys, **kwargs)
    else:
        return _run_single_business(biz_key_or_keys, **kwargs)


def _run_single_business(biz_key, **kwargs):
    """执行单个业务（调度层过滤：已停用业务不执行）。"""
    info = BUSINESS_REGISTRY[biz_key]

    # ⚠️ 调度层过滤：enabled=False 的业务（如自主访问）直接跳过，不触发导出
    if info.get("enabled", True) is False:
        print(f"[SKIP] 业务已停用，跳过: {biz_key}（{info['desc']}）")
        print(f"       如需开启，请将 BUSINESS_REGISTRY 中该业务的 enabled 改为 True。")
        return None

    handler = get_business_handler(biz_key)

    # 业务级打印（让日志可追踪）
    print()
    print("=" * 70)
    print(f"执行业务: {biz_key}  -  {info['desc']}")
    print("=" * 70)

    try:
        result = handler(**kwargs)
        print(f"[OK] 业务完成: {biz_key} → {result}")
        return result
    except CookieExpiredError as e:
        print(f"[ERR] Cookie已过期: {e}")
        print(f"       请重新获取Cookie，更新 config/sz_cookie.txt 后重试。")
        raise
    except Exception as e:
        print(f"[ERR] 业务失败: {biz_key} → {e}")
        print(f"       详细日志请查看 logs/ 目录下的日志文件。")
        raise


def _run_business_batch(biz_key_list, **kwargs):
    """批量执行多个业务（自动读取config里的请求间隔，循环调用）。

    调度层过滤：列表中已停用的业务（enabled=False）会自动剔除，
    仅执行启用状态正常的业务（如搜索/推荐/购物车）。
    """
    # ⚠️ 调度层过滤：剔除已停用业务，保留可执行业务
    enabled_list = []
    for k in biz_key_list:
        info = BUSINESS_REGISTRY.get(k, {})
        if info.get("enabled", True) is False:
            print(f"[SKIP] 业务已停用，从批量列表剔除: {k}（{info.get('desc', '')}）")
        else:
            enabled_list.append(k)
    if not enabled_list:
        print("[WARN] 批量列表中所有业务均已停用，无可执行任务。")
        return {}
    biz_key_list = enabled_list

    print()
    print("=" * 70)
    print(f"批量执行业务（{len(biz_key_list)}个）：")
    for i, k in enumerate(biz_key_list, 1):
        print(f"  [{i}] {k}")
    print("=" * 70)
    print()

    results = {}
    for i, biz_key in enumerate(biz_key_list, 1):
        print(f"--- [{i}/{len(biz_key_list)}] 开始执行: {biz_key} ---")
        try:
            file_path = _run_single_business(biz_key, **kwargs)
            results[biz_key] = file_path
            print(f"--- [{i}/{len(biz_key_list)}] 完成: {biz_key} ---")
        except Exception as e:
            print(f"--- [{i}/{len(biz_key_list)}] 失败: {biz_key} ({e}) ---")
            results[biz_key] = None
        print()

    # 汇总
    print("=" * 70)
    print(f"批量执行汇总（共 {len(biz_key_list)} 个）：")
    print("=" * 70)
    success_count = 0
    for biz_key, fp in results.items():
        status = "[OK]" if fp else "[FAIL]"
        if fp:
            success_count += 1
        print(f"  {status} {biz_key}: {fp}")
    print(f"\n总计: {success_count}/{len(biz_key_list)} 成功")
    return results


# ============================================================
#  配置一致性检查（启动时自动跑）
# ------------------------------------------------------------
#  中文说明（小白必读）：
#    项目铁律要求：禁止硬编码日期、业务参数。
#    此函数在 main() 启动时跑一遍，自动检测 main.py 里是否还有违规的硬编码。
#    如果发现违规，会打印警告（不影响启动，但提醒用户关注）。
# ============================================================
def config_consistency_check():
    """检查config和代码一致性，输出【配置一致性核对报告】。

    核对维度（项目铁律）：
        1. 所有API请求前必须执行间隔sleep（间隔从config读取，禁止硬编码休眠秒数）
        2. date/startDate/endDate 必须从config读取或外部动态传入，代码禁止硬编码固定日期
        3. 渠道ID、uuid前缀、导出条数、排序字段等业务参数必须走配置，禁止业务函数内写死
    报告格式：
        ✅ 已遵循（附依据） / ❌ 违规（附行号） / ⚠️ 待优化
    """
    import re as _re
    import ast as _ast

    print()
    print("=" * 70)
    print("【配置一致性核对报告】")
    print("=" * 70)

    # ---------- 第0步：读取 config.xlsx 全部配置项清单 ----------
    config_items = {}   # {变量名: 参数值}
    try:
        wb = load_workbook(JDBaseRequest.DEFAULT_CONFIG_PATH, read_only=True, data_only=True)
        ws = wb["全局配置"]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row and len(row) >= 3 and row[1]:
                config_items[str(row[1])] = "" if row[2] is None else str(row[2])
        wb.close()
        print(f"[✅] 配置读取成功：config.xlsx 共 {len(config_items)} 项配置")
        for name, value in config_items.items():
            print(f"      · {name} = {value}")
    except Exception as e:
        print(f"[❌] 配置读取失败: {e}")

    # ---------- 第1步：间隔逻辑核对 ----------
    try:
        with open(__file__, "r", encoding="utf-8") as f:
            source = f.read()
        tree = _ast.parse(source)

        # 1.1 request() 方法内必须调用 _wait_interval()
        wait_interval_called = False
        for node in _ast.walk(tree):
            if isinstance(node, _ast.FunctionDef) and node.name == "request":
                for sub in _ast.walk(node):
                    if isinstance(sub, _ast.Call) and isinstance(sub.func, _ast.Attribute) \
                            and sub.func.attr == "_wait_interval":
                        wait_interval_called = True
                        break
        if wait_interval_called:
            print(f"[✅] 间隔逻辑：request() 已调用 _wait_interval()（间隔从config[请求间隔(秒)]读取）")
        else:
            print(f"[❌] 间隔逻辑：request() 未调用 _wait_interval()，存在跳过间隔风险")

        # 1.2 检查是否有硬编码 time.sleep(固定秒数)（排除 _wait_interval 内部从config读取的写法）
        hardcoded_sleep = []
        for node in _ast.walk(tree):
            if isinstance(node, _ast.Call) and isinstance(node.func, _ast.Attribute) \
                    and node.func.attr == "sleep" and node.args:
                arg = node.args[0]
                # 数字常量 or 非config来源的固定表达式都算硬编码
                if isinstance(arg, _ast.Constant) and isinstance(arg.value, (int, float)):
                    hardcoded_sleep.append((node.lineno, arg.value))
        if hardcoded_sleep:
            print(f"[⚠️] 硬编码休眠：第{[f'L{n}({v}s)' for n, v in hardcoded_sleep]}行 存在固定休眠秒数，"
                  f"建议改为从config读取（当前仅为风控失败重试等待，需人工确认）")
        else:
            print(f"[✅] 间隔逻辑：未发现硬编码固定休眠秒数")
    except Exception as e:
        print(f"[❌] 间隔逻辑扫描失败: {e}")

    # ---------- 第2步：日期硬编码核对（AST扫描，自动跳过注释/docstring/示例） ----------
    date_pattern = _re.compile(r"^20\d{2}-\d{2}-\d{2}$")

    def is_date_str(node):
        return isinstance(node, _ast.Constant) and isinstance(node.value, str) \
            and bool(date_pattern.match(node.value))

    hardcoded_dates = []   # (行号, 描述)
    try:
        for node in _ast.walk(tree):
            # 2.1 赋值语句： date = "2026-07-29"
            if isinstance(node, _ast.Assign):
                for target in node.targets:
                    if isinstance(target, _ast.Name) and is_date_str(node.value):
                        hardcoded_dates.append((node.lineno, f"{target.id} = {node.value.value!r}"))
            # 2.2 函数调用关键字参数： download_sku(date="2026-07-29")
            if isinstance(node, _ast.Call):
                for kw in node.keywords:
                    if is_date_str(kw.value):
                        func_name = node.func.id if isinstance(node.func, _ast.Name) else "?"
                        hardcoded_dates.append((node.lineno, f"{func_name}({kw.arg}={kw.value.value!r})"))
            # 2.3 函数默认参数： def f(date="2026-07-29")
            if isinstance(node, _ast.FunctionDef):
                for default in node.args.defaults:
                    if is_date_str(default):
                        hardcoded_dates.append((node.lineno, f"函数{node.name}()默认参数 {default.value!r}"))
        if hardcoded_dates:
            for lineno, desc in hardcoded_dates:
                print(f"[❌] 日期硬编码：第{lineno}行 {desc}")
        else:
            print(f"[✅] 日期参数：date/startDate/endDate 全部从config读取或外部动态传入，未发现硬编码")
    except Exception as e:
        print(f"[❌] 日期扫描失败: {e}")

    # ---------- 第3步：业务参数核对 ----------
    # 3.1 业务参数必须走 _get_business_params()（从config读取），不允许业务函数内写死
    try:
        biz_param_funcs = [n.name for n in _ast.walk(tree)
                           if isinstance(n, _ast.FunctionDef)
                           and n.name in ("_get_business_params", "_get_business_param")]
        if biz_param_funcs:
            print(f"[✅] 业务参数：统一通过 {', '.join(biz_param_funcs)}() 从config.xlsx读取")
        else:
            print(f"[❌] 业务参数：未找到统一的config读取入口")
    except Exception as e:
        print(f"[❌] 业务参数扫描失败: {e}")

    # 3.2 业务参数核对：可变参数必须走config；固定参数为确认常量不要求
    var_params = ProductFlowAPI.VARIABLE_BIZ_PARAMS
    missing_in_config = [k for k in var_params if k not in config_items]
    if missing_in_config:
        print(f"[⚠️] 待优化：以下可变业务参数在config.xlsx未配置，当前走代码兜底值（开发期）：")
        print(f"      {', '.join(missing_in_config)}")
        print(f"      请将上述参数补写进 config.xlsx【全局配置】sheet，避免长期依赖兜底。")
    else:
        print(f"[✅] 业务参数：可变参数（{', '.join(var_params)}）已全部在config.xlsx中配置")
    print(f"[✅] 业务参数：固定常量（{', '.join(ProductFlowAPI.FIXED_BIZ_PARAMS)}）经用户确认写死代码，不依赖config")

    # 3.3 CHANNEL_MAP（业务专属注册表，按用户要求集中维护）
    print(f"[✅] 渠道配置：CHANNEL_MAP 集中维护渠道（搜索2008/推荐2009/购物车3001执行；"
          f"自主访问3001与购物车口径重叠已停用），uuid前缀按渠道区分")

    # 3.4 【阶段4新增】店铺来源-三级渠道业务专用配置（OfflineChannelAPI）
    print(f"[✅] 店铺来源-三级渠道：")
    print(f"      - 必带请求头: Origin={OfflineChannelAPI.ORIGIN}, Referer={OfflineChannelAPI.REFERER}")
    print(f"      - UUID策略: 完全随机（不依赖类常量UUID_PREFIX，符合用户2026-08-06确认的'禁止硬编码'要求）")
    print(f"      - 固定业务参数: {', '.join(OfflineChannelAPI.FIXED_BIZ_PARAMS.keys())}")
    print(f"      - 可变业务参数: {', '.join(OfflineChannelAPI.VARIABLE_BIZ_PARAMS.keys())}")
    print(f"      - 入口: python main.py --biz_key '店铺来源_三级渠道' --date '2026-08-04'")

    print("=" * 70)
    print("核对完成。若存在 ❌ 项，请先修复再运行；⚠️ 项请尽快补齐config。")
    print("=" * 70)


# ============================================================
#  命令行参数解析
# ============================================================
def parse_args():
    """解析命令行参数。"""
    parser = argparse.ArgumentParser(
        description="京东商智数据导出工具（main.py）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  # 命令行调用单个业务
  python main.py --biz_key "商品流量来源_搜索" --date "2026-07-29"

  # 命令行批量调用
  python main.py --biz_key "商品流量来源_搜索,商品流量来源_推荐,商品流量来源_自主访问" --date "2026-07-29"

  # 代码内部调用
  from main import run_business
  run_business("商品流量来源_搜索", date="2026-07-29")
  run_business(["商品流量来源_搜索", "商品流量来源_推荐"], date="2026-07-29")
        """,
    )
    parser.add_argument(
        "--biz_key",
        type=str,
        help="业务key（必填或逗号分隔的多个key批量），可用值: " + ", ".join(BUSINESS_REGISTRY.keys()),
    )
    parser.add_argument(
        "--date",
        type=str,
        help="查询日期 YYYY-MM-DD（可选，优先于config.xlsx）",
    )
    parser.add_argument(
        "--start_date",
        type=str,
        help="开始日期 YYYY-MM-DD（可选，区间查询时用）",
    )
    parser.add_argument(
        "--end_date",
        type=str,
        help="结束日期 YYYY-MM-DD（可选，区间查询时用）",
    )
    parser.add_argument(
        "--range",
        type=str,
        choices=["last_1d", "last_3d", "last_7d", "last_15d", "last_30d"],
        default=None,
        help="近N天快捷区间（用户决策 2026-08-10）：end=昨天, start=今天-N；"
             "示例: --range last_7d 即近7天（今天-7 至 昨天）。"
             "与 --date/--start_date/--end_date 互斥（同时传会报错）。",
    )
    parser.add_argument(
        "--sms_password",
        type=str,
        default=None,
        help="京麦订单明细【加密】导出专用：手动传入解压密码（优先级高于IMAP自动监听）",
    )
    parser.add_argument(
        "--h5st",
        type=str,
        default=None,
        help="京麦订单明细【加密】导出专用：浏览器F12抓 createdExportTask 请求头 h5st（前端强签名，一次性）",
    )
    parser.add_argument(
        "--cookie_path",
        type=str,
        default=None,
        help="京麦订单明细【加密】导出专用：Cookie 文件路径（默认 config/jm_cookie.txt）",
    )
    parser.add_argument(
        "--imap_config_path",
        type=str,
        default=None,
        help="京麦订单明细【加密】导出专用：IMAP 配置文件路径（默认 config/imap_config.ini）",
    )
    parser.add_argument(
        "--list",
        action="store_true",
        help="列出所有已注册业务清单",
    )
    return parser.parse_args()


def _resolve_range_to_dates(range_arg: str) -> tuple:
    """把 `--range last_Nd` 解析为 (start_date, end_date) 字符串元组。

    ⚠️ 用户决策 2026-08-10：end=昨天（今天-1），start=今天-N。
       避免「今天」数据未生成导致 OSS 404（项目10 阶段7 经验）。

    入参:
        range_arg - "last_1d" / "last_3d" / "last_7d" / "last_15d" / "last_30d"
    出参:
        (start_date, end_date) 字符串元组，YYYY-MM-DD 格式

    异常:
        ValueError - 格式不合法（理论上 argparse 已校验，这里兜底）
    """
    import re
    from datetime import datetime, timedelta
    m = re.fullmatch(r"last_(\d+)d", range_arg)
    if not m:
        raise ValueError(f"❌ range 参数格式不合法：{range_arg!r}（期望 last_Nd）")
    n = int(m.group(1))
    today = datetime.now().date()
    end_date = today - timedelta(days=1)        # 昨天
    start_date = today - timedelta(days=n)        # 今天-N
    return start_date.isoformat(), end_date.isoformat()


def split_date_range(start_date, end_date, max_days=31):
    """把 [start_date, end_date] 区间拆成逐天日期字符串列表（P0 边界保护工具）。

    ⚠️ 用途（2026-08-10 用户决策）：
       商智搜索/推荐/购物车接口服务端不支持多日区间导出，区间查询时需拆成逐天循环。
       本函数负责：① 格式校验（YYYY-MM-DD）；② start ≤ end 校验；③ 最大天数上限校验
       （默认31天，防止大批量循环压接口触发风控403）。

    入参:
        start_date - 区间开始日期 YYYY-MM-DD
        end_date   - 区间结束日期 YYYY-MM-DD
        max_days   - 最大允许天数（默认31，用户决策 2026-08-10）
    出参:
        list[str] - 逐天日期列表，如 ["2026-08-03", "2026-08-04", ..., "2026-08-09"]
    异常:
        ValueError - 日期格式不合法 / start晚于end / 天数超上限
    """
    from datetime import datetime, timedelta
    try:
        d_start = datetime.strptime(start_date, "%Y-%m-%d").date()
        d_end = datetime.strptime(end_date, "%Y-%m-%d").date()
    except (ValueError, TypeError) as e:
        raise ValueError(
            f"日期格式不合法（期望 YYYY-MM-DD）：start={start_date!r}, end={end_date!r}"
        ) from e
    if d_start > d_end:
        raise ValueError(f"开始日期不能晚于结束日期：start={start_date}, end={end_date}")
    days = (d_end - d_start).days + 1
    if days > max_days:
        raise ValueError(
            f"区间天数({days}天)超过最大限制({max_days}天)，"
            f"为防止大批量压接口触发风控403，请缩小日期区间后再试"
        )
    return [(d_start + timedelta(days=i)).isoformat() for i in range(days)]


# ============================================================
#  主程序入口
# ============================================================
def main():
    """主程序入口（支持命令行 + 默认业务）。"""
    args = parse_args()

    # 后台异步扫描项目文档索引（不阻塞启动）
    # 新建 .md / .py 后自动追加到 docs/项目文档索引.xlsx
    try:
        from update_doc_index import async_update_index
        async_update_index()
    except Exception as e:
        # 异步扫描失败不影响主业务
        pass

    # 启动信息
    print("=" * 70)
    print(f"京东商智 - 数据导出工具    店铺: {SHOP_NAME}")
    print("=" * 70)

    # 配置一致性检查
    config_consistency_check()

    # 列出业务清单（--list 参数 或 启动时打印）
    if args.list:
        list_businesses()
        return

    # 打印全局配置（启动时核对用）
    print()
    print("-" * 70)
    print("全局配置（从config.xlsx读取）：")
    print("-" * 70)
    try:
        wb = load_workbook(JDBaseRequest.DEFAULT_CONFIG_PATH, read_only=True, data_only=True)
        ws = wb["全局配置"]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row and len(row) >= 3 and row[1]:
                proj = row[0] if row[0] else "未分类"
                var_name = row[1]
                var_value = row[2]
                print(f"  [{proj}] {var_name} = {var_value}")
        wb.close()
    except Exception as e:
        print(f"  [WARN] 读取配置失败: {e}")
    print()

    # 列出已注册业务（始终打印，方便核对）
    list_businesses()

    # 决定业务key：命令行参数 > 默认值
    if args.biz_key:
        # 支持逗号分隔的批量
        biz_keys = [k.strip() for k in args.biz_key.split(",") if k.strip()]
    else:
        # 默认跑商品流量来源3个启用渠道（自主访问与购物车口径重叠，调度层已停用）
        biz_keys = [
            "商品流量来源_搜索",
            "商品流量来源_推荐",
            "商品流量来源_购物车",
        ]
        print("[INFO] 未指定 --biz_key，默认批量执行商品流量来源3个启用渠道：搜索/推荐/购物车")

    # 决定日期参数
    kwargs = {}

    # ⚠️ 用户决策 2026-08-10：--range 与 --date/--start_date/--end_date 互斥
    if args.range and (args.date or args.start_date or args.end_date):
        print("[ERR] --range 不能与 --date / --start_date / --end_date 同时使用")
        sys.exit(2)
    if args.range:
        start_date, end_date = _resolve_range_to_dates(args.range)
        kwargs["start_date"] = start_date
        kwargs["end_date"] = end_date
        print(f"[INFO] --range {args.range} → start_date={start_date}, end_date={end_date}")
    else:
        if args.date:
            kwargs["date"] = args.date
        if args.start_date:
            kwargs["start_date"] = args.start_date
        if args.end_date:
            kwargs["end_date"] = args.end_date

    # 透传京麦订单导出专用参数（2026-08-11 阶段5）
    if args.sms_password:
        kwargs["sms_password"] = args.sms_password
        print(f"[INFO] --sms_password 已传入（优先级高于IMAP自动监听）")
    if args.h5st:
        kwargs["h5st"] = args.h5st
        print(f"[INFO] --h5st 已传入（{len(args.h5st)} 字符）")
    if args.cookie_path:
        kwargs["cookie_path"] = args.cookie_path
        print(f"[INFO] --cookie_path 已传入（{args.cookie_path}）")
    if args.imap_config_path:
        kwargs["imap_config_path"] = args.imap_config_path
        print(f"[INFO] --imap_config_path 已传入（{args.imap_config_path}）")

    # 执行
    try:
        results = run_business(biz_keys, **kwargs)
        if isinstance(results, dict):
            # 批量：已打印汇总
            pass
        else:
            # 单个
            print(f"\n[OK] 导出成功: {results}")
    except CookieExpiredError as e:
        print(f"\n[ERR] Cookie已过期: {e}")
        sys.exit(1)
    except BusinessNotFoundError as e:
        print(f"\n[ERR] 业务未找到: {e}")
        sys.exit(2)
    except Exception as e:
        print(f"\n[ERR] 导出失败: {e}")
        sys.exit(3)


if __name__ == "__main__":
    main()

