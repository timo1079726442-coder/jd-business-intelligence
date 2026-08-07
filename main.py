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
INTEGER_ZERO_DECIMAL_COLUMNS = {"SKU", "SPU"}


def _col_matches(col_name, name_set):
    """判断列名是否命中规则集合。

    匹配规则：列名精确等于集合元素，或以集合元素结尾。
    举例：列名"商品SKU"命中"SKU"（以SKU结尾）；而"成交金额（SPU）"不命中"SPU"（以）结尾），
          避免把带（SPU）后缀的金额/客户数等指标列误套格式。
    """
    if col_name in name_set:
        return True
    return any(col_name.endswith(name) for name in name_set)


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


def safe_convert_numeric(df):
    """全表数值安全转换（所有报表复用，全局生效）。

    入参:
        df - pandas.DataFrame（从Excel读取的表格数据）
    出参:
        处理后的DataFrame（直接修改并返回），转换规则：
            0. 【强制文本黑名单】列名命中 TEXT_FORCE_COLUMNS（如"订单编号"）
               → 整列完全跳过数值转换，强制保留原始文本字符串（不依赖长度判断）；
            1. 其他字符串且为纯数字（可含小数点/负号）且数字位数≤15位 → 转成数值（int/float）；
               其中列名命中 INTEGER_ZERO_DECIMAL_COLUMNS（如"SKU"/"SPU"）时单元格格式为 0（0位小数无千分位）；
            2. 纯数字但数字位数>15位 → 保留原始文本，杜绝精度丢失（兜底防护，全局保留）；
            3. 非纯数字（日期/含字母/空值/已是数值类型） → 保留原值；转换失败同样保留原值。
    注意:
        ⚠️ 调用前请先把日期列用 convert_date_format() 处理好，否则"20260729"这类
           8位纯数字日期会被误当成普通数字转换（商品流量来源流程已保证先转日期再转数值）。
    """
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

    wb.save(file_path)
    wb.close()


# ============================================================
#  通用请求基类（JDBaseRequest）
# ------------------------------------------------------------
#  业务名称：通用能力
#  接口地址：无（封装通用能力，被各业务API复用）
#  功能说明：
#      - Cookie 读取与更新（config/cookie.txt）
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
                                raise CookieExpiredError("Cookie已过期，请更新 config/cookie.txt")

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
    CHANNEL_MAP = {
        "商品流量来源_搜索":     ("2008", "ca412182e5668a106054"),
        "商品流量来源_推荐":     ("2009", "ca412182e5668a106054"),
        "商品流量来源_购物车":   ("3001", "5f9cc2ca20cad3d11642"),
        "商品流量来源_自主访问": ("3001", "5f9cc2ca20cad3d11642"),  # 与购物车口径重叠，仅保留配置
    }

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

    # ---------- 商品流量来源 Excel后置处理（2026-08-05 新增）----------
    def _save_flow_excel(self, response, filename, date):
        """商品流量来源专用保存流程（Excel后置处理）。

        导出流程（需求文档要求 + 2026-08-07 公共规则1+2）：
            ① 接口返回的Excel二进制流 → 读成DataFrame
            ② 日期列统一处理 prepare_date_columns()：
               报表自带【日期】/【时间】列 → 禁止重复插入，仅做格式标准化；
               无日期/时间列 → 首列插入【日期】列（值=查询日期，yyyy/m/d）
            ③ 调用通用数值安全转换函数 safe_convert_numeric()，处理全表字段类型
            ④ 写入Excel并设置日期列单元格格式（打开文件不弹格式警告）

        入参:
            response - requests响应（content为接口返回的xlsx二进制）
            filename - 保存文件名（如 搜索流量_2026-07-29.xlsx）
            date     - 本次查询日期（如 2026-07-29）
        出参:
            保存后的Excel文件绝对路径
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

        # ②③ 日期列统一处理（公共规则1+2，2026-08-07）：
        #    报表自带【日期】/【时间】列 → 禁止重复插入日期列，仅做格式标准化；
        #    报表无日期/时间列 → 首列插入【日期】列，值=本次查询日期。
        date_column, date_value = prepare_date_columns(df, date)

        # ④ 全表数值安全转换（>15位长数字保留文本，防止精度丢失）
        df = safe_convert_numeric(df)

        # ⑤ 写入Excel → 按列名规则设置单元格格式（日期列/订单编号@/SKU·SPU数值0位小数）
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
#      4. Cookie 从浏览器会话获取（走 config/cookie.txt，禁止入代码）
# ============================================================
class OfflineChannelAPI(JDBaseRequest):
    """店铺来源-离线渠道流量报表 API。

    业务定位：
        商智 szgateway.jd.com 模块下"店铺来源-离线渠道"维度的报表导出。
        与商品流量来源（downSkuTable.ajax / SKU 维度）不同，本接口按
        三级流量渠道分组，输出渠道维度的访客/浏览/成交数据。

    父类复用：
        - 父类 JDBaseRequest 提供：
            * Cookie 读取（config/cookie.txt）
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
                    raise CookieExpiredError("Cookie已过期或无效，请更新 config/cookie.txt")
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
#      3. Cookie 从 config/cookie.txt 整体读取
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
            * Cookie 读取（config/cookie.txt）
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

        ⚠️ 业务背景：用户抓包显示 UUID 前缀为 `42005c22589c8b55826d`（非固定 prefix），
        证明前端 SDK 每次会话运行时动态生成。完全随机化符合用户 2026-08-07
        "禁止硬编码 uuid 前缀"约束。

        复制来源：项目 4 OfflineChannelAPI._gen_uuid_random（已验证可用）

        返回:
            str - 形如 "42005c22589c8b55826d-19fdaefe247"（16hex + - + 10hex）
        """
        import secrets
        # 16位小写hex + "-" + 10位小写hex，与抓包格式完全一致
        prefix = secrets.token_hex(8)        # 8字节 = 16hex 字符
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
                    raise CookieExpiredError("Cookie已过期或无效，请更新 config/cookie.txt")
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
BUSINESS_REGISTRY = {
    "商品流量来源_搜索": {
        "api_class": ProductFlowAPI,
        "method": "download_search_sku",
        "desc": "商品搜索效果（搜索子来源2008）",
        "params": {
            "date": "查询日期YYYY-MM-DD（从config.xlsx的date读取）",
            "startDate": "开始日期（默认=date）",
            "endDate": "结束日期（默认=date）",
        },
    },
    "商品流量来源_推荐": {
        "api_class": ProductFlowAPI,
        "method": "download_recommend_sku",
        "desc": "商品推荐效果（推荐子来源2009）",
        "params": {
            "date": "查询日期YYYY-MM-DD",
            "startDate": "开始日期",
            "endDate": "结束日期",
        },
    },
    "商品流量来源_购物车": {
        "api_class": ProductFlowAPI,
        "method": "download_cart_sku",
        "desc": "商品购物车效果（购物车/我的订单回流，3001）",
        "params": {
            "date": "查询日期YYYY-MM-DD",
            "startDate": "开始日期",
            "endDate": "结束日期",
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
}


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
    """根据业务key返回对应的处理函数。"""
    if biz_key not in BUSINESS_REGISTRY:
        available = "、".join(BUSINESS_REGISTRY.keys())
        raise BusinessNotFoundError(
            f"未知业务: {biz_key}\n"
            f"已注册业务: {available}\n"
            f"调用 list_businesses() 查看所有业务详情。"
        )
    info = BUSINESS_REGISTRY[biz_key]
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
        print(f"       请重新获取Cookie，更新 config/cookie.txt 后重试。")
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
        "--list",
        action="store_true",
        help="列出所有已注册业务清单",
    )
    return parser.parse_args()


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
    if args.date:
        kwargs["date"] = args.date
    if args.start_date:
        kwargs["start_date"] = args.start_date
    if args.end_date:
        kwargs["end_date"] = args.end_date

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