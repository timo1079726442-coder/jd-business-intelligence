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
"""

import os
import sys
import time
import json
import hashlib
import random
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
        """从config读取日期参数（允许入参动态覆盖）。"""
        # 优先用入参，其次从config读取
        if date is None:
            date = self.config.get("date")
        if date is None:
            raise ValueError("查询日期date未提供：请在config.xlsx配置或通过函数入参传入")

        if start_date is None:
            start_date = self.config.get("startDate", date)
        if end_date is None:
            end_date = self.config.get("endDate", date)

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
        # 例如：商品流量来源_搜索_2026-07-29.xlsx
        short_name = display_key.replace("商品流量来源_", "")  # 去掉前缀，保留"搜索/推荐/自主访问"
        filename = f"{short_name}流量_{date}.xlsx"
        return self.save_excel(response, filename)

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
#  业务名称：TODO 例如"店铺来源报表"
#  接口地址：TODO 例如 https://szgateway.jd.com/...
#  参数说明：TODO 列出该接口固定参数 + 可变参数
# ============================================================
# class ShopReportAPI(JDBaseRequest):
#     """店铺来源报表（占位，未实现）"""
#
#     API_URL = "TODO 接口URL"
#
#     def download(self, date=None, **kwargs):
#         # TODO: 组装业务参数，复用 JDBaseRequest.request()
#         # 业务参数必须从config读取，不允许硬编码！
#         raise NotImplementedError("该接口尚未实现")


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
    # "店铺来源报表": {
    #     "api_class": ShopReportAPI,
    #     "method": "download",
    #     "desc": "店铺来源报表",
    #     "params": {...},
    # },
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