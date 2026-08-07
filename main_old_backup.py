# -*- coding: utf-8 -*-
"""
main.py
FYA箱包旗舰店 - 京东商智数据导出 - 主程序入口

本文件集中存放：
    - 通用请求基类（JDBaseRequest）：Cookie管理 / 风控签名 / 30秒间隔 / 重试 / UA切换 / 日志 / Excel保存
    - 各业务API实现（按业务名分块）
    - 主程序入口 main()

代码组织原则（按全局agents.md第4条铁律）：
    - 不拆分大量独立py文件，所有接口集成在本文件
    - 每个API实现前用 ============ 分层注释隔离标记
    - 标注：接口业务名称、接口地址、参数说明，便于快速定位/修改/维护
"""

import os
import sys
import time
import json
import hashlib
import random
import logging
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


# ============================================================
#  通用请求基类（JDBaseRequest）
# ------------------------------------------------------------
#  业务名称：通用能力
#  接口地址：无（封装通用能力，被各业务API复用）
#  功能说明：
#      - Cookie 读取与更新（config/cookie.txt）
#      - 风控签名生成（User-mup / User-mnp / uuid）
#      - 30秒请求间隔控制
#      - 重试机制（最多3次，递增等待）
#      - Edge ↔ Chrome UA 自动切换
#      - 日志记录（按日期，文件+控制台）
#      - Excel 文件保存
#  复用方式：
#      各业务API类继承此类，直接调用 self.request(url, data) 即可。
# ============================================================
class JDBaseRequest:
    """京东商智API通用请求基类"""

    # ---------- 固定常量 ----------
    DEFAULT_REFERER = "https://sz.jd.com/szweb/sz/view/viewflow/flowPathDetailsNew.html"
    DEFAULT_ORIGIN = "https://sz.jd.com"
    DEFAULT_CONFIG_PATH = "config/config.xlsx"

    # 风控签名盐值默认值（config.xlsx 可覆盖）
    _DEFAULT_SIGN_SALT = "372ad2c2b6"
    UUID_PREFIX = "ca412182e5668a106054"
    UUID_RANDOM_DIGITS = 10

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

        # 控制参数
        self._last_request_time = 0
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
        通用请求方法（自动加风控签名、重试、UA切换）。

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
                self._wait_interval()
                risk_params = self._gen_risk_params(url, uuid_prefix=uuid_prefix)
                full_data = {**data, **risk_params}

                ua_name = "Edge" if self._current_ua_index == 0 else "Chrome"
                self.logger.info(f"发送请求 (第{attempt}/{self.MAX_RETRIES}次, UA={ua_name}): {url}")
                self.logger.debug(f"请求参数: {json.dumps(full_data, ensure_ascii=False)[:500]}")

                self._last_request_time = time.time()

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
#  业务接口 1：商品搜索效果 / 商品推荐效果（共享同一接口）
# ------------------------------------------------------------
#  业务名称：
#      - 商品搜索效果（lastSrcChannelId2=2008，搜索子来源）
#      - 商品推荐效果（lastSrcChannelId2=2009，推荐子来源）
#  接口地址：https://szgateway.jd.com/szpaas/szajax/shop/source/offlineFlowSource/downSkuTable.ajax
#  数据维度：店铺来源 → 搜索/推荐渠道 → SKU维度（按入店浏览量降序，最多5000条）
#  返回格式：Excel 二进制流（application/vnd.openxmlformats-officedocument.spreadsheetml.sheet）
#  参数说明：
#      业务参数（类常量，固定不变）：
#          interval=DAY, dateType=day
#          lastSrcChannelId1=2（搜索/推荐都是一级渠道2）
#          lastSrcChannelId2: 2008=搜索子来源, 2009=推荐子来源
#          groupType=skuId, attributes=skuId
#          sortField=jdr_sch_traffic_enter_shop__browse_page_cnt_shop_last_src, sortType=desc
#          limit=5000, compareType=hb
#      日期参数（每次可变）：
#          date / startDate / endDate → 优先用入参，其次从 config.xlsx 读取
#      关键发现（2026-08-04 第二次新增）：
#          商品搜索效果(2008)和商品推荐效果(2009)用同一个 downSkuTable.ajax 接口
#          只有 lastSrcChannelId2 不同，所以本类同时支持两个业务
# ============================================================
class ShopSourceAPI(JDBaseRequest):
    """店铺来源 - 搜索流量/推荐流量/购物车流量 - SKU维度 数据导出"""

    API_URL = "https://szgateway.jd.com/szpaas/szajax/shop/source/offlineFlowSource/downSkuTable.ajax"
    INTERVAL = "DAY"
    DATETYPE = "day"
    LAST_SRC_CHANNEL_ID1 = "2"       # 一级渠道：搜索/推荐/购物车都是2
    GROUP_TYPE = "skuId"
    ATTRIBUTES = "skuId"
    SORT_FIELD = "jdr_sch_traffic_enter_shop__browse_page_cnt_shop_last_src"
    SORT_TYPE = "desc"
    LIMIT = "5000"
    COMPARE_TYPE = "hb"

    # 渠道配置（lastSrcChannelId2 二级渠道ID + uuid前缀）
    # ----------------------------------------------------------------------
    # 中文说明（小白必读）：
    #   CHANNEL_MAP 是一个字典，key 是业务名（中文友好），
    #   value 是一个元组 (二级渠道ID, uuid前缀)。
    #
    #   二级渠道ID（lastSrcChannelId2）：京东商智后台给每个流量子来源分配的编号
    #       2008 = 搜索子来源 → 商品搜索效果
    #       2009 = 推荐子来源 → 商品推荐效果
    #       3001 = 购物车子来源 → 商品购物车效果
    #
    #   uuid前缀：京东风控校验用的随机ID前缀
    #       不同业务/页面前缀可能不一样！
    #       搜索/推荐：ca412182e5668a106054
    #       购物车  ：5f9cc2ca20cad3d11642
    #
    #   警告：uuid前缀一定要按渠道配置，不能写死成全局常量！
    # ----------------------------------------------------------------------
    CHANNEL_MAP = {
        "搜索":   ("2008", "ca412182e5668a106054"),
        "推荐":   ("2009", "ca412182e5668a106054"),
        "购物车": ("3001", "5f9cc2ca20cad3d11642"),
    }

    # 反向索引表：二级渠道ID → uuid前缀（用于向下兼容直接传channel_id2的场景）
    # ----------------------------------------------------------------------
    # 中文说明（小白必读）：
    #   这个字典是从 CHANNEL_MAP 自动生成的"反向索引"。
    #   作用：如果你直接传入二级渠道ID（比如 "3001"），也能找到对应的uuid前缀。
    #   自动构建：调用 _build_channel_id_index() 时会从 CHANNEL_MAP 反向生成。
    # ----------------------------------------------------------------------
    _CHANNEL_ID_INDEX = None  # 延迟到首次调用时构建

    @classmethod
    def _build_channel_id_index(cls):
        """从 CHANNEL_MAP 构建反向索引：{channel_id2: uuid_prefix}
        用于支持直接传入 channel_id2 的向下兼容场景。
        """
        index = {}
        for _channel_name, (channel_id2, uuid_prefix) in cls.CHANNEL_MAP.items():
            index[channel_id2] = uuid_prefix
        return index

    def _get_channel_config(self, channel):
        """获取渠道配置 (channel_id2, uuid_prefix)。

        支持两种调用方式（向下兼容）：
          1. 传业务名（推荐）：channel="购物车"
          2. 传二级渠道ID（兼容）：channel="3001"

        未注册时会抛错，并列出所有可用值。
        """
        # 方式1：业务名直接查
        if channel in self.CHANNEL_MAP:
            return self.CHANNEL_MAP[channel]

        # 方式2：二级渠道ID反向查（向下兼容老代码）
        if self._CHANNEL_ID_INDEX is None:
            self._CHANNEL_ID_INDEX = self._build_channel_id_index()
        if channel in self._CHANNEL_ID_INDEX:
            uuid_prefix = self._CHANNEL_ID_INDEX[channel]
            # 找出对应的业务名（用于日志）
            for name, (cid, _) in self.CHANNEL_MAP.items():
                if cid == channel:
                    self.logger.info(f"通过二级渠道ID '{channel}' 匹配到业务 '{name}'")
                    break
            return (channel, uuid_prefix)

        # 都不匹配：报错
        available_names = "、".join(self.CHANNEL_MAP.keys())
        available_ids = "、".join(self._CHANNEL_ID_INDEX.keys())
        raise ValueError(
            f"不支持的渠道: {channel}\n"
            f"可用业务名: {available_names}\n"
            f"可用二级渠道ID: {available_ids}"
        )

    def _get_uuid_for_channel(self, channel):
        """根据渠道名或渠道ID，返回对应的完整uuid（格式：前缀-10位随机数）。

        中文说明（小白必读）：
          uuid = uuid前缀 + "-" + 10位随机数字
          例如：ca412182e5668a106054-1234567890
        """
        _, uuid_prefix = self._get_channel_config(channel)
        random_min = 10 ** (self.UUID_RANDOM_DIGITS - 1)  # 1000000000
        random_max = 10 ** self.UUID_RANDOM_DIGITS - 1     # 9999999999
        return f"{uuid_prefix}-{random.randint(random_min, random_max)}"

    def download_sku(self, date=None, start_date=None, end_date=None, channel="搜索"):
        if date is None:
            date = self.config.get("date", "")
        if start_date is None:
            start_date = self.config.get("startDate", date)
        if end_date is None:
            end_date = self.config.get("endDate", date)

        channel_id2, uuid_prefix = self._get_channel_config(channel)

        # 中文说明：把channel归一化为友好业务名（用于日志和文件名）
        #   如果传入的是 "2009" 这种channel_id2，转换为 "推荐"
        #   如果传入的是 "购物车" 这种业务名，保持不变
        display_channel = self._resolve_channel_display_name(channel, channel_id2)

        data = {
            "date": date,
            "startDate": start_date,
            "endDate": end_date,
            "interval": self.INTERVAL,
            "dateType": self.DATETYPE,
            "lastSrcChannelId1": self.LAST_SRC_CHANNEL_ID1,
            "lastSrcChannelId2": channel_id2,
            "groupType": self.GROUP_TYPE,
            "attributes": self.ATTRIBUTES,
            "sortField": self.SORT_FIELD,
            "sortType": self.SORT_TYPE,
            "limit": self.LIMIT,
            "compareType": self.COMPARE_TYPE,
        }

        self.logger.info(f"下载店铺来源数据: 日期={date}, 渠道={display_channel}(id2={channel_id2}, uuid_prefix={uuid_prefix[:8]}...)")
        response = self.request(self.API_URL, data, uuid_prefix=uuid_prefix)
        # 文件名用友好业务名（即使传入的是channel_id2也能得到"推荐流量_xxx.xlsx"）
        filename = f"{display_channel}流量_{date}.xlsx"
        return self.save_excel(response, filename)

    def _resolve_channel_display_name(self, channel, channel_id2):
        """把channel归一化为友好业务名。
        中文说明（小白必读）：
          输入"购物车"→ 返回"购物车"（业务名，直接用）
          输入"3001"  → 返回"购物车"（通过channel_id2反查业务名）
          输入未注册的→ 返回channel本身（兜底）
        """
        # 已经是业务名
        if channel in self.CHANNEL_MAP:
            return channel
        # 是channel_id2，反查业务名
        for name, (cid, _) in self.CHANNEL_MAP.items():
            if cid == channel_id2:
                return name
        # 兜底：用原值
        return channel

    def download_search_sku(self, date=None, start_date=None, end_date=None):
        """便捷方法：导出搜索流量-SKU维度数据（商品搜索效果业务）"""
        return self.download_sku(date=date, start_date=start_date, end_date=end_date, channel="搜索")

    def download_recommend_sku(self, date=None, start_date=None, end_date=None):
        """便捷方法：导出推荐流量-SKU维度数据（商品推荐效果业务）"""
        return self.download_sku(date=date, start_date=start_date, end_date=end_date, channel="推荐")

    def download_cart_sku(self, date=None, start_date=None, end_date=None):
        """便捷方法：导出购物车流量-SKU维度数据（商品购物车效果业务）"""
        return self.download_sku(date=date, start_date=start_date, end_date=end_date, channel="购物车")


# ============================================================
#  业务接口 2：（占位 - 后续追加）
# ------------------------------------------------------------
#  业务名称：TODO 例如"首页流量-SKU维度"
#  接口地址：TODO 例如 https://szgateway.jd.com/...
#  参数说明：TODO 列出该接口固定参数 + 可变参数
# ============================================================
# class HomePageSourceAPI(JDBaseRequest):
#     """店铺来源 - 首页流量 - SKU维度（占位，未实现）"""
#
#     API_URL = "TODO 接口URL"
#     # TODO: 列出该接口的类常量（interval/channelId/sortField/limit 等）
#
#     def download_sku(self, date=None, start_date=None, end_date=None):
#         # TODO: 组装业务参数，复用 JDBaseRequest.request()
#         raise NotImplementedError("该接口尚未实现")


# ============================================================
#  业务接口 3：（占位 - 后续追加）
# ------------------------------------------------------------
#  业务名称：TODO
#  接口地址：TODO
#  参数说明：TODO
# ============================================================
# class CategorySourceAPI(JDBaseRequest):
#     """店铺来源 - 类目流量 - SKU维度（占位，未实现）"""
#     pass


# ============================================================
#  业务接口 4：（占位 - 后续追加）
# ------------------------------------------------------------
#  业务名称：TODO
#  接口地址：TODO
#  参数说明：TODO
# ============================================================
# class NewBizAPIxxxAPI(JDBaseRequest):
#     """TODO 业务名（占位，未实现）"""
#     pass


# ============================================================
#  业务接口调度入口
# ------------------------------------------------------------
#  按业务名分发到对应API类，集中管理后续新增接口
# ============================================================
def run_business(business_name, date=None, start_date=None, end_date=None, **kwargs):
    """
    业务分发器：
        business_name    业务名（与下方 MAPPING 中的 key 对应）
        date/start_date/end_date    日期参数
        **kwargs         其他业务特定参数
    """
    factory = {
        "商品搜索效果": lambda: ShopSourceAPI().download_search_sku(
            date=date, start_date=start_date, end_date=end_date
        ),
        "商品推荐效果": lambda: ShopSourceAPI().download_recommend_sku(
            date=date, start_date=start_date, end_date=end_date
        ),
        "商品购物车效果": lambda: ShopSourceAPI().download_cart_sku(
            date=date, start_date=start_date, end_date=end_date
        ),
        # "首页流量":  lambda: HomePageSourceAPI().download_sku(date=date, start_date=start_date, end_date=end_date),
        # "类目流量":  lambda: CategorySourceAPI().download_sku(date=date, start_date=start_date, end_date=end_date),
    }

    if business_name not in factory:
        raise ValueError(f"未知业务: {business_name}，可选: {list(factory.keys())}")

    print(f"[INFO] 执行业务: {business_name} (店铺 {SHOP_NAME})")
    return factory[business_name]()


# ============================================================
#  主程序入口
# ============================================================
def main():
    print("=" * 60)
    print(f"京东商智 - 数据导出工具    店铺: {SHOP_NAME}")
    print("=" * 60)

    # 查询日期（从 config.xlsx 读取，便于一处改全局生效）
    # 中文说明：从ShopSourceAPI实例的config字典读取date变量，遵循"配置集中管理"原则
    _api = ShopSourceAPI()
    date = _api.config.get("date", "2026-08-04")  # 占位默认
    business_name = "商品购物车效果"  # 本次要跑的商品购物车效果业务（也可改为其他业务）

    print(f"\n即将导出: 业务={business_name}, 日期={date}")
    print("-" * 60)

    try:
        file_path = run_business(business_name, date=date)
        print(f"\n[OK] 导出成功！")
        print(f"     文件: {file_path}")
    except CookieExpiredError as e:
        print(f"\n[ERR] Cookie已过期: {e}")
        print(f"       请重新获取Cookie，更新 config/cookie.txt 后重试。")
    except Exception as e:
        print(f"\n[ERR] 导出失败: {e}")
        print(f"       详细日志请查看 logs/ 目录下的日志文件。")


if __name__ == "__main__":
    main()


# ============================================================
# 【业务上下文备份 2026-08-05】—— 仅存档查阅，不作为运行入口
# 实际业务入口：main.py（本文件为旧版存档）
# ------------------------------------------------------------
# 一、本次需求变更记录（4项）
#   1. config.xlsx 优化：所有【变量参数】补充中文【说明】列，参数含义一目了然
#   2. 项目名统一：商品搜索效果 → 商品流量来源（搜索/推荐/购物车统一归属管理）
#   3. 业务执行规则调整：自主访问流量与购物车数据口径重叠（同3001），
#      执行任务时不触发自主访问导出，仅保留搜索/推荐/购物车三渠道执行；
#      自主访问注册配置保留不删除，仅调度层（BUSINESS_REGISTRY enabled=False）过滤不执行
#   4. 本次全部业务上下文需求变更同步备份至此文件，便于本地查阅回溯
#
# 二、渠道业务上下文
#   业务key              | 二级渠道ID | uuid前缀             | 状态
#   商品流量来源_搜索     | 2008       | ca412182e5668a106054 | 执行
#   商品流量来源_推荐     | 2009       | ca412182e5668a106054 | 执行
#   商品流量来源_购物车   | 3001       | 5f9cc2ca20cad3d11642 | 执行
#   商品流量来源_自主访问 | 3001       | 5f9cc2ca20cad3d11642 | 已停用(与购物车口径重叠)
#
# 三、接口信息
#   接口地址：https://szgateway.jd.com/szpaas/szajax/shop/source/offlineFlowSource/downSkuTable.ajax
#   请求方式：POST（application/x-www-form-urlencoded）
#   风控签名：User-mnp = MD5(URL路径 + uuid + 时间戳 + 盐值372ad2c2b6)
#   返回格式：Excel 二进制流
#   业务参数：可变参数（来自config.xlsx）interval/dateType/limit；
#             固定常量（2026-08-05经用户确认写死代码）lastSrcChannelId1/groupType/attributes/sortField/sortType/compareType
#   日期参数（来自config.xlsx或命令行动态传入）：date/startDate/endDate
#
# 四、后续业务规划（业务注册中心扩展，无需改动调度核心）
#   1. 店铺来源报表（商智）
#   2. 订单明细报表（京麦 seller-v10.shop.jd.com）
#   3. 售后订单报表（京麦）
#   4. 京准通推广数据报表（jzt.jd.com）
#   接入方式：定义API类继承JDBaseRequest → 在BUSINESS_REGISTRY注册 → run_business()触发
#
# 五、执行规则
#   默认执行：python main.py --date "YYYY-MM-DD"（搜索/推荐/购物车 3渠道）
#   单渠道执行：python main.py --biz_key "商品流量来源_购物车" --date "YYYY-MM-DD"
#   停用业务重新开启：将 BUSINESS_REGISTRY 中"商品流量来源_自主访问"的 enabled 改为 True
# ============================================================

# ============================================================
# 【业务上下文备份 2026-08-05（二）】Excel后置处理 + 日期修复 —— 仅存档查阅
# ------------------------------------------------------------
# 一、Excel后置处理（全局生效，所有报表复用）
#   背景：接口返回原始Excel无日期列 → 导出后处理，并统一数值/格式规范
#   新增3个公共工具函数（main.py 公共工具区，禁止硬编码业务逻辑）：
#     ① convert_date_format(date_str)
#        - 日期统一转 yyyy/m/d（月/日不补零），支持 8位纯数字/横杠/斜杠，带时间则时间原样保留
#        - 无法识别返回原值，不报错
#     ② safe_convert_numeric(df)
#        - 强制文本黑名单 TEXT_FORCE_COLUMNS={"订单编号"}：整列跳过转换保留文本（对齐京东订单导出风险提示）
#        - 整数0位小数白名单 INTEGER_ZERO_DECIMAL_COLUMNS={"SKU","SPU"}：允许转数字
#        - 纯数字>15位保留文本（防精度丢失）；转换失败保留原值
#        - 列名匹配用 _col_matches() 结尾匹配："商品SKU"命中"SKU"，但"成交金额（SPU）"不命中
#     ③ apply_column_formats(file_path, df, date_column, date_value)
#        - 订单编号列=@文本格式；SKU/SPU列=数值0格式（无千分位，仅对数字单元格）
#        - 日期列=真实datetime + yyyy/m/d（带时间用 yyyy/m/d hh:mm:ss），打开不弹格式警告
#   导出流程：读Excel(dtype=str防pandas自动转数值) → 转换日期 → 首列A插【日期】→ 数值安全转换 → 写回+设格式
#
# 二、⚠️ 日期同步修复（重大坑，2026-08-05）
#   现象：导出的 2026-07-30 数据与网页对不上；07-29 与 07-30 导出完全相同
#   根因：_get_date_params() 中 startDate/endDate 回落 config 旧值(07-29)，
#         --date 2026-07-30 时实际发送 date=07-30&startDate=07-29&endDate=07-29，
#         接口按 startDate~endDate 区间取数 → 返回 07-29 数据
#   修复：start/end 未显式传入时默认=date（三值一致）
#   验证：修复后 3渠道与网页导出行数/SKU/数值完全一致（搜索43/推荐4/购物车4行）
#   经验：--date 覆盖必须三日期同步；uuid前缀是前端动态生成（与数据无关）
#
# 三、Mock验证方案（可复用）
#   替换 JDBaseRequest.request 返回含渠道ID(2008/2009/3001)的模拟xlsx、
#   跳过30秒间隔、输出隔离 output_mock，走真实 run_business() 批量调度验证导出链路。
#   曾借此发现：pandas read_excel 默认把数字样式列转 int64 导致>15位保护失效 → 必须 dtype=str。
# ============================================================

# ============================================================
# 【业务上下文备份 2026-08-06 - 京麦订单明细【加密】导出项目｜状态：暂停归档】
#
# 1. 项目基本信息
#   - 项目名：京麦订单明细【加密】导出
#   - 业务域：京麦 seller-v10.shop.jd.com / 真实入口 shop.jd.com
#   - 项目状态：暂停归档（账号/IP 触发 601 风控限流）
#   - 启动时间：2026-08-06
#   - 暂停原因：连续多次 Chrome/Edge 抓包均收到 code=601「操作频繁，请稍后重试」，
#               判断非脚本 BUG，是账号/IP 临时限流。继续重试会加重风控标记。
#
# 2. 项目交付物（保存状态）
#   - jd_cdp_capture.py：抓包脚本（Playwright + CDP 全局监听），已就绪
#   - cdp_network_log.json：抓包日志（601 限流样本，17.5 MB，1641 条请求，2 条 exportCenterService）
#   - cdp_user_data/：浏览器用户数据目录（暂停前已清理，重启项目可重建）
#
# 3. 真实业务信息（关键发现）
#   - 真实入口页面：https://shop.jd.com/jdm/trade/tools/export/ExprotList
#   - 真实接口域名：sff.jd.com（不是 seller-v10.shop.jd.com）
#   - 接口路径模板：/api?v=1.0&appId=CQLEJWPYPFOVQBC8UFLQ&api=dsm.order.export.exportCenterService.<接口名>
#   - 关键 header：
#     * h5st（前端强签名，一次性，不可复用）
#     * dsm-eid（设备指纹）
#     * x-referer-page（来源页标识）
#     * x-rp-client=h5_2.4.0（客户端标识）
#     * dsm-platform=pc, dsm-lang=zh-CN, dsm-trace-id（追踪）
#     * Anti-Content（风控token，响应里 set-cookie 返回）
#   - 5 个目标接口：
#     1) countDown（前置限流校验）
#     2) createdExportTask（创建导出任务，⚠️结尾带ed，拼错返回301）
#     3) queryExportTaskInfo（轮询任务状态）
#     4) exportTaskPwdSend（申请密码短信，接口不返回密码明文）
#     5) export.action（GET 下载 zip 包）
#
# 4. 抓包架构（已确定）
#   - Playwright + launch_persistent_context（必须用持久化上下文，普通 launch 不支持 --user-data-dir）
#   - 全局监听 Network.requestWillBeSent + Network.responseReceived
#   - 全量抓包 + 关键词标记（exportCenterService），解析阶段再筛选
#   - 响应体 >10MB 截断防爆日志
#   - 30 分钟监听超时（防止忘按回车）
#   - 浏览器探测顺序：Chrome → Edge（项目恢复后可调整）
#   - 独立用户配置目录：cdp_user_data/（Chrome/Edge 不可共用，需清理）
#
# 5. 风控硬性约束（项目恢复时必须遵守）
#   - 601 触发后 30-120 分钟冷却；冷却期间任何导出请求都会重置冷却
#   - 禁止多端并发操作同一账号的导出模块
#   - 严禁代理/VPN/IP 池访问京麦
#   - 冷却无效可换手机热点换公网 IP
#   - h5st 必须真实浏览器实时生成，禁止硬编码
#   - 脚本不允许自动重试 601
#   - 抓包行为：登录后首页静置 1-2 分钟，缓慢操作
#
# 6. 项目恢复前置条件（必须全部满足）
#   - 收到明确【项目恢复指令】
#   - 账号完成 30-120 分钟冷却或更换干净公网 IP
#   - 重新运行 jd_cdp_capture.py 抓到 code=200 成功响应
#   - 拿到成功报文后再开发 jd_order_export.py
#
# 7. 拟开发的 jd_order_export.py 框架（待恢复后实施）
#   - 5 接口完整链路：countDown → createdExportTask → queryExportTaskInfo
#     → exportTaskPwdSend → export.action
#   - 可选 IMAP 模块（默认关闭）：
#     * iPhone 快捷指令监听京东短信 → 投递到 QQ 邮箱
#     * 程序 IMAP 读取 QQ 邮箱解析 taskId + password
#     * 关闭时打印 taskId + zip 路径，提示手动输入密码
#   - 边界处理：
#     * 时间跨度 >31 天直接拦截
#     * 同类型 10 分钟间隔 code=201 提示冷却
#     * 单 taskId 60s 密码申请间隔
#     * 邮箱轮询超时收不到密码保留 zip 退出
#     * 解压失败保留 zip 提示排查
#   - 敏感参数（IMAP 授权码、邮箱账号）从配置文件读取，不硬编码
# ============================================================


# ============================================================
# 【业务上下文备份 2026-08-07】店铺来源-三级渠道（离线流量报表）
# ------------------------------------------------------------
# 本区块为「项目 4：店铺来源-三级渠道」完整业务上下文存档，
# 不影响主程序运行；仅用于人工查阅 / 项目回溯 / 阶段性对比。
#
# 上线时间：2026-08-07
# 状态：✅ 已上线（业务类实现 + 容错 + 文档归档完成）
# 项目目标：按三级流量渠道分组，导出店铺来源离线日度流量报表
# ============================================================

# 1. 真实业务信息（来自用户 2026-08-06/07 抓包）
# ------------------------------------------------------------
# - 接口地址：https://szgateway.jd.com/szpaas/szajax/shop/source/offlineFlowSource/downTable.ajax
# - 请求方法：POST application/x-www-form-urlencoded
# - 返回内容：Excel 二进制（magic bytes: PK\x03\x04）
# - 真实入口页面：https://sz.jd.com/szweb/sz/view/viewflow/viewSourcesVNew.html
# - 必带 Header（缺失即拦截）：
#       Origin: https://sz.jd.com
#       Referer: https://sz.jd.com/szweb/sz/view/viewflow/viewSourcesVNew.html

# 2. 业务表单参数（13 项 = 9 固定 + 1 可变 + 3 风控动态）
# ------------------------------------------------------------
# 固定常量（用户确认固化为代码常量，不读 config）：
#   compareType=hb / interval=DAY / dateType=day / downType=day
#   groupType=lastSrcChannelId3 / attributes=lastSrcChannelId3
#   sortField=jdr_sch_traffic_enter_shop__visitor_cnt_shop_last_src
#   sortType=desc / lastSrcChannelId1=2
# 可变参数（从 config 读取）：
#   platformCate1=""（空=全品类）/ date / startDate / endDate
# 风控动态（运行时生成，不入代码）：
#   User-mup / User-mnp / uuid

# 3. 风控签名（与项目 1-3 复用 MD5 公式，盐值 372ad2c2b6 共用）
# ------------------------------------------------------------
# User-mnp = MD5(URL路径 + uuid + 时间戳 + 盐值)
# 算法来源：commons-a5562705.js 逆向（与项目1-3 同源）

# 4. UUID 完全随机（与项目1-3 关键差异）
# ------------------------------------------------------------
# 项目1-3：固定 prefix（如 ca412182e5668a106054）+ 随机后缀
# 项目 4：完全随机（prefix 也随机）
#
# 用户两次抓包（间隔 13 秒）：
#   抓包 1：uuid=f1d5ae161b41f4153fc0-19fd685e6a1
#   抓包 2：uuid=a31e066d8e94f4f39a3a-19fda02c2d4
# 前缀完全不同 → 必须完全随机化
#
# Python 实现：secrets.token_hex(8) + secrets.token_hex(5)
# 优势：加密随机 + 不依赖基类 UUID_PREFIX + 符合"禁止硬编码 uuid"

# 5. 容错与风控适配（阶段 4 新增）
# ------------------------------------------------------------
# - 重试循环：3 次递增等待 30/60/90 秒
# - UA 切换：每次重试前 Edge↔Chrome
# - 601 限流：不重试（避免加重风控，让用户决定）
# - Cookie 过期：抛 CookieExpiredError 立即停
# - 空响应拦截：HTTP 200 + <1KB → 视为失败
# - Excel 字节校验：magic bytes 不等于 PK\x03\x04 → 视为失败

# 6. 集成方式
# ------------------------------------------------------------
# - 新业务类：OfflineChannelAPI（继承 JDBaseRequest）
# - 注册入口：BUSINESS_REGISTRY["店铺来源_三级渠道"]
# - 调用命令：python main.py --biz_key "店铺来源_三级渠道" --date "2026-08-04"
# - 输出文件：output/店铺来源_三级渠道_YYYY-MM-DD.xlsx

# 7. 阶段交付节奏（5 阶段，每阶段输出总结等你确认）
# ------------------------------------------------------------
# 阶段 1：需求拆解 + 方案选型 + 风险梳理（基于用户 2 次抓包）
# 阶段 2：项目骨架搭建（评估 config/基类/JDBaseRequest 能力，零改动验证）
# 阶段 3：核心接口逻辑实现（OfflineChannelAPI 类 + UUID 完全随机 + 业务方法）
# 阶段 4：容错与风控适配（重试循环 + UA 切换 + 空响应拦截 + 风控业务码识别）
# 阶段 5：测试 + 文档归档（状态校验 + SKILL.md 沉淀 + 踩坑日志 + API 说明文档）

# 8. 踩坑要点（与 SKILL.md / 踩坑日志同步）
# ------------------------------------------------------------
# 1. UUID 策略分业务：项目1-3 固定 prefix，本项目完全随机
# 2. 必带 Header 缺失被拦截（Origin/Referer 必须传）
# 3. sortField 字段名按业务调整（浏览量 → 访客数）
# 4. 601 限流不重试（与京麦项目 SKILL 第八节一致）
# 5. 业务表单参数新增 downType / platformCate1（项目1-3 没有）

# 9. 关联文档
# ------------------------------------------------------------
# - 业务类实现：main.py（搜索 class OfflineChannelAPI）
# - 业务沉淀：.trae/skills/jd-api-analyze/SKILL.md 项目4
# - 踩坑记录：全局复利的踩坑日志.md 坑6（UUID 完全随机 vs 固定 prefix）
# - 接口说明：docs/API 实现逻辑说明.md 项目 4
# - 文档索引：docs/项目文档索引.xlsx（自动入库）
# ============================================================

