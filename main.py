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

    # 渠道配置（lastSrcChannelId2 + uuid前缀）
    # 2008 = 搜索子来源（商品搜索效果）→ uuid前缀 ca412182e5668a106054
    # 2009 = 推荐子来源（商品推荐效果）→ uuid前缀 ca412182e5668a106054
    # 3001 = 购物车子来源（商品购物车效果）→ uuid前缀 5f9cc2ca20cad3d11642
    # ⚠️ 关键发现：uuid前缀在不同渠道/页面可能不一样，需可配置
    CHANNEL_MAP = {
        "搜索": ("2008", "ca412182e5668a106054"),
        "推荐": ("2009", "ca412182e5668a106054"),
        "购物车": ("3001", "5f9cc2ca20cad3d11642"),
    }

    def _get_channel_config(self, channel):
        """获取渠道配置 (channel_id2, uuid_prefix)，未注册渠道抛错"""
        if channel not in self.CHANNEL_MAP:
            available = "、".join(self.CHANNEL_MAP.keys())
            raise ValueError(f"不支持的渠道: {channel}\n当前可用渠道: {available}")
        return self.CHANNEL_MAP[channel]

    def _get_uuid_for_channel(self, channel):
        """根据渠道返回对应的uuid前缀，组装成完整uuid"""
        _, uuid_prefix = self._get_channel_config(channel)
        random_min = 10 ** (self.UUID_RANDOM_DIGITS - 1)
        random_max = 10 ** self.UUID_RANDOM_DIGITS - 1
        return f"{uuid_prefix}-{random.randint(random_min, random_max)}"

    def download_sku(self, date=None, start_date=None, end_date=None, channel="搜索"):
        if date is None:
            date = self.config.get("date", "")
        if start_date is None:
            start_date = self.config.get("startDate", date)
        if end_date is None:
            end_date = self.config.get("endDate", date)

        channel_id2, uuid_prefix = self._get_channel_config(channel)

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

        self.logger.info(f"下载店铺来源数据: 日期={date}, 渠道={channel}(id2={channel_id2}, uuid_prefix={uuid_prefix[:8]}...)")
        response = self.request(self.API_URL, data, uuid_prefix=uuid_prefix)
        filename = f"{channel}流量_{date}.xlsx"
        return self.save_excel(response, filename)

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

    # 查询日期（可按需修改，或改为 sys.argv 接收）
    date = "2026-08-03"
    business_name = "商品购物车效果"  # 本次要跑的商品购物车效果业务（也可改为其他业务）
    date = "2026-08-04"  # 与抓包中的日期一致

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
