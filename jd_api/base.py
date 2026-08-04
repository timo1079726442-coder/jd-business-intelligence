# -*- coding: utf-8 -*-
"""
jd_api/base.py
京东商智API通用请求基类
功能：Cookie管理、风控签名生成、30秒间隔控制、重试机制、日志记录
所有具体API接口类继承此类，直接复用。
"""
import os
import time
import json
import hashlib
import random
import logging
from datetime import datetime
import requests


class JDBaseRequest:
    """京东商智API通用请求基类"""

    # ============ 固定常量（算法/标识，作为代码常量） ============
    # 请求来源页面
    DEFAULT_REFERER = "https://sz.jd.com/szweb/sz/view/viewflow/flowPathDetailsNew.html"
    # 站点域名
    DEFAULT_ORIGIN = "https://sz.jd.com"
    # config.xlsx配置文件路径
    DEFAULT_CONFIG_PATH = "config/config.xlsx"

    # 风控签名盐值 - 默认值（如果config.xlsx中没配置则用此值）
    # 盐值从commons.js逆向获得，是算法的一部分。但京东可能会更新，所以放在config.xlsx中方便用户手动修改
    _DEFAULT_SIGN_SALT = "372ad2c2b6"
    # UUID固定前缀（算法固定的输入）
    UUID_PREFIX = "ca412182e5668a106054"
    # UUID后半段随机数字位数（算法固定的位数）
    UUID_RANDOM_DIGITS = 10

    # 浏览器UA（硬编码，无需配置）
    UA_EDGE = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/144.0.0.0 Safari/537.36 Edg/144.0.0.0"
    UA_CHROME = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/144.0.0.0 Safari/537.36"
    SEC_CH_UA_EDGE = '"Not(A:Brand";v="8", "Chromium";v="144", "Microsoft Edge";v="144"'
    SEC_CH_UA_CHROME = '"Not(A:Brand";v="8", "Chromium";v="144", "Google Chrome";v="144"'

    def __init__(self, cookie_path=None, config_path=None, log_dir=None):
        """
        初始化

        参数:
            cookie_path : cookie.txt路径，默认为 config/cookie.txt
            config_path : config.xlsx路径，默认为 config/config.xlsx
            log_dir     : 日志目录，默认为 logs/
        """
        # 项目根目录
        project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

        # 加载配置文件
        self.config_path = config_path or os.path.join(project_root, self.DEFAULT_CONFIG_PATH)
        self.config = self._load_config()

        # Cookie路径（优先从配置读取）
        cookie_path_from_config = self.config.get("cookie文件路径")
        if cookie_path_from_config:
            # 配置中是相对路径，拼到项目根目录
            self.cookie_path = cookie_path_from_config if os.path.isabs(cookie_path_from_config) else os.path.join(project_root, cookie_path_from_config)
        else:
            self.cookie_path = cookie_path or os.path.join(project_root, "config", "cookie.txt")

        # 输出目录（从配置读取）
        output_dir_rel = self.config.get("输出目录", "output/")
        self.output_dir = output_dir_rel if os.path.isabs(output_dir_rel) else os.path.join(project_root, output_dir_rel)
        os.makedirs(self.output_dir, exist_ok=True)

        # 日志目录
        self.log_dir = log_dir or os.path.join(project_root, "logs")
        os.makedirs(self.log_dir, exist_ok=True)

        # 上次请求时间（用于30秒间隔控制）
        self._last_request_time = 0

        # 请求控制参数（从配置读取）
        self.REQUEST_INTERVAL = int(self.config.get("请求间隔(秒)", "30"))
        self.MAX_RETRIES = int(self.config.get("最大重试次数", "3"))
        self.REQUEST_TIMEOUT = int(self.config.get("请求超时(秒)", "30"))

        # 风控签名盐值（从配置读取，京东更新后可手动改配置）
        self.SIGN_SALT = self.config.get("签名盐值", self._DEFAULT_SIGN_SALT)

        # 当前UA索引：0=Edge, 1=Chrome
        self._ua_list = [
            (self.UA_EDGE, self.SEC_CH_UA_EDGE),
            (self.UA_CHROME, self.SEC_CH_UA_CHROME),
        ]
        self._current_ua_index = 0

        # 读取Cookie
        self.cookie_str = self._read_cookie()

        # 初始化日志
        self.logger = self._init_logger()
        self.logger.info(f"签名盐值: {self.SIGN_SALT}")

        # 创建Session（复用连接）
        self.session = requests.Session()
        self.session.headers.update(self._build_default_headers())

        self.logger.info(f"JDBaseRequest 初始化完成，当前UA: {'Edge' if self._current_ua_index == 0 else 'Chrome'}")

    def _load_config(self):
        """
        从 config.xlsx 加载配置

        返回:
            dict: {变量参数: 参数值} 的映射
        """
        from openpyxl import load_workbook

        if not os.path.exists(self.config_path):
            print(f"⚠️ 配置文件不存在: {self.config_path}，将使用默认值")
            return {}

        try:
            wb = load_workbook(self.config_path, read_only=True, data_only=True)
            config_dict = {}

            # 遍历所有Sheet
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                # 第一行是表头，从第二行开始读数据
                # 列：项目名(A) | 变量参数(B) | 参数值(C) | 说明(D)
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if row and len(row) >= 3:
                        # 项目名、变量参数、参数值
                        project_name = row[0]
                        var_name = row[1]
                        var_value = row[2]

                        if var_name and var_value is not None:
                            # 同一个变量参数可能被多个项目引用，用变量参数名作为key
                            config_dict[str(var_name)] = str(var_value)

            wb.close()
            print(f"✅ 配置文件加载成功: {len(config_dict)} 项配置")
            return config_dict

        except Exception as e:
            print(f"⚠️ 配置文件加载失败: {e}，将使用默认值")
            return {}

    # ==================== Cookie管理 ====================

    def _read_cookie(self):
        """从 config/cookie.txt 读取Cookie"""
        if not os.path.exists(self.cookie_path):
            raise FileNotFoundError(
                f"Cookie文件不存在: {self.cookie_path}\n"
                f"请将京东商智的Cookie保存到此文件中。"
            )
        with open(self.cookie_path, "r", encoding="utf-8") as f:
            cookie_str = f.read().strip()
        if not cookie_str:
            raise ValueError(f"Cookie文件为空: {self.cookie_path}")
        return cookie_str

    def refresh_cookie(self, cookie_str=None):
        """
        更新Cookie（Cookie过期时调用）

        参数:
            cookie_str: 新Cookie字符串。如果为None，则重新从文件读取。
        """
        if cookie_str:
            # 直接写入文件
            with open(self.cookie_path, "w", encoding="utf-8") as f:
                f.write(cookie_str)
            self.cookie_str = cookie_str
        else:
            # 从文件重新读取
            self.cookie_str = self._read_cookie()
        # 更新Session的Cookie
        self.session.headers["Cookie"] = self.cookie_str
        self.logger.info("Cookie已更新")

    # ==================== 风控参数生成 ====================

    def _gen_risk_params(self, url):
        """
        生成京东风控参数（User-mup, User-mnp, uuid）

        签名算法（从 commons-a5562705.js 逆向获得）:
            User-mnp = MD5(URL路径 + uuid + 时间戳 + 盐值)

        参数:
            url: 完整请求URL
        返回:
            dict: {"User-mup": "...", "User-mnp": "...", "uuid": "..."}
        """
        # 1. 当前毫秒时间戳
        timestamp = int(time.time() * 1000)

        # 2. 生成uuid（固定前缀 + N位随机数字，N从配置文件读取）
        # 注：京东真实uuid是由前端SDK运行时动态生成（ca412182e5668a106054-xxx），
        #    无法从静态JS文件中还原算法。当前用"固定前缀+随机数"模拟，测试通过。
        #    如果后续被拦截，考虑使用浏览器自动化(selenium/playwright)获取真实uuid。
        random_min = 10 ** (self.UUID_RANDOM_DIGITS - 1)
        random_max = 10 ** self.UUID_RANDOM_DIGITS - 1
        uuid_str = f"{self.UUID_PREFIX}-{random.randint(random_min, random_max)}"

        # 3. 提取URL路径（去掉域名部分）
        # JS中的处理: url.replace(/(http:|https:)?\/\/(.*?)\//, "/").replace(/\s+/g, "").replace(/\?.*/, "")
        # 即 "https://szgateway.jd.com/szpaas/..." → "/szpaas/..."
        from urllib.parse import urlparse
        parsed = urlparse(url)
        url_path = parsed.path  # 自动去掉域名和查询参数

        # 4. 拼接签名串并MD5加密
        sign_str = f"{url_path}{uuid_str}{timestamp}{self.SIGN_SALT}"
        user_mnp = hashlib.md5(sign_str.encode("utf-8")).hexdigest()

        return {
            "User-mup": str(timestamp),
            "User-mnp": user_mnp,
            "uuid": uuid_str,
        }

    # ==================== 请求间隔控制 ====================

    def _wait_interval(self):
        """
        确保两次请求之间间隔 ≥ 30秒
        如果距离上次请求不足30秒，自动等待补齐。
        """
        if self._last_request_time == 0:
            return  # 第一次请求，不需要等待

        elapsed = time.time() - self._last_request_time
        wait_time = self.REQUEST_INTERVAL - elapsed

        if wait_time > 0:
            self.logger.info(f"请求间隔控制：等待 {wait_time:.1f} 秒（距上次请求 {elapsed:.1f}秒，需≥{self.REQUEST_INTERVAL}秒）")
            time.sleep(wait_time)

    # ==================== 通用请求方法 ====================

    def _build_default_headers(self):
        """构建默认请求头（使用当前选中的UA）"""
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
        """切换UA（Edge↔Chrome），当风控拦截时自动调用"""
        old_name = "Edge" if self._current_ua_index == 0 else "Chrome"
        self._current_ua_index = 1 - self._current_ua_index  # 0↔1
        new_name = "Edge" if self._current_ua_index == 0 else "Chrome"
        # 更新Session的UA和sec-ch-ua
        ua, sec_ch_ua = self._ua_list[self._current_ua_index]
        self.session.headers["User-Agent"] = ua
        self.session.headers["sec-ch-ua"] = sec_ch_ua
        self.logger.info(f"UA切换: {old_name} → {new_name}")

    def request(self, url, data, method="POST", extra_headers=None):
        """
        通用请求方法（含30秒间隔控制、重试机制、UA切换）

        参数:
            url: 请求URL
            data: 请求体参数(dict)
            method: 请求方式，默认POST
            extra_headers: 额外的请求头(dict)，会合并到默认头之上
        返回:
            requests.Response 对象
        异常:
            所有重试失败后抛出最后一个异常
        """
        # 合并额外请求头
        headers = {}
        if extra_headers:
            headers.update(extra_headers)

        # 重试循环
        last_exception = None
        for attempt in range(1, self.MAX_RETRIES + 1):
            try:
                # 30秒间隔控制
                self._wait_interval()

                # 每次重试都重新生成风控参数（时间戳和uuid必须新鲜）
                risk_params = self._gen_risk_params(url)
                full_data = {**data, **risk_params}

                ua_name = "Edge" if self._current_ua_index == 0 else "Chrome"
                self.logger.info(f"发送请求 (第{attempt}/{self.MAX_RETRIES}次, UA={ua_name}): {url}")
                self.logger.debug(f"请求参数: {json.dumps(full_data, ensure_ascii=False)[:500]}")

                # 记录请求时间
                self._last_request_time = time.time()

                # 发送请求
                if method.upper() == "POST":
                    response = self.session.post(
                        url, data=full_data, headers=headers, timeout=self.REQUEST_TIMEOUT
                    )
                else:
                    response = self.session.get(
                        url, params=full_data, headers=headers, timeout=self.REQUEST_TIMEOUT
                    )

                # 检查是否被风控拦截
                content_type = response.headers.get("Content-Type", "")
                if "json" in content_type:
                    # JSON响应，检查是否有风控错误
                    try:
                        result = response.json()
                        if not result.get("success", True) and result.get("status", 0) < 0:
                            error_msg = result.get("message", "未知错误")
                            status_code = result.get("status")

                            # Cookie过期检测
                            if status_code in (302, -1) or "登录" in error_msg or "login" in error_msg.lower():
                                self.logger.error(f"Cookie可能已过期: {error_msg} (status={status_code})")
                                raise CookieExpiredError(f"Cookie已过期，请更新 config/cookie.txt")

                            # 风控拦截：切换UA后重试
                            self.logger.warning(f"风控拦截: {error_msg} (status={status_code})")
                            if attempt < self.MAX_RETRIES:
                                self._switch_ua()  # 切换UA
                                wait = self.REQUEST_INTERVAL * attempt
                                self.logger.info(f"等待 {wait}秒 后重试...")
                                time.sleep(wait)
                                continue

                    except json.JSONDecodeError:
                        pass

                # 请求成功
                self.logger.info(f"请求成功: HTTP {response.status_code}, {len(response.content)}字节, Content-Type: {content_type}")
                return response

            except CookieExpiredError:
                raise  # Cookie过期直接抛出，不重试

            except requests.exceptions.Timeout as e:
                last_exception = e
                self.logger.warning(f"第{attempt}次请求超时（{self.REQUEST_TIMEOUT}秒）")

            except Exception as e:
                last_exception = e
                self.logger.warning(f"第{attempt}次请求失败: {e}")

            # 重试前等待（非最后一次）
            if attempt < self.MAX_RETRIES:
                wait = self.REQUEST_INTERVAL * attempt
                self.logger.info(f"等待 {wait}秒 后重试...")
                time.sleep(wait)

        # 所有重试失败
        self.logger.error(f"所有 {self.MAX_RETRIES} 次重试均失败")
        raise last_exception

    # ==================== 日志记录 ====================

    def _init_logger(self):
        """初始化日志记录器"""
        logger = logging.getLogger(f"JD_{datetime.now().strftime('%Y%m%d')}")
        logger.setLevel(logging.DEBUG)

        # 避免重复添加handler
        if logger.handlers:
            return logger

        # 日志文件名（按日期）
        log_file = os.path.join(self.log_dir, f"jd_api_{datetime.now().strftime('%Y%m%d')}.log")

        # 文件handler（记录所有级别）
        file_handler = logging.FileHandler(log_file, encoding="utf-8")
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(
            logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
        )

        # 控制台handler（只记录INFO以上）
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        console_handler.setFormatter(
            logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S")
        )

        logger.addHandler(file_handler)
        logger.addHandler(console_handler)

        return logger

    # ==================== 保存文件 ====================

    def save_excel(self, response, filename):
        """
        将API返回的Excel二进制数据保存为文件

        参数:
            response: requests.Response 对象
            filename: 文件名（不含路径，如 "搜索流量_2026-08-03.xlsx"）
        返回:
            保存的文件完整路径
        """
        file_path = os.path.join(self.output_dir, filename)
        with open(file_path, "wb") as f:
            f.write(response.content)
        self.logger.info(f"Excel已保存: {file_path} ({len(response.content)}字节)")
        return file_path


# ==================== 自定义异常 ====================

class CookieExpiredError(Exception):
    """Cookie过期异常"""
    pass


class RiskControlError(Exception):
    """风控拦截异常"""
    pass


# ============================================================
# 本次改动内容总结（2026-08-04 第四次对话）
# ============================================================
#
# 【新增功能1】双UA切换机制
#   - 增加 UA_EDGE 和 UA_CHROME 两个常量
#   - 新增 _ua_list 列表和 _current_ua_index 状态
#   - 新增 _switch_ua() 方法（Edge ↔ Chrome 切换）
#   - _build_default_headers() 改为读取当前UA动态组装
#
# 【改动功能2】风控拦截自动UA切换
#   - request() 方法在检测到风控拦截时，调用 _switch_ua() 切换UA后重试
#   - 这样 Edge 被识别为爬虫时，自动换 Chrome 再试一次
#
# 【修复Bug】Timeout异常处理
#   - 原代码 `e = TimeoutError(...)` 会因 `e` 未在except块中定义而报错
#   - 改为 `last_exception = e`（except中已有e变量）
#
# 【改动逻辑（通俗版）】
#   - 你要求Edge和Chrome UA都写入，失败时切换
#   - 所以我加了"UA列表"和"切换按钮"
#   - 第一次请求用Edge（和你的抓包一致），如果被拦就自动切Chrome重试
#   - 就像你有两件衣服（Edge外套、Chrome外套），穿第一件被人认出来了就换第二件
#
# ============================================================
