# -*- coding: utf-8 -*-
"""
jd_cdp_capture.py
============================================================
京麦订单明细【加密】导出的【前置抓包脚本】。

业务背景：
    京麦 seller-v10.shop.jd.com 的订单导出页面嵌套 iframe，
    必须用 Playwright 启动真实 Chrome + CDP 全局监听 Network。
    脚本【只挂监听，绝不点击/操作页面】，所有页面交互由人完成。

使用流程：
    1. 运行本脚本：python jd_cdp_capture.py
    2. 浏览器自动弹出，已登录的 seller-v10.shop.jd.com 会话保持
    3. 人工在浏览器里：登录 → 打开订单导出页 → 新建导出任务 → 申请密码
    4. 回到终端按回车，脚本停止监听并保存 cdp_network_log.json
    5. 30 分钟无操作自动停止（防止忘按回车一直挂着）

⚠️ 本脚本不参与任何业务请求构造，纯粹是网络监听器。
"""

# ============================================================
# 一、import 导入区（每个import附带中文注释）
# ============================================================
import asyncio                                  # 异步事件循环，playwright 需要
import json                                     # 抓包结果最终落盘成 JSON
import os                                       # 操作系统能力：路径、目录、文件存在
import sys                                      # sys.exit 用于异常退出
import time                                     # 时间戳、监听超时计算
from pathlib import Path                        # 路径对象，跨平台友好
from datetime import datetime                   # 格式化的时间戳

from playwright.async_api import async_playwright  # 浏览器自动化引擎，提供 CDP 接入


# ============================================================
# 二、【业务可修改配置区】
#    所有可调参数集中在此处，业务逻辑代码不要硬编码
# ============================================================

# ---- 浏览器可执行文件探测顺序（按顺序找，找到就用） ----
# 2026-08-06 调整：京麦订单导出页 Chrome 触发 601 操作频繁；改用 Edge 测试
CHROME_CANDIDATE_PATHS = [
    # Microsoft Edge（优先 - 本次测试用）
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    # Google Chrome（兜底）
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
]

# ---- 独立用户配置目录（保存登录态，避免污染默认Chrome） ----
USER_DATA_DIR = Path(__file__).parent / "cdp_user_data"  # 项目根目录下

# ---- CDP 远程调试端口（Chrome DevTools 默认端口） ----
CDP_PORT = 9222

# ---- 抓包过滤关键词（命中请求会被打标记，不丢弃） ----
# 设计要点：首次抓包采【全量抓包+标记】，不丢弃任何请求，
# 避免漏掉签名/鉴权类未知接口；后续解析日志时再按业务筛选。
FILTER_KEYWORDS = [
    "exportCenterService",  # 京麦订单导出模块（5个目标接口全在这里）
]

# ---- 抓包输出文件路径 ----
CDP_LOG_FILE = Path(__file__).parent / "cdp_network_log.json"  # 项目根目录

# ---- 监听超时（秒）：超时未按回车自动停止抓包 ----
# 防止人工操作后忘记回车，脚本一直挂着
LISTEN_TIMEOUT_SECONDS = 30 * 60  # 30 分钟

# ---- 响应体大小上限（字节）：超过此大小只保存 headers，不保存 body ----
# 防止 zip 文件、报表文件把日志撑爆
RESPONSE_BODY_MAX_BYTES = 10 * 1024 * 1024  # 10 MB


# ============================================================
# 三、底层固定逻辑区
# ============================================================

def find_chrome_exe():
    """按候选顺序探测本地 Chrome / Edge，返回第一个存在的路径。

    返回：
        str: Chrome 可执行文件绝对路径
    抛出：
        FileNotFoundError: 全部候选都不存在时
    """
    for path in CHROME_CANDIDATE_PATHS:
        if path and os.path.exists(path):
            return path
    raise FileNotFoundError(
        "未找到本地 Chrome 或 Edge，请安装 Google Chrome 后再运行。\n"
        "已探测路径：\n  " + "\n  ".join(CHROME_CANDIDATE_PATHS)
    )


def is_url_match_keywords(url):
    """判断 URL 是否命中过滤关键词集合。

    入参：
        url: 请求 URL 字符串
    返回：
        bool: 命中任一关键词返回 True
    """
    if not url:
        return False
    return any(keyword in url for keyword in FILTER_KEYWORDS)


class CDPCapture:
    """CDP 全局抓包器。

    职责：
        1. 启动 Chrome（带独立用户配置）
        2. 建立 CDP Session，挂全局 Network 监听
        3. 收集请求/响应数据
        4. 等待回车或超时后落盘 JSON
    """

    def __init__(self):
        # 收集到的所有请求/响应记录
        self.records = []
        # 记录抓包开始时间
        self.start_time = None
        # Chrome 进程对象
        self.chrome_process = None
        # playwright 实例
        self.playwright = None
        # 浏览器上下文
        self.browser = None

    @staticmethod
    def _truncate_body(body_text, max_bytes=RESPONSE_BODY_MAX_BYTES):
        """响应体过大时截断处理。

        入参：
            body_text: 原始响应体字符串
            max_bytes: 最大字节数
        返回：
            (str, bool): (处理后的内容, 是否被截断)
        """
        if body_text is None:
            return None, False
        encoded = body_text.encode("utf-8", errors="replace")
        if len(encoded) <= max_bytes:
            return body_text, False
        # 超过阈值：截断到 max_bytes 并标记
        truncated = encoded[:max_bytes].decode("utf-8", errors="replace")
        return truncated + "\n...[TRUNCATED，超过10MB已丢弃剩余内容]...", True

    async def _on_request(self, request):
        """Network.requestWillBeSent 回调：记录请求元数据。

        入参：
            request: playwright Request 对象
        """
        url = request.url
        # 标记是否命中过滤关键词（不丢弃，全量保存）
        matched = is_url_match_keywords(url)

        record = {
            "type": "request",
            "timestamp": datetime.now().isoformat(),
            "method": request.method,
            "url": url,
            "headers": dict(request.headers),
            "post_data": request.post_data,  # POST body 字符串，可能为 None
            "matched_keywords": matched,     # 标记：是否命中 exportCenterService
        }
        self.records.append(record)

    async def _on_response(self, response):
        """Network.responseReceived 回调：记录响应状态与 body。

        入参：
            response: playwright Response 对象
        """
        url = response.url
        matched = is_url_match_keywords(url)

        # 尝试读取响应 body（受 10MB 阈值保护）
        body_text = None
        body_truncated = False
        try:
            body_bytes = await response.body()
            if body_bytes:
                # 先用 utf-8 解码，失败回退 latin-1
                try:
                    body_text = body_bytes.decode("utf-8")
                except UnicodeDecodeError:
                    body_text = body_bytes.decode("latin-1", errors="replace")
                body_text, body_truncated = self._truncate_body(body_text)
        except Exception as e:
            # 某些响应（如 event-stream、跨域）可能读不到 body，不致命
            body_text = f"<读取失败: {type(e).__name__}: {e}>"

        record = {
            "type": "response",
            "timestamp": datetime.now().isoformat(),
            "status": response.status,
            "url": url,
            "headers": dict(response.headers),
            "body": body_text,
            "body_truncated": body_truncated,
            "matched_keywords": matched,
        }
        self.records.append(record)

    def _print_summary(self):
        """控制台打印本次抓包统计摘要。"""
        total = len(self.records)
        req_count = sum(1 for r in self.records if r["type"] == "request")
        resp_count = sum(1 for r in self.records if r["type"] == "response")
        matched_count = sum(
            1 for r in self.records if r.get("matched_keywords")
        )
        elapsed = time.time() - self.start_time if self.start_time else 0

        print("\n" + "=" * 60)
        print("抓包监听结束，统计摘要：")
        print(f"  监听时长：{elapsed:.1f} 秒")
        print(f"  请求条数：{req_count}")
        print(f"  响应条数：{resp_count}")
        print(f"  命中关键词条数：{matched_count}（带标记未丢弃）")
        print(f"  输出文件：{CDP_LOG_FILE}")
        print(f"  文件大小：{CDP_LOG_FILE.stat().st_size / 1024:.1f} KB"
              if CDP_LOG_FILE.exists() else "  文件状态：未生成")
        print("=" * 60)

    def _save_log(self):
        """把所有记录按 JSON 格式落盘。"""
        output = {
            "capture_meta": {
                "start_time": datetime.fromtimestamp(self.start_time).isoformat()
                if self.start_time else None,
                "end_time": datetime.now().isoformat(),
                "filter_keywords": FILTER_KEYWORDS,
                "response_body_max_bytes": RESPONSE_BODY_MAX_BYTES,
                "capture_strategy": "全量抓包 + 命中标记（不丢弃）",
                "total_records": len(self.records),
            },
            "records": self.records,
        }
        with open(CDP_LOG_FILE, "w", encoding="utf-8") as f:
            json.dump(output, f, ensure_ascii=False, indent=2)

    async def run(self):
        """主流程：启动 Chrome → 监听 → 等待回车/超时 → 落盘。"""
        # 1. 准备用户配置目录
        USER_DATA_DIR.mkdir(parents=True, exist_ok=True)

        # 2. 探测 Chrome 路径
        chrome_exe = find_chrome_exe()
        browser_name = "Edge" if "msedge" in chrome_exe.lower() else "Chrome"
        print(f"[INFO] 使用浏览器：{browser_name} ({chrome_exe})")
        print(f"[INFO] 首次提示：本轮默认使用 Edge 测试京麦订单导出")

        # 3. 启动 playwright，连接系统 Chrome（不是内置 chromium）
        # ⚠️ playwright 限制：使用 user_data_dir 时必须用 launch_persistent_context()
        # 不能在 args 里加 --user-data-dir，否则会报错
        self.playwright = await async_playwright().start()
        context = await self.playwright.chromium.launch_persistent_context(
            user_data_dir=str(USER_DATA_DIR),     # 独立用户配置目录（保存登录态）
            executable_path=chrome_exe,           # 关键：用系统Chrome
            headless=False,                      # 关键：有头模式，可视化窗口
            args=[
                f"--remote-debugging-port={CDP_PORT}",  # 开启 CDP 远程调试
                "--no-first-run",                       # 跳过首次运行向导
                "--no-default-browser-check",           # 不询问是否设为默认浏览器
            ],
        )
        # launch_persistent_context 直接返回 context，需要额外创建 page
        page = await context.new_page()

        # 5. 挂全局 Network 监听
        page.on("request", self._on_request)
        page.on("response", self._on_response)

        # 6. 打开京麦订单导出页面（用户可以在此基础上继续手动操作）
        # 第一次进入触发加载，可观察到加载过程的请求
        # 2026-08-06 修正：之前抓包发现真实入口是 shop.jd.com，不是 seller-v10.shop.jd.com
        target_url = "https://shop.jd.com/jdm/trade/tools/export/ExprotList"
        print(f"[INFO] 打开目标页面：{target_url}")
        try:
            await page.goto(target_url, timeout=60000)
        except Exception as e:
            print(f"[WARN] 打开页面失败（不影响抓包）：{e}")
            print("[WARN] 请在浏览器中手动打开京麦订单导出页面")

        # 7. 抓包开始
        self.start_time = time.time()
        print("\n" + "=" * 60)
        print("【CDP 抓包已启动】")
        print("  现在请你手动操作 Chrome 浏览器页面完成抓包。")
        print(f"  监听超时：{LISTEN_TIMEOUT_SECONDS // 60} 分钟（超时自动停止）")
        print("  完成后请回到本控制台，按下回车键停止抓包。")
        print("=" * 60 + "\n")

        # 8. 等待回车 或 超时（后台任务监听超时）
        loop = asyncio.get_event_loop()

        def on_enter():
            """用户按回车时调用，停止 event loop。"""
            print("\n[INFO] 检测到回车，准备停止抓包...")
            loop.stop()

        # 使用 run_in_executor 在后台线程监听 stdin
        # 主线程跑 asyncio.wait_for 等待超时
        try:
            await asyncio.wait_for(
                loop.run_in_executor(None, input, "按回车停止抓包 > "),
                timeout=LISTEN_TIMEOUT_SECONDS,
            )
        except asyncio.TimeoutError:
            print(f"\n[INFO] 超过 {LISTEN_TIMEOUT_SECONDS // 60} 分钟未按回车，自动停止抓包")
        except (KeyboardInterrupt, EOFError):
            print("\n[INFO] 检测到中断信号，停止抓包")
        finally:
            # 9. 落盘 + 关闭浏览器
            self._save_log()
            self._print_summary()
            try:
                # launch_persistent_context 返回 context，关闭 context 即可
                await context.close()
            except Exception:
                pass
            try:
                await self.playwright.stop()
            except Exception:
                pass


def main():
    """入口函数。"""
    print("=" * 60)
    print("京麦订单明细【加密】导出 - CDP 抓包脚本")
    print("=" * 60)
    print("\n本脚本【只做网络监听】，不点击、不操作页面。")
    print("页面交互 100% 由你手动完成。\n")

    capture = CDPCapture()
    try:
        asyncio.run(capture.run())
    except FileNotFoundError as e:
        print(f"\n[ERROR] {e}")
        sys.exit(1)
    except Exception as e:
        print(f"\n[ERROR] 抓包异常：{type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
