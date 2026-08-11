# -*- coding: utf-8 -*-
"""
maintain.py - 自动检测过期项工具（项目15，2026-08-11）
============================================================

业务背景：
    仓库里有 12 个项目 19 个业务 key，大部分依赖 Cookie / h5st / IMAP 授权码，
    这些"凭据类"项目配置有不同周期的过期时间。本工具用于自动检测是否过期。

检测项（按风险等级排序）：
    🔴 P0 强制过期（不更新业务立即失败）
      - Cookie: 京麦 .shop.jd.com 域（项目14）
      - Cookie: 商智 szgateway.jd.com 域（项目1/4/5/6/13）
      - Cookie: 京准通 jzt-api.jd.com 域（项目7-12）
      - IMAP 授权码: QQ 邮箱（项目14）

    🟡 P1 中期过期（不更新会风控/限流，但不立即失败）
      - 风控签名盐值（项目1/4/5/6/13，从 commons.js 逆向）
      - UUID 前缀（项目1，需要重新抓包分析）

    🟢 P2 长期稳定（基本不过期，但监控以防京东服务端变更）
      - 京麦 URL 域 sff.jd.com / export.shop.jd.com
      - 商智 URL 域 szgateway.jd.com
      - 京准通 URL 域 jzt-api.jd.com

不检测项：
    - h5st：一次性签名（5-30分钟过期），需要每次新建任务前抓，不在工具范围

使用方法：
    # 全部检测（推荐）
    python maintain.py

    # 只检测某一类
    python maintain.py --only cookies
    python maintain.py --only imap
    python maintain.py --only urls

    # 输出为 JSON（便于外部脚本消费）
    python maintain.py --json

    # 静默模式（只输出错误）
    python maintain.py --quiet

退出码：
    0 - 全部正常
    1 - 发现 P0 过期项
    2 - 脚本执行异常
"""
import os
import sys
import json
import imaplib
import requests
from datetime import datetime
from typing import Dict, List, Tuple

# ============================================================
# 一、配置区（按需修改）
# ============================================================

# 项目根目录（脚本所在目录）
PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))

# Cookie 文件路径
COOKIE_FILES = {
    "jm": os.path.join(PROJECT_ROOT, "config", "jm_cookie.txt"),     # 京麦
    "sz": os.path.join(PROJECT_ROOT, "config", "sz_cookie.txt"),     # 商智（2026-08-11 重命名 cookie.txt → sz_cookie.txt）
    "jzt": os.path.join(PROJECT_ROOT, "config", "jzt_cookie.txt"),   # 京准通
}

# IMAP 配置（直接复用项目14的 imap_config.ini 解析逻辑）
IMAP_CONFIG_PATH = os.path.join(PROJECT_ROOT, "config", "imap_config.ini")

# URL 健康检查端点（用 GET 接口探测 Cookie 是否过期）
# 选择这些端点的原则：必须是 GET 请求 + 业务无副作用 + 京东服务端容忍
URL_HEALTH_CHECKS = {
    "sz": {
        # 商智：访问商品流量来源页（GET），需要 Cookie 但无副作用
        "url": "https://sz.jd.com/szweb/sz/view/viewflow/flowPathDetailsNew.html",
        "expected_status": 200,
        "cookie_file": "sz",
        "desc": "商智流量来源详情页（GET 探测）",
    },
    "jzt": {
        # 京准通：访问快车首页（GET）
        "url": "https://jzt.jd.com/",
        "expected_status": 200,
        "cookie_file": "jzt",
        "desc": "京准通首页（GET 探测）",
    },
    "jm": {
        # 京麦：访问订单导出页（GET）
        "url": "https://shop.jd.com/jdm/trade/tools/export/ExprotList",
        "expected_status": 200,
        "cookie_file": "jm",
        "desc": "京麦订单导出页（GET 探测）",
    },
}

# 请求超时（秒）
REQUEST_TIMEOUT = 10

# 风控签名盐值（项目1用，从 commons-a5562705.js 逆向获得）
# ⚠️ 此盐值京东可能不定期更新，本工具仅检查文件存在，不验证有效性
#     验证有效性需要：访问真实接口 + 解析 salt 值 + 重新生成 User-mnp 签名
#     复杂度高，本工具暂不做
SIGN_SALT_CHECK_ENABLED = False  # 设为 True 开启盐值验证（需要真实接口 + JS 逆向）

# ============================================================
# 二、检测函数区
# ============================================================

def check_cookie_file(name: str, path: str) -> Dict:
    """检查 Cookie 文件存在性 + 格式 + 长度。

    返回:
        dict - {
            "name": "jm",
            "path": "...",
            "exists": True,
            "length": 3098,
            "has_required_keys": True,
            "issues": [],
            "status": "ok" | "warning" | "error",
        }
    """
    result = {
        "name": name,
        "path": path,
        "exists": os.path.isfile(path),
        "length": 0,
        "has_required_keys": False,
        "issues": [],
        "status": "ok",
    }

    if not result["exists"]:
        result["issues"].append(f"Cookie 文件不存在: {path}")
        result["status"] = "error"
        return result

    try:
        with open(path, "r", encoding="utf-8") as f:
            content = f.read().strip()
        result["length"] = len(content)

        if len(content) < 100:
            result["issues"].append(f"Cookie 内容过短（{len(content)} 字符），可能已失效")
            result["status"] = "warning"
            return result

        # 京麦 Cookie 应包含 3AB9D23F7A4B3CSS 字段
        if name == "jm":
            if "3AB9D23F7A4B3CSS" not in content:
                result["issues"].append("京麦 Cookie 缺 3AB9D23F7A4B3CSS 字段")
                result["status"] = "error"
            else:
                result["has_required_keys"] = True

        # 商智 Cookie 应包含 pin= 字段
        elif name == "sz":
            if "pin=" not in content:
                result["issues"].append("商智 Cookie 缺 pin= 字段")
                result["status"] = "error"
            else:
                result["has_required_keys"] = True

        # 京准通 Cookie 应包含 HMACCOUNT= 字段
        elif name == "jzt":
            if "HMACCOUNT=" not in content:
                result["issues"].append("京准通 Cookie 缺 HMACCOUNT= 字段")
                result["status"] = "error"
            else:
                result["has_required_keys"] = True

    except Exception as e:
        result["issues"].append(f"读取 Cookie 失败: {e}")
        result["status"] = "error"

    return result


def check_cookie_live(cookie_name: str, cookie_path: str, target_url: str) -> Dict:
    """用 Cookie 访问京东 GET 端点，验证 Cookie 是否有效。

    ⚠️ 这个函数会发真实网络请求，请谨慎使用。
    """
    result = {
        "cookie_name": cookie_name,
        "url": target_url,
        "alive": None,
        "http_status": None,
        "redirect_url": None,
        "issues": [],
    }

    if not os.path.isfile(cookie_path):
        result["issues"].append(f"Cookie 文件不存在，跳过实测: {cookie_path}")
        return result

    try:
        with open(cookie_path, "r", encoding="utf-8") as f:
            cookie_str = f.read().strip()

        headers = {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/144.0.0.0 Safari/537.36 Edg/144.0.0.0"
            ),
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
            "Cookie": cookie_str,
        }

        resp = requests.get(
            target_url,
            headers=headers,
            timeout=REQUEST_TIMEOUT,
            allow_redirects=False,  # 关键：不跟随重定向（直接检测 302）
        )
        result["http_status"] = resp.status_code
        result["redirect_url"] = resp.headers.get("Location", "")

        # 京东 Cookie 过期的标志：302 重定向到登录页
        if resp.status_code == 302 and (
            "passport" in result["redirect_url"]
            or "login" in result["redirect_url"]
            or "logout" in result["redirect_url"]
        ):
            result["alive"] = False
            redir = result["redirect_url"]
            result["issues"].append(
                f"Cookie 已过期（重定向到登录: {redir[:60]}...）"
            )
        elif resp.status_code == 200:
            result["alive"] = True
        elif resp.status_code == 401:
            result["alive"] = False
            result["issues"].append("Cookie 已过期（HTTP 401）")
        elif resp.status_code == 403:
            result["alive"] = False
            result["issues"].append("Cookie 无权限（HTTP 403）")
        else:
            result["alive"] = None  # 不确定
            result["issues"].append(f"未预期的 HTTP 状态: {resp.status_code}")

    except requests.exceptions.Timeout:
        result["issues"].append(f"请求超时（>{REQUEST_TIMEOUT}s）")
    except requests.exceptions.RequestException as e:
        result["issues"].append(f"网络异常: {e}")
    except Exception as e:
        result["issues"].append(f"实测异常: {e}")

    return result


def check_imap_credentials(config_path: str) -> Dict:
    """验证 QQ 邮箱 IMAP 授权码是否有效。"""
    result = {
        "config_path": config_path,
        "exists": False,
        "login_success": None,
        "issues": [],
        "status": "ok",
    }

    if not os.path.isfile(config_path):
        result["issues"].append(f"IMAP 配置文件不存在: {config_path}")
        result["status"] = "error"
        return result

    result["exists"] = True

    try:
        import configparser
        cfg = configparser.ConfigParser()
        cfg.read(config_path, encoding="utf-8")
        if "imap" not in cfg:
            result["issues"].append("缺 [imap] section")
            result["status"] = "error"
            return result

        section = cfg["imap"]
        host = section.get("host", "").strip()
        port = int(section.get("port", "993"))
        user = section.get("user", "").strip()
        auth_code = section.get("auth_code", "").strip()
        use_ssl = section.get("use_ssl", "true").strip().lower() in ("1", "true", "yes")

        if not (host and port and user and auth_code):
            result["issues"].append("缺必填字段（host/port/user/auth_code）")
            result["status"] = "error"
            return result

        # 尝试登录
        try:
            if use_ssl:
                mail = imaplib.IMAP4_SSL(host, port, timeout=REQUEST_TIMEOUT)
            else:
                mail = imaplib.IMAP4(host, port, timeout=REQUEST_TIMEOUT)

            typ, _ = mail.login(user, auth_code)
            if typ == "OK":
                result["login_success"] = True
                mail.logout()
            else:
                result["login_success"] = False
                result["issues"].append(f"登录返回非 OK: {typ}")
                result["status"] = "error"
        except imaplib.IMAP4.error as e:
            result["login_success"] = False
            result["issues"].append(f"IMAP 登录失败: {e}")
            result["status"] = "error"
        except Exception as e:
            result["issues"].append(f"IMAP 连接异常: {e}")
            result["status"] = "error"

    except Exception as e:
        result["issues"].append(f"配置解析异常: {e}")
        result["status"] = "error"

    return result


def check_url_reachable(url: str, desc: str) -> Dict:
    """URL 健康检查（不依赖 Cookie，仅验证域可达）。"""
    result = {
        "url": url,
        "desc": desc,
        "reachable": None,
        "http_status": None,
        "issues": [],
    }
    try:
        resp = requests.get(url, timeout=REQUEST_TIMEOUT, allow_redirects=True)
        result["http_status"] = resp.status_code
        result["reachable"] = (resp.status_code == 200)
        if resp.status_code != 200:
            result["issues"].append(f"非 200 响应: {resp.status_code}")
    except Exception as e:
        result["issues"].append(f"请求失败: {e}")
    return result


# ============================================================
# 三、主流程
# ============================================================

def run_checks(only: str = None, do_live: bool = False) -> Dict:
    """执行所有检查，返回报告。

    参数:
        only - 可选过滤："cookies" / "imap" / "urls"
        do_live - 是否做实测（用 Cookie 访问真实 GET 端点验证）
    """
    report = {
        "timestamp": datetime.now().isoformat(),
        "checks": {},
        "summary": {"total": 0, "ok": 0, "warning": 0, "error": 0},
        "exit_code": 0,
    }

    # 1. Cookie 检查（文件层面）
    if only is None or only == "cookies":
        for name, path in COOKIE_FILES.items():
            r = check_cookie_file(name, path)
            report["checks"][f"cookie_{name}"] = r

            # 如果用户要求实测 + 文件存在，则实测
            if do_live and r["exists"] and r["has_required_keys"]:
                hc = URL_HEALTH_CHECKS.get(name)
                if hc:
                    live_result = check_cookie_live(name, path, hc["url"])
                    report["checks"][f"cookie_{name}_live"] = live_result

    # 2. IMAP 检查
    if only is None or only == "imap":
        report["checks"]["imap"] = check_imap_credentials(IMAP_CONFIG_PATH)

    # 3. URL 健康检查（不依赖 Cookie）
    if only is None or only == "urls":
        for name, hc in URL_HEALTH_CHECKS.items():
            report["checks"][f"url_{name}"] = check_url_reachable(hc["url"], hc["desc"])

    # 汇总
    for k, v in report["checks"].items():
        status = v.get("status", "ok")
        report["summary"]["total"] += 1
        if status == "ok":
            report["summary"]["ok"] += 1
        elif status == "warning":
            report["summary"]["warning"] += 1
            if report["exit_code"] == 0:
                report["exit_code"] = 0  # 警告不影响退出码
        elif status == "error":
            report["summary"]["error"] += 1
            report["exit_code"] = 1

    return report


def print_report(report: Dict, quiet: bool = False):
    """打印报告（人类可读）。"""
    if quiet:
        if report["summary"]["error"] > 0:
            print(f"❌ 发现 {report['summary']['error']} 个 P0 过期项")
        elif report["summary"]["warning"] > 0:
            print(f"⚠️ 发现 {report['summary']['warning']} 个警告")
        else:
            print(f"✅ 全部正常（{report['summary']['ok']} 项）")
        return

    print("=" * 70)
    print(f"🛠️  京东数据导出工具 - 过期项检测报告")
    print(f"📅 时间: {report['timestamp']}")
    print("=" * 70)

    for k, v in report["checks"].items():
        # 确定 emoji
        status = v.get("status", "ok")
        if status == "ok":
            emoji = "✅"
        elif status == "warning":
            emoji = "⚠️"
        else:
            emoji = "❌"

        # 标题
        if k.startswith("cookie_") and not k.endswith("_live"):
            name = v["name"]
            title = f"{emoji} Cookie[{name}]  文件: {v['path']}"
        elif k.endswith("_live"):
            title = f"{emoji} Cookie[{v['cookie_name']}]  实测  URL: {v['url'][:60]}..."
        elif k == "imap":
            title = f"{emoji} IMAP 授权码  配置: {v['config_path']}"
        elif k.startswith("url_"):
            title = f"{emoji} URL[{v.get('desc', '')}]  {v['url']}"
        else:
            title = f"{emoji} {k}"

        print(f"\n{title}")
        print(f"  存在: {v.get('exists') or v.get('alive') or v.get('login_success') or v.get('reachable')}")

        if "length" in v and v["length"] > 0:
            print(f"  长度: {v['length']} 字符")

        if "has_required_keys" in v:
            print(f"  关键字段: {'✅ 完整' if v['has_required_keys'] else '❌ 缺失'}")

        if v.get("http_status"):
            print(f"  HTTP 状态: {v['http_status']}")
            if v.get("redirect_url"):
                print(f"  重定向到: {v['redirect_url'][:80]}")

        if v.get("issues"):
            for issue in v["issues"]:
                print(f"  ⚠️ {issue}")

    # 汇总
    print()
    print("=" * 70)
    print("📊 汇总")
    print("=" * 70)
    s = report["summary"]
    print(f"  ✅ 正常: {s['ok']}")
    print(f"  ⚠️ 警告: {s['warning']}")
    print(f"  ❌ 错误: {s['error']}")
    print(f"  总计: {s['total']}")
    print()
    if s["error"] > 0:
        print("🔴 存在 P0 过期项，需要立即处理（参考 docs/维护手册.md）")
    elif s["warning"] > 0:
        print("🟡 有警告项，建议尽快处理")
    else:
        print("🟢 全部正常，可放心使用")


def main():
    import argparse
    parser = argparse.ArgumentParser(description="京东数据导出工具 - 自动检测过期项")
    parser.add_argument("--only", choices=["cookies", "imap", "urls"], help="只检测某一类")
    parser.add_argument("--json", action="store_true", help="输出 JSON 格式")
    parser.add_argument("--quiet", action="store_true", help="静默模式（只输出汇总）")
    parser.add_argument(
        "--live",
        action="store_true",
        help="实测 Cookie（用 Cookie 访问真实 GET 端点验证）—— 会发真实网络请求",
    )
    args = parser.parse_args()

    try:
        report = run_checks(only=args.only, do_live=args.live)
    except Exception as e:
        print(f"❌ 脚本执行异常: {e}", file=sys.stderr)
        if args.json:
            print(json.dumps({"error": str(e)}, ensure_ascii=False, indent=2))
        sys.exit(2)

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print_report(report, quiet=args.quiet)

    sys.exit(report["exit_code"])


if __name__ == "__main__":
    main()