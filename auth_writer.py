# -*- coding: utf-8 -*-
"""影刀 RPA 写鉴权 JSON 的辅助工具（项目19，2026-08-13）

职责：
    1. 影刀拟人登录后，调用本工具把 Cookie/h5st 写入正确路径
    2. 自动加 captured_at 毫秒时间戳（auth_loader 用它判断 30 分钟过期）
    3. 校验写入路径在 config/{shop_id}/ 下（防止误写）
    4. 提供 CLI 命令行接口（影刀「执行命令」动作直接调用）

设计原则：
    - 影刀**不能**直接用 Python 写 JSON（太复杂），所以用 CLI 包装
    - 影刀把抓到的值通过命令行参数传入，本脚本负责写文件 + 加时间戳
    - 写之前做"目录存在"检查，缺则创建

使用示例（影刀命令行）：
    # 写 Cookie
    python auth_writer.py cookie --shop "FYA箱包旗舰店" --biz jm --json '{"url":"...","cookies":[...]}'

    # 写 h5st（自动加 captured_at 毫秒时间戳）
    python auth_writer.py h5st --shop "FYA箱包旗舰店" --value "20260813...;abc123;..." --ua "Mozilla/5.0..." --biz_domain "shop.jd.com"

    # 写全部（3 域 Cookie + 1 个 h5st）
    python auth_writer.py all --shop "FYA箱包旗舰店" --jm_cookie_file C:\\tmp\\jm_raw.json --h5st "xxx" --ua "yyy"
"""
import os
import sys
import json
import time
import argparse

# 项目根目录
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_DIR = os.path.join(BASE_DIR, "config")

# biz_type 映射（与 auth_loader.py 保持一致）
BIZ_TYPE_MAP = {
    "sz": "sz_cookie",  # 商智
    "jzt": "jzt_cookie",  # 京准通
    "jm": "jm_cookie",  # 京麦
}

# h5st 子类型（按业务页面区分，2026-08-14 实测：不同页面 h5st 不能跨业务）
# 实证：售后页 h5st 跑订单明细 → 服务端 code=1001 未登录
H5ST_KEY_MAP = {
    "jm_order": "h5st_jm_order.json",       # 京麦订单明细（项目14）
    "jm_after_sale": "h5st_jm_after_sale.json",  # 京麦售后明细（项目16）
    "jzt": "h5st_jzt.json",                 # 京准通（项目1，项目8-12 不需要）
}


def validate_shop_id(shop_id: str) -> str:
    """校验店铺名（防路径注入）"""
    if not shop_id or ".." in shop_id or "/" in shop_id or "\\" in shop_id:
        raise ValueError(f"非法的 shop_id: {shop_id!r}（不能含 .. / \\）")
    return shop_id


def get_shop_dir(shop_id: str) -> str:
    """获取/创建店铺目录"""
    shop_id = validate_shop_id(shop_id)
    shop_dir = os.path.join(CONFIG_DIR, shop_id)
    os.makedirs(shop_dir, exist_ok=True)
    return shop_dir


def write_cookie(shop_id: str, biz_type: str, cookie_data) -> str:
    """写 Cookie JSON

    参数:
        shop_id - 店铺 ID
        biz_type - 业务类型: sz/jzt/jm
        cookie_data - 两种格式都支持：
            1) dict: {"url": "...", "cookies": [...]}
            2) str: 已经是 JSON 字符串

    返回:
        str - 写入的文件绝对路径
    """
    if biz_type not in BIZ_TYPE_MAP:
        raise ValueError(f"未知 biz_type={biz_type!r}，合法值: {list(BIZ_TYPE_MAP.keys())}")
    if isinstance(cookie_data, str):
        cookie_data = json.loads(cookie_data)
    if "cookies" not in cookie_data:
        raise ValueError(f"cookie_data 必须含 'cookies' 字段，实际 keys: {list(cookie_data.keys())}")

    shop_dir = get_shop_dir(shop_id)
    filename = f"{BIZ_TYPE_MAP[biz_type]}.json"
    file_path = os.path.join(shop_dir, filename)
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(cookie_data, f, ensure_ascii=False, indent=2)
    print(f"[OK] Cookie 已写入: {file_path}（{len(cookie_data['cookies'])} 个字段）")
    return file_path


def write_h5st(shop_id: str, h5st_value: str, ua: str = "", biz_domain: str = "", h5st_key: str = "jm_order") -> str:
    """写 h5st JSON（自动加 captured_at 毫秒时间戳）

    参数:
        shop_id - 店铺 ID
        h5st_value - h5st 字符串值
        ua - User-Agent（可选）
        biz_domain - 业务域（可选，如 shop.jd.com）
        h5st_key - h5st 子类型（必填）：
            "jm_order"      → 写入 h5st_jm_order.json（京麦订单明细）
            "jm_after_sale" → 写入 h5st_jm_after_sale.json（京麦售后明细）
            "jzt"           → 写入 h5st_jzt.json（京准通）

    返回:
        str - 写入的文件绝对路径

    关键（2026-08-14 实测）：
        不同业务页面的 h5st 不能跨业务复用（售后页 h5st 跑订单明细 → code=1001 未登录）
        必须从对应业务页面分别抓 h5st 单独存
    """
    if not h5st_value:
        raise ValueError("h5st_value 不能为空")
    if h5st_key not in H5ST_KEY_MAP:
        raise ValueError(f"未知 h5st_key={h5st_key!r}，合法值: {list(H5ST_KEY_MAP.keys())}")

    data = {
        "h5st": h5st_value,
        "captured_at": int(time.time() * 1000),  # ⚠️ 关键：13 位毫秒时间戳
        "ua": ua,
        "biz_domain": biz_domain,
        "h5st_key": h5st_key,  # ⚠️ 标记这是哪种 h5st
    }

    shop_dir = get_shop_dir(shop_id)
    filename = H5ST_KEY_MAP[h5st_key]
    file_path = os.path.join(shop_dir, filename)
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"[OK] h5st 已写入: {file_path}")
    print(f"     类型: {h5st_key}")
    print(f"     captured_at: {data['captured_at']}（{time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(data['captured_at']/1000))}）")
    return file_path


def write_from_raw_file(shop_id: str, raw_cookie_path: str, biz_type: str = None) -> str:
    """从影刀临时文件读 Cookie 原始数据，再写到 config/{shop_id}/

    使用场景：影刀「保存 Cookie」动作会先存到临时文件，影刀再调本脚本
    """
    # ⚠️ 兼容 Windows UTF-8 BOM（PowerShell echo 会写 BOM）
    with open(raw_cookie_path, "r", encoding="utf-8-sig") as f:
        raw_data = json.load(f)
    # 推断 biz_type（按文件名前缀，如 jm_cookie.json → jm）
    if not biz_type:
        basename = os.path.basename(raw_cookie_path).lower()
        for bt, prefix in BIZ_TYPE_MAP.items():
            if basename.startswith(prefix):
                biz_type = bt
                break
    if not biz_type:
        raise ValueError(
            f"无法从文件名 {os.path.basename(raw_cookie_path)!r} 推断 biz_type，"
            f"请显式传 --biz (sz/jzt/jm)"
        )
    return write_cookie(shop_id, biz_type, raw_data)


# ============================ h5st 提取（影刀监听结果）============================

# 不同 h5st_key 默认的 API 关键词（URL 包含此关键词即认为是目标请求）
DEFAULT_API_KEYWORDS = {
    "jm_order": ["createdExportTask"],          # 京麦订单明细
    "jm_after_sale": ["ExportDsmService.createExportTask"],  # 京麦售后明细
    "jzt": ["customreport/v2/report/report"],   # 京准通 add 接口
}


def extract_h5st_from_list(
    input_file: str,
    shop_id: str,
    h5st_key: str,
    api_keyword: str = None,
    ua: str = "",
) -> str:
    """从影刀监听结果文件提取 h5st，写入对应 JSON 文件

    输入文件格式（影刀"获取网页监听结果"输出的）：
        - JSON 字符串数组（每个元素是一个请求的字典）
        - 或 单个请求字典
        - 或 已转换的字符串（影刀 req_list 变量）

    提取逻辑：
        1. 解析输入文件为请求列表
        2. 找到 URL 含 api_keyword 的 POST 请求
        3. 从请求头 'h5st' 字段提取值
        4. 调 write_h5st 写入对应 JSON 文件

    参数:
        input_file  - 影刀监听结果 JSON 文件路径
        shop_id     - 店铺 ID
        h5st_key    - h5st 子类型
        api_keyword - URL 关键词（默认按 h5st_key 选）
        ua          - User-Agent（可选）

    返回:
        str - 写入的 h5st JSON 文件绝对路径
    """
    # ⚠️ 兼容 Windows UTF-8 BOM（PowerShell echo 会写 BOM）
    with open(input_file, "r", encoding="utf-8-sig") as f:
        content = f.read().strip()

    if not content:
        raise ValueError(f"输入文件 {input_file} 为空")

    # 尝试解析 JSON
    try:
        data = json.loads(content)
    except json.JSONDecodeError as e:
        raise ValueError(f"输入文件不是合法 JSON：{e}\n前 200 字符：{content[:200]!r}")

    # 统一为列表格式
    if isinstance(data, dict):
        req_list = [data]
    elif isinstance(data, list):
        req_list = data
    else:
        raise ValueError(f"输入 JSON 不是 dict 或 list，实际类型：{type(data).__name__}")

    # 选 API 关键词
    if not api_keyword:
        candidates = DEFAULT_API_KEYWORDS.get(h5st_key, ["createdExportTask"])
        api_keyword = candidates[0]

    print(f"[INFO] h5st_key={h5st_key}, api_keyword={api_keyword!r}")
    print(f"[INFO] 共 {len(req_list)} 个请求，搜索 URL 含 {api_keyword!r} 的 POST 请求")

    # 找目标请求（POST + URL 含 api_keyword）
    target_req = None
    for req in req_list:
        url = req.get("url", "") or req.get("requestUrl", "")
        method = (req.get("method", "POST") or "POST").upper()
        if method == "POST" and api_keyword in url:
            target_req = req
            break

    if not target_req:
        raise ValueError(
            f"未找到 URL 含 {api_keyword!r} 的 POST 请求。"
            f"共 {len(req_list)} 个请求，请检查影刀监听结果是否正确保存。"
        )

    print(f"[OK] 找到目标请求：URL={target_req.get('url', '')[:120]}...")

    # 从请求头提取 h5st（兼容多种头字段命名）
    headers = target_req.get("headers", {}) or target_req.get("requestHeaders", {})
    if isinstance(headers, str):
        # 字符串格式："h5st: value\r\n..." → 解析
        import re
        m = re.search(r"h5st\s*[:=]\s*([^\r\n;]+)", headers)
        h5st_value = m.group(1).strip() if m else ""
    elif isinstance(headers, dict):
        # dict 格式（影刀 CDP 输出常见）
        h5st_value = (
            headers.get("h5st")
            or headers.get("H5st")
            or headers.get("H5ST")
            or ""
        )
        if isinstance(h5st_value, list):
            h5st_value = h5st_value[0] if h5st_value else ""
        h5st_value = str(h5st_value).strip()
    else:
        h5st_value = ""

    if not h5st_value:
        raise ValueError(
            f"目标请求未找到 h5st 头。请求 URL={target_req.get('url', '')[:120]}\n"
            f"headers 字段 keys: {list(headers.keys()) if isinstance(headers, dict) else type(headers).__name__}"
        )

    # 推断 biz_domain
    biz_domain = ""
    url = target_req.get("url", "")
    if "shop.jd.com" in url:
        biz_domain = "shop.jd.com"
    elif "jzt.jd.com" in url:
        biz_domain = "jzt.jd.com"
    elif "sz.jd.com" in url:
        biz_domain = "sz.jd.com"

    print(f"[OK] h5st 提取成功：长度 {len(h5st_value)} 字符，biz_domain={biz_domain}")

    # 写文件（复用 write_h5st）
    return write_h5st(shop_id, h5st_value, ua=ua, biz_domain=biz_domain, h5st_key=h5st_key)


# ============================ CLI 入口 ============================

def main(argv=None):
    parser = argparse.ArgumentParser(
        prog="auth_writer",
        description="影刀 RPA 写鉴权 JSON 辅助工具（项目19）",
    )
    subparsers = parser.add_subparsers(dest="command", help="子命令")

    # 子命令 1: cookie
    cookie_parser = subparsers.add_parser("cookie", help="写单个 Cookie JSON")
    cookie_parser.add_argument("--shop", required=True, help="店铺 ID")
    cookie_parser.add_argument("--biz", required=True, choices=["sz", "jzt", "jm"], help="业务类型")
    cookie_parser.add_argument("--json", required=True, help="Cookie JSON 字符串（DevTools 导出格式）")

    # 子命令 2: h5st
    h5st_parser = subparsers.add_parser("h5st", help="写 h5st JSON（自动加 captured_at）")
    h5st_parser.add_argument("--shop", required=True, help="店铺 ID")
    h5st_parser.add_argument("--value", required=True, help="h5st 字符串值")
    h5st_parser.add_argument("--ua", default="", help="User-Agent（可选）")
    h5st_parser.add_argument("--biz_domain", default="", help="业务域（可选）")
    h5st_parser.add_argument(
        "--h5st_key",
        required=True,
        choices=["jm_order", "jm_after_sale", "jzt"],
        help="h5st 子类型：jm_order=订单明细 / jm_after_sale=售后明细 / jzt=京准通",
    )

    # 子命令 3: raw（从临时文件读）
    raw_parser = subparsers.add_parser("raw", help="从影刀临时文件读 Cookie 再写")
    raw_parser.add_argument("--shop", required=True, help="店铺 ID")
    raw_parser.add_argument("--raw_file", required=True, help="影刀临时 Cookie 文件路径")
    raw_parser.add_argument("--biz", default=None, choices=["sz", "jzt", "jm"], help="业务类型（可选，不传则从文件名推断）")

    # 子命令 4: h5st_extract（从影刀监听结果 JSON 提取 h5st）
    extract_parser = subparsers.add_parser(
        "h5st_extract",
        help="从影刀监听到的请求列表中提取 h5st（推荐方案 B）",
    )
    extract_parser.add_argument("--input_file", required=True, help="影刀监听结果保存的 JSON 文件路径")
    extract_parser.add_argument("--shop", required=True, help="店铺 ID")
    extract_parser.add_argument(
        "--h5st_key",
        required=True,
        choices=["jm_order", "jm_after_sale", "jzt"],
        help="h5st 子类型（决定写哪个文件）",
    )
    extract_parser.add_argument(
        "--api_keyword",
        default=None,
        help="用于匹配请求 URL 的关键词（如 'createdExportTask' / 'add' / 'ExportDsmService'），默认按 h5st_key 自动选",
    )
    extract_parser.add_argument(
        "--ua",
        default="",
        help="User-Agent（可选，影刀提供则记录到 h5st.json）",
    )

    args = parser.parse_args(argv)

    if not args.command:
        parser.print_help()
        return 1

    try:
        if args.command == "cookie":
            write_cookie(args.shop, args.biz, args.json)
        elif args.command == "h5st":
            write_h5st(args.shop, args.value, args.ua, args.biz_domain, args.h5st_key)
        elif args.command == "raw":
            write_from_raw_file(args.shop, args.raw_file, args.biz)
        elif args.command == "h5st_extract":
            extract_h5st_from_list(
                args.input_file,
                args.shop,
                args.h5st_key,
                args.api_keyword,
                args.ua,
            )
        return 0
    except Exception as e:
        print(f"[ERR] {e}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
