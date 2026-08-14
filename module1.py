# -*- coding: utf-8 -*-
"""影刀 RPA 调用的 Python 入口（项目19，2026-08-13）

职责：
    1. 影刀通过命令行调用本脚本（也支持 Python 代码直接 import）
    2. 设置环境变量 AUTH_LOADER=1 + SHOP_ID=店铺名，启用多店铺鉴权
    3. 接收影刀传的 --shop / --biz_keys / --date / --h5st 等参数
    4. 调用 main.py 的 run_business() 跑业务
    5. **按退出码约定返回**（影刀根据退出码决定下一步）：
        0 = 全部成功
        1 = 业务失败（数据问题，不重抓）
        2 = 鉴权过期（影刀应重抓鉴权）
        3 = 系统错误（未预期异常）

使用示例（影刀命令行）：
    python module1.py --shop "FYA箱包旗舰店" --biz_keys "商品流量来源_搜索,京麦售后明细_完整一键导出" --date "2026-08-13" --h5st "manual_or_use_authloader"

    # 不传 --h5st（依赖 AuthLoader 从 h5st.json 自动读，需要影刀先写）
    python module1.py --shop "FYA箱包旗舰店" --biz_keys "A,B,C" --date "2026-08-13"

使用示例（Python 代码）：
    import sys
    from module1 import main as module1_main
    sys.exit(module1_main([
        "--shop", "FYA箱包旗舰店",
        "--biz_keys", "A,B,C",
        "--date", "2026-08-13"
    ]))
"""
import os
import sys
import argparse
import traceback

# 启用 AuthLoader（多店铺鉴权从 JSON 读，必须在 import main 之前）
# 影刀调用前已 set SHOP_ID，这里兜底设默认
if "AUTH_LOADER" not in os.environ:
    os.environ["AUTH_LOADER"] = "1"
if "AUTH_RPA_CLI" not in os.environ:
    os.environ["AUTH_RPA_CLI"] = "0"  # 影刀环境下不调 RPA（避免循环）

# 退出码常量（影刀根据这些值决定下一步）
EXIT_SUCCESS = 0        # 全部成功
EXIT_BIZ_FAIL = 1       # 业务失败（数据问题）
EXIT_AUTH_EXPIRED = 2   # 鉴权过期（Cookie/h5st）
EXIT_SYSTEM_ERROR = 3   # 系统错误


def parse_module_args(argv=None):
    """解析命令行参数

    参数:
        argv - 参数列表（默认 sys.argv[1:]）

    返回:
        argparse.Namespace
    """
    parser = argparse.ArgumentParser(
        prog="module1",
        description="影刀 RPA 调用的 Python 入口（项目19）",
    )
    parser.add_argument(
        "--shop",
        required=True,
        help="店铺 ID（必填，会覆盖 SHOP_ID 环境变量）",
    )
    parser.add_argument(
        "--biz_keys",
        required=True,
        help="业务 key 列表（逗号分隔），如 '商品流量来源_搜索,京麦售后明细_完整一键导出'",
    )
    parser.add_argument(
        "--date",
        default=None,
        help="查询日期 YYYY-MM-DD（与 --start_date/--end_date 互斥）",
    )
    parser.add_argument(
        "--start_date",
        default=None,
        help="开始日期 YYYY-MM-DD",
    )
    parser.add_argument(
        "--end_date",
        default=None,
        help="结束日期 YYYY-MM-DD",
    )
    parser.add_argument(
        "--range",
        default=None,
        choices=["last_1d", "last_3d", "last_7d", "last_15d", "last_30d"],
        help="快捷区间（与 --date/--start_date/--end_date 互斥）",
    )
    parser.add_argument(
        "--h5st",
        default=None,
        help="[兼容旧版] 单个 h5st 值，对所有京麦业务生效（不推荐，建议用 --jm_order_h5st/--jm_after_sale_h5st/--jzt_h5st 分别传）",
    )
    parser.add_argument(
        "--jm_order_h5st",
        default=None,
        help="京麦订单明细 h5st（项目14）。覆盖从 h5st_jm_order.json 自动读",
    )
    parser.add_argument(
        "--jm_after_sale_h5st",
        default=None,
        help="京麦售后明细 h5st（项目16）。覆盖从 h5st_jm_after_sale.json 自动读",
    )
    parser.add_argument(
        "--jzt_h5st",
        default=None,
        help="京准通 h5st（项目1）。覆盖从 h5st_jzt.json 自动读",
    )
    parser.add_argument(
        "--cookie_path",
        default=None,
        help="Cookie 文件路径（不传则用 AuthLoader 按 shop_id 推断）",
    )
    parser.add_argument(
        "--sms_password",
        default=None,
        help="京麦项目14 解压密码（手动传入，优先级高于 IMAP）",
    )
    parser.add_argument(
        "--imap_config_path",
        default=None,
        help="IMAP 配置文件路径（项目14 IMAP 监听）",
    )
    return parser.parse_args(argv)


def main(argv=None):
    """影刀调用的主入口

    参数:
        argv - 命令行参数列表

    返回:
        int - 退出码（0/1/2/3）
    """
    args = parse_module_args(argv)

    # 1. 强制设置 SHOP_ID（影刀可能忘了设）
    os.environ["SHOP_ID"] = args.shop
    # 注：环境变量已在 import 前设默认值（line 30-32），这里再次覆盖

    # 2. 打印启动信息（影刀可读 stdout 判断）
    print(f"[SHOP] {args.shop} / 业务: {args.biz_keys} / 日期: {args.date or args.start_date or args.range}")
    print(f"[AUTH_LOADER] enabled=1, SHOP_ID={args.shop}, h5st={'传' if args.h5st else '从JSON读'}")

    # 3. 准备传给 main.run_business() 的参数
    kwargs = {}
    if args.date:
        kwargs["date"] = args.date
    if args.start_date:
        kwargs["start_date"] = args.start_date
    if args.end_date:
        kwargs["end_date"] = args.end_date
    if args.range:
        kwargs["range"] = args.range
    if args.h5st:
        kwargs["h5st"] = args.h5st
    if args.jm_order_h5st:
        kwargs["jm_order_h5st"] = args.jm_order_h5st
    if args.jm_after_sale_h5st:
        kwargs["jm_after_sale_h5st"] = args.jm_after_sale_h5st
    if args.jzt_h5st:
        kwargs["jzt_h5st"] = args.jzt_h5st
    if args.cookie_path:
        kwargs["cookie_path"] = args.cookie_path
    if args.sms_password:
        kwargs["sms_password"] = args.sms_password
    if args.imap_config_path:
        kwargs["imap_config_path"] = args.imap_config_path

    # 4. 解析业务列表
    biz_keys = [k.strip() for k in args.biz_keys.split(",") if k.strip()]
    if not biz_keys:
        print("[ERR] --biz_keys 解析为空，至少需要 1 个业务 key")
        return EXIT_BIZ_FAIL

    # 5. ⚠️ 2026-08-14 升级：按 biz_key 自动选 h5st
    #   实证：不同业务页面的 h5st 不能跨业务复用
    #   如果用户没显式传 --xxx_h5st，从 h5st_xxx.json 自动读
    if any("京麦" in k and "售后" in k for k in biz_keys):
        # 含项目16 售后明细：需要 jm_after_sale h5st
        if not kwargs.get("jm_after_sale_h5st"):
            try:
                from auth_loader import AuthLoader
                _auth = AuthLoader(shop_id=args.shop)
                kwargs["jm_after_sale_h5st"] = _auth.get_h5st(check_expire=True, h5st_key="jm_after_sale")
                print(f"✅ 自动读取 jm_after_sale h5st")
            except Exception as e:
                print(f"[WARN] 自动读取 jm_after_sale h5st 失败：{e}")
    if any(k.startswith("京麦订单明细") for k in biz_keys):
        # 含项目14 订单明细：需要 jm_order h5st
        if not kwargs.get("jm_order_h5st"):
            try:
                from auth_loader import AuthLoader
                _auth = AuthLoader(shop_id=args.shop)
                kwargs["jm_order_h5st"] = _auth.get_h5st(check_expire=True, h5st_key="jm_order")
                print(f"✅ 自动读取 jm_order h5st")
            except Exception as e:
                print(f"[WARN] 自动读取 jm_order h5st 失败：{e}")
    if any(k == "京准通快车自定义报表" for k in biz_keys):
        # 含项目1 强校验京准通：需要 jzt h5st
        if not kwargs.get("jzt_h5st"):
            try:
                from auth_loader import AuthLoader
                _auth = AuthLoader(shop_id=args.shop)
                kwargs["jzt_h5st"] = _auth.get_h5st(check_expire=True, h5st_key="jzt")
                print(f"✅ 自动读取 jzt h5st")
            except Exception as e:
                print(f"[WARN] 自动读取 jzt h5st 失败：{e}")

    # 6. 调 main.run_business()（延迟 import，避免循环引用）
    try:
        from main import run_business
    except ImportError as e:
        print(f"[ERR] import main 失败：{e}")
        traceback.print_exc()
        return EXIT_SYSTEM_ERROR

    # 6. 跑业务
    try:
        results = run_business(biz_keys, **kwargs)
        # 判断成功/失败
        if isinstance(results, dict):
            # 批量：检查所有业务是否成功
            failed = [k for k, v in results.items() if not v]
            if failed:
                print(f"[PARTIAL_FAIL] 失败业务: {failed}")
                return EXIT_BIZ_FAIL
            else:
                print(f"[OK] 全部 {len(biz_keys)} 个业务成功")
                return EXIT_SUCCESS
        else:
            # 单个：返回文件路径即成功
            if results:
                print(f"[OK] 业务成功: {results}")
                return EXIT_SUCCESS
            else:
                print(f"[FAIL] 业务失败")
                return EXIT_BIZ_FAIL

    # 7. 鉴权过期（影刀必须重抓鉴权）
    except (CookieExpiredError, H5stExpiredError) as e:
        print(f"[AUTH_EXPIRED] {e}")
        return EXIT_AUTH_EXPIRED

    # 8. 业务失败（不重抓）
    except (ValueError, RuntimeError) as e:
        print(f"[FAIL] 业务失败: {e}")
        return EXIT_BIZ_FAIL

    # 9. 系统错误（兜底）
    except Exception as e:
        print(f"[SYSTEM_ERROR] 未预期异常: {e}")
        traceback.print_exc()
        return EXIT_SYSTEM_ERROR


# ============================ 异常 import（在 main 之外定义避免循环）============================
# 从 auth_loader 复用 CookieExpiredError / H5stExpiredError（项目18 定义）
try:
    from auth_loader import CookieExpiredError, H5stExpiredError
except ImportError:
    # 兜底：AuthLoader 不可用时定义临时异常
    class CookieExpiredError(Exception):
        pass
    class H5stExpiredError(Exception):
        pass


# ============================ CLI 入口 ============================
if __name__ == "__main__":
    sys.exit(main())
