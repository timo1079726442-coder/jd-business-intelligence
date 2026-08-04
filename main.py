# -*- coding: utf-8 -*-
"""
main.py
京东商智数据导出 - 主程序入口
功能：调用店铺来源API，导出搜索流量SKU数据为Excel
"""
import sys
import os

# 添加项目根目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from jd_api.shop_source import ShopSourceAPI
from jd_api.base import CookieExpiredError


def main():
    """主程序入口"""
    print("=" * 60)
    print("京东商智 - 店铺来源数据导出工具")
    print("店铺: FYA8888")
    print("=" * 60)

    # 查询日期（可修改）
    date = "2026-08-03"

    print(f"\n即将导出: {date} 的搜索流量SKU数据")
    print("-" * 60)

    try:
        # 创建API实例（自动读取Cookie、初始化日志）
        api = ShopSourceAPI()

        # 调用搜索流量接口（自动处理风控签名+30秒间隔+重试+UA切换）
        file_path = api.download_search_sku(date=date)

        print(f"\n✅ 导出成功！")
        print(f"   文件: {file_path}")

    except CookieExpiredError as e:
        print(f"\n❌ Cookie已过期: {e}")
        print(f"   请重新获取Cookie，更新 config/cookie.txt 文件后重试。")

    except Exception as e:
        print(f"\n❌ 导出失败: {e}")
        print(f"   详细日志请查看 logs/ 目录下的日志文件。")


if __name__ == "__main__":
    main()


# ============================================================
# 本次改动内容总结（2026-08-04 第四次对话）
# ============================================================
#
# 【新建文件】主程序入口
#   - 用户的"启动按钮"
#   - 告诉程序：要导出哪天的数据（date="2026-08-03"）
#
# 【使用方法】
#   1. 直接运行：python main.py
#   2. 或修改 main() 中的 date 变量改成其他日期
#
# 【代码做了什么】
#   - 创建 ShopSourceAPI 实例（自动读Cookie/初始化日志）
#   - 调用 download_search_sku(date="2026-08-03")
#   - 自动处理：风控签名+30秒间隔+失败重试+UA切换
#   - 把结果保存到 output/搜索流量_2026-08-03.xlsx
#
# 【改动逻辑（通俗版）】
#   - base.py 是"工具箱"
#   - shop_source.py 是"工人"
#   - main.py 就是"按钮"，按一下工人就开始干活
#
# ============================================================
