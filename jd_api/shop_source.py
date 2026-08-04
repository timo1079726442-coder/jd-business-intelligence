# -*- coding: utf-8 -*-
"""
jd_api/shop_source.py
京东商智 - 店铺来源数据API
功能：获取店铺流量来源中各渠道的SKU维度数据

接口的固定参数（API_URL/渠道ID/排序字段等）作为代码常量
日期参数(date/startDate/endDate)从config.xlsx读取，方便换日期
"""
from .base import JDBaseRequest


class ShopSourceAPI(JDBaseRequest):
    """店铺来源数据API"""

    # ============ 接口固定参数（代码常量） ============
    # 接口URL
    API_URL = "https://szgateway.jd.com/szpaas/szajax/shop/source/offlineFlowSource/downSkuTable.ajax"

    # 时间粒度
    INTERVAL = "DAY"
    DATETYPE = "day"

    # 渠道ID
    LAST_SRC_CHANNEL_ID1 = "2"  # 一级渠道：搜索
    LAST_SRC_CHANNEL_ID2 = "2008"  # 二级渠道：搜索子来源

    # 聚合维度
    GROUP_TYPE = "skuId"
    ATTRIBUTES = "skuId"

    # 排序字段（按入店浏览量降序）
    SORT_FIELD = "jdr_sch_traffic_enter_shop__browse_page_cnt_shop_last_src"
    SORT_TYPE = "desc"

    # 返回上限
    LIMIT = "5000"

    # 对比类型
    COMPARE_TYPE = "hb"

    def download_sku(self, date=None, start_date=None, end_date=None, channel="搜索"):
        """
        下载店铺来源-SKU维度数据

        参数:
            date      (str): 查询日期，格式 YYYY-MM-DD（如 "2026-08-03"）
                              默认为None，从config.xlsx读取
            start_date(str): 开始日期，默认从config读取。也可传入
            end_date  (str): 结束日期，默认从config读取。也可传入
            channel   (str): 渠道名称，默认"搜索"
        返回:
            str: 保存的Excel文件路径
        """
        # 日期参数：优先用传入的参数，否则从配置读取
        if date is None:
            date = self.config.get("date", "")
        if start_date is None:
            start_date = self.config.get("startDate", date)
        if end_date is None:
            end_date = self.config.get("endDate", date)

        # 组装业务参数
        data = {
            "date": date,
            "startDate": start_date,
            "endDate": end_date,
            "interval": self.INTERVAL,
            "dateType": self.DATETYPE,
            "lastSrcChannelId1": self.LAST_SRC_CHANNEL_ID1,
            "lastSrcChannelId2": self.LAST_SRC_CHANNEL_ID2,
            "groupType": self.GROUP_TYPE,
            "attributes": self.ATTRIBUTES,
            "sortField": self.SORT_FIELD,
            "sortType": self.SORT_TYPE,
            "limit": self.LIMIT,
            "compareType": self.COMPARE_TYPE,
        }

        self.logger.info(f"下载店铺来源数据: 日期={date}, 渠道={channel}")

        # 调用通用请求方法
        response = self.request(self.API_URL, data)

        # 保存Excel
        filename = f"{channel}流量_{date}.xlsx"
        file_path = self.save_excel(response, filename)

        return file_path

    def download_search_sku(self, date=None, start_date=None, end_date=None):
        """
        便捷方法：下载搜索流量-SKU维度数据

        参数:
            date      (str): 查询日期，None则从配置读取
            start_date(str): 开始日期，None则从配置读取
            end_date  (str): 结束日期，None则从配置读取
        返回:
            str: 保存的Excel文件路径
        """
        return self.download_sku(
            date=date,
            start_date=start_date,
            end_date=end_date,
            channel="搜索",
        )


# ============================================================
# 本次改动内容总结（2026-08-04 第八次对话）
# ============================================================
#
# 【改动内容】接口固定参数改回代码常量，日期从配置读取
#   - 删除config.xlsx中的搜索流量SKU接口业务参数（API_URL/interval/channelId等）
#   - 新增日期配置（date/startDate/endDate）到config.xlsx
#   - shop_source.py的API_URL等固定参数改回类常量
#
# 【改动逻辑】
#   - 你要求"日期变量放配置，接口固定参数不放"
#   - 接口业务参数每次调用都一样（interval=DAY, channelId=2等），属于固定逻辑
#   - 日期是用户每次可能改的（查不同天的数据），属于易变参数
#
# ============================================================