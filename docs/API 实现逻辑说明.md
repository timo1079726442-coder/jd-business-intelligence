# API 实现逻辑说明

> 本文档详细说明每个API接口的实现逻辑，包括字段映射、动态参数来源、风控处理流程。
> **写代码前请先确认本文档无误，确认后再开始编码。**

---

## 项目代码结构（2026-08-04 第二次更新）

按全局铁律第4条要求，所有代码集中在 `main.py`：
- `class JDBaseRequest`（通用基类）
- `class ShopSourceAPI(JDBaseRequest)`（店铺来源，支持搜索/推荐两个业务）
- `def run_business(business_name, ...)`（业务分发器）
- `def main()`（主入口）

---

## 接口1：店铺来源-SKU维度（搜索流量 + 推荐流量）

### 一、接口基本信息

| 项目 | 内容 |
|------|------|
| 接口名称 | 店铺来源-SKU维度（搜索流量 + 推荐流量） |
| 请求方式 | POST |
| 接口URL | `https://szgateway.jd.com/szpaas/szajax/shop/source/offlineFlowSource/downSkuTable.ajax` |
| 域 | szgateway.jd.com |
| 返回格式 | Excel (.xlsx) |
| 单次返回上限 | 5000条SKU |
| 请求间隔 | ≥30秒（防风控） |
| 代码位置 | `main.py` → `class ShopSourceAPI` |

**支持业务**：
- `商品搜索效果`（lastSrcChannelId2=2008）→ `download_search_sku()`
- `商品推荐效果`（lastSrcChannelId2=2009）→ `download_recommend_sku()`

---

### 二、字段映射表

#### 2.1 Header字段

| 序号 | 字段名 | 值 | 类型 | 来源 | 说明 |
|------|--------|-----|------|------|------|
| 1 | User-Agent | Edge UA（默认）| 静态 | 配置 | 优先Edge，失败切换Chrome |
| 2 | Accept | 浏览器默认 | 静态 | 固定 | 模拟正常浏览器 |
| 3 | Accept-Language | zh-CN,zh;q=0.9,... | 静态 | 固定 | 中文优先 |
| 4 | Accept-Encoding | gzip, deflate, br, zstd | 静态 | 固定 | 支持压缩 |
| 5 | Content-Type | application/x-www-form-urlencoded | 静态 | 固定 | 表单提交 |
| 6 | Origin | https://sz.jd.com | 静态 | 固定 | |
| 7 | Referer | https://sz.jd.com/szweb/sz/view/viewflow/flowPathDetailsNew.html | 静态 | 固定 | |
| 8 | sec-ch-ua | Edge对应sec-ch-ua | 静态 | 配置 | 配合UA |
| 9 | sec-ch-ua-mobile | ?0 | 静态 | 固定 | |
| 10 | sec-ch-ua-platform | "Windows" | 静态 | 固定 | |
| 11 | Upgrade-Insecure-Requests | 1 | 静态 | 固定 | |
| 12 | Cookie | 完整Cookie字符串 | ⚠️ **动态** | config/cookie.txt | 见Cookie分析 |
| 13 | Content-Length | 自动计算 | ⚠️ 动态 | requests自动 | 根据请求体 |

#### 2.2 Cookie字段（风控相关）

| 字段名 | 类型 | 说明 |
|--------|------|------|
| pin | ⚠️ 动态 | 账号名（如FYA8888） |
| pinId | ⚠️ 动态 | 账号对应ID |
| __jdu / __jda / __jdb / __jdc | ⚠️ 动态 | 京东统计ID（含时间戳） |
| shshshfpa / shshshfpx / shshshfpb | ⚠️ **风控** | 京东风控设备指纹 |
| wlfstk_smdl / thor / light_key / flash | ⚠️ **风控** | 京东风控核心字段 |
| 3AB9D23F7A4B3C9B / 3AB9D23F7A4B3CSS | ⚠️ **风控** | 加密风控字段 |
| areaId / ipLoc-djd / PCSYCityID | 静态 | 地区（登录后稳定）|
| _base_ / unick / _tp | 静态 | 一次性会话标识 |
| _pst | ⚠️ 动态 | 等于pin |
| __USE_NEW_PAGEFRAME__ / __USE_NEW_PAGEFRAME_VERSION__ | 静态 | 页面框架标识 |
| 其他 | 静态 | 基本不变 |

**Cookie使用方式**：从 `config/cookie.txt` 整体读取，**不解析**，作为字符串整体塞入Cookie头。

#### 2.3 Body（表单参数）

| 字段名 | 类型 | 值 | 说明 |
|--------|------|-----|------|
| date | ⚠️ **动态** | 用户输入 | 如 2026-08-03 |
| startDate | ⚠️ **动态** | 用户输入 | 单日查询时=date |
| endDate | ⚠️ **动态** | 用户输入 | 单日查询时=date |
| interval | 静态 | DAY | |
| dateType | 静态 | day | |
| **lastSrcChannelId1** | 静态 | 2 | 一级渠道（搜索/推荐都是2）|
| **lastSrcChannelId2** | ⚠️ **业务关键** | **2008=搜索, 2009=推荐** | 二级渠道ID |
| groupType | 静态 | skuId | |
| attributes | 静态 | skuId | |
| sortField | 静态 | jdr_sch_traffic_enter_shop__browse_page_cnt_shop_last_src | 入店浏览量 |
| sortType | 静态 | desc | |
| limit | 静态 | 5000 | |
| compareType | 静态 | hb | 环比 |
| **User-mup** | ⚠️ **风控** | int(time.time()*1000) | 13位毫秒时间戳 |
| **User-mnp** | ⚠️ **风控** | MD5(URL路径+uuid+时间戳+盐值) | MD5签名 |
| **uuid** | ⚠️ **风控** | ca412182e5668a106054-{10位随机数} | 模拟前端SDK |

---

### 三、风控签名算法（User-mnp）

```
签名原文 = URL路径 + uuid + 时间戳 + 盐值
签名结果 = MD5(签名原文)
```

**示例**：
- URL路径：`/szpaas/szajax/shop/source/offlineFlowSource/downSkuTable.ajax`
- uuid：`ca412182e5668a106054-1234567890`
- 时间戳：`1785857258121`
- 盐值：`372ad2c2b6`
- 原文：`/szpaas/szajax/shop/source/offlineFlowSource/downSkuTable.ajaxca412182e5668a106054-12345678901785857258121372ad2c2b6`
- 签名：`MD5(...)` = `68bb84ea004652df1010bc4e5064c40e`

盐值放在 `config.xlsx`（项目"商品搜索效果"分组），京东更新后可手动修改。

---

### 四、业务调用

```python
# 商品搜索效果
api = ShopSourceAPI()
api.download_search_sku(date="2026-08-03")  # 生成 推荐流量_2026-08-03.xlsx

# 商品推荐效果
api = ShopSourceAPI()
api.download_recommend_sku(date="2026-08-03")  # 生成 推荐流量_2026-08-03.xlsx

# 通过业务分发器
from main import run_business
run_business("商品搜索效果", date="2026-08-03")
run_business("商品推荐效果", date="2026-08-03")
```

---

### 五、配置项（config.xlsx）

| 项目名 | 变量参数 | 默认值 | 说明 |
|--------|----------|--------|------|
| 全局 | cookie文件路径 | config/cookie.txt | |
| 全局 | 输出目录 | output/ | |
| 全局 | 日志目录 | logs/ | |
| 全局 | 请求间隔(秒) | 30 | |
| 全局 | 最大重试次数 | 3 | |
| 全局 | 请求超时(秒) | 30 | |
| 商品搜索效果 | 签名盐值 | 372ad2c2b6 | 京东更新时可改 |
| 商品搜索效果 | date | 2026-08-03 | ⚠️ 商品推荐效果复用此配置 |
| 商品搜索效果 | startDate | 2026-08-03 | |
| 商品搜索效果 | endDate | 2026-08-03 | |

**关键设计**：商品推荐效果复用商品搜索效果的日期配置（用户要求），不新增配置项。

---

### 六、踩坑经验

1. **lastSrcChannelId2 容易混淆**：
   - 2008 = 搜索子来源 → 商品搜索效果
   - 2009 = 推荐子来源 → 商品推荐效果
   - 仅这一个字段差异，但属于两个独立业务项目

2. **风控盐值会更新**：
   - 京东不定期更新 commons.js 中的盐值
   - 配置文件中的盐值要可手动改（已实现）

3. **30秒间隔必须遵守**：
   - 不遵守立即触发风控拦截
   - 间隔可在 config.xlsx 中调整

4. **UA切换是兜底方案**：
   - 默认Edge，遇到风控自动切Chrome
   - 不要禁用此功能

---

### 七、变更记录

| 日期 | 改动 | 作者 |
|------|------|------|
| 2026-08-04 | 初次实现商品搜索效果（2008） | 第一次对话 |
| 2026-08-04 | 重构：所有代码合并到 main.py | GitHub commit bc331c7 |
| 2026-08-04 | 新增商品推荐效果（2009），复用搜索的日期配置 | 本次更新 |