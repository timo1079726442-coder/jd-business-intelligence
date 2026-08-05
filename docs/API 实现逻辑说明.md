# API 实现逻辑说明

> 本文档详细说明每个API接口的实现逻辑，包括字段映射、动态参数来源、风控处理流程。
> **写代码前请先确认本文档无误，确认后再开始编码。**

---

## 项目代码结构（2026-08-05 更新）

按全局铁律第4条要求，所有代码集中在 `main.py`：
- `class JDBaseRequest`（通用基类：Cookie/风控签名/30秒间隔/重试/UA切换/日志/Excel保存）
- `class ProductFlowAPI(JDBaseRequest)`（商品流量来源：搜索2008/推荐2009/购物车3001）
- 公共工具函数（所有报表复用，禁止硬编码业务逻辑）：
  - `convert_date_format(date_str)` —— 通用日期格式转换（8位纯数字/横杠/斜杠 → yyyy/m/d，时间保留）
  - `safe_convert_numeric(df)` —— 全表数值安全转换（订单编号强制文本、SKU/SPU转数字、>15位保留文本）
  - `apply_column_formats(...)` —— 按列名批量设置Excel单元格格式（订单编号@/SKU·SPU格式0/日期格式）
- `BUSINESS_REGISTRY`（业务注册中心）+ `run_business(biz_key_or_list, **kwargs)`（统一调度，支持批量）
- `config_consistency_check()`（启动自动跑，输出✅/❌/⚠️配置一致性核对报告）
- `main()`（主入口，命令行：`--biz_key / --date / --start_date / --end_date / --list`）

---

## 接口1：店铺来源-SKU维度（搜索流量 / 推荐流量 / 购物车流量）

### 一、接口基本信息

| 项目 | 内容 |
|------|------|
| 接口名称 | 店铺来源-SKU维度（offlineFlowSource） |
| 请求方式 | POST |
| 接口URL | `https://szgateway.jd.com/szpaas/szajax/shop/source/offlineFlowSource/downSkuTable.ajax` |
| 域 | szgateway.jd.com |
| 返回格式 | Excel (.xlsx) |
| 单次返回上限 | 5000条SKU（limit） |
| 请求间隔 | ≥30秒（防风控，config可调） |
| 代码位置 | `main.py` → `class ProductFlowAPI` |

**支持业务（CHANNEL_MAP）**：

| 业务key | lastSrcChannelId2 | uuid前缀 | 状态 |
|---------|-------------------|----------|------|
| 商品流量来源_搜索 | 2008 | ca412182e5668a106054 | 执行 |
| 商品流量来源_推荐 | 2009 | ca412182e5668a106054 | 执行 |
| 商品流量来源_购物车 | 3001 | 5f9cc2ca20cad3d11642 | 执行 |
| 商品流量来源_自主访问 | 3001 | 5f9cc2ca20cad3d11642 | 已停用（enabled=False，与购物车数据口径重叠）|

> ⚠️ uuid前缀为前端每次会话**动态生成**（与数据无关，无需跟随更新）；如3001最新抓包为 `d6f270911983d45006dd`。

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

**Cookie使用方式**：从 `config/cookie.txt` 整体读取，**不解析**，作为字符串整体塞入Cookie头。过期需手动抓包更新。

#### 2.3 Body（表单参数）

| 字段名 | 类型 | 值 | 说明 |
|--------|------|-----|------|
| date | ⚠️ **动态** | 用户输入 | 查询日期，如 2026-07-30 |
| startDate | ⚠️ **动态** | = date | ⚠️ 与date一致（2026-08-05修复，见踩坑）|
| endDate | ⚠️ **动态** | = date | ⚠️ 与date一致（2026-08-05修复，见踩坑）|
| interval | 静态 | DAY | 时间粒度（config可变）|
| dateType | 静态 | day | 日期类型（config可变）|
| **lastSrcChannelId1** | 静态 | 2 | 一级渠道（搜索/推荐/购物车都是2）|
| **lastSrcChannelId2** | ⚠️ **业务关键** | **2008=搜索, 2009=推荐, 3001=购物车** | 二级渠道ID |
| groupType | 静态 | skuId | |
| attributes | 静态 | skuId | |
| sortField | 静态 | jdr_sch_traffic_enter_shop__browse_page_cnt_shop_last_src | 按入店浏览量 |
| sortType | 静态 | desc | |
| limit | 静态 | 5000 | 条数上限（config可变）|
| compareType | 静态 | hb | 环比 |
| **User-mup** | ⚠️ **风控** | int(time.time()*1000) | 13位毫秒时间戳 |
| **User-mnp** | ⚠️ **风控** | MD5(URL路径+uuid+时间戳+盐值) | MD5签名 |
| **uuid** | ⚠️ **风控** | {渠道uuid前缀}-{10位随机数} | 按渠道生成 |

> 说明：interval/dateType/limit 为**可变参数**（从config.xlsx读取，缺省兜底+警告）；
> lastSrcChannelId1/groupType/attributes/sortField/sortType/compareType 为**固定常量**（经用户确认写死代码 FIXED_BIZ_PARAMS）。

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

盐值放在 `config.xlsx`（项目"商品流量来源"分组），京东更新后可手动修改。

---

### 四、业务调用

```python
# 命令行调用（推荐）
python main.py --date "2026-07-30"                                          # 默认批量：搜索/推荐/购物车
python main.py --biz_key "商品流量来源_购物车" --date "2026-07-30"            # 单渠道
python main.py --biz_key "商品流量来源_搜索,商品流量来源_推荐" --date "2026-07-30"  # 多渠道批量

# 代码内部调用
from main import run_business
run_business("商品流量来源_搜索", date="2026-07-30")
run_business(["商品流量来源_搜索", "商品流量来源_推荐"], date="2026-07-30")
```

---

### 五、配置项（config.xlsx，2026-08-05 精简后）

| 项目名 | 变量参数 | 参数值 | 说明 |
|--------|----------|--------|------|
| 全局 | cookie文件路径 | config/cookie.txt | |
| 全局 | 输出目录 | output/ | |
| 全局 | 日志目录 | logs/ | |
| 全局 | 请求间隔(秒) | 30 | |
| 全局 | 最大重试次数 | 3 | |
| 全局 | 请求超时(秒) | 30 | |
| 商品流量来源 | 签名盐值 | 372ad2c2b6 | 京东更新时可改 |
| 商品流量来源 | date | 2026-07-29 | 默认查询日期 |
| 商品流量来源 | startDate / endDate | = date | 单日查询默认=date |
| 商品流量来源 | interval | DAY | 可变参数 |
| 商品流量来源 | dateType | day | 可变参数 |
| 商品流量来源 | limit | 5000 | 可变参数 |

**可变参数**（interval/dateType/limit）从config读取；**固定常量**（lastSrcChannelId1/groupType/attributes/sortField/sortType/compareType）经用户确认固化在代码 `FIXED_BIZ_PARAMS`。

---

### 六、Excel后置处理（2026-08-05 新增，全局生效）

1. 接口返回Excel二进制 → `pd.read_excel(dtype=str, na_filter=False)`（先按文本读，防pandas自动转数值）
2. `convert_date_format(date)` 把查询日期转成 `2026/7/30` 格式
3. 首列A插入【日期】列
4. `safe_convert_numeric(df)`：订单编号列强制文本；SKU/SPU转数字；>15位纯数字保留文本；其余可转则转
5. `apply_column_formats()`：日期列真实datetime+`yyyy/m/d`；订单编号列`@`文本；SKU/SPU列数值`0`位小数
6. 保存到 output/ 目录

后续京麦订单明细/售后/京准通等报表可直接复用上述公共工具函数。

---

### 七、踩坑经验

1. **lastSrcChannelId2 容易混淆**：
   - 2008 = 搜索子来源、2009 = 推荐子来源、3001 = 购物车/我的订单回流
   - 3001 与「自主访问」数据口径重叠，自主访问业务已停用（enabled=False）

2. **⚠️ date/startDate/endDate 三值必须一致（2026-08-05 重大坑）**：
   - 现象：导出的 2026-07-30 与网页对不上；07-29 与 07-30 导出完全相同
   - 根因：原逻辑 start/end 回落 config 旧值，`--date 2026-07-30` 实际发送
     `date=07-30&startDate=07-29&endDate=07-29` → 接口按 **startDate~endDate 区间**取数，返回07-29数据
   - ✅ **修复方案**（`_get_date_params()` 当前逻辑）：
     ```
     date      : 入参(--date)优先，其次 config.xlsx 的 date
     startDate : 入参(--start_date)优先，未传时默认=date
     endDate   : 入参(--end_date)优先，未传时默认=date
     ```
     即：只要用命令行 `--date` 指定新日期，三个日期自动同步为该日期，不再受 config 旧值影响。
   - ✅ **验证结果**：修复后 3渠道（搜索43行/推荐4行/购物车4行）与网页导出行数、SKU集合、数值**完全一致**
   - 使用方式：日常跑任意日期直接 `python main.py --date "YYYY-MM-DD"` 即可，无需改 config；
     如需区间查询用 `--start_date/--end_date` 显式指定。

3. **pandas 读取长数字精度丢失**：
   - `pd.read_excel()` 默认把数字样式列自动转 int64 → >15位保护失效
   - 必须 `read_excel(..., dtype=str, na_filter=False)`

4. **uuid前缀是动态的**：前端每次会话生成（3001出现过 5f9cc2ca... / d6f27091...），与数据无关，无需更新代码

5. **风控盐值会更新**：京东不定期更新，盐值在config可手动改

6. **30秒间隔必须遵守**：不遵守立即触发风控拦截；间隔在config可调

7. **UA切换是兜底方案**：默认Edge，遇风控自动切Chrome，不要禁用

---

### 八、变更记录

| 日期 | 改动 |
|------|------|
| 2026-08-04 | 初次实现商品搜索效果（2008） |
| 2026-08-04 | 重构：所有代码合并到 main.py（commit bc331c7） |
| 2026-08-04 | 新增商品推荐效果（2009），复用搜索的日期配置 |
| 2026-08-04 | 新增商品购物车效果（3001），CHANNEL_MAP改元组支持按渠道uuid前缀 |
| 2026-08-05 | 业务注册中心v2.0：BUSINESS_REGISTRY + run_business批量 + argparse命令行 + 配置一致性检查 |
| 2026-08-05 | 3001渠道执行名改回「购物车」；自主访问停用(enabled=False)；config项目名统一「商品流量来源」 |
| 2026-08-05 | config精简：6项固定参数固化代码常量，仅留 interval/dateType/limit 可变 |
| 2026-08-05 | Excel后置处理：日期列插入/数值安全转换/单元格格式 + 公共工具函数封装 |
| 2026-08-05 | 对齐京东订单导出风险：订单编号强制文本黑名单 + SKU/SPU数值0位小数 |
| 2026-08-05 | ⚠️ 修复日期参数同步Bug：--date 时 start/end 默认=date，数据与网页核对完全一致 |
