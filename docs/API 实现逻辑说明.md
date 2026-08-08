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

### 六、Excel后置处理（2026-08-05 新增，2026-08-07 升级公共规则）

**公共规则（所有报表统一，全局生效）**：

1. **日期列智能新增（规则1）**：
   - 报表**已存在**【日期】/【时间】列（列名="日期"或"时间"，或以"日期"结尾）→ **禁止重复插入日期列**，仅做日期格式标准化转换
   - 报表**无任何**日期/时间列 → 在**首列插入【日期】列**，填入导出业务对应的日期
   - ⚠️ 识别刻意**不用**"以'时间'结尾"匹配：防止"最近上架时间"等业务时间字段被误判
   - 统一入口：公共函数 `prepare_date_columns(df, date)` → 返回 `(date_column, date_value)`

2. **日期格式统一（规则2）**：目标格式 `yyyy/m/d`（例 `2026/7/29`）
   - `20260729` → `2026/7/29`
   - `2026-07-29` → `2026/7/29`
   - `2026-07-29 13:45:59` → `2026/7/29 13:45:59`（带时分秒保留时间部分，只改写日期分隔符）

3. **数值转换规则（规则3）**：
   - **SKU / SPU 字段**：转为数字，单元格小数位数设 0（`INTEGER_ZERO_DECIMAL_COLUMNS={SKU,SPU}`）
   - **订单编号列**：强制跳过转换、保持文本（`TEXT_FORCE_COLUMNS={订单编号}` → 格式 `@`），防止长数字科学计数、末尾归零
   - **其余指标**（访客、浏览、成交金额等）：批量转为数值；纯数字>15位保留文本（兜底防精度丢失）

**处理流程**：
1. 接口返回Excel二进制 → `pd.read_excel(dtype=str, na_filter=False)`（先按文本读，防pandas自动转数值）
2. `prepare_date_columns(df, date)`：规则1智能插入/标准化 + 规则2格式统一
3. `safe_convert_numeric(df)`：规则3数值安全转换（订单编号文本/SKU·SPU数字/其余数值/>15位保文本）
4. 写入Excel + `apply_column_formats()`：日期列真实datetime+`yyyy/m/d`（带时间用`yyyy/m/d hh:mm:ss`）；订单编号`@`文本；SKU/SPU数值`0`位小数
5. 保存到 output/ 目录

**输出目录规则（规则4）**：按业务模块建文件夹 + 日期子文件夹
- `output/{业务模块}/{date}/{文件名}`，如 `output/搜索流量/2026-08-01/搜索流量_2026-08-01.xlsx`
- 业务模块名 = 文件名去掉 `_{date}.xlsx` 后缀的主体（公共函数 `build_business_output_path()` 统一构造）

**落地实例（2026-08-07）**：
- 商品明细导出（原表自带【时间】列）→ 不再插入【日期】列，仅标准化【时间】列为 `2026/7/29`
- 店铺来源-三级渠道（原表自带【时间】列）→ 同上，消除冗余
- 搜索/推荐/购物车（原表无日期/时间列）→ 仍插入【日期】列（行为不变）

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

---

## 项目 4：店铺来源-三级渠道（离线流量报表）｜2026-08-07 上线

### 业务定位
按三级流量渠道（`lastSrcChannelId3`）分组，导出店铺来源离线日度流量报表。**与项目 1-3 的"商品流量来源（SKU 维度）"完全不同**——本项目是"渠道流量来源"维度。

### 接口
| 字段 | 内容 |
|------|------|
| **URL** | `https://szgateway.jd.com/szpaas/szajax/shop/source/offlineFlowSource/downTable.ajax` |
| **方法** | POST |
| **Content-Type** | `application/x-www-form-urlencoded` |
| **页面入口** | `https://sz.jd.com/szweb/sz/view/viewflow/viewSourcesVNew.html` |

### 必带 Header（用户 2026-08-06 强调，缺失即拦截）
```python
ORIGIN = "https://sz.jd.com"
REFERER = "https://sz.jd.com/szweb/sz/view/viewflow/viewSourcesVNew.html"
```

### 业务表单参数（13 项 = 9 固定 + 1 可变 + 3 风控动态）

| 参数 | 值 | 类型 | 说明 |
|------|-----|------|------|
| `compareType` | `hb` | ❌ 固定 | 对比方式：环比 |
| `interval` | `DAY` | ❌ 固定 | 聚合粒度 |
| `dateType` | `day` | ❌ 固定 | 日期类型 |
| `downType` | `day` | ❌ 固定 | 下载类型（本项目独有）|
| `groupType` | `lastSrcChannelId3` | ❌ 固定 | 分组维度 |
| `attributes` | `lastSrcChannelId3` | ❌ 固定 | 返回字段 |
| `sortField` | `jdr_sch_traffic_enter_shop__visitor_cnt_shop_last_src` | ❌ 固定 | 排序字段（候选 A，进店访客数降序）|
| `sortType` | `desc` | ❌ 固定 | 排序方式 |
| `lastSrcChannelId1` | `2` | ❌ 固定 | 一级渠道 |
| `platformCate1` | `""`（默认空）| ✅ 可变 | 平台品类 1（本项目独有），空=全品类 |
| `date` / `startDate` / `endDate` | `2026-08-04` | ✅ 可变 | 三值必须一致 |
| `User-mup` | 毫秒时间戳 | ⚠️ 风控 | 每次调用 `int(time.time()*1000)` |
| `User-mnp` | MD5 签名 | ⚠️ 风控 | 算法见下 |
| `uuid` | 完全随机 | ⚠️ 风控 | **与项目1-3不同，必须完全随机化** |

### 风控签名算法（与项目1-3 复用）
```
User-mnp = MD5(URL路径 + uuid + 时间戳 + "372ad2c2b6")
```
- 盐值 `372ad2c2b6` 从 config.xlsx【全局配置】读取（与项目1-3 共用）
- 失败排查：参考 SKILL.md 第三节"签名验证失败排查指南"

### UUID 完全随机生成（与项目1-3 关键差异）
```python
def _gen_uuid_random(self):
    import secrets
    prefix = secrets.token_hex(8)        # 8字节 = 16hex
    suffix = secrets.token_hex(5)        # 5字节 = 10hex
    return f"{prefix}-{suffix}"
```
- 用户抓包两次（间隔 13 秒）：`f1d5ae161b41f4153fc0` → `a31e066d8e94f4f39a3a`，**前缀完全不同**
- 与项目1-3（prefix 固定 `ca412182e5668a106054`）模式不同
- **不依赖基类 UUID_PREFIX**（符合用户 2026-08-06 "禁止硬编码 uuid" 约束）

### 入口命令
```bash
# 列出所有业务（含本项目）
python main.py --list

# 跑本项目（指定日期）
python main.py --biz_key "店铺来源_三级渠道" --date "2026-08-04"

# 批量跑（按日期分多次跑）
for d in 2026-08-04 2026-08-05 2026-08-06; do
    python main.py --biz_key "店铺来源_三级渠道" --date "$d"
    sleep 30  # 严格30秒间隔，防止风控
done
```

### 容错与风控适配（阶段 4 新增）
| 场景 | 行为 |
|------|------|
| HTTP 200 + < 1KB | 视为空响应，抛 RuntimeError，重试 |
| HTTP 200 + magic bytes ≠ `PK\x03\x04` | 视为非 Excel，记录前 200 字节，重试 |
| HTTP 401 | 抛 CookieExpiredError，停止重试 |
| HTTP 403 | 警告 + 重试（UA 切换兜底） |
| 业务码 601（操作频繁） | **不重试**（避免加重风控，让用户决定）|
| 业务码 302/-1（含"登录"） | 抛 CookieExpiredError，停止重试 |
| 业务码 -407/-402（签名错误）| 重试兜底（UA 切换）|
| 超时（30 秒）| 重试 |
| 其他网络异常 | 重试（UA 切换）|

### 重试机制（与基类 request() 对齐）
- 最多 3 次（`MAX_RETRIES`，从 config 读取）
- 递增等待：`30 / 60 / 90 秒`（`REQUEST_INTERVAL * attempt`）
- 每次重试前切换 UA（Edge ↔ Chrome）

### 输出
- 文件名：`店铺来源_三级渠道_YYYY-MM-DD.xlsx`（保存到 `output/店铺来源_三级渠道/{date}/` 子目录，AGENTS.md Excel规则4）
- Excel 后置处理：复用基类 `_save_flow_excel`（日期列插入 / 数值安全转换 / 单元格格式）

### 与项目1-3 的核心差异汇总
| 维度 | 项目1-3（downSkuTable.ajax）| 项目4（downTable.ajax）|
|------|--------------------------|----------------------|
| 业务类 | ProductFlowAPI | **OfflineChannelAPI** |
| 分组维度 | SKU（id2=2008/2009/3001）| **三级渠道**（id3）|
| 业务目的 | 商品效果分析 | 渠道投产分析 |
| Referer | flowPathDetailsNew.html | **viewSourcesVNew.html** |
| 业务参数 | 9 项 | **12 项**（+downType / +platformCate1）|
| 排序字段 | 进店浏览量 | **进店访客数** |
| UUID 策略 | 固定 prefix + 随机后缀 | **完全随机** |
| 鉴权签名 | UUID_PREFIX（基类常量）| **业务内完全随机** |

### 入口检查清单（写代码前确认）
- [x] 业务类继承 JDBaseRequest
- [x] API_URL / ORIGIN / REFERER 类常量固定
- [x] FIXED_BIZ_PARAMS 9 项（含 lastSrcChannelId1 / groupType / attributes / sortField / sortType / compareType / interval / dateType / downType）
- [x] VARIABLE_BIZ_PARAMS 1 项（platformCate1，默认空）
- [x] _gen_uuid_random() 完全随机（不依赖 UUID_PREFIX）
- [x] _gen_risk_params_random() 复用 MD5 算法 + 全局 SIGN_SALT
- [x] download_offline_channel() 主方法：日期解析 + 参数组装 + 必带 header + 重试循环 + 风控识别 + Excel 校验 + 后置保存
- [x] _check_business_code() 风控业务码识别
- [x] _validate_excel_response() 空响应 + magic bytes 校验
- [x] BUSINESS_REGISTRY 注册 `"店铺来源_三级渠道"`
- [x] config_consistency_check() 报告新增 3.4 节

### 关联文件
- 业务类：`main.py`（搜索 `class OfflineChannelAPI`）
- 业务沉淀：`.trae/skills/jd-api-analyze/SKILL.md` 项目4
- 踩坑记录：`全局复利的踩坑日志.md` 坑6（UUID 策略差异）
- 文档索引：自动入库到 `docs/项目文档索引.xlsx`

---

## 项目 5：商品明细导出（ProductDetailAPI）｜2026-08-07 上线

### 业务定位
在「商品分析-商品明细」页面按**二级/三级类目 + 渠道**维度导出商品明细流量报表。与项目1-3（SKU维度）、项目4（三级渠道维度）不同——本项目按**商品维度**导出明细，且为**GET 请求**（参数全拼 URL）。

### 接口
| 字段 | 内容 |
|------|------|
| **URL** | `https://sz.jd.com/sz/api/productDetail/exportProList.ajax` |
| **方法** | **GET**（参数全拼 URL query string，与项目1-4 的 POST 完全不同）|
| **域名** | `sz.jd.com`（项目1-4 是 `szgateway.jd.com`）|
| **页面入口** | `https://sz.jd.com/szweb/sz/view/productAnalysis/productDetail.html` |
| **Sec-Fetch-Site** | `same-origin`（项目4 是 `same-site`，必须显式覆盖基类默认值）|
| **成功响应** | HTTP 200 + `Content-Disposition: attachment` + body 前 4 字节 `PK\x03\x04` |

### 必带 Header
```python
REFERER = "https://sz.jd.com/szweb/sz/view/productAnalysis/productDetail.html"
extra_headers = {"Referer": REFERER, "Sec-Fetch-Site": "same-origin"}
```

### 业务参数（GET query，7 项 = 3 固定 + 3 可变 + 日期 3 值）
| 参数 | 值 | 类型 | 说明 |
|------|-----|------|------|
| `type` | `0` | ❌ 固定 | FIXED_BIZ_PARAMS |
| `categoryType` | `0` | ❌ 固定 | FIXED_BIZ_PARAMS |
| `downloadType` | `dayList` | ❌ 固定 | 下载类型（日列表）|
| `second` | `999999` | ✅ 可变 | 二级类目，默认全类目，CLI `--second` 可覆盖 |
| `third` | `""` | ✅ 可变 | 三级类目，默认空不限 |
| `channel` | `99` | ✅ 可变 | 渠道，默认全部 |
| `isMonitored` | `undefined` | ✅ 可变 | 是否监控商品 |
| `date` / `startDate` / `endDate` | `2026-08-07` | ✅ 可变 | 三值必须一致（复用踩坑经验）|
| `User-mup` | 毫秒时间戳 | ⚠️ 风控 | 每次调用 `int(time.time()*1000)` |
| `User-mnp` | MD5 签名 | ⚠️ 风控 | 算法与项目1-4 相同 |
| `uuid` | 完全随机 | ⚠️ 风控 | 16hex-10hex，与项目4 相同 |

> 注意：`isMonitored=undefined` 是**字符串**，拼 URL 时 requests 会原样编码为 `undefined`（真实抓包如此）。

### 风控签名算法（复用全局盐值）
```
User-mnp = MD5(URL路径 + uuid + 时间戳 + "372ad2c2b6")
```
- 盐值从 config.xlsx【全局配置】读取，与项目1-4 共用
- uuid 完全随机：`secrets.token_hex(8) + "-" + secrets.token_hex(5)`（与项目4 相同，抓包证实前缀每次不同）

### 目录规则（本项目独有，业务子目录）
```
output/商品明细/{date}/{原始文件名}
例：output/商品明细/2026-08-07/商品明细导出_2026-08-07.xlsx
```
- 原始文件名从响应头 `Content-Disposition` 解析（`filename*=UTF-8''...` 需 URL 解码；`filename="..."` 直接提取）
- 解析失败时兜底命名：`{second or '全类目'}_{date}_商品明细.xlsx`
- Excel 后置处理：日期列插入 + 数值安全转换 + 单元格格式（复用公共工具函数）

### 容错与风控适配（阶段 4 修复，3 处核心缺陷）
| # | 缺陷 | 原行为 | 修复后 |
|---|------|--------|--------|
| 1 | **601 限流"不重试"失效** | 601 抛 RuntimeError → 被 `except Exception` 捕获 → **继续重试 3 次**（加重风控）| 601 抛 `RiskControlError` → 循环内 `except RiskControlError: raise` **直接抛出不重试**（单测验证只请求 1 次）|
| 2 | **最后一次失败误保存错误内容** | 判断"response 有内容 → 继续保存"，若最后一次响应是 HTML 错误页（>1KB）会**误判成功保存错误内容** | 新增 `success` 标记，只有 `break` 才算成功，否则一律抛 RuntimeError |
| 3 | **文本型 601 无法识别** | 非 json 的 HTML 错误页（含"操作频繁"）只当普通失败重试 3 次 | magic 校验分支检测"操作频繁/频繁"字样 → 抛 `RiskControlError` 停止重试 |

> 修复1 同时**同步应用到项目4**（同属"601不重试"全局风控硬约束，非业务功能改动，与京麦 SKILL 第八节一致）。

### 异常抛出规则（阶段 4 定版）
| 异常 | 触发场景 | 处理 |
|------|----------|------|
| `CookieExpiredError` | HTTP 401；业务码 302/-1；message 含"登录/login" | **立即停止**，提示更新 config/cookie.txt |
| `RiskControlError` | 业务码 601；文本含"操作频繁/频繁" | **不重试**，直接抛出，提示冷却 30-120 分钟 |
| `RuntimeError` | 业务码 -407/-402；空响应(<1KB)；非 Excel(magic 不匹配)；缺 attachment | **重试兜底**（UA 切换 + 递增等待）|
| `requests.exceptions.Timeout` | 请求超时（30 秒）| 重试 |

### 测试结论（阶段 4，mock 网络单测 24/24 通过，临时脚本已删）
| 类别 | 覆盖点 | 结果 |
|------|--------|------|
| 风控码识别 | 601→RiskControlError / 302·登录→CookieExpiredError / -407→RuntimeError / 正常放行 / 非json跳过 | 6/6 ✅ |
| 双重校验 | 正常通过 / 缺attachment / HTML错误页 / 文本型601 / 空响应 | 5/5 ✅ |
| 文件名解析 | UTF-8中文解码 / 普通文件名 / 空头兜底 | 3/3 ✅ |
| 随机参数 | uuid两次不同 / mnp 32位md5 / mup时间戳 / uuid 16hex-10hex格式 | 4/4 ✅ |
| 全失败兜底 | 抛 RuntimeError 且不误保存 | 2/2 ✅ |
| 601 不重试 | 抛 RiskControlError 且只请求1次 | 2/2 ✅ |
| 成功路径 | 路径含 `output/商品明细/{date}/` 子目录 | 2/2 ✅ |

### 入口命令
```bash
# 列出业务
python main.py --list

# 单日全类目导出
python main.py --biz_key "商品明细导出" --date "2026-08-07"

# 指定二级类目
python main.py --biz_key "商品明细导出" --date "2026-08-07" --second "12345"
```

### 与项目4 的核心差异汇总
| 维度 | 项目4（downTable.ajax）| 项目5（exportProList.ajax）|
|------|----------------------|--------------------------|
| 业务类 | OfflineChannelAPI | **ProductDetailAPI** |
| 请求方式 | POST（表单）| **GET（参数拼 URL）** |
| 域名 | szgateway.jd.com | **sz.jd.com** |
| Sec-Fetch-Site | same-site | **same-origin** |
| 必带 Header | Origin + Referer | **Referer**（同源无需 Origin）|
| 分组维度 | 三级渠道 | **商品明细（类目）** |
| 业务参数 | 12 项表单 | **7 项 query**（type/categoryType/downloadType/second/third/channel/isMonitored）|
| 响应校验 | magic 字节 | **Content-Disposition attachment + magic 字节双重校验** |
| 文件名 | 代码固定拼接 | **从 Content-Disposition 解析原始名** |
| 保存目录 | output/店铺来源_三级渠道/{date}/ 子目录 | **output/商品明细/{date}/ 子目录** |
| 601 处理 | 阶段4 同步修复为不重试 | **不重试（RiskControlError）** |

### 变更记录
| 日期 | 改动 |
|------|------|
| 2026-08-07 | 阶段3：ProductDetailAPI 完整实现 + BUSINESS_REGISTRY 注册（第6业务）+ 修复 `_save_flow_excel` 误插类 Bug |
| 2026-08-07 | 阶段4：601 改抛 RiskControlError 不重试（含项目4 同步）；success 标记防误保存；文本型 601 识别；单测 24/24 通过 |

---

## 项目 6：商品流失分析（LossProductAPI）｜2026-08-07 上线

### 业务定位
在「竞争分析-竞争流失-商品流失分析」页面导出**流失商品明细报表**（哪些商品引起本店成交客户流失到竞品店）。与项目5 同属 `sz.jd.com` 域，但**首次出现 `.xls`（OLE2复合文档）响应**，是核心差异。

### 接口
| 字段 | 内容 |
|------|------|
| **URL** | `https://sz.jd.com/sz/api/competitionAnalysis/exportLossProList.ajax` |
| **方法** | **POST**（表单 `application/x-www-form-urlencoded`，与项目5 的 GET 不同）|
| **域名** | `sz.jd.com` |
| **页面入口** | `https://sz.jd.com/sz/view/competitionAnalysis/lossAnalysiss.html` |
| **Sec-Fetch-Site** | `same-origin`（同项目5，覆盖基类默认）|
| **成功响应** | HTTP 200 + `Content-Disposition: attachment;charset=utf-8` + body 前 4 字节 **`\xD0\xCF\x11\xE0`（.xls 魔数）** |
| **响应文件名头** | `filename=商品流失分析_全部渠道_SPU_20260805_20260805.xls`（**UTF-8 字节直放**，非 URL 编码）|

### 必带 Header
```python
ORIGIN  = "https://sz.jd.com"
REFERER = "https://sz.jd.com/sz/view/competitionAnalysis/lossAnalysiss.html"
extra_headers = {"Origin": ORIGIN, "Referer": REFERER, "Sec-Fetch-Site": "same-origin"}
```

### 业务表单参数（POST，固定 2 项 + 日期 3 值 + 风控 3 项）
| 参数 | 值 | 类型 | 说明 |
|------|-----|------|------|
| `indChannel` | `99` | ❌ 固定 | 渠道（99=全部渠道，用户确认固化 FIXED_BIZ_PARAMS）|
| `unitType` | `0` | ❌ 固定 | 维度（0=SPU，用户确认固化）|
| `date` / `startDate` / `endDate` | `2026-08-05` | ✅ 可变 | 三值必须一致（复用踩坑经验）|
| `User-mup` | 毫秒时间戳 | ⚠️ 风控 | `int(time.time()*1000)` |
| `User-mnp` | MD5 签名 | ⚠️ 风控 | `MD5(URL路径 + uuid + 时间戳 + 盐值372ad2c2b6)` |
| `uuid` | 完全随机 | ⚠️ 风控 | `secrets.token_hex(8) + "-" + secrets.token_hex(5)`（16hex-10hex，同项目4/5）|

### .xls 读取与转存（本项目核心新增）
- **响应魔数**：`.xls` 是 `\xD0\xCF\x11\xE0`（OLE2复合文档），`.xlsx` 是 `PK\x03\x04`（zip）
- **公共函数** `read_excel_bytes(content)`（main.py 顶部）：按魔数自动选引擎
  - `PK\x03\x04` → 默认 openpyxl 引擎
  - `\xD0\xCF\x11\xE0` → `engine="xlrd"`（需 `pip install xlrd>=2.0.1`，pandas 3.0 要求）
  - 都不匹配 → `ValueError`（调用方重试兜底）
- **保存流程** `_save_excel_to_path()`：`read_excel_bytes` → `prepare_date_columns`（规则1+2）→ `safe_convert_numeric`（规则3）→ 写 `.xlsx` + `apply_column_formats`
- 输出统一**转存为 .xlsx**（用户确认决策），后缀 `xls → xlsx` 替换

### ⚠️ 文件名编码踩坑（阶段5 真实导出发现，已修复）
| 现象 | 根因 | 修复 |
|------|------|------|
| 真实导出文件名乱码 `鍟嗗搧娴佸け鍒嗘瀽_鍏ㄩ儴娓犻亾...` | 服务器 filename 字节是 **UTF-8**（header 声明 charset=utf-8），requests 按 latin-1 解码成 U+00xx 字符；此前按 GBK 解码 UTF-8 字节 → 产生"鍟嗗搧"乱码 | `_parse_content_disposition_filename` 改为 **UTF-8 优先解码**、GBK 回退：`raw.encode('latin-1')` 还原字节 → 先 `decode('utf-8')` 且无 `\ufffd` 替换符即返回 → 失败再 `decode('gbk')` 回退 |
| mock 测试用 GBK 字节构造 → 当时没暴露 | mock 与真实服务器编码不一致 | 验证脚本覆盖 4 场景（真实UTF-8 / GBK兼容 / filename*URL编码 / 纯ASCII），4/4 通过 |

### 容错与风控适配（阶段 4，与项目4/5 对齐）
| 场景 | 行为 |
|------|------|
| HTTP 200 + < 1KB | 空响应拦截 → RuntimeError 重试 |
| 魔数 ≠ `PK\x03\x04` 且 ≠ `\xD0\xCF\x11\xE0` | 非 Excel → RuntimeError 重试（**双魔数校验**，本项目独有）|
| 缺 `Content-Disposition: attachment` | 疑似风控伪装 → RuntimeError 重试 |
| HTTP 401 / 业务码 302·-1·登录 | `CookieExpiredError` 立即停止 |
| 业务码 601 / 文本"操作频繁" | `RiskControlError` **不重试**直接抛出 |
| `success` 标记 | 只有 break 才算成功，防最后一次失败误保存 |
| 重试 | 3 次递增等待 30/60/90s + UA 切换 Edge↔Chrome |

### 测试结论
| 阶段 | 结果 |
|------|------|
| 阶段4 mock 单测 | **30/30 通过**（xls 读取 / GBK·UTF-8 文件名 / 双魔数校验 / 601 不重试 / success 标记 / 转存 xlsx / 公共规则）|
| 阶段5 真实导出 | HTTP 200 + 8192 字节 + **文件名正常**（UTF-8 解码修复后）→ 转存 xlsx 6340 字节 |
| 数据核对 | 14 行 × 13 列（日期/商品名称/商品ID/流失成交金额/流失成交客户数/流失率/关注后流失人数/加购后流失人数/关注后跳失人数/加购后跳失人数/直接跳失人数/引起流失的商品数/引起流失的店铺数），与抓包响应一致；商品ID 保留文本、日期列首列插入 |

### 入口命令
```bash
python main.py --biz_key "商品流失分析" --date "2026-08-05"
```

### 与项目5 的核心差异汇总
| 维度 | 项目5（exportProList.ajax）| 项目6（exportLossProList.ajax）|
|------|---------------------------|------------------------------|
| 业务类 | ProductDetailAPI | **LossProductAPI** |
| 请求方式 | GET（参数拼 URL）| **POST（表单）** |
| 页面入口 | productAnalysis/productDetail.html | **competitionAnalysis/lossAnalysiss.html** |
| 响应格式 | .xlsx（PK 魔数）| **.xls（\xD0\xCF\x11\xE0 魔数，首次出现）** |
| filename 头 | `filename*=UTF-8''`（URL 编码）| **`filename=`（UTF-8 字节直放）** |
| 读取引擎 | openpyxl | **xlrd（read_excel_bytes 自动识别）** |
| 保存 | 直接保存 | **xls 读取 → 转存 .xlsx** |
| 固定业务参数 | type/categoryType/downloadType | **indChannel=99 / unitType=0** |

### 变更记录
| 日期 | 改动 |
|------|------|
| 2026-08-07 | 阶段1-2：需求拆解 + 骨架零改动验证（30/30 mock 通过）|
| 2026-08-07 | 阶段3：LossProductAPI 完整实现（POST + xls 读取 + 业务子目录）+ BUSINESS_REGISTRY 注册（第7业务）|
| 2026-08-07 | 阶段4：双魔数校验 + RiskControlError + success 标记 + 601 不重试 |
| 2026-08-07 | 阶段5：真实导出发现 filename **UTF-8 编码**乱码 → `_parse_content_disposition_filename` UTF-8 优先解码 + GBK 回退；重测通过；文档归档 |

---

## 项目 7：京准通快车自定义报表导出（JZTKuaicheAPI）｜2026-08-07 上线骨架

### 业务定位
京准通广告投放后台（`jzt.jd.com/home`）的快车广告报表导出。**与项目1-6 商智域完全不同**——鉴权用 h5st 请求头（非商智 User-mnp/uuid），三步异步流程（创建任务→轮询列表→CDN 下载），Cookie 与商智/京麦不互通。本阶段（阶段3骨架）仅 3 个接口方法，不含轮询循环与 Excel 解析。

### 接口基础信息
| 字段 | 内容 |
|------|------|
| **base URL** | `https://jzt-api.jd.com/dataCenter/customreport/v2/report` |
| **请求方式** | POST（创建任务）/ GET（任务列表、CDN 下载）|
| **Content-Type** | `application/json;charset=UTF-8` |
| **业务页面入口** | `https://jzt.jd.com/home`（快车-自定义报表）|

### 三个核心接口
| 序号 | 方法 | URL | 用途 |
|------|------|-----|------|
| 1 | `create_export_task(start_date, end_date) -> str` | `POST /add?requestFrom=0&businessFrom=1` | 创建导出任务，返回 `reportId` |
| 2 | `get_task_list() -> dict` | `GET /list?requestFrom=0&businessFrom=1` | 查询任务列表原始 dict（含 status / downloadUrl）|
| 3 | `download_report(task_id, save_filename) -> str` | `GET <downloadUrl>`（CDN）| 根据 task_id 查找 downloadUrl，二进制保存 |

### 必带 Header（业务约束）
```python
ORIGIN = "https://jzt.jd.com"
REFERER = "https://jzt.jd.com"
siteId = "0"
User-Agent = "<浏览器抓包 UA>"
Content-Type = "application/json;charset=UTF-8"
Cookie = "<jzt.jd.com 域完整 Cookie>"
h5st = "<浏览器F12抓 add 接口请求头复制>"
```
- ⚠️ **不写 uuid**：京准通域未校验 uuid 字段（项目7 抓包已确认）
- ⚠️ **UA 与 h5st 绑定**：禁用基类 UA 切换（切换会致 h5st 失效），业务内固定 UA

### Cookie 与 h5st（**严禁硬编码**）
- Cookie 文件路径：`config/jzt_cookie.txt`（**独立文件**，与 `config/cookie.txt` 商智 Cookie 不互通）
- Cookie 文件不存在 → `FileNotFoundError`（强制用户抓包填入）
- h5st 通过 `__init__(h5st=...)` 外部传入；脚本不实现 JS 签名（复杂度高）
- h5st 过期（业务码 601）→ 提示用户重新抓包更新

### 业务 Payload 模板（最小字段版）
- 完整 payload 体量大，阶段3 用类内常量 `JZT_KUAICHE_PAYLOAD_TEMPLATE` 存最少字段版
- 动态注入：`startTime` / `endTime` / `startTimeStr` / `endTimeStr` / `tempName` / `reportName`
- 延后引入：完整 payload 抽到 `templates/*.json`（阶段迭代再做）

### 业务参数关键点
| 参数 | 决策 | 说明 |
|------|------|------|
| `caliberSettings` | 转化周期 15 天 + 点击 + 不含赠品 + 成交订单 | 4 项默认勾选 |
| `customDimension` | 基础维度（产品线/计划/单元）+ 细分维度（营销目标/搜索词/关键词）| 字段最少版 |
| `customDimensionOptions` | pin=subUser=99936530475=FYA8888 | 当前账号 |
| `customIndex` | 基础数据（展现/点击/点击率/花费）+ 转化数据（直接订单数/金额/总订单数/金额/ROI）| 9 项核心指标 |
| `daily=1` | 日维度 | |
| `version="JZT_V9"` | 报表版本 | 抓包得到 |

### 与现有架构的核心差异（**不继承 JDBaseRequest**）
| 维度 | 商智域（项目1-6）| 本项目（项目7）|
|------|----------------|--------------|
| 鉴权参数 | Cookie + User-mnp/uuid | **Cookie + h5st**（不依赖 User-mnp）|
| 请求方式 | 表单 / JSON | **JSON** |
| 鉴权签名 | MD5(URL+uuid+ts+salt) | **h5st 一次性签名（外部传入）** |
| Cookie 文件 | `config/cookie.txt` | **`config/jzt_cookie.txt`**（独立）|
| 流程模型 | 同步请求-响应 | **异步三步：创建任务 → 轮询 → 下载** |
| 重试模型 | 基类 30 秒间隔 + 3 次重试 | **不适配**（h5st 短期有效，重试加重风控）|
| UA 切换 | 基类 Edge↔Chrome 切换 | **禁用**（UA 与 h5st 绑定）|
| Excel 后置 | `safe_convert_numeric` + `apply_column_formats` | **阶段5 再补**（基础骨架不含）|

### 异常抛出规则（基础骨架版，阶段4 会扩展）
| 异常 | 触发场景 | 处理 |
|------|----------|------|
| `FileNotFoundError` | `config/jzt_cookie.txt` 不存在 | 立即停止，提示抓包填入 |
| `ValueError` | h5st 为空 / Cookie 文件为空 | 立即停止 |
| `RuntimeError` | 业务码 601（h5st 过期）/ 业务码非0 / HTTP raise_for_status | 立即停止，提示用户排查 |
| `requests.exceptions.Timeout` | 请求超时（30s 创建/轮询，60s 下载）| **未捕获**（让调用方处理，阶段4 加入重试）|

### 调用示例（写进代码注释）
```python
from main import JZTKuaicheAPI

# h5st 从浏览器抓 add 接口获取填入
api = JZTKuaicheAPI(h5st="抓包得到h5st字符串")
task_id = api.create_export_task("2026-08-07", "2026-08-07")
print(f"任务ID: {task_id}")
task_data = api.get_task_list()  # 阶段3 骨架：调用方手动轮询
# api.download_report(task_id, "report_0807.xlsx")
```

### 入口命令
```bash
# 注册已生效，可用 --list 验证
python main.py --list

# ⚠️ 阶段3 骨架不直接支持 --biz_key 入口（需手动编排 create→poll→download）
# 阶段4/5 才会包装成 --biz_key 一键调度
```

### 阶段3 验证结论
| 项 | 结果 |
|------|------|
| AST 语法校验 | ✅ 通过 |
| 类方法归属（`__dict__` 校验，5 方法）| ✅ create_export_task / get_task_list / download_report / _build_payload / __init__ 全部在类内 |
| BUSINESS_REGISTRY 注册 | ✅ 第8业务 `京准通快车自定义报表` 已注册 |
| `python main.py --list` | ✅ 显示「[8] 京准通快车自定义报表」|
| payload 时间字段注入 | ✅ 6 个时间字段正确动态填充 |
| 模板字段体量 | 18 个顶层字段（最小字段版）|

### 阶段交付承诺
- ✅ 阶段1：需求拆解文档（已在对话中输出）
- ✅ 阶段2：抓包确认（用户决策「不写uuid」「合并main.py」，落地为阶段3 代码骨架）
- ✅ 阶段3：JZTKuaicheAPI 类骨架 + BUSINESS_REGISTRY 注册 + 验证全过（本文档）
- ⏳ 阶段4：容错适配（轮询循环 / CDN 403 重刷 URL / 401 Cookie 过期 / 业务码-407/-402）
- ⏳ 阶段5：mock 单测 → 真实跑通 → docs/SKILL.md 归档 → GitHub 推送

### 变更记录
| 日期 | 改动 |
|------|------|
| 2026-08-07 | 阶段1：需求拆解 + 5 阶段交付计划（与商智项目1-6 同套流程）|
| 2026-08-07 | 阶段2：用户决策汇总（合并 main.py / 不写 uuid / payload 类内常量 / Cookie 独立文件 / 兜底 2026-08-07）|
| 2026-08-07 | 阶段3：JZTKuaicheAPI 类骨架（不继承 JDBaseRequest，3 接口方法）+ 模板常量 JZT_KUAICHE_PAYLOAD_TEMPLATE + BUSINESS_REGISTRY 第8业务 + 验证脚本全过 |
