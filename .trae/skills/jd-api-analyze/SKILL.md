---
name: jd-api-analyze
description: 京东全系业务接口分析（商智 szgateway / 京麦 seller-v10.shop / 京准通 jzt）+ 抓包解析、参数拆解、风险识别、输出带详细注释的可运行python代码；承载京东全系业务沉淀（接口规范、踩坑经验、迭代记录、代码集成规范、归档规则等）
---

# 京东全系业务接口分析 + 项目知识沉淀 Skill

> 本文件按需加载，覆盖京东全系业务（商智/京麦/京准通）。
> 加载本Skill的场景：京东任何业务抓包分析、生成/重构京东接口代码、查阅京东项目历史/踩坑经验/接口规范。
>
> ⚠️ **Skill体量监控**：当前 < 2000行，**维持单Skill+内部分区模式**。后续若膨胀超过2000行再评估拆分子Skill。

---

## 📌 京东各业务域名与抓包备注（通用头部）

### 1. 商智后台（szgateway.jd.com）
- **域名**：`szgateway.jd.com`
- **业务**：店铺经营数据分析报表接口（流量、来源、SKU效果、活动效果等）
- **鉴权**：通过京东账号体系登录后获取Cookie，含 thor / light_key / User-mnp 等风控字段

### 2. 新版京麦商家工作台（seller-v10.shop.jd.com）
- **域名**：`seller-v10.shop.jd.com`
- **业务**：订单管理、售后管理、店铺业务报表
- **登录鉴权域名**：`passport.shop.jd.com`
- **Cookie归属**：所有 `.shop.jd.com` 子域共享 Cookie（与 szgateway.jd.com 不互通）
- **抓包备注**：Cookie 抓自 `seller-v10.shop.jd.com` 或 `passport.shop.jd.com` 都可

### 3. 京准通广告投放后台（jzt.jd.com/home）
- **域名**：`jzt.jd.com/home`
- **业务**：快车、海投、DMP、广告数据报表接口
- **Cookie**：⚠️ **与京麦Cookie不完全互通**，抓包时必须抓取本域名下 Cookie，不可直接复用京麦Cookie

---

# 【通用基础规则】（京东全系业务共用）

## 一、京东接口通用分析方法

### 强制执行步骤
#### 步骤1：接口基础总览
输出：请求方式、完整URL、业务用途、所属模块；对接口路径内英文名词给出通俗释义。

#### 步骤2：请求头Headers拆解
表格输出，列：字段名｜示例值｜业务含义｜标记类型
标记类型：✅可动态修改 / ❌固定不可修改 / ⚠️鉴权会过期
重点识别Cookie核心签名字段：thor、light_key、User‑mnp、3AB9D会话串，标注过期风险。

#### 步骤3：Body表单/参数逐字段拆解
表格输出：参数名｜示例｜业务含义｜是否可动态修改｜小白注释。
重点标记前端JS加密无法逆向生成的参数 User‑mnp，明确告知限制：Python无法本地复现，只能浏览器抓包复制。

#### 步骤4：风险点汇总
逐条输出：鉴权过期风险、参数格式约束、条数上限、是否支持分页、防盗链校验、访问频率风控、账号隔离。

#### 步骤5：输出完整可运行Python代码
代码强制结构：
1. import导入区，每个import附带注释，说明库用途
2. 请求头配置区
3. ✂【业务可修改配置区】集中放日期、渠道ID、条数等业务变量
4. 表单参数组装区
5. 请求执行 + 结果打印

代码约束：
- 英文关键字、http参数、库方法全部附带中文注释，小白可读。
- 动态签名、Cookie只留变量位置，注释提醒用户从浏览器抓包替换，严禁写死。
- 高亮标注 User‑mnp 的限制说明。

### 输出约束
表格清晰，不可省略风险警告；京东接口牢记 User‑mnp 为浏览器JS加密参数，无法逆向生成。

---

## 二、风控签名算法（User-mnp）完整机制

### 签名公式
```
User-mnp = MD5(URL路径 + uuid + 时间戳 + 盐值)
```

### 三个风控参数说明

| 参数 | 作用 | 怎么生成 |
|------|------|----------|
| URL路径 | 请求的接口地址路径 | 固定值（从URL手动提取）|
| uuid | 随机标识符 | 随机生成，格式 `ca412182e5668a106054-数字` |
| 时间戳 | 当前时间 | 代码自动获取当前毫秒时间 |
| 盐值 | 固定盐值（密钥） | 京东前端写死的，从JS代码中提取 |

### 通俗解释
就像你去银行办业务，工作人员要核对你的身份证号+当前时间+一个暗号，组合在一起做一个加密处理，生成一个"防伪码"。京东服务端收到请求后用同样的算法验证这个码对不对，对的话才给你数据。

### uuid方案调查结论
- **HTML页面、jmsdkPC.min.js、sgm-web-4.2.0.js、commons.js** 里都没有完整的uuid生成算法
- **vendors.js** 有引用但生成逻辑不在此
- **结论**：uuid是前端SDK运行时动态生成（`window.$frontend_monitor.getUUID(url)`），无法从静态JS还原
- **当前方案**：保留"固定前缀+10位随机数"（已测试通过）
- **未来升级方向**（如果uuid被拦）：
  - selenium/playwright启动真实浏览器提取uuid（最准确）
  - 继续观察先用着
  - 逆向JM SDK追踪uuid算法

---

## 三、签名验证失败排查指南

### 排查步骤

| 步骤 | 检查项 | 排查方法 | 解决方式 |
|------|--------|----------|----------|
| 1 | 时间戳是否新鲜 | 看日志中的 `User-mup` 值是不是当前时间（毫秒级） | 如果是几小时前的，说明时间戳生成逻辑坏了 |
| 2 | UUID格式 | 看日志中的 `uuid`，应该是 `ca412182e5668a106054-数字` 格式 | 如果格式不对，检查 `UUID_PREFIX` 常量 |
| 3 | URL路径处理 | 手动算一下MD5，看和 `User-mnp` 是不是一样 | 用下面的Python代码手动算 |
| 4 | 盐值是否正确 | 检查 `SIGN_SALT` 常量 | 直接对比代码常量 |

### 手动验证MD5签名的Python代码

```python
import hashlib

# 复制日志中的3个值
url_path = "/szpaas/szajax/shop/source/offlineFlowSource/downSkuTable.ajax"  # ← 从URL手动提取
uuid_str = "ca412182e5668a106054-1234567890"  # ← 从日志复制
timestamp = "1785828813987"  # ← 从日志复制
salt = "372ad2c2b6"

# 拼接 + MD5
sign_str = f"{url_path}{uuid_str}{timestamp}{salt}"
expected = hashlib.md5(sign_str.encode("utf-8")).hexdigest()
print(f"期望签名: {expected}")
# 和日志中的 User-mnp 对比，如果不一致就是拼接顺序错了
```

### 拼接顺序可能出错的情况
```
顺序1（当前用的，多数情况）：url_path + uuid + timestamp + salt
顺序2（备用）：uuid + url_path + timestamp + salt
顺序3（备用）：url_path + timestamp + uuid + salt
```

---

## 四、Cookie字段类型与风控标识（京东通用）

### Cookie字段分类

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

### Cookie使用方式
从 `config/cookie.txt` 整体读取，**不解析**，作为字符串整体塞入Cookie头。

### Cookie存放逻辑
- Cookie是京东识别你身份的凭证，相当于你的"登录通行证"
- 不直接写在代码里，而是放在单独的文件中，原因：
  1. Cookie会过期（几小时到几天），需要经常更新，放文件里方便替换
  2. 避免Cookie混在代码里不好管理
  3. 上传GitHub时可以排除这个文件，保护账号安全

---

## 五、参数分析规范（每个API接口必做）

### 接口参数完整分析维度
每一个API接口，必须完整分析输出：
- **静态固定参数**
- **全部动态参数**（动态参数范围包含：请求头、请求体、表单参数）
- **Cookie内部的动态/风控变量也必须识别标注**

### 参数类型标记规则
- ✅ 可动态修改（业务可调，如日期、渠道ID）
- ❌ 固定不可修改（写死在代码里）
- ⚠️ 鉴权会过期（Cookie、token、User-mnp）
- ⚠️ 风控参数（必须用算法生成）

### 字段类型说明（API清单用）
- **硬编码**：固定写死不变
- **动态参数**：每次查询要改
- **风控参数**：必须用算法生成
- **配置化**：有默认值但可改

---

## 六、代码集成规范（业务项目命名）

### 文件结构原则
1. **不拆分大量独立py文件**：所有项目接口代码集中在 `main.py` 一个文件
2. **按业务名称命名**：每个API类/方法按业务名命名（如 `商品搜索效果`、`商品推荐效果`）
3. **基类复用**：通用能力封装在 `JDBaseRequest` 基类中（Cookie管理、风控签名、间隔、重试、UA切换、日志、Excel保存）
4. **业务分发器**：`run_business(business_name, ...)` 统一调度各业务

### 代码结构
```
main.py
├── class JDBaseRequest（通用基类）
├── class ShopSourceAPI(JDBaseRequest)（业务API）
├── def run_business(business_name, ...)(业务分发器)
└── def main()（主程序入口）
```

### 配置集中管理
- 业务可修改参数集中在 `config/config.xlsx`
- 不可修改的固定参数作为类常量写在代码里
- Cookie、token等鉴权参数单独文件管理（`config/cookie.txt`）

---

## 七、30秒间隔控制 + 重试机制

### 为什么需要30秒间隔
- 不遵守立即触发风控拦截
- 间隔可在 `config.xlsx` 中调整

### 重试机制
- 最多重试3次
- 递增等待（30秒→60秒→90秒）
- Edge失败自动切换Chrome再试

### 双UA切换机制
- 默认Edge（和用户抓包一致）
- 失败时自动切Chrome
- 不要禁用此功能（兜底方案）

---

# 【商智 szgateway.jd.com 模块】（店铺经营数据分析）

> 已完成项目归档，按时间顺序保留所有历史上下文。

## 项目1：商品搜索效果（已完成 2026-08-04）

### 项目业务说明
- **业务名**：`商品搜索效果`
- **接口**：`https://szgateway.jd.com/szpaas/szajax/shop/source/offlineFlowSource/downSkuTable.ajax`
- **功能**：导出店铺来源中"搜索"渠道的SKU维度流量数据
- **输出**：Excel文件（按入店浏览量降序，最多5000条SKU）

### 当前实现文件（重构后）
- `main.py` - **集中所有代码**（基类 + 商品搜索效果 + 商品推荐效果 + 业务分发器 + 主入口）
- `config/config.xlsx` - 项目配置（全局+商品搜索效果参数）
- `docs/API 实现逻辑说明.md` - 接口实现逻辑文档
- `create_config_xlsx.py` - 配置文件生成脚本

> 注：远程GitHub上有过一次重构（commit bc331c7），把 `jd_api/base.py` 和 `jd_api/shop_source.py` 合并到 `main.py`。

### 返回数据说明
- 文件格式：Excel (.xlsx)
- 工作表名：`店铺来源_商品效果`
- 数据量：45条SKU记录 × 68列字段
- 主要字段包括：
  - 商品名称、商品SKU
  - 访客数、访客数环比、访客数占比
  - 浏览量、浏览量环比、人均浏览量
  - 平均停留时长、UV价值
  - 加购客户数、加购转化率
  - 成交客户数、成交转化率、成交金额
  - 客单价、下单客户数、下单金额
  - SPU维度数据（加购/成交/下单等）
  - 店铺关注人数
  - 新老客户数据（整体/180天/两年）

---

## 项目2：商品推荐效果（已完成 2026-08-04）

### 项目业务说明
- **业务名**：`商品推荐效果`
- **接口**：**与商品搜索效果完全相同**
- **关键差异**：只有 `lastSrcChannelId2=2009`（搜索是2008）
- **功能**：导出店铺来源中"推荐"渠道的SKU维度流量数据
- **输出**：Excel文件（按入店浏览量降序，最多5000条SKU）

### 关键发现（渠道ID理解）
- `lastSrcChannelId1=2` 是一级渠道（搜索/推荐都是2）
- `lastSrcChannelId2` 是二级渠道：
  - `2008` = 搜索子来源 → 商品搜索效果
  - `2009` = 推荐子来源 → 商品推荐效果
- 两个业务用同一个接口 + 同一个URL，只有 `lastSrcChannelId2` 一个字段不同

### 改动重点
1. **main.py - ShopSourceAPI 类重构**：
   - 新增 `CHANNEL_MAP = {"搜索": "2008", "推荐": "2009"}`
   - 删除原来的 `LAST_SRC_CHANNEL_ID2` 类常量
   - 新增 `_get_channel_id2(channel)` 方法
   - `download_sku()` 中的 `lastSrcChannelId2` 改为从映射获取
   - 新增 `download_recommend_sku()` 便捷方法
2. **main.py - 业务分发器**：`run_business` 的 `factory` 中新增 `"商品推荐效果"` 映射
3. **main.py - main() 函数**：默认 `business_name = "商品推荐效果"`
4. **配置文件**：未改动（复用商品搜索效果的 date/startDate/endDate 配置）

### 踩坑要点
1. **渠道ID混淆**：2008=搜索、2009=推荐，二级渠道区分实际来源
2. **复用配置的判断**：当时考虑过是否给商品推荐效果单独加日期配置
   - 用户回答："配置文件复用商品搜索效果的date配置"
3. **集成vs新建的判断**：当时考虑过新建独立的类
   - 用户回答："只加渠道映射，复用现有代码（推荐）"

### 测试结果
```
[INFO] 下载店铺来源数据: 日期=2026-08-03, 渠道=推荐(id2=2009)
[INFO] 请求成功: HTTP 200, 7661字节
[INFO] Excel已保存: output/推荐流量_2026-08-03.xlsx (7661字节)
✅ 导出成功！
```

---

## 项目3：商品购物车效果（已完成 2026-08-04）

### 项目业务说明
- **业务名**：`商品购物车效果`
- **接口**：**与商品搜索效果/推荐效果完全相同**
- **关键差异**：除了 `lastSrcChannelId2=3001`，**uuid前缀也不同**（5f9cc2ca20cad3d11642）
- **功能**：导出购物车来源的SKU维度流量数据
- **输出**：Excel文件（按入店浏览量降序，最多5000条SKU）

### 关键发现（uuid前缀差异）
| 业务 | lastSrcChannelId2 | uuid前缀 |
|------|-------------------|----------|
| 商品搜索效果 | 2008 | `ca412182e5668a106054` |
| 商品推荐效果 | 2009 | `ca412182e5668a106054` |
| **商品购物车效果** | **3001** | **`5f9cc2ca20cad3d11642`** ⭐ |

⚠️ **重要发现**：uuid前缀在不同业务/页面可能不一样，**必须按渠道可配置**，不能全局写死。

### 改动重点
1. **main.py - ShopSourceAPI 类重构**：
   - `CHANNEL_MAP` 改为 `(channel_id2, uuid_prefix)` 元组
   - 删除 `_get_channel_id2()`，改用 `_get_channel_config()` 返回元组
   - 新增 `_get_uuid_for_channel(channel)` 按渠道生成uuid
   - 基类 `_gen_risk_params(url, uuid_prefix=None)` 增加 uuid_prefix 参数
   - 基类 `request(url, data, ..., uuid_prefix=None)` 透传 uuid_prefix
2. **main.py - 新增便捷方法 `download_cart_sku()`**
3. **main.py - 业务分发器**：注册 `"商品购物车效果"` 映射
4. **main.py - main() 函数**：默认 `business_name = "商品购物车效果"`，日期改为 `2026-08-04`
5. **配置文件**：未改动（复用商品搜索效果的 date/startDate/endDate 配置 + 现有签名盐值 372ad2c2b6）

### 踩坑要点
1. **uuid前缀不一致**：购物车(5f9cc2ca)与搜索/推荐(ca412182)不同
   - 解决：CHANNEL_MAP每项绑定独立uuid前缀，按渠道动态选择
2. **基类签名方法扩展**：`_gen_risk_params()` 增加可选参数 `uuid_prefix`
   - 不破坏向后兼容（旧调用方式仍可用类常量UUID_PREFIX）
3. **元组解包**：`channel_id2, uuid_prefix = self._get_channel_config(channel)`

### 测试结果
```
[INFO] 下载店铺来源数据: 日期=2026-08-04, 渠道=购物车(id2=3001, uuid_prefix=5f9cc2ca...)
[INFO] 请求成功: HTTP 200, 7651字节
[INFO] Excel已保存: output/购物车流量_2026-08-04.xlsx (7651字节)
✅ 导出成功！
```

---

### 新增业务接入规范（v2.0 业务注册中心，2026-08-05）

后续新增业务（店铺来源报表/订单明细/售后订单/京准通广告报表）统一按以下步骤接入，**禁止大改调度核心**：
1. 定义业务API类（继承 `JDBaseRequest`，复用 Cookie/签名/间隔/重试/UA切换/日志/Excel保存）
2. 可变业务参数从 config.xlsx【全局配置】读取（`_get_business_params()` 兜底+警告）；固定常量须经用户确认后方可固化为代码常量（如 FIXED_BIZ_PARAMS），禁止私自写死
3. 在 `BUSINESS_REGISTRY` 注册：`"业务key": {"api_class": 类, "method": "方法名", "desc": 描述, "params": {...}}`
4. 调用：命令行 `python main.py --biz_key "业务key" --date "2026-07-29"`；代码内 `run_business("业务key", date=...)`
5. 新渠道注意 uuid 前缀差异（商智各页面可能不同，按 CHANNEL_MAP 元组维护）

### 配置一致性核对要点（v2.0）
- `config_consistency_check()` 启动自动跑，输出 ✅/❌/⚠️ 报告
- 日期硬编码用 **AST 扫描**（自动跳过 docstring/epilog 示例日期，只查真实赋值/传参/默认值）
- 间隔控制：`request()` 内必须调用 `_wait_interval()`；`_last_request_time` 为类属性，批量跨实例共享，严禁跳过

---

# 【京麦 seller-v10.shop.jd.com 模块（订单/售后）】

> 📝 暂无项目开发记录，待后续添加。

---

# 【京准通 jzt.jd.com 模块（广告报表）】

> 📝 暂无项目开发记录，待后续添加。

---

## 八、迭代更新记录（时间倒序）

| 日期 | 改动概要 |
|------|----------|
| 2026-08-04 | 初始化项目，破解京东风控签名，API测试成功，创建规范文档 |
| 2026-08-04 | 创建项目API清单Excel，含5个工作表共115条记录 |
| 2026-08-04 | 创建API实现逻辑说明文档 + 通用请求基类base.py（含30秒间隔/重试/日志），测试通过 |
| 2026-08-04 | 加入Edge/Chrome双UA切换机制 + 编写shop_source.py接口实现 + main.py主程序入口，测试通过 |
| 2026-08-04 | uuid方案调查：通过API获取不可行，京东用运行时SDK动态生成uuid，无法从静态JS还原。保留"固定前缀+随机数"方案并在代码中加详细说明 |
| 2026-08-04 | 重构：所有jd_api/*.py代码合并到main.py（GitHub commit bc331c7） |
| 2026-08-04 | 新增商品推荐效果业务（复用downSkuTable.ajax接口），复用搜索的日期配置 |
| 2026-08-04 | Agent体系改造：精简agents.md + 部署3个Skill（jd-api-analyze/git-safe-operate/python-code-gen） |
| 2026-08-04 | 旧agents.md业务内容精准拆分：备份为agents_old_backup.md，业务SOP全部迁移至jd-api-analyze SKILL |
| 2026-08-04 | Skill管理规则升级：京东全系业务统一复用jd-api-analyze/SKILL.md，按业务分区隔离（商智/京麦/京准通） |
| 2026-08-04 | **新增商品购物车效果项目**：CHANNEL_MAP改为元组支持按渠道uuid前缀（购物车3001用5f9cc2ca20cad3d11642，搜索/推荐用ca412182e5668a106054）|
| 2026-08-04 | **CHANNEL_MAP增强**：增加反向索引 `_CHANNEL_ID_INDEX`，支持直接传channel_id2调用（向下兼容）；新增 `_resolve_channel_display_name()` 让channel_id2传入也能得到友好文件名 |
| 2026-08-04 | **文件夹命名注释**：为所有目录创建 `FOLDER_NAME.md` 说明（10个文件夹），根目录加 `FOLDERS.md` 总索引，便于快速理解项目结构 |
| 2026-08-05 | **整改v2.0（业务注册中心架构）**：① 新增 `BUSINESS_REGISTRY` 注册表 + `run_business(biz_key_or_list, **kwargs)` 统一调度（支持单/批量）；② `parse_args()` argparse 命令行调用 `python main.py --biz_key "xx" --date "2026-07-29"`；③ 移除全部业务参数硬编码（interval/limit/sortField等9项改走config.xlsx，代码仅留开发期兜底+警告）；④ 业务名"商品购物车效果"改为"商品自主访问效果"（3001，购物车/我的订单回流）；⑤ 新增 `config_consistency_check()` 配置一致性核对报告（✅/❌/⚠️，AST扫描自动跳过docstring/epilog示例日期）；⑥ `_last_request_time` 改类级共享，批量跨实例严格30秒间隔；⑦ 修复 `get_business_handler` 缺self实例化Bug；⑧ 3渠道批量测试3/3成功 |
| 2026-08-05 | **config优化+渠道执行规则调整**：① config.xlsx项目名"商品搜索效果"统一为"商品流量来源"，9项业务参数补齐【说明】列中文注释；② **3001渠道执行名改回"商品流量来源_购物车"**（自主访问流量与购物车数据口径重叠，统一以"购物车"命名执行）；③ 自主访问保留注册配置但 `enabled=False`，调度层（`_run_single_business`/`_run_business_batch`）过滤不执行，可随时改回True开启；④ 默认批量执行搜索/推荐/购物车3渠道，实测3/3成功；⑤ 业务上下文备份至 main_old_backup.py（仅存档） |
| 2026-08-05 | **config精简+固定参数固化**：6项固定业务参数（lastSrcChannelId1/groupType/attributes/sortField/sortType/compareType）经用户确认移出config.xlsx，固化为代码常量 `FIXED_BIZ_PARAMS`（不读config、不警告）；config仅保留可变参数 interval/dateType/limit（`VARIABLE_BIZ_PARAMS`，缺省兜底+警告）及日期/全局配置；`config_consistency_check` 3.2同步更新，实测参数组装完整无警告 |
| 2026-08-05 | **Excel后置处理（商品流量来源落地+通用工具预留）**：① 新增2个公共工具函数 `convert_date_format()`（日期统一转 yyyy/m/d 不补零，支持8位纯数字/横杠/斜杠±时间，失败返原值）和 `safe_convert_numeric()`（全表数值转换，>15位纯数字保留文本防精度丢失，失败保留原值）；② 商品流量来源导出改为：读Excel→转换日期→首列A插入【日期】列→数值安全转换→写回；③ 日期列写入真实datetime并设 `yyyy/m/d`（带时间用 `yyyy/m/d hh:mm:ss`）单元格格式，打开不弹格式警告；④ 其余报表仅预留工具函数暂不改动；⑤ 实测购物车渠道输出验证通过（A列datetime+数值转换正确） |

---

## 九、归档规则（防止agents.md膨胀）

### ✅ 永久写入位置（按需加载）
- `.trae/skills/jd-api-analyze/SKILL.md` - 京东全系业务接口分析/规范/踩坑/历史
- `.trae/skills/git-safe-operate/SKILL.md` - Git安全操作规则
- `.trae/skills/python-code-gen/SKILL.md` - Python代码生成规范
- `docs/API实现逻辑说明.md` - 每个API接口的完整实现逻辑
- `agents_old_backup.md` - 旧版完整对话历史（仅人工查阅）

### ❌ 禁止写入
- 根目录 `agents.md` / `AGENTS.md`（只放全局底线规则，禁止堆业务）
- Skill.md禁止粘贴完整原始对话流水（只留精炼要点，原始流水归档到 agents_old_backup.md）