---
name: jd-api-analyze
description: 解析京东商智szgateway抓包POST/HTTP报文，接口拆解、风险识别、输出带详细注释的可运行python代码；同时承载本项目所有业务SOP、踩坑经验、迭代记录、接口规范、代码集成规范、归档规则、参数分析规范等知识沉淀
---

# 京东商智接口分析 + 项目知识沉淀 Skill

> 本文件按需加载，承担原 `agents.md` 中所有业务内容（接口规范、踩坑经验、迭代记录、代码集成规范、归档规则、参数分析规范等）。
> 加载本Skill的场景：分析京东商智抓包、生成/重构接口代码、查阅项目历史、确认接口字段含义、查阅踩坑经验、确认代码集成规范。

---

## 一、京东商智接口分析（粘贴抓包时执行）

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
表格清晰，不可省略风险警告；szgateway接口牢记 User‑mnp 为浏览器JS加密参数，无法逆向生成。

---

## 二、风控签名算法（User-mnp）完整机制

### 签名公式
```
User-mnp = MD5(URL路径 + uuid + 时间戳 + 盐值)
```

### 三个风控参数说明

| 参数 | 作用 | 怎么生成 |
|------|------|----------|
| URL路径 | 请求的接口地址路径 | 固定值 `/szpaas/szajax/shop/source/offlineFlowSource/downSkuTable.ajax` |
| uuid | 随机标识符 | 随机生成，格式 `ca412182e5668a106054-数字` |
| 时间戳 | 当前时间 | 代码自动获取当前毫秒时间 |
| `372ad2c2b6` | 固定盐值（密钥） | 京东前端写死的，从JS代码中提取 |

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
| 4 | 盐值是否正确 | 检查 `SIGN_SALT` 常量是不是 `372ad2c2b6` | 直接对比代码常量 |

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

如果按以上代码算出来的 `expected` 和日志中 `User-mnp` 不一致，尝试以下拼接顺序：

```
顺序1（当前用的，多数情况）：url_path + uuid + timestamp + salt
顺序2（备用）：uuid + url_path + timestamp + salt
顺序3（备用）：url_path + timestamp + uuid + salt
```

### 排查流程图

```
签名验证失败
  │
  ├─ Step1: 看日志中3个风控参数是否都存在？
  │    ├─ 缺少 → 检查 _gen_risk_params() 是否被调用
  │    └─ 都有 → 进入Step2
  │
  ├─ Step2: 手动计算MD5，对比日志中的User-mnp
  │    ├─ 一致 → 签名算法没问题，是其他原因（如Cookie过期）
  │    └─ 不一致 → 拼接顺序错了，尝试其他顺序
  │
  └─ Step3: 用排除法
       ├─ 换一个URL试试（如首页接口）
       ├─ 换一个盐值试试
       └─ 重新去京东商智页面抓一次最新的JS代码对比盐值
```

---

## 四、Cookie字段类型与风控标识（京东商智通用）

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

## 八、已完成项目归档（项目经验沉淀）

### 项目1：商品搜索效果（已完成 2026-08-04）

#### 项目业务说明
- 业务名：`商品搜索效果`
- 接口：`https://szgateway.jd.com/szpaas/szajax/shop/source/offlineFlowSource/downSkuTable.ajax`
- 功能：导出店铺来源中"搜索"渠道的SKU维度流量数据
- 输出：Excel文件（按入店浏览量降序，最多5000条SKU）

#### 当前实现文件（重构后）
- `main.py` - **集中所有代码**（基类 + 商品搜索效果 + 商品推荐效果 + 业务分发器 + 主入口）
- `config/config.xlsx` - 项目配置（全局+商品搜索效果参数）
- `docs/API 实现逻辑说明.md` - 接口实现逻辑文档
- `create_config_xlsx.py` - 配置文件生成脚本

> 注：远程GitHub上有过一次重构（commit bc331c7），把 `jd_api/base.py` 和 `jd_api/shop_source.py` 合并到 `main.py`。

#### 返回数据说明
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

### 项目2：商品推荐效果（已完成 2026-08-04）

#### 项目业务说明
- 业务名：`商品推荐效果`
- 接口：**与商品搜索效果完全相同**
- 关键差异：只有 `lastSrcChannelId2=2009`（搜索是2008）
- 功能：导出店铺来源中"推荐"渠道的SKU维度流量数据
- 输出：Excel文件（按入店浏览量降序，最多5000条SKU）

#### 关键发现（渠道ID理解）
- `lastSrcChannelId1=2` 是一级渠道（搜索/推荐都是2）
- `lastSrcChannelId2` 是二级渠道：
  - `2008` = 搜索子来源 → 商品搜索效果
  - `2009` = 推荐子来源 → 商品推荐效果
- 两个业务用同一个接口 + 同一个URL，只有 `lastSrcChannelId2` 一个字段不同

#### 改动内容
1. **main.py - ShopSourceAPI 类重构**：
   - 新增 `CHANNEL_MAP = {"搜索": "2008", "推荐": "2009"}`
   - 删除原来的 `LAST_SRC_CHANNEL_ID2` 类常量
   - 新增 `_get_channel_id2(channel)` 方法
   - `download_sku()` 中的 `lastSrcChannelId2` 改为从映射获取
   - 新增 `download_recommend_sku()` 便捷方法
2. **main.py - 业务分发器**：`run_business` 的 `factory` 中新增 `"商品推荐效果"` 映射
3. **main.py - main() 函数**：默认 `business_name = "商品推荐效果"`
4. **配置文件**：未改动（复用商品搜索效果的 date/startDate/endDate 配置）

#### 踩坑经验
1. **渠道ID混淆**：第一次看到2009以为是搜索的别称，容易和2008搞混
   - 正确理解：2008=搜索、2009=推荐，是搜索流量的两个不同子来源
2. **复用配置的判断**：当时考虑过是否给商品推荐效果单独加日期配置
   - 用户回答："配置文件复用商品搜索效果的date配置"
   - 这是个很好的设计：两个业务常配合使用，复用避免分散
3. **集成vs新建的判断**：当时考虑过新建独立的类
   - 用户回答："只加渠道映射，复用现有代码（推荐）"
   - 因为URL和所有参数都一样，强行新建类是过度设计

#### 测试结果
```
[INFO] 下载店铺来源数据: 日期=2026-08-03, 渠道=推荐(id2=2009)
[INFO] 请求成功: HTTP 200, 7661字节
[INFO] Excel已保存: output/推荐流量_2026-08-03.xlsx (7661字节)
✅ 导出成功！
```

---

## 九、迭代更新记录（时间倒序）

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

---

## 十、归档规则（防止agents.md膨胀）

### ✅ 永久写入位置（按需加载）
- `.trae/skills/jd-api-analyze/SKILL.md` - 接口分析/规范/踩坑/历史
- `.trae/skills/git-safe-operate/SKILL.md` - Git安全操作规则
- `.trae/skills/python-code-gen/SKILL.md` - Python代码生成规范
- `docs/API实现逻辑说明.md` - 每个API接口的完整实现逻辑
- `agents_old_backup.md` - 旧版全局agents.md（仅人工查阅）

### ❌ 禁止写入
- 根目录 `agents.md` / `AGENTS.md`（只放全局底线规则，禁止堆业务）