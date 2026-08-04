# API 实现逻辑说明

> 本文档详细说明每个API接口的实现逻辑，包括字段映射、动态参数来源、风控处理流程。
> **写代码前请先确认本文档无误，确认后再开始编码。**

---

## 接口1：店铺来源-搜索流量-SKU维度

### 一、接口基本信息

| 项目 | 内容 |
|------|------|
| 接口名称 | 店铺来源-搜索流量-SKU维度 |
| 请求方式 | POST |
| 接口URL | `https://szgateway.jd.com/szpaas/szajax/shop/source/offlineFlowSource/downSkuTable.ajax` |
| 域 | szgateway.jd.com |
| 返回格式 | Excel (.xlsx) |
| 单次返回上限 | 5000条SKU |
| 请求间隔 | ≥30秒（防风控） |
| 代码位置 | `jd_api/shop_source.py` → `class ShopSourceAPI` → `def download_search_sku()` |

---

### 二、字段映射表

#### 2.1 Header字段

| 序号 | 字段名 | 值 | 类型 | 来源 | 说明 |
|------|--------|-----|------|------|------|
| 1 | Cookie | shshshfpa=...; thor=... | 动态参数 | 读取 `config/cookie.txt` | 登录凭证，过期需更新 |
| 2 | User-Agent | Mozilla/5.0 ...Edge/144 | 硬编码 | 代码常量 `DEFAULT_UA` | 固定Edge浏览器UA |
| 3 | Referer | https://sz.jd.com/szweb/...flowPathDetailsNew.html | 硬编码 | 代码常量 `DEFAULT_REFERER` | 商智流量页面地址 |
| 4 | Origin | https://sz.jd.com | 硬编码 | 代码常量 `DEFAULT_ORIGIN` | 固定站点域名 |
| 5 | Content-Type | application/x-www-form-urlencoded | 硬编码 | 代码常量 | 表单提交格式 |
| 6 | Host | szgateway.jd.com | 自动生成 | requests库自动填充 | 无需手动设置 |
| 7 | Accept | text/html,... | 硬编码 | 代码常量 | 标准浏览器头 |
| 8 | Accept-Language | zh-CN,zh;q=0.9 | 硬编码 | 代码常量 | 中文环境 |
| 9 | sec-ch-ua | "Not(A:Brand";v="8"... | 硬编码 | 代码常量 | 浏览器品牌标识 |
| 10 | sec-ch-ua-mobile | ?0 | 硬编码 | 代码常量 | PC端 |
| 11 | sec-ch-ua-platform | "Windows" | 硬编码 | 代码常量 | Windows平台 |

#### 2.2 Body字段 - 业务参数

| 序号 | 字段名 | 示例值 | 类型 | 来源 | 说明 |
|------|--------|--------|------|------|------|
| 1 | date | 2026-08-03 | 动态参数 | 用户输入/函数参数 | 查询日期，格式 YYYY-MM-DD |
| 2 | startDate | 2026-08-03 | 动态参数 | 用户输入/函数参数 | 开始日期，单日查询与date相同 |
| 3 | endDate | 2026-08-03 | 动态参数 | 用户输入/函数参数 | 结束日期，单日查询与date相同 |
| 4 | interval | DAY | 配置化 | 代码默认值，可覆盖 | 时间粒度，DAY/WEEK/MONTH |
| 5 | dateType | day | 配置化 | 代码默认值，可覆盖 | 和interval配套 |
| 6 | lastSrcChannelId1 | 2 | 配置化 | 代码默认值，可覆盖 | ⭐一级渠道：2=搜索 |
| 7 | lastSrcChannelId2 | 2008 | 配置化 | 代码默认值，可覆盖 | ⭐二级渠道：搜索子来源 |
| 8 | groupType | skuId | 硬编码 | 代码常量 | 聚合维度：按SKU |
| 9 | attributes | skuId | 硬编码 | 代码常量 | 返回属性：SKU |
| 10 | sortField | jdr_sch_traffic_... | 配置化 | 代码默认值，可覆盖 | 排序字段：入店浏览量 |
| 11 | sortType | desc | 硬编码 | 代码常量 | 降序排列 |
| 12 | limit | 5000 | 配置化 | 代码默认值，可覆盖 | 返回上限 |
| 13 | compareType | hb | 硬编码 | 代码常量 | 环比对比 |

#### 2.3 Body字段 - 风控参数

| 序号 | 字段名 | 类型 | 来源 | 生成方式 |
|------|--------|------|------|----------|
| 1 | User-mup | 风控参数 | 代码动态生成 | `int(time.time() * 1000)` 当前毫秒时间戳 |
| 2 | User-mnp | 风控参数 | 代码动态生成 | `MD5(url_path + uuid + timestamp + "372ad2c2b6")` |
| 3 | uuid | 风控参数 | 代码动态生成 | `f"ca412182e5668a106054-{随机数字}"` |

---

### 三、动态参数来源说明

```
用户输入                    代码生成                    配置文件
┌──────────┐            ┌──────────────┐          ┌──────────────┐
│ date     │            │ User-mup     │          │ Cookie       │
│ startDate│            │ User-mnp     │          │ (cookie.txt) │
│ endDate  │            │ uuid         │          └──────────────┘
└──────────┘            └──────────────┘
```

| 参数来源 | 参数列表 | 说明 |
|----------|----------|------|
| **用户输入** | date, startDate, endDate | 每次调用时传入要查询的日期 |
| **代码生成** | User-mup, User-mnp, uuid | 每次请求自动生成，无需用户关心 |
| **配置文件** | Cookie | 从 config/cookie.txt 读取，过期时更新此文件 |
| **代码常量** | UA, Referer, Origin, groupType, sortType 等 | 固定写死，不变化 |
| **代码默认值(可覆盖)** | interval, dateType, lastSrcChannelId1/2, sortField, limit | 有默认值，调用时可传参覆盖 |

---

### 四、风控处理流程

#### 4.1 风控参数生成流程

```
步骤1: 获取当前时间戳
  timestamp = int(time.time() * 1000)
  示例: 1785826928129

步骤2: 生成随机uuid
  uuid = f"ca412182e5668a106054-{random.randint(1000000000, 9999999999)}"
  示例: ca412182e5668a106054-6319166040

步骤3: 提取URL路径
  url_path = "/szpaas/szajax/shop/source/offlineFlowSource/downSkuTable.ajax"
  (从完整URL中去掉域名部分)

步骤4: 拼接签名串
  sign_str = url_path + uuid + str(timestamp) + "372ad2c2b6"
  示例: /szpaas/szajax/shop/source/offlineFlowSource/downSkuTable.ajaxca412182e5668a106054-63191660401785826928129372ad2c2b6

步骤5: MD5加密
  user_mnp = hashlib.md5(sign_str.encode("utf-8")).hexdigest()
  示例: 1d9b71bbaa36e44014a168407273daa0

步骤6: 组装到请求体
  data["User-mup"] = str(timestamp)
  data["User-mnp"] = user_mnp
  data["uuid"] = uuid
```

#### 4.2 签名算法关键信息

| 项目 | 值 |
|------|-----|
| 算法 | MD5 |
| 拼接顺序 | URL路径 → uuid → 时间戳 → 盐值 |
| 盐值(密钥) | `372ad2c2b6` |
| 盐值来源 | 从京东前端 `commons-a5562705.js` 逆向获得 |
| URL路径处理规则 | 去掉 `http://` 或 `https://` 和域名，保留路径部分，去掉查询参数 |

#### 4.3 风控失败处理

| 错误码 | 错误信息 | 原因 | 处理方式 |
|--------|----------|------|----------|
| -407 | 不安全的请求 | User-mnp签名错误或缺失 | 检查签名算法是否正确 |
| -402 | 不安全的请求 | 缺少风控参数 | 确保三个风控参数都已传入 |
| 302 | 跳转登录页 | Cookie过期 | 更新 config/cookie.txt |
| 其他 | - | 网络或服务异常 | 重试机制（最多3次） |

---

### 五、请求完整流程

```
调用 download_search_sku(date="2026-08-03")
  │
  ├── 1. 读取Cookie (config/cookie.txt)
  │
  ├── 2. 组装业务参数 (date, startDate, endDate, channelId等)
  │
  ├── 3. 生成风控参数 (User-mup, User-mnp, uuid)
  │     ├── 获取当前时间戳
  │     ├── 生成随机uuid
  │     └── MD5签名计算
  │
  ├── 4. 组装请求头 + 请求体
  │
  ├── 5. 检查请求间隔 (距上次请求是否≥30秒)
  │     └── 不足30秒则自动等待补齐
  │
  ├── 6. 发送POST请求
  │
  ├── 7. 检查响应
  │     ├── 成功 → 保存Excel到 output/ 目录
  │     ├── 风控拦截 → 记录日志，重试（最多3次）
  │     └── Cookie过期 → 提示用户更新Cookie
  │
  ├── 8. 记录请求日志 (时间、参数、结果)
  │
  └── 9. 返回Excel文件路径
```

---

### 六、返回数据处理

| 项目 | 说明 |
|------|------|
| 返回格式 | Excel二进制流 (application/vnd.openxmlformats-officedocument.spreadsheetml.sheet) |
| 保存方式 | 直接写入 .xlsx 文件 |
| 文件命名 | `搜索流量_{date}.xlsx` |
| 保存目录 | `output/` |
| 数据结构 | 1行表头 + N行SKU数据，共68列字段 |
| 工作表名 | `店铺来源_商品效果` |

---

### 七、代码结构预览

```
jd_api/
├── __init__.py
├── base.py              # 通用请求基类
│   ├── class JDBaseRequest:
│   │   ├── __init__()           # 初始化：读Cookie、设UA等
│   │   ├── _read_cookie()       # 读取cookie.txt
│   │   ├── _gen_risk_params()   # 生成风控参数(User-mup/mnp/uuid)
│   │   ├── _wait_interval()     # 30秒间隔控制
│   │   ├── _request()           # 通用请求方法(含重试)
│   │   └── _log()               # 日志记录
│
└── shop_source.py       # 店铺来源API
    ├── class ShopSourceAPI(JDBaseRequest):
    │   ├── download_search_sku()        # 搜索流量-SKU维度
    │   ├── download_homepage_sku()      # 首页流量(后续)
    │   └── download_category_sku()      # 类目流量(后续)
```

---

## 待确认事项

1. ✅ 字段映射是否准确？
2. ✅ 风控参数生成逻辑是否正确？
3. ✅ 请求流程是否合理？
4. ✅ 代码结构是否满意？

**确认无误后，我将开始编写 `base.py` 和 `shop_source.py` 代码。**
