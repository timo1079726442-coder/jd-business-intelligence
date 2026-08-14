# 影刀 RPA 集成指南（项目19，2026-08-13）

> 适用版本：main.py ≥ b2ab6f3（项目18 AuthLoader 已入库）
> 目的：让影刀 RPA 全自动跑京东数据导出，包含多店铺切换、鉴权抓取、Python 业务调用

---

## 1. 整体架构

```
┌─────────────────────────────────────────────────────────────┐
│                  影刀 RPA 主流程                              │
│   ┌─────────────────────────────────────────────────┐      │
│   │  外层循环：遍历 config.xlsx「店铺账号」sheet     │      │
│   │  ┌──────────────────────────────────────┐      │      │
│   │  │  子流程 A：拟人登录 + 抓 3 域鉴权   │      │      │
│   │  │   - 拟人打开 3 个京东域名            │      │      │
│   │  │   - 拟人输入账号密码点登录          │      │      │
│   │  │   - 保存 3 个 Cookie JSON            │      │      │
│   │  │   - 抓 h5st 写 h5st.json             │      │      │
│   │  └──────────────────────────────────────┘      │      │
│   │  ┌──────────────────────────────────────┐      │      │
│   │  │  子流程 B：调用 Python 跑业务        │      │      │
│   │  │   - 调 module1.main(店铺名, 业务列表)│      │      │
│   │  │   - 监听退出码                        │      │      │
│   │  │   - 退出码=2 → 重抓鉴权              │      │      │
│   │  │   - 退出码=1 → 业务失败记录         │      │      │
│   │  │   - 退出码=0 → 成功                  │      │      │
│   │  └──────────────────────────────────────┘      │      │
│   │  循环到下一店铺                                  │      │
│   └─────────────────────────────────────────────────┘      │
└─────────────────────────────────────────────────────────────┘
                            ↓
            ┌───────────────────────────────┐
            │  Python 端（项目18 + 项目19）  │
            │   - auth_loader 读 JSON 鉴权  │
            │   - main.py 跑业务            │
            │   - 返回退出码给影刀          │
            └───────────────────────────────┘
```

---

## 2. 账号密码文件：config.xlsx

**位置**：`D:\CODE\trae\traespace\FYA箱包旗舰店\config\config.xlsx`
**Sheet 名**：`店铺账号`

**实际结构**（你截图读取的）：

| 店铺名 | 账号 | 密码 | 启用 |
|--------|------|------|------|
| FYA箱包旗舰店 | FYA8888 | 3.1415926 | 是 |
| Miyo箱包旗舰店 | miyo-周 | 3.1415926 | 是 |
| OTA箱包旗舰店 | ota8888 | 3.1415926 | 是 |

> ⚠️ 密码是**演示占位**，真实使用请**用 Excel 保护 + 不在日志打印**（AGENTS.md 风控规则第 4 条）

**影刀读取方式**：
1. 打开 Excel → 选中「店铺账号」sheet
2. 循环变量 `row_index` 从 2 到 4
3. 读 `第 row_index 行 第 A 列`（店铺名）→ 存 `shop_name`
4. 读 `第 row_index 行 第 B 列`（账号）→ 存 `account`
5. 读 `第 row_index 行 第 C 列`（密码）→ 存 `password`
6. 读 `第 row_index 行 第 D 列`（启用）→ 判断 = "是" 才执行

---

## 3. 影刀完整流程（按你截图的 1-29 步改造）

### 阶段 A：读取店铺账号（外层循环准备）

```
[步骤 1] 打开 Excel
         路径: D:\CODE\trae\traespace\FYA箱包旗舰店\config\config.xlsx
         保存对象: excel_instance

[步骤 2] 读取店铺总数
         读「店铺账号」sheet 第 1 列，排除表头，得到店铺数 N=3

[步骤 3] 设置外层循环变量 row_index = 2（从第 2 行开始，跳过表头）
```

### 阶段 B：读单店铺账号密码 + 拟人登录（外层循环体）

```
[步骤 4] 读店铺名（A 列）
[步骤 5] 读账号（B 列）→ 存变量 account
[步骤 6] 读密码（C 列）→ 存变量 password
[步骤 7] 读启用（D 列），判断 = "是" 才继续，否则跳到 [步骤 30]

[步骤 8] 关闭 Excel（密码已读完）

[步骤 9] 打开网页：https://passport.jd.com/new/login.aspx
         保存对象: web_page_login

[步骤 10] 等待 1.5 秒（拟人延迟）
[步骤 11] 拟人输入账号（输入框 → 填 account）
[步骤 12] 等待 0.8-1.2 秒（拟人延迟）
[步骤 13] 拟人输入密码（密码框 → 填 password）
[步骤 14] 等待 0.5 秒
[步骤 15] 拟人点击「登录」按钮
[步骤 16] 等待 2.5 秒（登录提交 + 跳转）

[步骤 17] 保存 Cookie 到「店铺名」目录
          文件: D:\CODE\trae\traespace\FYA箱包旗舰店\config\店铺名\jm_cookie.json
          格式: DevTools 导出格式（cookies[] 含 name/value/domain/path/expires）
```

### 阶段 C：触发 h5st 抓取（拟人点击业务页面）

```
[步骤 18] 跳转新网址：https://shop.jd.com/jdm/trade/tools/export/ExprotList
         保存对象: web_page_h5st

[步骤 19] 等待 1.5 秒
[步骤 20] 拟人点击「订单导出」按钮（触发 createdExportTask 请求）
[步骤 21] 等待 1.5 秒

[步骤 22] 获取网络监听结果（CDP 抓包）
         找到 POST 请求 URL 含 "createdExportTask" 的那条
         读请求头 "h5st" 字段

[步骤 23] 保存 h5st 到 h5st.json
         文件: D:\CODE\trae\traespace\FYA箱包旗舰店\config\店铺名\h5st.json
         格式: {"h5st": "<值>", "captured_at": <毫秒时间戳>, "ua": "<UA>", "biz_domain": "shop.jd.com"}
         ⚠️ 关键：captured_at 必须是毫秒时间戳（13 位），auth_loader 用它判断 30 分钟过期
```

### 阶段 D：调用 Python 跑业务（核心）

```
[步骤 24] 调模块 module1 的 main 方法
         参数: shop_name=店铺名, biz_keys=本店铺要跑的业务列表, date=2026-08-13
         例: python module1.py --shop "FYA箱包旗舰店" --biz_keys "商品流量来源_搜索,京麦售后明细_完整一键导出" --date "2026-08-13"
```

**退出码约定**（**关键**）：

| 退出码 | 含义 | 影刀动作 |
|--------|------|----------|
| **0** | 全部成功 | 跳 [步骤 28]（关闭 web_page + 记录成功） |
| **1** | 业务失败（数据问题） | 跳 [步骤 27]（记录失败原因，跳到下一业务，不重抓） |
| **2** | **鉴权过期**（Cookie/h5st） | 跳 [步骤 9]（重新拟人登录，重抓 3 域鉴权） |
| **3** | 系统错误 | 跳 [步骤 27]（记录系统错，跳到下一业务） |

```
[步骤 25] 判断退出码
         退出码 = 0 → [步骤 28]
         退出码 = 2 → [步骤 9]（重抓鉴权）
         退出码 = 1 或 3 → [步骤 27]（记录失败）
```

### 阶段 E：循环 + 关闭

```
[步骤 26] 关闭 web_page_h5st（释放浏览器）

[步骤 27] 写导出历史到「店铺账号」sheet 第 E 列
         成功 → "✅ 成功"
         业务失败 → "❌ 业务失败: <error>"
         鉴权过期 → "🔄 鉴权过期，已重抓"
         系统错 → "💥 系统错: <error>"

[步骤 28] row_index = row_index + 1
[步骤 29] 如果 row_index <= 4，回到 [步骤 4]（外层循环）
         否则 → [步骤 30]

[步骤 30] End 主流程
```

### 阶段 F：可选 - 重抓鉴权子流程

如果你想用子流程组织"重新登录"逻辑（更清晰）：

```
[子流程 C1] 重新拟人登录（参数：shop_name, account, password）
[子流程 C2] 重新抓 Cookie（参数：shop_name）
[子流程 C3] 重新抓 h5st（参数：shop_name, biz_domain）
```

影刀在退出码=2 时跳到 C1。

---

## 4. JSON 格式（**关键**）

### 4.1 Cookie JSON（DevTools 导出格式）

影刀"保存 Cookie"动作输出的文件格式：

```json
{
  "url": "https://shop.jd.com/jdm/trade/orders/order-list",
  "cookies": [
    {
      "name": "shshshfpa",
      "value": "7c2dc6b6-7ebd-1317-4b9f-9080f62257a4-1713936809",
      "domain": ".jd.com",
      "path": "/",
      "secure": false,
      "httpOnly": false,
      "expires": 1821121158,
      "sessionCookie": false
    },
    {
      "name": "pin",
      "value": "FYA8888",
      "domain": ".jd.com",
      "path": "/",
      "secure": false,
      "httpOnly": false,
      "expires": -1,
      "sessionCookie": true
    }
  ]
}
```

> 关键字段说明：
> - `expires`: unix 时间戳（秒），auth_loader 用 `now > expires` 判断过期
> - `sessionCookie: true`: 关闭浏览器即失效，auth_loader 跳过检查
> - `expires: -1`: 永不过期（很多京东字段都是）

### 4.2 h5st JSON（⚠️ 2026-08-14 改造：3 个独立文件）

影刀写文件的内容（**必须是这个格式**）：

```json
{
  "h5st": "20260814145001718;ijn5jin75aebjn54;...",
  "captured_at": 1786690196718,
  "ua": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) ... Chrome/144.0.0.0 Safari/537.36 Edg/144.0.0.0",
  "biz_domain": "shop.jd.com",
  "h5st_key": "jm_order"
}
```

> ⚠️ 关键：
> - `captured_at` 必须是**毫秒时间戳**（13 位），不是秒（10 位）
> - 30 分钟过期判断：`now * 1000 - captured_at > 30 * 60 * 1000`
> - `h5st_key` 标记 h5st 类型（`jm_order` / `jm_after_sale` / `jzt`）

### 4.2.1 ⚠️ 3 个独立 h5st 文件（2026-08-14 实测）

**关键发现**：同店同账号下，**不同业务页面的 h5st 不能跨业务复用**！

| 业务 | h5st 文件 | 抓取页面 | appId | 用于 |
|------|----------|---------|-------|------|
| **订单明细** | `h5st_jm_order.json` | [ExprotList?exportTaskType=0](https://shop.jd.com/jdm/trade/tools/export/ExprotList?exportTaskType=0) 点订单导出 | `CQLEJWPYPFOVQBC8UFLQ` | 项目14（订单明细5 个业务） |
| **售后明细** | `h5st_jm_after_sale.json` | [售后明细页](https://shop.jd.com/jdm/trade/after-sale/independent-after-sale/list?tabCode=all) 点售后导出 | `BHPQ4MHJBUOQZKTFTRNS` | 项目16（售后明细） |
| **京准通** | `h5st_jzt.json` | [jzt.jd.com](https://jzt.jd.com/home) 任意按钮 | - | 项目1（快车自定义） |

**实测结论**（2026-08-14）：
- ✅ 售后页 h5st 跑售后明细：成功（`code=200, msg='成功', data=True`）
- ❌ 售后页 h5st 跑订单明细：失败（`code=1001 未登录`）
- ✅ 订单页 h5st 跑售后明细：成功（同样 `code=200`）

**结论**：**每个 h5st 必须从对应业务页面触发抓取，不能跨业务复用**！

### 4.2.2 影刀端抓 3 个 h5st 的命令

```bash
# 1. 京麦订单 h5st（在订单导出页抓）
python auth_writer.py h5st --shop "FYA箱包旗舰店" \
  --value "<订单页抓的h5st>" \
  --biz_domain "shop.jd.com" \
  --h5st_key jm_order

# 2. 京麦售后 h5st（在售后明细页抓）
python auth_writer.py h5st --shop "FYA箱包旗舰店" \
  --value "<售后页抓的h5st>" \
  --biz_domain "shop.jd.com" \
  --h5st_key jm_after_sale

# 3. 京准通 h5st（在 jzt.jd.com 抓）
python auth_writer.py h5st --shop "FYA箱包旗舰店" \
  --value "<jzt页抓的h5st>" \
  --biz_domain "jzt.jd.com" \
  --h5st_key jzt
```

---

## 5. Python 端接口（module1.py）

影刀调用入口已封装在 `module1.py`（项目19 新增），调用方式：

```python
# module1.py 作为子进程被影刀启动
# 影刀命令行参数: python module1.py --shop "FYA箱包旗舰店" --biz_keys "A,B,C" --date "2026-08-13"

# 也可被 Python 直接 import 调用
from module1 import main
sys.exit(main(["--shop", "FYA箱包旗舰店", "--biz_keys", "A,B,C", "--date", "2026-08-13"]))
```

**退出码**（**重要**，影刀用它判断下一步）：

| 退出码 | 含义 | 触发条件 |
|--------|------|----------|
| 0 | 全部业务成功 | 跑完无异常 |
| 1 | 业务失败 | 数据问题（如 Excel 解析错、空数据、API 业务码非 0） |
| 2 | **鉴权过期** | `CookieExpiredError` / `H5stExpiredError` |
| 3 | 系统错误 | 未预期异常（Python bug、网络断） |

**stdout 输出**（影刀可读取判断）：
- 启动时打印 `[SHOP] FYA箱包旗舰店 / 业务: A,B,C / 日期: 2026-08-13`
- 业务失败时打印 `[FAIL] 业务 A: <错误信息>`
- 鉴权过期时打印 `[AUTH_EXPIRED] Cookie/h5st 过期，请重抓`

---

## 6. 鉴权文件目录结构

```
config/
├── _global/
├── FYA箱包旗舰店/                ← 店铺 1
│   ├── jm_cookie.json          ← 影刀 [步骤 17] 写
│   ├── jzt_cookie.json         ← 备用（京准通业务需要时写）
│   ├── sz_cookie.json          ← 备用（商智业务需要时写）
│   └── h5st.json               ← 影刀 [步骤 23] 写（带 captured_at）
├── Miyo箱包旗舰店/              ← 店铺 2
│   ├── jm_cookie.json
│   └── h5st.json
├── OTA箱包旗舰店/                ← 店铺 3
│   ├── jm_cookie.json
│   └── h5st.json
└── config.xlsx                 ← 账号密码
```

> `.gitignore` 已自动排除 `config/*/jm_cookie.json` 等 4 个 JSON 文件（项目18 已配）

---

## 7. 关键注意事项

### 7.1 退出码监听（**最易错**）

影刀在 [步骤 24] 用「调模块」动作时，**要勾选「捕获输出」+「返回退出码」**。如果用 subprocess 模式：
```bash
python module1.py --shop "X" --biz_keys "A,B" --date "2026-08-13"
echo $LASTEXITCODE
```

影刀的 IF 条件用 `$LASTEXITCODE` 或「进程退出码」变量判断。

### 7.2 鉴权过期自动重试

影刀应该**对同一店铺最多重试 3 次**（避免死循环）：

```
[步骤 24.5] 鉴权重试计数器 retry_count = 0
[步骤 25.1] 退出码=2 且 retry_count < 3 → 跳 [步骤 9]，retry_count += 1
[步骤 25.2] 退出码=2 且 retry_count >= 3 → 跳 [步骤 27]，记录"鉴权失败"
```

### 7.3 拟人延迟（防风控）

关键操作间必加延迟（**不能省**）：
- 输入框 → 密码框：0.8-1.2 秒
- 密码框 → 登录按钮：0.5 秒
- 登录 → 跳转：2.0-3.0 秒
- 跳转 → 抓 h5st：1.0-2.0 秒
- 抓 h5st → 关闭：0.5 秒

> 京东风控会识别"零延迟操作"是机器人。**延迟随机化更稳**（影刀支持随机延迟）。

### 7.4 错误兜底

如果影刀崩溃/断网：
- 重启后能从「店铺账号」sheet 第 E 列读上次状态
- 跳过「✅ 成功」的行，只跑「❌ 失败」/「空」
- 这是**幂等性**保证

---

## 8. 完整流程图（影刀视角）

```
┌──────────────────────────────────────────────────────────┐
│  开始                                                      │
└──────────────────────────────────────────────────────────┘
                          ↓
┌──────────────────────────────────────────────────────────┐
│  打开 config.xlsx「店铺账号」                              │
│  row_index = 2                                            │
│  N = 4（最大行）                                          │
└──────────────────────────────────────────────────────────┘
                          ↓
┌──────────────────────────────────────────────────────────┐
│  WHILE row_index < N:                                     │
│    ├── 读 shop_name, account, password, enabled         │
│    ├── IF enabled != "是": row_index++ ; CONTINUE        │
│    ├── 关闭 Excel                                         │
│    ├── [子流程 A: 拟人登录 + 抓 3 域鉴权]                 │
│    │   ├── 打开 https://passport.jd.com                  │
│    │   ├── 拟人输入 account + password                   │
│    │   ├── 拟人点击登录                                  │
│    │   ├── 保存 jm_cookie.json                           │
│    │   ├── 跳转 shop.jd.com 业务页面                     │
│    │   ├── 拟人点击「订单导出」（触发 createExportTask） │
│    │   ├── 抓请求头 h5st                                 │
│    │   └── 保存 h5st.json（含 captured_at 毫秒时间戳）   │
│    ├── retry_count = 0                                   │
│    ├── REPEAT:                                            │
│    │   ├── 调用 module1.main(shop_name, biz_keys, date) │
│    │   ├── 退出码 = 0 → BREAK（成功）                    │
│    │   ├── 退出码 = 2 且 retry_count < 3                 │
│    │   │   → retry_count++                               │
│    │   │   → [子流程 A: 重抓鉴权]                       │
│    │   ├── 退出码 = 2 且 retry_count >= 3               │
│    │   │   → 记录"鉴权失败"，BREAK                      │
│    │   ├── 退出码 = 1 或 3                               │
│    │   │   → 记录"业务失败/系统错"，BREAK              │
│    ├── 写导出历史到第 E 列                               │
│    ├── row_index++                                        │
└──────────────────────────────────────────────────────────┘
                          ↓
┌──────────────────────────────────────────────────────────┐
│  关闭 Excel                                               │
│  结束                                                      │
└──────────────────────────────────────────────────────────┘
```

---

## 9. 快速开始

### 9.1 影刀端最小配置（10 步可跑）

```
[1] 打开 config.xlsx → excel_instance
[2] 设置循环变量 row_index=2, max_row=4
[3] WHILE row_index <= max_row:
[4]   读 A 列 → shop_name，B 列 → account，C 列 → password
[5]   打开 https://passport.jd.com/new/login.aspx → web_page
[6]   拟人输入 account → 等待 1 秒 → 拟人输入 password → 等待 0.5 秒
[7]   拟人点击「登录」→ 等待 2.5 秒
[8]   跳转 https://shop.jd.com/jdm/trade/tools/export/ExprotList
[9]   拟人点击「订单导出」→ 等待 1.5 秒
[10]  CDP 抓 createExportTask 请求头 h5st → 写 h5st.json
[11]  调用 module1.main(shop_name, biz_keys, date) → 退出码
[12]  写导出历史到第 E 列
[13]  row_index++，回到 [3]
```

### 9.2 Python 端最小配置

```bash
# 安装依赖（首次）
pip install openpyxl requests

# 启用 AuthLoader
set AUTH_LOADER=1
set SHOP_ID=FYA箱包旗舰店

# 测试 module1 接口
python module1.py --shop "FYA箱包旗舰店" --biz_keys "京麦售后明细_完整一键导出" --date "2026-08-12" --h5st "manual_test"
echo %ERRORLEVEL%  # 应该输出 0/1/2/3 之一
```

---

## 10. 故障排查

| 现象 | 原因 | 修复 |
|------|------|------|
| 影刀调 module1 后所有业务都失败 | SHOP_ID 没设 | 影刀在调用前 `set SHOP_ID=店铺名` |
| 退出码=2 但重试 3 次都失败 | 影刀重抓没生效 | 检查 [子流程 A] 是否真把 JSON 写到了正确路径 |
| 退出码=0 但 Excel 报表没生成 | Python 业务正常但数据为空 | 看 [步骤 27] 写"✅ 成功"，查 `output/{业务}/{date}/` 目录 |
| 影刀卡在"拟人输入"步骤 | 元素定位不到 | 加"等待元素出现"+ 2 秒超时 |
| 影刀跑得很快（< 1 秒/业务） | 没加拟人延迟 | 必加 0.5-2.5 秒随机延迟，否则被风控 |

---

## 11. 相关文件

- [module1.py](file:///d:/CODE/trae/traespace/FYA箱包旗舰店/module1.py) - 影刀调用入口（项目19）
- [auth_loader.py](file:///d:/CODE/trae/traespace/FYA箱包旗舰店/auth_loader.py) - 鉴权加载器（项目18）
- [main.py](file:///d:/CODE/trae/traespace/FYA箱包旗舰店/main.py) - 业务主程序
- [auth_writer.py](file:///d:/CODE/trae/traespace/FYA箱包旗舰店/auth_writer.py) - JSON 写入辅助（项目19）
- [一键全部跑指南.md](file:///d:/CODE/trae/traespace/FYA箱包旗舰店/docs/一键全部跑指南.md) - 命令行参考

---

## 12. 更新记录

| 日期 | 改动 |
|------|------|
| 2026-08-13 | 创建文档，含多店铺循环、退出码约定、JSON 格式、完整流程图 |
