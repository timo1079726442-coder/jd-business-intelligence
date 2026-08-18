# 京准通快车自定义报表｜业务沉淀 Skill

> 本文件按需加载，覆盖京准通 jzt.jd.com 快车自定义报表业务。
> 上线时间：2026-08-07（阶段3 骨架 → 阶段10 完整三步流程）
> 代码位置：main.py `class JZTKuaicheAPI`（行 2547-3079）

---

## 一、接口总览（三步异步）

| # | 步骤 | 方法 | 请求方式 | URL | 说明 |
|---|---|---|---|---|---|
| 1 | 创建导出任务 | `create_export_task()` | POST | `/dataCenter/customreport/v2/report/add?requestFrom=0&businessFrom=1` | 返回 task_id（reportId）|
| 2 | 查询任务列表 | `get_task_list()` | GET | `/dataCenter/customreport/v2/report/list?requestFrom=0&businessFrom=1` | 返回 `data.data[]` 嵌套任务列表 |
| 3 | CDN 下载 | `download_report()` | GET | `/dataCenter/customreport/v2/report/downloadById?{id,name,startDay,endDay,pin,fileName,...}` | 返回 urlCsv，再 GET urlCsv 拿 CSV 流 |

**完整流程由 `run_full_export()` 一键封装**：
```python
# ⚠️ 2026-08-15 修订：京准通仅 Cookie 鉴权，不需要 h5st（list/add 均 HTTP 200）
api = JZTKuaicheAPI()
api.run_full_export(start_date="2026-08-07", end_date="2026-08-07")
# → 自动跑完：add → 轮询 → downloadById → urlCsv → 保存 output/京准通快车/*.csv
```

---

## 二、状态码对照表（关键修正：2026-08-07）

| jzt-api `subscribeState` | atoms-api `status` | 含义 | 处理 |
|---|---|---|---|
| **0** | **0** | 排队处理中（任务已提交未完成） | 继续轮询 |
| **2** | **2** | ✅ 报表生成完成，可以下载 | 调 downloadById |
| -1 | - | 失败（占位，待运行日志采集更多枚举） | 立即停 |

⚠️ **重要修正**：旧版代码 `SUBSCRIBE_STATE_OK=0` 是错误推测，0 实际是"排队中"。完成态是 `2`。

源码常量：
```python
SUBSCRIBE_STATE_QUEUED = 0   # 排队中
SUBSCRIBE_STATE_OK = 2       # ✅ 完成
SUBSCRIBE_STATE_FAIL = -1    # 失败（占位）
```

### 响应双字段判定（必须同时满足）
```python
success = ret.get("success", True)   # 缺省视为 True（兼容旧响应）
code = ret.get("code")               # 0 或 1 都视为成功
if success and code in (0, 1):
    return ret  # 成功
```

### list 响应双层 data 嵌套
```python
# ⚠️ 真实响应：list 数据是 data.data[]（双层 data 嵌套）
records = ret.get("data", {}).get("data", [])
```

---

## 三、Payload 模板说明

### 关键字段（main.py `JZT_KUAICHE_PAYLOAD_TEMPLATE`）

| 字段 | 类型 | 说明 | 风险点 |
|---|---|---|---|
| `checkSum` | int | **1114112**（硬编码，京东前端 JS 动态计算） | ⚠️ 京东若更新算法，需用 `page.evaluate()` 提取真实值 |
| `startTime` / `endTime` | int | **毫秒时间戳**（北京时区 00:00:00） | 不是日期字符串！转换：`int(dt.timestamp() * 1000)` |
| `startTimeStr` / `endTimeStr` | str | YYYY-MM-DD（用于展示） | — |
| `tempName` / `reportName` | str | 报表名（**1-30 字符**） | 重复会报错！加 HHMM 后缀确保唯一 |
| `version` | str | "JZT_V9" | 固定 |
| `daily` | int | 1=日报 | 固定 |

### 店铺 ID 硬编码（P4 待办）

⚠️ 当前 payload 模板写死勾选账号 `FYA8888（99936530475）`。

**换店铺需修改的位置**：
```
main.py → JZT_KUAICHE_PAYLOAD_TEMPLATE["customDimensionOptions"][0]["options"][0]["options"]
→ 把对应店铺的 {"checked": True, ...} 改成 True，其他改 False
```

**示例**：
```python
{"checked": True, "desc": "FYA8888", "flag": True, "hidden": False, "key": "99936530475", "value": "FYA8888"},
# ↑ 改成新店铺的 key + value，并把 checked 设为 True，其他账号设为 False
```

**中期优化**：入参 `account_pin_list`，运行时动态勾选（暂未实现，需你确认 jzt-api 是否会拒绝动态拼接）。

---

## 四、鉴权要点（与商智/京麦关键差异）

| 维度 | 商智 | **京准通（本项目）** |
|---|---|---|
| 鉴权字段 | User-mnp（MD5） | **仅 Cookie**（不需要 h5st，2026-08-15 实测：list/add 均 HTTP 200）|
| Cookie 文件 | `config/cookie.txt` | `config/jzt_cookie.txt`（**不互通**） |
| 抓包 URL | sz.jd.com 任意页 | **必须** jzt.jd.com 域 |
| UA 切换 | Edge ↔ Chrome 自动 | **固定 UA**（京准通无 h5st 绑定约束，保留习惯） |
| wlfstk_smdl | 可选 | 可选（与项目4-6一致，缺失仅警告） |

### Cookie 抓取路径
```
浏览器登录 https://jzt.jd.com/home
→ F12 → Network → 任意请求 → Request Headers → Cookie
→ 整段复制写入 config/jzt_cookie.txt
```

⚠️ **不需要抓 h5st**（2026-08-15 实测反证）：京准通 jzt-api 的 list/add 接口不带 h5st 均返回 HTTP 200（add 返 code=400 参数错误，非登录/风控错误），仅 Cookie 即可跑通三步链路。曾误判为"必须 h5st"，已在 main.py / auth_loader.py / auth_writer.py 全量移除 jzt h5st 支持。

---

## 五、OSS 下载链路（阶段8-9 实证）

```
downloadById 返回 JSON { data: { urlCsv: "https://storage.jd.com/...?Expires=..." } }
↓ GET urlCsv
首次可能返回 404 NoSuchKey（OSS 预热延迟 ~10s）
↓ 重试循环（MAX_DOWNLOAD_RETRY=3）
   ├─ 随机退避 3-10 秒
   ├─ 重新调 downloadById 拿新 urlCsv（**禁止缓存旧 urlCsv**）
   └─ 再 GET urlCsv
↓ 200 成功 → 保存到 output/京准通快车/*.csv
```

⚠️ **关键约束**：urlCsv 是一次性 OSS 预签名链接，每次失败必须**重新调 downloadById**，不能用旧 urlCsv。

---

## 六、容错参数（用户决策 2026-08-07）

| 常量 | 值 | 含义 |
|---|---|---|
| `POLL_INTERVAL` | 3 | 轮询间隔（秒） |
| `MAX_POLL_TIMES` | 15 | 轮询最大次数（45秒超时）|
| `MAX_DOWNLOAD_RETRY` | 3 | CDN 404 重试最大次数 |

⚠️ 大报表可能超过 45 秒超时；运行观察后考虑调整 `MAX_POLL_TIMES`。

---

## 七、调度器接入（阶段10）

### 注册方式
```python
"京准通快车自定义报表": {
    "api_class": JZTKuaicheAPI,            # 兼容老调用
    "method": "run_full_export",            # 实例化后也支持直调
    "callable": _run_jzt_kuaiche_full,      # ⭐ 调度器优先使用
    ...
}
```

### 调度器两种注册机制
1. **标准方式**：`api_class + method`（无参 __init__ → 实例化 → 调方法）
2. **callable 方式**：自定义函数（用于特殊业务需要传额外构造参数）

`get_business_handler()` 检测 `info.get("callable")` 优先返回自定义函数。

### 推荐调用方式
```python
from main import run_business

run_business(
    "京准通快车自定义报表",
    h5st="2026-08-07 从F12抓的h5st值",
    start_date="2026-08-06",
    end_date="2026-08-06",
    cookie_path="config/jzt_cookie.txt",   # 可选
)
```

---

## 八、checkSum 风险预案（P5 待办）

`checkSum=1114112` 是京东前端 JS 计算，京东改动算法会直接报参数错误。

**预案**（若 checkSum 失效）：
```python
# 用 Playwright 启动真实 Chrome，访问 jzt.jd.com 报表页面
from playwright.sync_api import sync_playwright
with sync_playwright() as p:
    browser = p.chromium.launch()
    page = browser.new_page()
    page.goto("https://jzt.jd.com/...")
    real_checksum = page.evaluate("window.ParamsSign.sign(JSON.stringify(payload))")
    print(f"真实 checkSum: {real_checksum}")
```

---

## 九、atoms-api 备选调试接口

⚠️ atoms-api 仅作**调试备选**，主链路继续使用 jzt-api。

| 字段 | jzt-api | atoms-api |
|---|---|---|
| 域 | jzt-api.jd.com | atoms-api.jd.com |
| 状态字段 | `subscribeState` (int) | `status` (int) |
| 业务头 | 基础 | 含 `loginMode`、`language` |
| 响应格式 | success+code 双字段 | 同左 |
| 已知状态 | 0=排队, 2=完成 | 同左 |

---

## 十、待办（与本任务清单同步）

| 优先级 | 项 | 状态 |
|---|---|---|
| P0 | 常量 `SUBSCRIBE_STATE_OK=2`（后修正为**纯探针策略**）| ✅ 已修正 |
| P0 | 封装完整三步流程 | ✅ 已加 `run_full_export` |
| P0 | OSS 404 重试 3-10s 随机退避 | ✅ 已实现 |
| P1 | 调度器接入 | ✅ 已加 `callable` 注册方式 |
| P2 | 本文档 | ✅ 已写 |
| P3 | jd_cdp_capture.py 抓包指引 | ✅ 已补（2026-08-09）|
| P4 | payload 店铺 ID 动态化 | 📋 暂不实现中期优化，仅文档标注 |
| P5 | checkSum 风险预案 | ✅ 预案已写（本SKILL.md 第八章），待 Playwright 实际触发验证 |
| P6 | 单元测试 | ✅ `_test_jzt_kuaiche.py` 10/10 通过（不入仓，.gitignore 排除）|
| - | 任务失败/取消 subscribeState 枚举采集 | ⏳ 待运行日志采集（已知 0/2 含义，-1 占位失败）|
| - | 大报表是否超时 | ⏳ 实际业务观察（MAX_POLL_TIMES=15×3s=45s）|
| - | 项目8响应判定 bug 修复 | ✅ str(code) 兼容字符串"1" |
| - | Excel 后置处理 | ✅ 日期/数值/格式三规则已对齐 |
| - | 项目9 全站营销单品计划上线 | ✅ 同步两步流程跑通 |
| - | Cookie 管理规则 | ✅ 已写入 AGENTS.md |

---

## 十一、变更记录

| 日期 | 阶段 | 改动概要 |
|---|---|---|
| 2026-08-07 | 阶段3 | 骨架上线，3 接口方法（add/list/downloadById），含 h5st 可选 |
| 2026-08-07 | 阶段4 | 容错配置（轮询+重试）+ Cookie 过期识别 |
| 2026-08-07 | 阶段6 | list 响应字段适配（双层 data.data[]）+ add 响应双格式兼容 |
| 2026-08-07 | 阶段8 | downloadById 链路打通（拿 urlCsv） |
| 2026-08-07 | 阶段9 | OSS 404 预热延迟重试（固定 3s）|
| 2026-08-07 | 阶段10 | **完整流程封装** `run_full_export` + 调度器 callable 接入 + SUBSCRIBE_STATE_OK=2 修正 + OSS 重试改 3-10s 随机退避 + 本文档 |
| 2026-08-09 | 阶段11 | **真实跑通 8/6 数据**——发现 subscribeState=0 长期不变但报表已生成，弃用状态字段改用**纯 downloadById 探针策略**；修复 OSS 404 重试「同 urlCsv 退避」不再重调 downloadById；`_random` → `random` |
| 2026-08-09 | 阶段12 | **Excel 后置处理接入**——新增 `_post_process_csv_to_xlsx`（复用 `prepare_date_columns` / `safe_convert_numeric` / `apply_column_formats`）；OUTPUT_SUBDIR 改为「京准通快车效果自定义」；扩展 `_col_matches` 支持「商品定向SKU ID」类含空格的列名；新增 `METRIC_BLOCKLIST` 防止「SKU金额」被误套 0 位小数格式 |
| 2026-08-09 | 阶段13 | **项目8 京准通快车订单效果明细上线**——新接口 `POST /reweb/msa/effect/order/download`（同步两步）；新增 `JZTKuaicheOrderEffectAPI`；输出 `output/京准通快车订单效果明细/{date}/`；修复 `str(code)` 兼容字符串"1"的隐藏 bug |
| 2026-08-09 | 阶段14 | **项目9 京准通全站营销单品计划上线**——新接口 `POST /reweb/swa/account/campaign/download`；新增 `JZTQuanZhanCampaignAPI`；字段类型严格匹配抓包（giftFlag 等用字符串""，campaignTypes 是列表）；输出 `output/京准通全站营销单品计划/{date}/` |
| 2026-08-09 | 阶段15 | **Cookie 管理规则写入 AGENTS.md**——默认保留本地 Cookie，仅 601/切账号/分享时才删 |
| 2026-08-09 | 阶段16 | **代码提交 GitHub**——commit `c678c16`（3业务+规则+Excel扩展）+ `3fac25c`（.gitignore 补 `_test_*.py` 排除）|
| 2026-08-10 | 阶段17 | **项目10/11/12（京准通全站营销）上线**——单品推广效果 `JZTQuanZhanEffectAPI` / 全店计划 `JZTQuanZhanCampaignAllStoreAPI` / 全店推广效果 `JZTQuanZhanEffectAllStoreAPI`；完整业务沉淀已归档至 `.trae/skills/jd-api-analyze/SKILL.md`（本快车专项文档不重复记录） |
| 2026-08-10 | 阶段18 | **p7 测试适配**——同步 `wait_for_task_ready` 纯 downloadById 探针策略（08-09 重构）断言：重算 session.get 消耗（轮1不探针，5轮=5 list+4 探针+1 后置 list）、新增探针失败工厂、修正保存名断言 `京准通快车效果自定义_{startTimeStr}.xlsx`；`tests/test_jzt_kuaiche_p7.py` 42/42 通过 |

---

## 十二、关联文件

- 主代码：[main.py](../main.py)（class JZTKuaicheAPI，行 2547-3079）
- 调度器入口：[main.py](../main.py)（`_run_jzt_kuaiche_full`，行 3196+）
- 注册表：[main.py](../main.py)（`BUSINESS_REGISTRY["京准通快车自定义报表"]`）
- 抓包脚本（待补京准通指引）：[jd_cdp_capture.py](../jd_cdp_capture.py)
- 全局规则：[AGENTS.md](../AGENTS.md)
- 全局踩坑：[全局复利的踩坑日志.md](../全局复利的踩坑日志.md)