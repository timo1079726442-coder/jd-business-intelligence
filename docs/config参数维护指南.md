# config.xlsx 可调整参数维护指南（2026-08-21 新增）

> 本文档是**轻量日常维护**手册，与 `docs/维护手册.md`（Cookie / h5st 重抓）互补。
>
> **适用场景**：你日常只需调整这 4 项参数，重抓鉴权请看维护手册。

---

## 一、这 4 项是什么？

所有可调整参数都在 **`config/config.xlsx`** 的「全局配置」sheet：

| 组 | key | 当前值 | 默认调整频率 | 用途 |
|----|-----|--------|-------------|------|
| 全局 | `shop_interval_seconds` | 60 | 几乎不动 | 多店串跑时每跑完一个店铺的等待秒数（防京东 IP 风控） |
| IMAP | `auth_code_last4` | `ihfj` | 90 天一次 | QQ 邮箱 IMAP 授权码的后 4 位掩码（一致性校验用） |
| IMAP | `auth_code_expire_date` | `2026-11-21` | 90 天一次 | IMAP 授权码建议过期日（<14 天 WARN，<7 天阻塞） |
| IMAP | `auth_code_note` | `2026-08-21 用户授权` | 90 天一次 | 授权码备注：变更时间 + 授权人 |

---

## 二、什么时候需要改？

### 2.1 `shop_interval_seconds`（店间冷却）

**触发场景**：跑三店串跑时收到京东 IP 风控告警（`code=601`、`code=403` 等）

**怎么改**：
1. 打开 `config/config.xlsx` → 「全局配置」sheet
2. 找 `shop_interval_seconds = 60` 这一行
3. 改 value（推荐 60~120 之间）
4. 保存 xlsx 即可生效（无需重启任何服务）

**示例值**：
- 60（默认，常规场景）
- 90（中等风控）
- 120（高风控，建议早晚高峰避让）

**优先级**：CLI 参数 `--shop-interval` 最高。例如临时跳过冷却：
```bash
python rpa_run.py --shop-interval 0
```

---

### 2.2 IMAP 三项（QQ 邮箱授权码）

**触发场景**：京麦订单明细报"IMAP 登录失败"或维护手册检查告警

**前置知识**：
- QQ 邮箱授权码**不是 QQ 密码**，是 16 位独立字符串
- 真实授权码存在 `config/imap_config.ini` 的 `[imap].auth_code`（**不入仓**，本地鉴权）
- xlsx 只存**后 4 位掩码 + 过期日 + 备注**，用于**一致性校验 + 过期提醒**

**完整更新流程**：

#### 步骤 1：登录 QQ 邮箱网页版
1. 浏览器打开 [https://mail.qq.com/](https://mail.qq.com/)
2. 登录 `1079726442@qq.com`

#### 步骤 2：生成新授权码
1. 顶部「设置」→ 「账户」
2. 滚到「POP3/IMAP/SMTP/Exchange/CardDAV/CalDAV服务」
3. 找到「IMAP/SMTP服务」→ 点「开启」→ 验证密保
4. 复制新生成的 **16 位授权码**

#### 步骤 3：写入 ini（真实鉴权，不入仓）
1. 打开 `config/imap_config.ini`
2. 找到 `[imap]` section
3. 把 `auth_code=...` 改成新值
4. 保存

#### 步骤 4：同步更新 xlsx（轻量元信息）
打开 `config/config.xlsx` →「全局配置」sheet，改 3 项：

| key | 新值示例 | 说明 |
|-----|----------|------|
| `auth_code_last4` | 新授权码的**后 4 位** | 一致性校验 |
| `auth_code_expire_date` | 今天 + 90 天 | 建议过期日 |
| `auth_code_note` | `2026-11-21 用户授权` | 变更备注 |

#### 步骤 5：验证
```bash
python imap_config_loader.py
```

预期输出：
```
📬 [IMAP 配置健康检查] ...
  邮箱账号：1079726442@qq.com
  授权码后4位：****XXXX（xlsx 登记：****XXXX）
  备注：2026-XX-XX 用户授权
  ✅ 授权码有效，剩余 90 天（到期 2026-XX-XX）
```

**`❌ 警告：xlsx 登记的后4位 (XXXX) 与 ini 实际值 (YYYY) 不一致`** → 说明步骤 3 和 4 没同步，重新检查。

---

## 三、刷新报表（统一入口，2026-08-21 新增）

跑全店全业务时，建议用 `refresh_reports.py`，它会自动按业务类型分流：

| 业务类型 | 行为 |
|---------|------|
| **区间类**（京麦订单/售后、京准通 6 业务、商品明细、商品流失、店铺来源_三级渠道） | **覆盖** — 按近 30 天区间重跑，xlsx 同名覆盖、DB Upsert |
| **逐日类**（商智 4 个流量来源、关键词分析 day 粒度） | **补齐** — xlsx ∪ DB 任一缺失则按日补，已有的跳过（防风控） |

### 3.1 常用命令

```bash
# 全店全业务（默认行为：区间覆盖 + 逐日补齐）
python refresh_reports.py

# 单店
python refresh_reports.py --shop "MIYO箱包旗舰店"

# 指定日期窗口
python refresh_reports.py --start_date 2026-07-22 --end_date 2026-08-20

# 指定业务
python refresh_reports.py --biz_keys "京麦订单明细_完整一键导出,商智关键词分析"

# 只跑 xlsx（不入库）
python refresh_reports.py --skip-db

# 只跑 DB（不下 xlsx）
python refresh_reports.py --skip-xlsx

# 干跑（只打印计划）
python refresh_reports.py --dry-run
```

### 3.2 与其他工具的区别

| 工具 | xlsx | DB | 区间/逐日 | 巡检缺失 |
|------|------|----|----------|---------|
| `rpa_run.py` | ✅ | ❌ | 一刀切 30 天 | ✅（仅 xlsx） |
| `daily_update.py` | ❌ | ✅ | ✅ 区分 | ❌（全跑） |
| `fill_missing.py` | ❌ | ✅ | ❌（默认逐日） | ✅（仅 DB） |
| **`refresh_reports.py`** | ✅ | ✅ | ✅ 区分 | ✅（xlsx + DB） |

**推荐**：
- 影刀定时跑（凌晨）→ 用 `daily_update.py`（DB 入库优先，xlsx 备份由影刀调 `refresh_reports.py --skip-db`）
- 手动一次跑全店全业务 → 用 `refresh_reports.py`
- 只想补某些缺失日期 → 用 `fill_missing.py --biz_keys "..."`

### 3.3 退出码（与影刀约定一致）

| 退出码 | 含义 |
|--------|------|
| 0 | 全部成功 |
| 1 | 部分业务失败 |
| 2 | 鉴权过期（Cookie / h5st） |
| 3 | 参数错误 |

---

## 四、修改 xlsx 时的注意事项

### 4.1 不要用 WPS / Excel 直接打开改
Windows 上 WPS / Excel 打开 xlsx 时会**创建锁文件**，导致 Python 程序下次写文件时报 `PermissionError`。

**推荐方式**：
- **查看**：用 WPS / Excel 打开（只读 OK）
- **修改**：建议用 Python 脚本（参考 `imap_config_loader.py` 的写法）
- 或关闭 WPS / Excel 后再用 WPS 改

### 4.2 改完后建议立刻跑一次验证
```bash
python imap_config_loader.py     # 验证 IMAP
python rpa_run.py --dry-run      # 验证全局配置（含 shop_interval_seconds）
```

### 4.3 不要误删 xlsx「需填入内容」列
2026-08-21 项目新增的「需填入内容」列**只是注释**，删除会导致：
- 失去配置说明
- 程序加载**不受影响**（所有加载器只读前几列）

但删了后下次维护就不知道哪项是动态值哪项是可调值了，**建议保留**。

---

## 五、完整参数清单（速查）

| key | 类型 | 默认 | 改不改 | 备注 |
|-----|------|------|--------|------|
| `shop_interval_seconds` | ✅ 可调 | 60 | 风控严时调大 | 详见 § 2.1 |
| `auth_code_last4` | ✅ 可调 | `ihfj` | 90 天换 IMAP 时同步 | 详见 § 2.2 |
| `auth_code_expire_date` | ✅ 可调 | `2026-11-21` | 90 天换 IMAP 时同步 | 详见 § 2.2 |
| `auth_code_note` | ✅ 可调 | `2026-08-21 用户授权` | 90 天换 IMAP 时同步 | 详见 § 2.2 |
| 其他 24 项 | 🔒 动态/固定 | — | **不要动** | 见 xlsx「需填入内容」列 |

---

## 六、关联文档

- **鉴权重抓**（Cookie / h5st）：`docs/维护手册.md`
- **配置文件位置**：`config/config.xlsx`（5 个 sheet，详见 xlsx「需填入内容」列）
- **业务实现**：`main.py`、`rpa_run.py`
- **报表刷新统一入口**：`refresh_reports.py`（详见 § 三）
- **健康检查工具**：`imap_config_loader.py`（CLI 跑一次即出报告）

---

**版本**：v1.0（2026-08-21 项目23 新增）
**下次更新触发**：新增可调整参数时（如增加 `--concurrency` 等 CLI 配置项）