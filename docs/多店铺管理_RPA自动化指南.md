# 多店铺管理 & RPA 自动抓取指南（项目17，2026-08-12）

## 1. 目标

让让可以通过 **影刀 RPA** 自动抓取 Cookie + h5st，按店铺 ID 写入 `config/{店铺ID}/` 目录，Python 端自动读取最新值并用于接口调用。

## 2. 目录结构

```
config/
├── FYA旗舰店/                  # 店铺A（pin=FYAxxxx）
│   ├── sz_cookie.txt          # 商智 Cookie（项目1/4/5/6/13）
│   ├── jzt_cookie.txt         # 京准通 Cookie（项目7-12）
│   ├── jm_cookie.txt          # 京麦 Cookie（项目14/16）
│   └── h5st.txt                # h5st（30 分钟过期）
├── FYA箱包旗舰店/              # 店铺B（pin=FYA8888，当前测试用）
│   ├── sz_cookie.txt
│   ├── jzt_cookie.txt
│   ├── jm_cookie.txt
│   └── h5st.txt
└── shops.db                   # SQLite 数据库（不入仓）
```

## 3. SQLite 数据库表结构

### 3.1 shops 表（店铺配置）

```sql
CREATE TABLE shops (
    shop_id          TEXT PRIMARY KEY,        -- 店铺标识（如 "FYA旗舰店"）
    pin              TEXT NOT NULL,           -- Cookie里的pin字段（验证用）
    shop_name        TEXT,                     -- 中文店名
    sz_cookie_path   TEXT,                     -- config/{shop_id}/sz_cookie.txt
    jzt_cookie_path  TEXT,                     -- config/{shop_id}/jzt_cookie.txt
    jm_cookie_path   TEXT,                     -- config/{shop_id}/jm_cookie.txt
    h5st_path        TEXT,                     -- config/{shop_id}/h5st.txt
    created_at       TEXT,
    updated_at       TEXT
);
```

### 3.2 h5st_log 表（h5st 抓取记录 + 过期检查）

```sql
CREATE TABLE h5st_log (
    shop_id          TEXT PRIMARY KEY,
    h5st_value       TEXT NOT NULL,
    captured_at      TEXT,                     -- 抓取时间
    expires_at       TEXT                      -- 30 分钟后过期
);
```

### 3.3 export_history 表（导出历史记录）

```sql
CREATE TABLE export_history (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    shop_id          TEXT NOT NULL,
    biz_key          TEXT NOT NULL,            -- 业务key（如"京麦售后明细_完整一键导出"）
    task_id          TEXT,                      -- 京东返回的taskId
    status           TEXT,                      -- success / fail / code=201
    file_path        TEXT,                      -- 落盘xlsx路径
    row_count        INTEGER,                   -- 数据行数
    created_at       TEXT
);
```

## 4. Python 端用法（已实现 config_manager.py）

```python
from config_manager import get_config_manager

cm = get_config_manager()

# 4.1 自动扫描 config/ 目录注册所有店铺（首次运行必须）
new_shops = cm.auto_register_from_disk()
print(f"新注册店铺: {new_shops}")

# 4.2 查所有店铺
for shop in cm.list_shops():
    print(f"{shop['shop_id']} (pin={shop['pin']})")

# 4.3 读取某店铺的 Cookie 路径
cookie_path = cm.get_cookie_path("FYA箱包旗舰店", "jm")

# 4.4 读取 h5st（自动检查 30 分钟过期）
h5st = cm.get_h5st("FYA箱包旗舰店")
if cm.is_h5st_expired("FYA箱包旗舰店"):
    print("⚠️ h5st 已过期，请让 RPA 重新抓取")

# 4.5 保存 RPA 抓到的 h5st（DB + 文件双写）
cm.save_h5st("FYA箱包旗舰店", "<h5st_value>")

# 4.6 保存 RPA 抓到的 Cookie
cm.save_cookie("FYA箱包旗舰店", "sz", "<cookie_string>")

# 4.7 记录导出历史
cm.log_export("FYA箱包旗舰店", "京麦售后明细_完整一键导出",
              status="success", task_id="105884780328",
              file_path="output/...", row_count=1)
```

## 5. RPA 端集成（影刀 RPA）

### 5.1 你已完成的步骤

```
1. 打开网页 → https://szjd.com/sz/view/dealAnalysis/...
2. 保存 Cookie → D:\tmp\cookie\xiazai.cookie.json
3. 关闭网页
```

### 5.2 需要新增的步骤

#### 步骤1：保存 Cookie 到正确路径

```
输入参数:
    shop_id  = 店铺标识（如 "FYA旗舰店"）
    biz_type = "sz" / "jzt" / "jm"

动作:
    ① 读取当前 Cookie 内容（从 xiazai.cookie.json）
    ② 拼接成字符串格式: "pin=XXX; me_js_token=YYY; ..."
    ③ 写入 D:\CODE\trae\traespace\FYA箱包旗舰店\config\{shop_id}\{biz_type}_cookie.txt
    ④ （可选）调用 Python 验证:
        python -c "from config_manager import get_config_manager; cm = get_config_manager(); print(cm.save_cookie('{shop_id}', '{biz_type}', '{cookie}'))"
```

#### 步骤2：抓 h5st（最难）

⚠️ **h5st 是京东前端 JS 计算出来的强签名，不能从 DOM 读**。

3 种方案（按实施难度）：

**方案A：拦截网络请求**（推荐 ✅）

```
影刀指令：网络请求拦截
配置:
    监听URL:    *createdExportTask*
    或:         *dsm.seller.afs.bff.ExportDsmService*
    或:         *dsm.order.export.exportCenterService*
    提取字段:   请求头 "h5st" 字段

动作:
    ① 用户触发一次导出（点击页面上的"导出"按钮）
    ② 影刀拦截 createdExportTask 请求
    ③ 提取 h5st 字段值
    ④ 写入 D:\CODE\trae\traespace\FYA箱包旗舰店\config\{shop_id}\h5st.txt
    ⑤ 通知 Python:
        python -c "from config_manager import get_config_manager; cm = get_config_manager(); print(cm.save_h5st('{shop_id}', '{h5st}'))"
```

**方案B：执行 JS 取 h5st**

```
影刀指令：执行JavaScript
代码:
    // 让浏览器计算一次 h5st 然后返回
    // 难点：需要知道 h5st 计算逻辑 + 注入时间戳
    // （不建议，h5st 每次请求都变，需要精确控制时间戳）
```

**方案C：RPA 通知人工抓包**

```
影刀指令：弹窗提醒 + 暂停流程
内容: "请打开 F12 → Network → 点击导出 → 复制 h5st → 粘贴到下方输入框"
输入框: 用户粘贴的 h5st
动作: 写入 config/{shop_id}/h5st.txt + 调用 Python cm.save_h5st
```

**建议先用方案C（最简单），方案A是长期目标。**

## 6. 完整 RPA 流程（建议）

```
┌─────────────────────────────────────┐
│  启动（参数：shop_id, biz_type）  │
└─────────────────────────────────────┘
            ↓
┌─────────────────────────────────────┐
│ 打开登录页（如 szjd.com）           │
│ 跳转至 Cookie 抓取页面              │
│ 保存 Cookie → config/{shop}/{type}_cookie.txt │
└─────────────────────────────────────┘
            ↓
┌─────────────────────────────────────┐
│ 通知：跳转到含 createdExportTask 的页面 │
│ 让用户触发一次导出（人工点击）       │
└─────────────────────────────────────┘
            ↓
┌─────────────────────────────────────┐
│ 拦截网络请求，提取 h5st               │
│ 写入 config/{shop}/h5st.txt         │
└─────────────────────────────────────┘
            ↓
┌─────────────────────────────────────┐
│ 关闭网页                              │
└─────────────────────────────────────┘
            ↓
┌─────────────────────────────────────┐
│ 调用 Python 同步数据库（cm.save_*）  │
└─────────────────────────────────────┘
            ↓
┌─────────────────────────────────────┐
│ ✅ 完成！主程序可调用：               │
│     cm.get_h5st(shop_id)             │
│     cm.get_cookie_path(shop_id, type) │
└─────────────────────────────────────┘
```

## 7. main.py 集成（下一步要做）

将 main.py 中硬编码的 `config/sz_cookie.txt` / `config/jzt_cookie.txt` / `config/jm_cookie.txt` 改为：

```python
# 旧（硬编码）
api = JingMaiAfterSaleExportAPI(h5st=h5st, cookie_path="config/jm_cookie.txt")

# 新（通过 ConfigManager）
from config_manager import get_config_manager
cm = get_config_manager()
shop_id = "FYA箱包旗舰店"  # 可 CLI 参数指定
cookie_path = cm.get_cookie_path(shop_id, "jm")
h5st_value = cm.get_h5st(shop_id) or h5st_from_cli  # CLI 优先级 > DB
api = JingMaiAfterSaleExportAPI(h5st=h5st_value, cookie_path=cookie_path)
```

## 8. 验证

```bash
# 演示用：扫描自动注册
python config_manager.py scan

# 演示用：列所有店铺
python config_manager.py list

# 演示用：检查 h5st 是否过期
python config_manager.py h5st "FYA箱包旗舰店"
```

## 9. 后续优化

- [ ] main.py 集成 ConfigManager（替换硬编码路径）
- [ ] RPA 完整流程脚本（Python 调用影刀 + 影刀调用 Python 双工）
- [ ] 多店铺批量跑（遍历所有店铺执行同一业务）
- [ ] h5st 过期自动通知（前端 websocket 或 Telegram 机器人）
- [ ] 导出历史仪表板（SQLite 统计每店铺每天导出次数）