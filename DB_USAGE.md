# 数据库集成使用说明（项目21，2026-08-14）

> 适用版本：main.py ≥ 5645769
> 目的：将 20 个京东数据报表（Excel）持久化到本地 SQLite 数据库，支持每日自动滚动更新与缺失补录。

---

## 1. 数据库位置

`data/jd_report.db`（相对项目根目录）

- 由 `db_utils.get_db_path()` 读取（默认 `data/jd_report.db`，可在 `config.xlsx` 修改）
- 目录不存在时自动创建
- **不入仓**（`.gitignore` 已排除 `data/`）

---

## 2. 表命名规则

`biz_{业务英文短名}`，例如：

| 业务 | 表名 |
|---|---|
| 商智关键词分析 | `biz_keyword_analysis` |
| 店铺来源_三级渠道 | `biz_offline_channel` |
| 商品流量来源_搜索 | `biz_traffic_search` |
| 商品流量来源_推荐 | `biz_traffic_recommend` |
| 商品流量来源_购物车 | `biz_traffic_cart` |
| 商品流量来源_自主访问 | `biz_traffic_selfvisit` |
| 商品明细导出 | `biz_product_detail` |
| 商品流失分析 | `biz_loss_product` |
| 京准通快车自定义报表 | `biz_jzt_kuaiche` |
| 京准通快车订单效果明细 | `biz_jzt_kuaiche_order_effect` |
| 京准通全站营销单品计划 | `biz_jzt_quanzhan_campaign` |
| 京准通全站营销单品推广效果 | `biz_jzt_quanzhan_effect` |
| 京准通全站营销全店计划 | `biz_jzt_quanzhan_campaign_all_store` |
| 京准通全站营销全店推广效果 | `biz_jzt_quanzhan_effect_all_store` |
| 京麦订单明细_完整一键导出 | `biz_jm_order_full` |
| 京麦售后明细_完整一键导出 | `biz_jm_after_sale_full` |

---

## 3. 通用字段与 Schema

每张表都包含：

| 字段 | 类型 | 含义 |
|---|---|---|
| `id` | INTEGER PRIMARY KEY AUTOINCREMENT | 自增 ID |
| `report_date` | TEXT NOT NULL | 业务日期，格式 `YYYY-MM-DD` |
| `etl_time` | TEXT NOT NULL | 入库时间戳，格式 `YYYY-MM-DD HH:MM:SS` |
| `granularity` | TEXT | **仅关键词分析表有**（`day` / `month`） |
| `[业务动态列]` | TEXT | 业务自身字段（数量按报表决定）|
| `[业务主键列]` | TEXT | 单主键或复合主键 |

**唯一约束**：`(report_date, 主键列)`（关键词分析额外含 `granularity`）

### Schema 示例

```sql
-- 关键词分析
CREATE TABLE biz_keyword_analysis (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    report_date TEXT NOT NULL,
    etl_time TEXT NOT NULL,
    "日期" TEXT,
    "关键词" TEXT,
    "访客数" TEXT,
    ...,
    granularity TEXT,
    UNIQUE(report_date, "关键词")
);

-- 三级渠道（复合主键）
CREATE TABLE biz_offline_channel (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    report_date TEXT NOT NULL,
    etl_time TEXT NOT NULL,
    "时间" TEXT,
    "一级来源" TEXT,
    "二级来源" TEXT,
    "三级来源" TEXT,
    "访客数" TEXT,
    ...,
    UNIQUE(report_date, "一级来源", "二级来源", "三级来源")
);
```

---

## 4. 主键配置

### 4.1 配置位置

`config/config.xlsx`「全局配置」sheet 中，按业务短写分组：

| 配置 key | 值格式 | 说明 |
|---|---|---|
| `<业务名> primary_key_column` | 列名（单主键）| 单列主键 |
| `<业务名> primary_key_columns` | 列名1,列名2,...（复合主键）| 多列复合主键，逗号分隔 |

### 4.2 已配置的业务（MVP）

| 业务名 | 主键配置 |
|---|---|
| 关键词分析 | `primary_key_column = 关键词` |
| 三级渠道 | `primary_key_columns = 一级来源,二级来源,三级来源` |

### 4.3 主键自动推断 fallback

若 `config.xlsx` 无配置，`db_utils.infer_primary_key()` 自动推断：
1. 列名包含关键词 `ID / 编号 / 名称 / 关键词 / SKU`
2. 否则取第一列文本字段
3. 失败返回 `None`（业务入库会报错）

---

## 5. 工具脚本

### 5.1 `daily_update.py`

**每日定时任务入口**，由影刀 RPA 凌晨自动触发。

| 参数 | 默认值 | 说明 |
|---|---|---|
| `--biz_keys` | MVP 2 业务（逗号分隔）| 要跑的业务列表 |
| `--start_date` | 今天-30天 | 起始日期（YYYY-MM-DD）|
| `--end_date` | 今天-1天 | 结束日期（YYYY-MM-DD，不含今天）|
| `--interval` | 0 | 逐日循环间隔秒数（防风控）|
| `--dry-run` | False | 仅预览计划，不实际下载 |

**行为**：
- 对每个业务，按 `BIZ_FEATURES` 配置判断「区间/逐日」
- `supports_range=True` → 一次性区间调用
- `supports_range=False` → 逐日循环调用
- 完成后自动入库到 DB（覆盖当天数据）

**示例**：
```bash
# 默认近 30 天跑 MVP 2 业务
python daily_update.py

# 自定义窗口
python daily_update.py --start_date 2026-07-01 --end_date 2026-08-13

# 干跑（仅预览计划）
python daily_update.py --dry-run

# 防风控：逐日循环间隔 30 秒
python daily_update.py --interval 30
```

---

### 5.2 `fill_missing.py`

**缺失日期巡查与补录**，适合周末/节假日停跑后周一补齐。

| 参数 | 默认值 | 说明 |
|---|---|---|
| `--biz_keys` | MVP 2 业务 | 要巡查的业务列表 |
| `--start_date` | 今天-30天 | 起始日期 |
| `--end_date` | 今天-1天 | 结束日期 |
| `--granularity` | None | 粒度筛选（day/month）|
| `--interval` | 0 | 逐日循环间隔秒数 |
| `--dry-run` | False | 仅扫描缺失，不下载 |

**行为**：
1. 读 DB 中各业务的已有 `report_date`
2. 与预期日期列表比对，得到缺失集合
3. 对缺失日期，调用业务下载函数（单日）逐条补录
4. 已存在的日期不重复下载

**示例**：
```bash
# 默认近 30 天巡查
python fill_missing.py

# 扫描更早历史（项目上线前的基础数据）
python fill_missing.py --start_date 2026-06-01

# 仅巡查关键词分析的 day 粒度
python fill_missing.py --biz_keys "商智关键词分析" --granularity day

# 仅显示缺失，不下载
python fill_missing.py --dry-run
```

---

## 6. `config.xlsx` 新增配置项

| 配置项 | 默认值 | 说明 |
|---|---|---|
| `db_path` | `data/jd_report.db` | SQLite 数据库路径（相对项目根目录或绝对路径）|
| `enable_db_storage` | `True` | 是否启用 DB 入库（`False` 时只导 Excel 不入库）|
| `<业务> primary_key_column` | - | 单主键列名 |
| `<业务> primary_key_columns` | - | 复合主键列名（逗号分隔）|

---

## 7. 数据库 CLI 工具

`db_utils.py` 自带 CLI，调试用：

```bash
# 查看 DB 路径 + 是否启用
python db_utils.py --db-info

# 列出所有表 + 行数
python db_utils.py --list-tables

# 查指定业务已有日期集合
python db_utils.py --list-dates "商智关键词分析"
python db_utils.py --list-dates "店铺来源_三级渠道"
```

---

## 8. 退出码约定（影刀监听）

| 退出码 | 含义 | 影刀动作 |
|---|---|---|
| **0** | 全部成功 | 继续 |
| **1** | 部分失败 | 记录失败项，继续 |
| **2** | 鉴权过期（Cookie / h5st）| 重抓鉴权（影刀子流程）|
| **3** | 系统错误 / 参数错误 | 终止，告警 |

---

## 9. 日志

- 所有操作日志写入 `logs/jd_api_YYYYMMDD.log`（沿用现有日志）
- 调度脚本（`daily_update.py` / `fill_missing.py`）的进度信息输出到 stdout
- DB 入库失败仅写 ERROR 日志，不抛异常（不影响 Excel 导出）

---

## 10. 后续扩展（第二批/第三批）

| 批 | 业务 | 预计改动 |
|---|---|---|
| 第二批 | 商智剩余 5 个（流量来源 4 + 商品明细 + 商品流失）| 5 个业务类各加 DB 入库 + config 加主键 |
| 第三批 | 京准通 6 个 + 京麦 7 个 | 同上 + h5st 透传 + IMAP 透传 |

---

## 11. 相关文件

| 文件 | 作用 |
|---|---|
| [db_utils.py](file:///d:/CODE/trae/traespace/FYA箱包旗舰店/jd_fya/db_utils.py) | 数据库工具集（连接/建表/Upsert/查询/CLI）|
| [daily_update.py](file:///d:/CODE/trae/traespace/FYA箱包旗舰店/jd_fya/daily_update.py) | 每日自动更新入口（影刀凌晨调用）|
| [fill_missing.py](file:///d:/CODE/trae/traespace/FYA箱包旗舰店/jd_fya/fill_missing.py) | 缺失日期巡查与补录入口 |
| [main.py](file:///d:/CODE/trae/traespace/FYA箱包旗舰店/jd_fya/main.py) | 业务主程序（含 DB 入库集成点）|
| [config/config.xlsx](file:///d:/CODE/trae/traespace/FYA箱包旗舰店/jd_fya/config/config.xlsx) | DB 配置 + 业务主键配置 |

---

## 12. 更新记录

| 日期 | 改动 |
|---|---|
| 2026-08-14 | 创建文档（MVP 上线，含 2 业务 + 完整 CLI 说明）|