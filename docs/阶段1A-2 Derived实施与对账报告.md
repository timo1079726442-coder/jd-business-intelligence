# 阶段 1A-2 Derived 实施与对账报告

## 1. 结论

MIYO 的 Derived/Enriched v1 已以可回滚 MySQL View + 纯 Python 回归实现；Derived DDL 未向 Raw/Standard 表增加或写入派生字段，也未把 Excel 派生字段伪装成原始字段。五条已确认的历史异常登记为 `LEGACY_EXCEPTION`，不建立 override 覆盖机制。

Metric Service、`/api/summary` 和第一版 MIYO 全店页面已切换到新 View，并返回稳定 Schema 与指标版本。由于旧 Excel 与当前 MySQL 并非同一数据快照，最终 Excel→MySQL 数值对账目前不能判定为通过，后续需要先统一数据快照再继续扩展店铺。

## 2. 五条异常收尾

详见 `docs/MIYO异常样本收尾报告.md`。结论如下：

| 字段 | Excel 行 | 订单号 | Excel 值 | 当前规则重算 | 分类 |
|---|---:|---|---:|---:|---|
| CA / order_split_amount | 3035 | 3387201004685880 | 0.5 | 1.0 | LEGACY_EXCEPTION |
| CA / order_split_amount | 3036 | 3387201004685883 | 0.5 | 1.0 | LEGACY_EXCEPTION |
| CB / order_count_flag | 3036 | 3387201004685883 | 0 | 1 | LEGACY_EXCEPTION |
| CB / order_count_flag | 3116 | 3387209016483151 | 0 | 1 | LEGACY_EXCEPTION |
| CB / order_count_flag | 3719 | 3387296010314002 | 0 | 1 | LEGACY_EXCEPTION |

没有发现可证明的人工修改、稳定替代规则、浮点误差或可推广的例外模式。登记文件为 `validation/legacy_exceptions/miyo_order_derived.json`。

## 3. Derived v1 规则

| Derived 字段 | 真实 Excel 依据 | 实现 |
|---|---|---|
| `order_refund_aggregated_amount` | BZ：同店铺同订单，首次订单行汇总售后退款金额；后续行记 0 | MySQL View，按 `shop_pin + 订单号` 聚合 |
| `order_split_amount` | CA：`ROUND(商家应收 × 京东价 / 同订单非零京东价合计, 2)` | MySQL View，窗口分区 |
| `order_count_flag` | CB：京东价非零、状态白名单、付款确认时间非空、销售订单、稳定排序首行 | MySQL View，`ROW_NUMBER()` |
| `shipment_status_normalized` | CC：状态白名单优先；完成类状态且无售后申请则已出库，否则取首条售后出库状态 | MySQL View，按售后 `id` 确定首条 |

时间口径保持旧 Excel：订单指标使用 AC「付款确认时间」，不是下单时间。店铺隔离使用 `shop_pin`，订单分区键为 `shop_pin + 订单号`，不依赖 Excel 行号。

## 4. 实现文件

- `sql/create_derived_views.sql`：可重复执行的 View DDL。
- `sql/drop_derived_views.sql`：回滚 DDL。
- `tools/apply_derived_views.py`：读取本地 MySQL 配置并建立 View。
- `derived/order_v1.py`：同规则纯 Python 实现，用于离线回归和未来 ETL。
- `metrics/derived_registry.json`：Derived 字段注册、依赖、分区、排序、粒度和版本元数据。
- `tests/test_derived_order_v1.py`：金额拆分、退款多行汇总、店铺隔离、售后状态回退测试。
- `metrics/service.py`：从 `vw_derived_jm_order_full` 查询，不再读取 Standard 中不存在的 BZ/CA/CB/CC 假字段。
- `dashboard/index_v1.html`：消费稳定 API Schema 的 MIYO 全店页面。

## 5. 验证结果

### 5.1 Derived 与数据库结构

- View 行数：3,668。
- 覆盖店铺：3 个（MIYO、OTA、FYA），本阶段只验证 MIYO 业务口径。
- 纯 Python 回归：3/3 通过。
- View 不向 Raw/Standard 表写入数据，删除 `vw_derived_jm_order_full` 即可回滚。验证期间复用了项目既有的 SQLite→Standard 重建脚本，重建后订单/售后行数分别复核为 3,668/801；Derived 本身没有向 Standard 增加字段。

### 5.2 API

临时端口调用 `/api/summary?shop=MIYO&start=2026-08-13&end=2026-08-19` 成功，返回：

- `shop / period / comparison_period / metrics / data_quality / meta`。
- `meta.metric_version = legacy_miyo_v1`。
- `meta.derived_rule_version = legacy_miyo_v1`。
- `meta.generated_at`、`meta.data_updated_at` 均存在。

### 5.3 Excel 对账阻断项

同一日期窗口 2026-08-13 至 2026-08-19 的结果：

| 指标 | 旧 Excel 全店复盘 | 当前 MySQL Derived/API | 差异 |
|---|---:|---:|---:|
| 成交金额 | 51,876.69 | 98,197.20 | +46,320.51 |
| 退款金额 | 2,358.00 | 29,609.60 | +27,251.60 |
| 净成交额 | 49,518.69 | 68,587.60 | +19,068.91 |
| 总订单数 | 96 | 145 | +49 |
| 退货订单数 | 4 | 28 | +24 |

订单集合检查显示：旧 Excel 有 108 个订单，MySQL 有 173 个订单；旧 Excel 没有独有订单，MySQL 多出 65 个订单。共同订单也存在少量金额快照差异。因此当前差异首先应归类为“历史 Excel 与 MySQL 数据覆盖/快照不一致”，不能归类为 Derived 公式错误，也不能通过补默认值解决。

## 6. 当前未完成与下一步

1. 暂不扩展 FYA/OTA 的指标对账；先补齐或确认 MIYO 对账窗口对应的原始数据快照。
2. 对齐快照后，重新执行 Excel→Derived→Metric→API 对账；只有差异收敛后才锁定全店看板数值。
3. 推广花费仍标记 `pending_confirmation`，没有确认三张推广表合计就是最终生产口径。
4. 本阶段没有重建 Raw/Standard、没有修改历史 Excel、没有实现商品/推广/AI 页面。
