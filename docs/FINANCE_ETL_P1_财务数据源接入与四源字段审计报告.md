# FINANCE-ETL-P1 财务数据源接入与四源字段审计

## 范围与安全边界

本轮只新增财务导出隔离链路：已验证的财务导出 XLSX → Python Cleaner → SQLite finance staging → MySQL Finance Fact。没有修改订单、售后、商智、京准通、Metric、Dashboard、Agent 或既有 Pipeline；没有执行正式退款/利润指标计算。鉴权文件、Cookie、sdtoken、h5st、签名 URL 和财务单元格内容均未写入日志、报告或 Git。

参数化日期 Smoke Test 使用 MIYO 的小窗口 `2026-09-01` 至 `2026-09-03`，待结算与已结算均返回 HTTP 200、业务成功、`taskStatus=3`，并下载可读 XLSX。Raw 文件按 SHA-256 和 metadata 归档；原始字节不修改。

## 字段与职责映射

| 来源 | 事实职责 | 日期语义 | 当前落点 |
|---|---|---|---|
| 京麦订单明细 | order/payment/product/status/shipping | `payment_confirmed_time` 仍来自京麦订单付款确认时间 | 既有 `biz_jm_order_full`，本轮只读引用 |
| 京麦售后明细 | aftersale apply/type/status/process | `aftersale_apply_time` = 售后申请时间；不等于退款成功 | 既有 `biz_jm_after_sale_full`，本轮只读引用 |
| 财务待结算 | expected settlement candidate | `settleTime` → `expected_settlement_time`；不解释为实际结算或退款成功 | `finance_pending_transaction` / `finance_pending_fee_detail` |
| 财务已结算 | actual finance settlement fact | `finishTime`（导出列“结算时间”）→ `financial_settled_time`；不解释为退款成功 | `finance_settled_transaction` / `finance_settled_fee_detail` |

两张 Sheet 始终分开处理：`交易汇总` 是交易/结算粒度，`费用明细` 是费用行粒度。ID 作为字符串保留，金额先转 Decimal 文本，日期标准化，空值保持 NULL，异常行写入专用 `finance_quarantine`，并带 `source_file/source_sheet/source_row/run_id/file_hash/reason/error_type/row_fingerprint` 血缘。

## 实际数据与 DQ

| 报告 | Sheet | 源 XLSX 非空数据行 | Cleaner 入库行 | 源内精确重复 | 隔离/Quarantine |
|---|---|---:|---:|---:|---:|
| pending | 交易汇总 | 72 | 72 | 0 | 0 |
| pending | 费用明细 | 198 | 198 | 0 | 0 |
| settled | 交易汇总 | 3583 | 3583 | 0 | 0 |
| settled | 费用明细 | 6852 | 6829 | 23 | 23 |

已结算费用明细的 23 条是源文件内精确重复，未被强行合并到业务金额。清洗时未发现空 `order_id` 与 `rf_busi_id` 同时为空的有效入库行，也未发现金额/日期解析错误。费用明细还包含 109 个未进入当前 canonical 映射的可选导出列；这些列被保留为 `source_*` 原文列，不丢失，待后续单独确认业务语义后再映射。

### 最终行数 reconciliation

此前 Smoke Test 的已结算“交易汇总”计数使用 `max_row - 1`，把第 2 行表头再次计入，因此显示 3584；按实际数据区第 3 行至末行重新计数为 3583。该差异是计数器的表头偏差，不是缺失、重复或静默丢行。

已结算费用明细严格闭合：`SOURCE_ROWS=6852 = ACCEPTED_ROWS=6829 + QUARANTINED_ROWS=23 + INTENTIONALLY_IGNORED_ROWS=0`。23 条分类为 `duplicate=23`、`missing_key=0`、`parse_error=0`、`summary_row=0`、`other=0`；每条均保留 `source_file/source_sheet/source_row/run_id/reason/error_type`。

本轮重放曾产生幂等审计记录 `idempotent_replay=10682`，这些记录未作为业务接受行或源异常行计入闭合，也未删除原始 23 条 duplicate quarantine 证据。最终四张 SQLite/MySQL 表均回到 72、198、3583、6829，且再次重放不增长。

## SQLite / MySQL

SQLite 新增四张隔离 staging 表：

- `finance_pending_transaction`
- `finance_pending_fee_detail`
- `finance_settled_transaction`
- `finance_settled_fee_detail`

MySQL 新增四张隔离 Finance Fact 表：

- `fact_finance_pending_transaction`
- `fact_finance_pending_fee_detail`
- `fact_finance_settled_transaction`
- `fact_finance_settled_fee_detail`

四张 MySQL 表均保留 `shop_key/report_type/bill_status/task_type/source_file/source_sheet/source_row/run_id/file_hash/row_fingerprint`，并保留后续关系所需的 `order_id/rf_busi_id/statement_id/bank_flow`（存在于相应报告粒度）。去重只针对本轮 replay 产生的相同业务指纹重复接受行，SQLite/MySQL 两侧均先生成回滚快照；quarantine 原始证据未改动。重复执行相同文件后 SQLite 与 MySQL 行数仍为 72、198、3583、6829，幂等校验通过。

## MIYO 四源 Join 审计

审计基于 `shop_pin=miyo-周` 与 `shop_key=MIYO`，统计口径在结果中明确区分分母：

- 京麦订单去重订单号：1238；京麦售后去重订单号：424；订单与售后交集：412，订单侧命中率 33.28%。
- 待结算财务交易去重订单号：58；其中 58 个能命中京麦订单，财务侧命中率 100%。
- 已结算财务交易去重订单号：1457；其中 1010 个能命中京麦订单，财务侧命中率 69.32%。
- 售后订单号与 pending/settled 财务订单号并集交集：208 / 424 = 49.06%。
- pending 与 settled 候选订单交集：54 / 58 = 93.10%（仅候选关系，不代表同一结算事实）。
- `rf_busi_id = order_id`：pending 65/72（90.28%），settled 3233/3583（90.23%）。`rf_busi_id != order_id`：pending 7/72，settled 350/3583；该差异只作为逆向业务候选，绝不直接定义为退款。

## `finishTime` 语义门禁

真实已结算 XLSX 的导出字段名称是“结算时间”，而创建条件使用 `finishTime/finishTimeS/finishTimeE`；既有查询视图历史上出现过“付款完成时间”字样。字段存在且能落入 `financial_settled_time`，但这组命名冲突尚未得到平台官方语义证明。因此：

`FINANCE_FINISH_TIME_SEMANTIC_STATUS = PARTIAL`

当前禁止把 `financial_settled_time` 与京麦 `payment_confirmed_time` 合并，也禁止把它用作 `refund_success_time`。

## 退款与秒退边界

财务费用分类/费用名称已可追溯，但不同费用行的业务含义尚未完成退款成功证据冻结；本次样本中仅发现极少量带“退”字样的候选费用行，不能据此建立正式退款映射。因此：

- `REFUND_FINANCIAL_EVIDENCE = PARTIAL`
- `INSTANT_REFUND_METRIC_READY = false`
- 售后申请时间、`是否闪退订单`、`financial_settled_time` 均未被当作退款成功时间。

## 代码、测试与回归

新增：`finance/etl.py`、`finance/__init__.py`、`tools/finance_ingest.py`、`tests/test_finance_etl.py`。参数化 Smoke Test 仍位于 ignored runtime 临时目录；没有新增鉴权文件或业务数据到 Git。

- 新增财务 Cleaner/幂等/Quarantine lineage 测试：3/3 PASS。
- 既有订单/Metric/Product/Attribution/Analysis/Pipeline 目标测试：48/48 PASS。
- 全量 `unittest discover` 运行 299 个既有测试时出现 1 个 Cookie mock 断言差异与 23 个既有 Windows 编码/鉴权夹具错误；这些均在本轮改动之外，没有修改夹具。
- `CROSS_SHOP_CONTAMINATION = 0`：本轮只归档/入库 MIYO，四张事实表均以 `shop_key=MIYO` 写入并有店铺索引。

## 最终结论

```ini
FINANCE_ETL_P1 = PASS_WITH_OPEN_GAPS
FINANCE_EXPORT_SOURCE = READY
FINANCE_PENDING_ETL = READY
FINANCE_SETTLED_ETL = READY
AUTH_PROVIDER = ShadowBot/RPA finance auth provider (existing verified contract)
DYNAMIC_DATE_EXPORT = PASS
RAW_ARCHIVE = PASS
PENDING_XLSX = PASS
SETTLED_XLSX = PASS
PENDING_TRANSACTION_ROWS = 72
PENDING_FEE_ROWS = 198
SETTLED_TRANSACTION_ROWS = 3583
SETTLED_FEE_ROWS = 6852 source / 6829 accepted / 23 quarantined
SQLITE_FINANCE = PASS
MYSQL_FINANCE = PASS
ORDER_FINANCE_MATCH_RATE = pending 100.00%; settled 69.32% (finance-order denominator)
AFTERSALE_FINANCE_MATCH_RATE = 49.06% (aftersale-order denominator)
FINANCE_FINISH_TIME_SEMANTIC_STATUS = PARTIAL
REFUND_FINANCIAL_EVIDENCE = PARTIAL
INSTANT_REFUND_METRIC_READY = false
IDEMPOTENCY = PASS
CROSS_SHOP_CONTAMINATION = 0
TESTS = 51 relevant PASS; full existing suite retains pre-existing fixture failures
REGRESSION = PASS for targeted existing order/metric/product/attribution/pipeline tests
FILES_CHANGED = finance/etl.py; finance/__init__.py; tools/finance_ingest.py; tests/test_finance_etl.py; this report
MIGRATIONS = 4 SQLite finance tables + 4 MySQL finance fact tables; only these task-owned finance tables were extended/deduplicated; rollback snapshots retained
GIT_STATUS = pre-existing worktree changes retained; new finance source files are the only task additions
FINANCE_FINISH_TIME_SEMANTIC_STATUS = PARTIAL
REFUND_FINANCIAL_EVIDENCE = PARTIAL
INSTANT_REFUND_METRIC_READY = false
OPEN_GAPS = 1) finishTime 平台业务语义仍未完全确认; 2) refund_success_time 尚未确认; 3) 正式退款分类规则尚未冻结; 4) 秒退指标不可正式计算; 5) 财务事实尚未正式改写 Profit Metric; 6) 财务事实尚未正式改写 Dashboard Metric
```

本轮完成后停止，不进入正式财务 Metric、Profit、Dashboard、Agent 或财务 ETL 调度接入。

## FINANCE-ETL-P1 Git closeout evidence

```ini
FINANCE_ETL_P1_GIT_CLOSEOUT = READY
ROW_RECONCILIATION = PASS
SETTLED_TRANSACTION_SOURCE = 3583 actual data rows (previous 3584 was header-counting error)
SETTLED_TRANSACTION_ACCEPTED = 3583
SETTLED_TRANSACTION_EXCLUDED_REASON = 0; no silent loss
SETTLED_FEE_SOURCE = 6852
SETTLED_FEE_ACCEPTED = 6829
SETTLED_FEE_QUARANTINED = 23 duplicate rows
SETTLED_FEE_IGNORE_REASON = intentionally_ignored=0; missing_key=0; parse_error=0
LINEAGE = source_file + source_sheet + source_row + run_id + file_hash + row_fingerprint
QUARANTINE = persisted finance_quarantine; duplicate=23; idempotent_replay audit rows retained
SECRET_CHECK = PASS
RAW_BUSINESS_DATA_COMMITTED = false
AUTH_FILES_COMMITTED = false
DEDUPE_DRY_RUN = PASS
SQLITE_DEDUPE = PASS
MYSQL_DEDUPE = PASS
SQLITE_IDEMPOTENCY_AFTER_FIX = PASS
MYSQL_IDEMPOTENCY_AFTER_FIX = PASS
QUARANTINE_INTACT = PASS
CROSS_SHOP_CONTAMINATION = 0
```
