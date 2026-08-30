# Python现有运行链路审计

审计日期：2026-08-30。此文档只描述入口职责，不替代业务接口说明。

| 入口 | 当前真实职责 | 结论 |
|---|---|---|
| `main.py` | 业务注册中心与采集器调用边界；`run_business()` 找到业务 handler 并执行。部分业务会自动从 xlsx 入 SQLite。 | 保留为 Collector 层，不作为人工/RPA 的日常入口。 |
| `run_daily.py` | 读取“业务下载配置”，按京麦 rolling 30d、区间合并、逐日补齐，调用 handler 后归档 SQLite 和 Excel 总表。 | 保留为既有调度实现；新入口复用它的策略、h5st 与 xlsx→SQLite 归档辅助函数。 |
| `daily_update.py` | 早期每日更新入口，按窗口调用 `main.run_business()`。 | Legacy；不再推荐直接调用。 |
| `refresh_reports.py` | 旧的 xlsx + SQLite 双巡检/补数入口；包含重复的 Excel flush 兜底。 | Legacy 运维入口；不再作为 RPA 调用目标。 |
| `fill_missing.py` | 早期 SQLite 缺失日期补录，主要逐日执行。 | Legacy；其缺口概念由新 Data Plan 统一。 |
| `run_recent_30d.py` | 早期“近 30 天”批量辅助脚本。 | Legacy；不应再单独用它拉取业务。 |
| `rpa_run.py` | 旧的多店批次运行器；本身不应承担影刀业务规则。 | Legacy；保留兼容，不再推荐。 |
| `rpa_run_business.bat` | 指向旧 Trae 路径和 `module1.py` 的影刀辅助批处理。 | 明确过期；本阶段不修改影刀或该批处理。 |
| `tools/build_mysql_standard_tables.py` | SQLite→MySQL 全量迁移。 | 高风险：含 `DROP TABLE IF EXISTS`，禁止由生产流水线调用。 |

## 当前唯一推荐入口

`run_pipeline.py`

它是 Orchestrator（编排器），不复制京东采集逻辑：

```text
config.xlsx → Data Plan → main.py handler
                     ↓
              run_daily 既有归档辅助
                     ↓
             SQLite → MySQL 增量同步
                     ↓
        Derived View 健康检查 → Metric smoke test
                     ↓
         ETL 元数据 + runtime/runs/<run_id>.json
```

旧入口保留是为了可回滚与排障；以后人工和影刀均应切换为调用 `run_pipeline.py`，而不是删除旧脚本。
