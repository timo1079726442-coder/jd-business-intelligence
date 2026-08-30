# MIYO 生产快照差异报告

## 窗口与粒度

窗口：2026-08-13 至 2026-08-19。Legacy Excel distinct_order_count=108；当前 MySQL Derived distinct_order_count=173，额外订单=65。MySQL 窗口明细 row_count 为订单行数，不等同订单数。

- 当前 MySQL row_count：177
- 当前 MySQL distinct_order_count：173
- 额外订单 valid_order_count（order_count_flag=1）：49

## 分类统计

- LATER_BACKFILL：65
- 窗口内重复候选：0 个订单，0 行超出业务行唯一组合。
- 跨 shop_pin 串店订单：0。
- 额外订单均可在 SQLite 追踪，当前分类为 LATER_BACKFILL 候选；现有字段不足以进一步区分后续补数与 rolling 30d 状态更新。

## 判断

额外订单中，SQLite 可追踪的记录归为 LATER_BACKFILL 候选；无 SQLite 记录的归为 SOURCE_SNAPSHOT_CHANGED 候选。DUPLICATE_CANDIDATE 仅表示订单行数大于业务行唯一组合数，需要人工抽样确认，不能直接删除。当前查询未发现跨 shop_pin 的订单串店证据；Raw 层在当前 MySQL 中没有可识别的独立 Raw 表，因此 mysql_raw_exists 记录为 false，需后续建立批次元数据后补足来源链路。

明细见 `docs/MIYO生产快照差异明细.csv`。未修改、删除或覆盖任何订单数据。