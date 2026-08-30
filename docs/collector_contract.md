# Collector 接口约定

本阶段只定义协议，不重写全部采集器。新增业务应返回统一 `CollectorResult`，让调度器只负责配置、日期计划、调用和结果记录。

## 结果字段

```text
status        SUCCESS_WITH_DATA | SUCCESS_EMPTY | FAILED_AUTH | FAILED_REQUEST | FAILED_PARSE | FAILED_DATABASE
records       解析后的业务记录列表（空表时为空列表）
date_range    {start, end}
error_type    可选错误分类
error_message 脱敏后的可读错误
metadata      request_id、来源、耗时等非敏感信息
```

## 约束

- `SUCCESS_EMPTY` 表示平台成功返回但没有业务记录，不等同于失败。
- `FAILED_AUTH` 不得把 Cookie、h5st、密码或完整 token 写入错误消息。
- Raw 层保存原始返回，Standard 层负责日期、店铺、金额和唯一键标准化。
- Collector 不计算全店经营指标；指标由 Metric Service 统一计算。
- 新业务先注册到 `config/config.xlsx`「业务清单」，再绑定 `collector_key`、日期字段、店铺字段、唯一键及 Raw/Standard 表。
