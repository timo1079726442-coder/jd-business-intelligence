/*
  Derived / Enriched v1
  - 不修改 Raw/Standard 表；仅建立可删除、可重建的统一视图。
  - 规则来源：MIYO 旧 Excel 订单列表 BZ/CA/CB/CC 及售后 AX/AY。
  - 视图按 shop_pin 隔离，FYA/OTA 只要字段口径一致即可复用。
*/
CREATE OR REPLACE VIEW `vw_derived_jm_order_full` AS
WITH `order_base` AS (
    SELECT
        o.*,
        NULLIF(TRIM(o.`订单号`), '') AS `derived_order_no`,
        NULLIF(TRIM(o.`付款确认时间`), '') AS `derived_payment_text`,
        STR_TO_DATE(REPLACE(REPLACE(TRIM(o.`付款确认时间`), '/', '-'), 'T', ' '), '%Y-%m-%d %H:%i:%s') AS `derived_payment_at`,
        CAST(NULLIF(REPLACE(REPLACE(REPLACE(REPLACE(TRIM(o.`京东价`), ',', ''), '￥', ''), '¥', ''), '元', ''), '') AS DECIMAL(20, 4)) AS `derived_price_num`,
        CAST(NULLIF(REPLACE(REPLACE(TRIM(SUBSTRING_INDEX(REPLACE(o.`商家应收`, '（', '('), '(', 1)), '件', ''), ',', ''), '') AS DECIMAL(20, 4)) AS `derived_allocation_base_num`
    FROM `std_biz_jm_order_full` o
),
`order_ranked` AS (
    SELECT
        b.*,
        ROW_NUMBER() OVER (
            PARTITION BY b.`shop_pin`, b.`derived_order_no`
            ORDER BY b.`id`, COALESCE(b.`derived_payment_at`, '1000-01-01 00:00:00'),
                     COALESCE(b.`订单号`, ''), COALESCE(b.`商品ID`, ''), COALESCE(b.`商家SKUID`, '')
        ) AS `derived_order_line_rank`,
        SUM(CASE WHEN COALESCE(b.`derived_price_num`, 0) <> 0 THEN b.`derived_price_num` ELSE 0 END) OVER (
            PARTITION BY b.`shop_pin`, b.`derived_order_no`
        ) AS `derived_order_price_sum`
    FROM `order_base` b
),
`after_refund` AS (
    SELECT
        a.`shop_pin`, NULLIF(TRIM(a.`订单号`), '') AS `derived_order_no`,
        COALESCE(SUM(CAST(NULLIF(REPLACE(REPLACE(REPLACE(REPLACE(TRIM(a.`退款金额`), ',', ''), '￥', ''), '¥', ''), '元', ''), '') AS DECIMAL(20, 4))), 0) AS `derived_refund_sum`
    FROM `std_biz_jm_after_sale_full` a
    WHERE NULLIF(TRIM(a.`订单号`), '') IS NOT NULL
    GROUP BY a.`shop_pin`, NULLIF(TRIM(a.`订单号`), '')
),
`after_ranked` AS (
    SELECT
        a.`shop_pin`, NULLIF(TRIM(a.`订单号`), '') AS `derived_order_no`,
        a.`出库状态` AS `derived_after_shipment_status`,
        a.`售后申请时间` AS `derived_after_application_at`,
        ROW_NUMBER() OVER (
            PARTITION BY a.`shop_pin`, NULLIF(TRIM(a.`订单号`), '')
            ORDER BY a.`id`
        ) AS `derived_after_rank`
    FROM `std_biz_jm_after_sale_full` a
    WHERE NULLIF(TRIM(a.`订单号`), '') IS NOT NULL
),
`after_first` AS (
    SELECT * FROM `after_ranked` WHERE `derived_after_rank` = 1
)
SELECT
    r.*,
    CASE WHEN r.`derived_order_line_rank` = 1 AND COALESCE(r.`derived_price_num`, 0) <> 0 THEN COALESCE(ar.`derived_refund_sum`, 0) ELSE 0 END AS `order_refund_aggregated_amount`,
    CASE
        WHEN COALESCE(r.`derived_price_num`, 0) = 0 THEN 0
        ELSE ROUND(COALESCE(r.`derived_allocation_base_num`, 0) * r.`derived_price_num` / NULLIF(r.`derived_order_price_sum`, 0), 2)
    END AS `order_split_amount`,
    CASE
        WHEN COALESCE(r.`derived_price_num`, 0) <> 0
         AND r.`订单状态` IN ('完成', '等待确认收货', '(删除)等待确认收货')
         AND COALESCE(r.`derived_payment_text`, '') <> ''
         AND r.`订单类型` = '销售订单'
         AND r.`derived_order_line_rank` = 1 THEN 1
        ELSE 0
    END AS `order_count_flag`,
    CASE
        WHEN r.`订单状态` IN ('(删除)暂停', '(删除)等待付款确认', '(删除)延迟付款确认', '等待出库', '(删除)等待出库', '(删除)新订单') THEN '未出库'
        WHEN r.`订单状态` IN ('完成', '等待确认收货', '(删除)等待确认收货', '(暂停)等待确认收货', '(锁定)等待确认收货')
         AND COALESCE(r.`derived_payment_text`, '') <> ''
         AND COALESCE(TRIM(af.`derived_after_application_at`), '') = '' THEN '已出库'
        ELSE COALESCE(af.`derived_after_shipment_status`, '')
    END AS `shipment_status_normalized`,
    af.`derived_after_application_at` AS `after_sale_application_at_derived`,
    DATE(r.`derived_payment_at`) AS `payment_confirmed_date`,
    CONCAT('shop=', COALESCE(r.`shop_pin`, ''), '|order=', COALESCE(r.`derived_order_no`, '')) AS `derived_partition_key`,
    CONCAT(COALESCE(DATE_FORMAT(r.`derived_payment_at`, '%Y-%m-%d %H:%i:%s'), ''), '|', COALESCE(r.`订单号`, ''), '|', COALESCE(r.`商品ID`, ''), '|', COALESCE(r.`商家SKUID`, ''), '|id=', r.`id`) AS `derived_ordering_key`,
    'order_line' AS `derived_grain`,
    'legacy_miyo_v1' AS `derived_rule_version`
FROM `order_ranked` r
LEFT JOIN `after_refund` ar
  ON ar.`shop_pin` = r.`shop_pin` AND ar.`derived_order_no` = r.`derived_order_no`
LEFT JOIN `after_first` af
  ON af.`shop_pin` = r.`shop_pin` AND af.`derived_order_no` = r.`derived_order_no`;
