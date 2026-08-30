from decimal import Decimal

from derived.order_v1 import derive_order_rows


def test_split_amount_and_order_count_use_order_partition():
    rows = derive_order_rows([
        {"shop_pin": "MIYO", "order_no": "o1", "price": "100", "allocation_base": "100", "status": "完成", "order_type": "销售订单", "payment_confirmed_at": "2026-08-27 10:00:00", "product_id": "p1", "sku_id": "s1", "source_id": "2"},
        {"shop_pin": "MIYO", "order_no": "o1", "price": "100", "allocation_base": "100（颜色）", "status": "完成", "order_type": "销售订单", "payment_confirmed_at": "2026-08-27 10:01:00", "product_id": "p2", "sku_id": "s2", "source_id": "3"},
    ], [])
    assert [row["order_split_amount"] for row in rows] == [Decimal("50.00"), Decimal("50.00")]
    assert [row["order_count_flag"] for row in rows] == [1, 0]


def test_refund_isolated_by_shop_and_sums_all_after_sale_rows():
    rows = derive_order_rows([
        {"shop_pin": "MIYO", "order_no": "o1", "price": "100", "allocation_base": "100", "status": "完成", "order_type": "销售订单", "payment_confirmed_at": "2026-08-27", "product_id": "p", "source_id": "1"},
        {"shop_pin": "OTA", "order_no": "o1", "price": "100", "allocation_base": "100", "status": "完成", "order_type": "销售订单", "payment_confirmed_at": "2026-08-27", "product_id": "p", "source_id": "2"},
    ], [
        {"shop_pin": "MIYO", "order_no": "o1", "refund_amount": "10", "shipment_status": "", "application_at": "", "source_id": "1"},
        {"shop_pin": "MIYO", "order_no": "o1", "refund_amount": "2.5", "shipment_status": "", "application_at": "", "source_id": "2"},
        {"shop_pin": "OTA", "order_no": "o1", "refund_amount": "7", "shipment_status": "", "application_at": "", "source_id": "1"},
    ])
    assert rows[0]["order_refund_aggregated_amount"] == Decimal("12.5")
    assert rows[1]["order_refund_aggregated_amount"] == Decimal("7")


def test_shipment_status_follows_after_sale_fallback():
    rows = derive_order_rows([
        {"shop_pin": "MIYO", "order_no": "o1", "price": "100", "allocation_base": "100", "status": "完成", "order_type": "销售订单", "payment_confirmed_at": "2026-08-27", "product_id": "p", "source_id": "1"},
        {"shop_pin": "MIYO", "order_no": "o2", "price": "100", "allocation_base": "100", "status": "等待出库", "order_type": "销售订单", "payment_confirmed_at": "2026-08-27", "product_id": "p", "source_id": "2"},
    ], [
        {"shop_pin": "MIYO", "order_no": "o1", "refund_amount": "0", "shipment_status": "售后中", "application_at": "2026-08-27", "source_id": "1"},
    ])
    assert rows[0]["shipment_status_normalized"] == "售后中"
    assert rows[1]["shipment_status_normalized"] == "未出库"
