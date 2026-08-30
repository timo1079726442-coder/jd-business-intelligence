"""订单 Derived v1：复现 MIYO 旧 Excel 的 BZ/CA/CB/CC 规则。

生产查询由 MySQL 视图承载；本模块提供同一规则的纯 Python 实现，供回归测试、
离线对账和未来 ETL 使用。输入使用标准英文键，避免绑定某一店铺或 Excel 列号。
"""
from __future__ import annotations

from collections import defaultdict
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Any

RULE_VERSION = "legacy_miyo_v1"
ORDER_COUNT_STATUSES = {"完成", "等待确认收货", "(删除)等待确认收货"}
UNSHIPPED_STATUSES = {"(删除)暂停", "(删除)等待付款确认", "(删除)延迟付款确认", "等待出库", "(删除)等待出库", "(删除)新订单"}
SHIPPED_STATUSES = {"完成", "等待确认收货", "(删除)等待确认收货", "(暂停)等待确认收货", "(锁定)等待确认收货"}


def _text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _decimal(value: Any) -> Decimal:
    text = _text(value).replace(",", "").replace("￥", "").replace("¥", "").replace("元", "")
    if not text:
        return Decimal("0")
    try:
        return Decimal(text)
    except InvalidOperation:
        return Decimal("0")


def _quantity(value: Any) -> Decimal:
    text = _text(value).replace("（", "(").split("(", 1)[0].replace("件", "").replace(",", "")
    return _decimal(text)


def _sort_key(row: dict[str, Any]) -> tuple[int, str, str, str, str]:
    """明确的稳定顺序：源记录 ID 优先，再用时间/订单/商品字段消歧。"""
    try:
        source_id = int(_text(row.get("source_id")))
    except ValueError:
        source_id = 0
    return (source_id, _text(row.get("payment_confirmed_at")), _text(row.get("order_no")), _text(row.get("product_id")), _text(row.get("sku_id")))


def derive_order_rows(order_rows: list[dict[str, Any]], after_rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """按店铺+订单隔离计算四个订单派生字段。

    订单输入键：shop_pin/order_no/price/allocation_base/status/order_type/payment_confirmed_at/
    product_id/sku_id/source_id；售后输入键：shop_pin/order_no/refund_amount/
    shipment_status/application_at/source_id。
    """
    refund_by_order: defaultdict[tuple[str, str], Decimal] = defaultdict(Decimal)
    first_after: dict[tuple[str, str], dict[str, Any]] = {}
    for row in sorted(after_rows, key=lambda item: (_text(item.get("shop_pin")), _text(item.get("order_no")), _text(item.get("source_id")))):
        key = (_text(row.get("shop_pin")), _text(row.get("order_no")))
        if not key[1]:
            continue
        refund_by_order[key] += _decimal(row.get("refund_amount"))
        first_after.setdefault(key, row)

    grouped: defaultdict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
    for row in order_rows:
        grouped[(_text(row.get("shop_pin")), _text(row.get("order_no")))].append(row)

    results: list[dict[str, Any]] = []
    for key, rows in grouped.items():
        rows = sorted(rows, key=_sort_key)
        price_sum = sum((_decimal(row.get("price")) for row in rows if _decimal(row.get("price")) != 0), Decimal("0"))
        after = first_after.get(key, {})
        for rank, row in enumerate(rows, start=1):
            price = _decimal(row.get("price")); allocation_base = _quantity(row.get("allocation_base", row.get("quantity")))
            split = Decimal("0") if price == 0 or price_sum == 0 else (allocation_base * price / price_sum).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            count_flag = int(price != 0 and _text(row.get("status")) in ORDER_COUNT_STATUSES and _text(row.get("payment_confirmed_at")) and _text(row.get("order_type")) == "销售订单" and rank == 1)
            status = _text(row.get("status"))
            if status in UNSHIPPED_STATUSES:
                shipment = "未出库"
            elif status in SHIPPED_STATUSES and _text(row.get("payment_confirmed_at")) and not _text(after.get("application_at")):
                shipment = "已出库"
            else:
                shipment = _text(after.get("shipment_status"))
            results.append({
                **row,
                "order_refund_aggregated_amount": refund_by_order[key] if rank == 1 and price != 0 else Decimal("0"),
                "order_split_amount": split,
                "order_count_flag": count_flag,
                "shipment_status_normalized": shipment,
                "derived_partition_key": f"shop={key[0]}|order={key[1]}",
                "derived_order_line_rank": rank,
                "derived_rule_version": RULE_VERSION,
            })
    return results
