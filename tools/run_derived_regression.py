"""无需 pytest 也可执行 Derived v1 的最小回归断言。"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from tests.test_derived_order_v1 import (
    test_refund_isolated_by_shop_and_sums_all_after_sale_rows,
    test_shipment_status_follows_after_sale_fallback,
    test_split_amount_and_order_count_use_order_partition,
)

for test in (test_split_amount_and_order_count_use_order_partition, test_refund_isolated_by_shop_and_sums_all_after_sale_rows, test_shipment_status_follows_after_sale_fallback):
    test()
    print(f"PASS {test.__name__}")
