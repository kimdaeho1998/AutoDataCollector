from __future__ import annotations

from dataclasses import dataclass
from datetime import date

import pytest

from sales_data_collector.models import (
    ProductDetailSalesRecord,
    ProductDetailSalesResult,
)
from sales_data_collector.product_detail_batch import (
    ProductDetailBatchCollector,
    ProductDetailBatchStatus,
)


START_DATE = date(2026, 9, 1)
END_DATE = date(2026, 9, 1)


@dataclass(frozen=True)
class FakeStore:
    magic_store_id: str
    store_name: str


def make_result(
    store: FakeStore,
    *,
    quantity: int,
    sales: int,
    product_name: str = "테스트상품",
) -> ProductDetailSalesResult:
    return ProductDetailSalesResult(
        store_id=store.magic_store_id,
        store_name=store.store_name,
        period_start=START_DATE,
        period_end=END_DATE,
        records=[
            ProductDetailSalesRecord(
                store_id=store.magic_store_id,
                store_name=store.store_name,
                period_start=START_DATE,
                period_end=END_DATE,
                product_name=product_name,
                sales_quantity=quantity,
                sales_amount=sales,
                unit_price=None,
                classification_name=None,
            )
        ],
        source_total_sales=sales,
        source_total_quantity=quantity,
    )


class SequencedClient:
    def __init__(self, outcomes):
        self.outcomes = dict(outcomes)
        self.calls: list[str] = []

    def get_product_detail_sales(
        self,
        *,
        start_date,
        end_date,
        brand_idx,
        brand_name,
        store_idx,
        store_name,
    ):
        self.calls.append(store_idx)

        outcome = self.outcomes[store_idx]

        if isinstance(outcome, Exception):
            raise outcome

        return outcome


class EmptyClient:
    def get_product_detail_sales(
        self,
        *,
        start_date,
        end_date,
        brand_idx,
        brand_name,
        store_idx,
        store_name,
    ):
        return ProductDetailSalesResult(
            store_id=store_idx,
            store_name=store_name,
            period_start=start_date,
            period_end=end_date,
            records=[],
            source_total_sales=0,
            source_total_quantity=0,
        )


class FakeResponse:
    def __init__(self, status_code: int):
        self.status_code = status_code


class FatalHttpError(RuntimeError):
    def __init__(self, status_code: int):
        super().__init__(f"HTTP {status_code}")
        self.response = FakeResponse(status_code)


def test_batch_isolates_one_store_failure_and_continues():
    store_a = FakeStore("A", "A점")
    store_b = FakeStore("B", "B점")
    store_c = FakeStore("C", "C점")

    client = SequencedClient(
        {
            "A": make_result(
                store_a,
                quantity=2,
                sales=10_000,
                product_name="상품A",
            ),
            "B": RuntimeError("STORE_PARSE_FAILED"),
            "C": make_result(
                store_c,
                quantity=3,
                sales=20_000,
                product_name="상품C",
            ),
        }
    )

    collector = ProductDetailBatchCollector(
        client,
        brand_idx="B23020300001",
        brand_name="선비꼬마김밥",
    )

    result = collector.collect(
        stores=[store_a, store_b, store_c],
        start_date=START_DATE,
        end_date=END_DATE,
    )

    assert client.calls == ["A", "B", "C"]

    assert result.attempted_store_count == 3
    assert result.success_count == 2
    assert result.empty_count == 0
    assert result.failure_count == 1

    assert result.items[0].status is ProductDetailBatchStatus.SUCCESS
    assert result.items[1].status is ProductDetailBatchStatus.ERROR
    assert result.items[2].status is ProductDetailBatchStatus.SUCCESS

    assert result.items[1].store_id == "B"
    assert result.items[1].error_type == "RuntimeError"
    assert result.items[1].reason == "STORE_PARSE_FAILED"

    assert result.total_quantity == 5
    assert result.total_sales == 30_000

    assert [record.product_name for record in result.records] == [
        "상품A",
        "상품C",
    ]


def test_batch_preserves_negative_product_values():
    store = FakeStore("NEG", "음수점")

    client = SequencedClient(
        {
            "NEG": make_result(
                store,
                quantity=-2,
                sales=-8_000,
                product_name="취소상품",
            )
        }
    )

    collector = ProductDetailBatchCollector(
        client,
        brand_idx="B23020300001",
        brand_name="선비꼬마김밥",
    )

    result = collector.collect(
        stores=[store],
        start_date=START_DATE,
        end_date=END_DATE,
    )

    assert result.success_count == 1
    assert result.failure_count == 0

    assert result.total_quantity == -2
    assert result.total_sales == -8_000

    assert result.records[0].sales_quantity == -2
    assert result.records[0].sales_amount == -8_000


def test_batch_marks_empty_product_result_without_failure():
    store = FakeStore("EMPTY", "빈매장")

    collector = ProductDetailBatchCollector(
        EmptyClient(),
        brand_idx="B23020300001",
        brand_name="선비꼬마김밥",
    )

    result = collector.collect(
        stores=[store],
        start_date=START_DATE,
        end_date=END_DATE,
    )

    assert result.attempted_store_count == 1
    assert result.success_count == 0
    assert result.empty_count == 1
    assert result.failure_count == 0

    assert result.items[0].status is ProductDetailBatchStatus.EMPTY
    assert result.records == ()
    assert result.total_quantity == 0
    assert result.total_sales == 0


@pytest.mark.parametrize("status_code", [403, 429])
def test_batch_stops_immediately_on_fatal_http_access_error(status_code):
    store_a = FakeStore("A", "A점")
    store_b = FakeStore("B", "B점")

    client = SequencedClient(
        {
            "A": FatalHttpError(status_code),
            "B": make_result(
                store_b,
                quantity=1,
                sales=4_000,
            ),
        }
    )

    collector = ProductDetailBatchCollector(
        client,
        brand_idx="B23020300001",
        brand_name="선비꼬마김밥",
    )

    with pytest.raises(FatalHttpError):
        collector.collect(
            stores=[store_a, store_b],
            start_date=START_DATE,
            end_date=END_DATE,
        )

    # B must never be requested after 403/429.
    assert client.calls == ["A"]


def test_batch_stops_immediately_on_captcha():
    store_a = FakeStore("A", "A점")
    store_b = FakeStore("B", "B점")

    client = SequencedClient(
        {
            "A": RuntimeError("CAPTCHA_REQUIRED"),
            "B": make_result(
                store_b,
                quantity=1,
                sales=4_000,
            ),
        }
    )

    collector = ProductDetailBatchCollector(
        client,
        brand_idx="B23020300001",
        brand_name="선비꼬마김밥",
    )

    with pytest.raises(RuntimeError, match="CAPTCHA"):
        collector.collect(
            stores=[store_a, store_b],
            start_date=START_DATE,
            end_date=END_DATE,
        )

    assert client.calls == ["A"]


def test_batch_rejects_reversed_period_without_client_call():
    store = FakeStore("A", "A점")

    client = SequencedClient(
        {
            "A": make_result(
                store,
                quantity=1,
                sales=4_000,
            )
        }
    )

    collector = ProductDetailBatchCollector(
        client,
        brand_idx="B23020300001",
        brand_name="선비꼬마김밥",
    )

    with pytest.raises(
        ValueError,
        match="PRODUCT_DETAIL_INVALID_PERIOD",
    ):
        collector.collect(
            stores=[store],
            start_date=date(2026, 9, 2),
            end_date=date(2026, 9, 1),
        )

    assert client.calls == []
