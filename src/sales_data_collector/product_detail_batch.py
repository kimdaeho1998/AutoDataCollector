from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from enum import Enum
from typing import Protocol, Sequence

from .models import ProductDetailSalesRecord, ProductDetailSalesResult


class ProductDetailClient(Protocol):
    """Minimum client contract required by ProductDetailBatchCollector."""

    def get_product_detail_sales(
        self,
        *,
        start_date: date,
        end_date: date,
        brand_idx: str,
        brand_name: str,
        store_idx: str,
        store_name: str,
    ) -> ProductDetailSalesResult:
        ...


class ProductDetailStore(Protocol):
    """Minimum store contract required by ProductDetailBatchCollector."""

    magic_store_id: str
    store_name: str


class ProductDetailBatchStatus(str, Enum):
    SUCCESS = "SUCCESS"
    EMPTY = "EMPTY"
    ERROR = "ERROR"


@dataclass(frozen=True)
class ProductDetailBatchItem:
    """Collection outcome for exactly one MagicERP store."""

    store_id: str
    store_name: str
    status: ProductDetailBatchStatus
    result: ProductDetailSalesResult | None = None
    error_type: str | None = None
    reason: str | None = None


@dataclass(frozen=True)
class ProductDetailBatchResult:
    """Aggregate result of a sequential Product Detail batch collection."""

    period_start: date
    period_end: date
    items: tuple[ProductDetailBatchItem, ...]

    @property
    def attempted_store_count(self) -> int:
        return len(self.items)

    @property
    def success_count(self) -> int:
        return sum(
            item.status is ProductDetailBatchStatus.SUCCESS
            for item in self.items
        )

    @property
    def empty_count(self) -> int:
        return sum(
            item.status is ProductDetailBatchStatus.EMPTY
            for item in self.items
        )

    @property
    def failure_count(self) -> int:
        return sum(
            item.status is ProductDetailBatchStatus.ERROR
            for item in self.items
        )

    @property
    def successful_results(self) -> tuple[ProductDetailSalesResult, ...]:
        return tuple(
            item.result
            for item in self.items
            if (
                item.status is ProductDetailBatchStatus.SUCCESS
                and item.result is not None
            )
        )

    @property
    def empty_results(self) -> tuple[ProductDetailSalesResult, ...]:
        return tuple(
            item.result
            for item in self.items
            if (
                item.status is ProductDetailBatchStatus.EMPTY
                and item.result is not None
            )
        )

    @property
    def failures(self) -> tuple[ProductDetailBatchItem, ...]:
        return tuple(
            item
            for item in self.items
            if item.status is ProductDetailBatchStatus.ERROR
        )

    @property
    def records(self) -> tuple[ProductDetailSalesRecord, ...]:
        return tuple(
            record
            for item in self.items
            if (
                item.status is ProductDetailBatchStatus.SUCCESS
                and item.result is not None
            )
            for record in item.result.records
        )

    @property
    def total_quantity(self) -> int:
        return sum(
            result.source_total_quantity
            for result in self.successful_results
        )

    @property
    def total_sales(self) -> int:
        return sum(
            result.source_total_sales
            for result in self.successful_results
        )


class ProductDetailBatchCollector:
    """
    Sequentially collect raw Product Detail data for multiple stores.

    Contract:
    - concurrency is intentionally 1.
    - one ordinary store failure does not discard other store results.
    - raw Product Detail records are preserved without menu filtering.
    - signed quantity/sales values are preserved unchanged.
    - HTTP 403 / HTTP 429 / CAPTCHA-like access failures stop the batch.
    """

    def __init__(
        self,
        client: ProductDetailClient,
        *,
        brand_idx: str,
        brand_name: str,
    ) -> None:
        self.client = client
        self.brand_idx = brand_idx
        self.brand_name = brand_name

    def collect(
        self,
        *,
        stores: Sequence[ProductDetailStore],
        start_date: date,
        end_date: date,
    ) -> ProductDetailBatchResult:
        if end_date < start_date:
            raise ValueError("PRODUCT_DETAIL_INVALID_PERIOD")

        items: list[ProductDetailBatchItem] = []

        for store in stores:
            try:
                result = self.client.get_product_detail_sales(
                    start_date=start_date,
                    end_date=end_date,
                    brand_idx=self.brand_idx,
                    brand_name=self.brand_name,
                    store_idx=store.magic_store_id,
                    store_name=store.store_name,
                )

                status = (
                    ProductDetailBatchStatus.SUCCESS
                    if result.records
                    else ProductDetailBatchStatus.EMPTY
                )

                items.append(
                    ProductDetailBatchItem(
                        store_id=store.magic_store_id,
                        store_name=store.store_name,
                        status=status,
                        result=result,
                    )
                )

            except Exception as exc:
                # Authentication/rate-limit/anti-bot failures are not
                # ordinary store-level data failures. Continuing requests
                # could worsen the situation, so stop immediately.
                if _is_fatal_access_error(exc):
                    raise

                items.append(
                    ProductDetailBatchItem(
                        store_id=store.magic_store_id,
                        store_name=store.store_name,
                        status=ProductDetailBatchStatus.ERROR,
                        error_type=type(exc).__name__,
                        reason=str(exc),
                    )
                )

        return ProductDetailBatchResult(
            period_start=start_date,
            period_end=end_date,
            items=tuple(items),
        )


def _is_fatal_access_error(exc: Exception) -> bool:
    """
    Return True for access-control/rate-limit failures that must stop
    further ERP requests.
    """

    response = getattr(exc, "response", None)
    status_code = getattr(response, "status_code", None)

    if status_code in {403, 429}:
        return True

    message = str(exc).upper()

    return "CAPTCHA" in message
