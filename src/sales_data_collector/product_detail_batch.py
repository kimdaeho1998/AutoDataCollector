from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from enum import Enum
from typing import Protocol, Sequence

from .models import ProductDetailSalesRecord, ProductDetailSalesResult
from .mapping.store_normalizer import normalize_store_name


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


class ProductDetailDuplicateStoreError(RuntimeError):
    """Base error for duplicate ERP store resolution failures."""


class ProductDetailEmptyAmbiguousStoreError(
    ProductDetailDuplicateStoreError
):
    """No duplicate candidate contains Product Detail rows."""


class ProductDetailAmbiguousDataStoreError(
    ProductDetailDuplicateStoreError
):
    """More than one duplicate candidate contains Product Detail rows."""


class ProductDetailDuplicateProbeError(
    ProductDetailDuplicateStoreError
):
    """A duplicate candidate could not be safely inspected."""


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

        resolved_stores, precollected = (
            self._resolve_duplicate_stores(
                stores=stores,
                start_date=start_date,
                end_date=end_date,
            )
        )

        items: list[ProductDetailBatchItem] = []

        for store in resolved_stores:
            try:
                result = precollected.get(
                    store.magic_store_id
                )

                if result is None:
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

    def _resolve_duplicate_stores(
        self,
        *,
        stores: Sequence[ProductDetailStore],
        start_date: date,
        end_date: date,
    ) -> tuple[
        list[ProductDetailStore],
        dict[str, ProductDetailSalesResult],
    ]:
        """
        Resolve only the confirmed duplicate ERP identity for 서정리역점.

        General store matching remains unchanged.

        Resolution contract:
        - one 서정리역점 candidate: use normally.
        - multiple candidates:
          * exactly one candidate with parsed raw product rows -> select.
          * zero candidates with rows -> fail closed.
          * multiple candidates with rows -> fail closed.
          * ordinary probe error -> fail closed.
        - values themselves are never used as the data-presence test.
          Zero and negative quantity/sales remain valid.
        """

        store_list = list(stores)

        seojeongri_candidates = [
            store
            for store in store_list
            if _is_seojeongri_store(
                store.store_name
            )
        ]

        if len(seojeongri_candidates) <= 1:
            return store_list, {}

        data_candidates: list[
            tuple[
                ProductDetailStore,
                ProductDetailSalesResult,
            ]
        ] = []

        for candidate in seojeongri_candidates:
            try:
                result = self.client.get_product_detail_sales(
                    start_date=start_date,
                    end_date=end_date,
                    brand_idx=self.brand_idx,
                    brand_name=self.brand_name,
                    store_idx=candidate.magic_store_id,
                    store_name=candidate.store_name,
                )

            except Exception as exc:
                if _is_fatal_access_error(exc):
                    raise

                if _is_empty_product_detail_probe_error(exc):
                    # Valid MagicERP Product Detail page with no
                    # Product rows for this duplicate candidate.
                    #
                    # Treat this candidate as "no data", not as
                    # a probe failure. Do not synthesize zero values.
                    continue

                raise ProductDetailDuplicateProbeError(
                    "PRODUCT_DETAIL_DUPLICATE_PROBE_ERROR: "
                    f"store_id={candidate.magic_store_id}; "
                    f"reason={exc}"
                ) from exc

            # Source-data presence is based only on parsed rows.
            #
            # Do NOT test quantity/sales > 0:
            # signed and zero-valued source rows are valid.
            if result.records:
                data_candidates.append(
                    (
                        candidate,
                        result,
                    )
                )

        if not data_candidates:
            raise ProductDetailEmptyAmbiguousStoreError(
                "PRODUCT_DETAIL_EMPTY_AMBIGUOUS_STORE: "
                "서정리역점"
            )

        if len(data_candidates) > 1:
            ids = ",".join(
                candidate.magic_store_id
                for candidate, _ in data_candidates
            )

            raise ProductDetailAmbiguousDataStoreError(
                "PRODUCT_DETAIL_AMBIGUOUS_DATA_STORE: "
                f"서정리역점; ids={ids}"
            )

        selected_store, selected_result = (
            data_candidates[0]
        )

        resolved: list[ProductDetailStore] = []

        seojeongri_added = False

        for store in store_list:
            if not _is_seojeongri_store(
                store.store_name
            ):
                resolved.append(store)
                continue

            if (
                not seojeongri_added
                and store.magic_store_id
                == selected_store.magic_store_id
            ):
                resolved.append(store)
                seojeongri_added = True

        return (
            resolved,
            {
                selected_store.magic_store_id:
                    selected_result
            },
        )


def _is_seojeongri_store(store_name: str) -> bool:
    """
    Return True only for the confirmed duplicate ERP identity
    '서정리역점'.

    The existing project normalizer is consulted first, but this
    Product Detail boundary also tolerates the MagicERP brand prefix
    still being present.

    No generic suffix such as (신), (전), 1호점, 직영 is removed.
    """

    normalized = normalize_store_name(
        store_name
    )

    candidates = {
        " ".join(str(normalized).split()),
        " ".join(str(store_name).split()),
    }

    brand_prefix = "선비꼬마김밥"

    canonical_names: set[str] = set()

    for value in candidates:
        value = value.strip()

        if value.startswith(brand_prefix):
            value = value[len(brand_prefix):].strip()

        canonical_names.add(
            "".join(value.split())
        )

    return canonical_names == {"서정리역점"} or (
        "서정리역점" in canonical_names
    )




def _is_empty_product_detail_probe_error(
    exc: Exception,
) -> bool:
    """
    Return True only for the confirmed MagicERP Product Detail
    no-data response shape encountered while resolving a duplicate store.

    This helper is intentionally narrow.

    It must NOT convert arbitrary ParseError/malformed HTML into EMPTY.
    """

    message = str(exc)

    required_fragments = (
        "product detail response contains no parseable product rows",
        "title='::::MagicERP::::'",
        "table_count=0",
        "'startDate'",
        "'endDate'",
        "'brandidx'",
        "'storeidx'",
    )

    return all(
        fragment in message
        for fragment in required_fragments
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
