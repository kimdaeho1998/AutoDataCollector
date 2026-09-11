from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from enum import Enum
from typing import Protocol, Sequence

from .excel.product_projection import (
    project_product_detail,
)
from .excel.product_workbook_dry_run import (
    ProductWorkbookDryRunResult,
    dry_run_product_workbook_plan,
)
from .excel.product_workbook_plan import (
    ProductWorkbookPlan,
    normalize_product_store_name,
    build_product_workbook_plan,
)
from .product_detail_batch import (
    ProductDetailBatchCollector,
    ProductDetailBatchStatus,
)


class WorksheetLike(Protocol):
    max_row: int

    def cell(
        self,
        *,
        row: int,
        column: int,
    ):
        ...


class ErpStoreLike(Protocol):
    magic_store_id: str
    store_name: str


class ProductWorkbookBatchStatus(str, Enum):
    READY = "READY"
    SAME_VALUE = "SAME_VALUE"
    ERP_NOT_FOUND = "ERP_NOT_FOUND"
    ERP_AMBIGUOUS = "ERP_AMBIGUOUS"
    EMPTY = "EMPTY"
    COLLECTION_ERROR = "COLLECTION_ERROR"
    BLOCKED = "BLOCKED"


@dataclass(frozen=True)
class ProductWorkbookStoreTarget:
    store_name: str
    sales_row: int
    quantity_row: int
    ratio_row: int


@dataclass(frozen=True)
class ProductWorkbookBatchItem:
    store_name: str
    status: ProductWorkbookBatchStatus

    erp_store_id: str | None = None
    erp_store_name: str | None = None

    plan: ProductWorkbookPlan | None = None
    dry_run: ProductWorkbookDryRunResult | None = None

    reason: str | None = None


@dataclass(frozen=True)
class ProductWorkbookBatchResult:
    period_start: date
    period_end: date

    excel_store_count: int
    erp_store_count: int

    items: tuple[
        ProductWorkbookBatchItem,
        ...
    ]

    @property
    def ready_count(self) -> int:
        return sum(
            item.status
            is ProductWorkbookBatchStatus.READY
            for item in self.items
        )

    @property
    def same_value_count(self) -> int:
        return sum(
            item.status
            is ProductWorkbookBatchStatus.SAME_VALUE
            for item in self.items
        )

    @property
    def erp_not_found_count(self) -> int:
        return sum(
            item.status
            is ProductWorkbookBatchStatus.ERP_NOT_FOUND
            for item in self.items
        )

    @property
    def erp_ambiguous_count(self) -> int:
        return sum(
            item.status
            is ProductWorkbookBatchStatus.ERP_AMBIGUOUS
            for item in self.items
        )

    @property
    def empty_count(self) -> int:
        return sum(
            item.status
            is ProductWorkbookBatchStatus.EMPTY
            for item in self.items
        )

    @property
    def collection_error_count(self) -> int:
        return sum(
            item.status
            is ProductWorkbookBatchStatus.COLLECTION_ERROR
            for item in self.items
        )

    @property
    def blocked_count(self) -> int:
        return sum(
            item.status
            is ProductWorkbookBatchStatus.BLOCKED
            for item in self.items
        )

    @property
    def writable_items(
        self,
    ) -> tuple[
        ProductWorkbookBatchItem,
        ...
    ]:
        return tuple(
            item
            for item in self.items
            if item.status
            is ProductWorkbookBatchStatus.READY
        )

    @property
    def has_failures(self) -> bool:
        return any(
            item.status
            in {
                ProductWorkbookBatchStatus.ERP_AMBIGUOUS,
                ProductWorkbookBatchStatus.COLLECTION_ERROR,
                ProductWorkbookBatchStatus.BLOCKED,
            }
            for item in self.items
        )


def read_product_workbook_store_targets(
    worksheet: WorksheetLike,
) -> tuple[
    ProductWorkbookStoreTarget,
    ...
]:
    """
    Read existing Product Workbook 3-row store blocks.

    Contract:
      row N     : A=store name, D=매출
      row N + 1 : D=건수
      row N + 2 : D=비율

    No worksheet mutation.
    """

    sales_label = "\uB9E4\uCD9C"
    quantity_label = "\uAC74\uC218"
    ratio_label = "\uBE44\uC728"

    targets: list[
        ProductWorkbookStoreTarget
    ] = []

    row = 1

    while row <= worksheet.max_row:

        if row + 2 > worksheet.max_row:
            break

        name = worksheet.cell(
            row=row,
            column=1,
        ).value

        if not (
            isinstance(name, str)
            and name.strip()
        ):
            row += 1
            continue

        current_sales_label = worksheet.cell(
            row=row,
            column=4,
        ).value

        current_quantity_label = worksheet.cell(
            row=row + 1,
            column=4,
        ).value

        current_ratio_label = worksheet.cell(
            row=row + 2,
            column=4,
        ).value

        if (
            current_sales_label == sales_label
            and current_quantity_label
            == quantity_label
            and current_ratio_label
            == ratio_label
        ):
            targets.append(
                ProductWorkbookStoreTarget(
                    store_name=name.strip(),
                    sales_row=row,
                    quantity_row=row + 1,
                    ratio_row=row + 2,
                )
            )

            row += 3
            continue

        row += 1

    normalized = [
        normalize_product_store_name(
            item.store_name
        )
        for item in targets
    ]

    if len(normalized) != len(
        set(normalized)
    ):
        raise ValueError(
            "PRODUCT_WORKBOOK_DUPLICATE_STORE_TARGET"
        )

    return tuple(
        targets
    )


def _is_seojeongri(
    value: str,
) -> bool:

    normalized = (
        normalize_product_store_name(
            value
        )
    )

    return (
        "".join(
            normalized.split()
        )
        == "\uC11C\uC815\uB9AC\uC5ED\uC810"
    )


class ProductWorkbookBatchOrchestrator:
    """
    Product Detail -> Workbook dry-run orchestration.

    This class never saves a workbook.

    Responsibilities:
    - read Excel Product store targets
    - map Excel targets to ERP stores
    - preserve the confirmed duplicate ERP candidates for
      Seojeongri so ProductDetailBatchCollector can resolve
      them by actual Product Detail row presence
    - collect Product Detail sequentially
    - project raw Product rows
    - build Workbook plans
    - execute Workbook dry-run
    """

    def __init__(
        self,
        batch_collector: ProductDetailBatchCollector,
    ) -> None:
        self.batch_collector = (
            batch_collector
        )

    def run(
        self,
        *,
        worksheet: WorksheetLike,
        erp_stores: Sequence[
            ErpStoreLike
        ],
        start_date: date,
        end_date: date,
    ) -> ProductWorkbookBatchResult:

        if end_date < start_date:
            raise ValueError(
                "PRODUCT_WORKBOOK_INVALID_PERIOD"
            )

        targets = (
            read_product_workbook_store_targets(
                worksheet
            )
        )

        target_by_name = {
            normalize_product_store_name(
                target.store_name
            ): target
            for target in targets
        }

        erp_by_name: dict[
            str,
            list[ErpStoreLike],
        ] = {}

        for store in erp_stores:

            key = (
                normalize_product_store_name(
                    store.store_name
                )
            )

            if not key:
                continue

            erp_by_name.setdefault(
                key,
                [],
            ).append(
                store
            )

        pre_items: dict[
            str,
            ProductWorkbookBatchItem,
        ] = {}

        collection_stores: list[
            ErpStoreLike
        ] = []

        for target in targets:

            key = (
                normalize_product_store_name(
                    target.store_name
                )
            )

            candidates = (
                erp_by_name.get(
                    key,
                    [],
                )
            )

            if not candidates:

                pre_items[key] = (
                    ProductWorkbookBatchItem(
                        store_name=(
                            target.store_name
                        ),
                        status=(
                            ProductWorkbookBatchStatus
                            .ERP_NOT_FOUND
                        ),
                        reason=(
                            "ERP_STORE_NOT_FOUND"
                        ),
                    )
                )

                continue

            if len(candidates) == 1:

                collection_stores.append(
                    candidates[0]
                )

                continue

            if _is_seojeongri(
                target.store_name
            ):
                # Preserve every confirmed duplicate candidate.
                # ProductDetailBatchCollector resolves exactly
                # one candidate by parsed Product Detail rows.
                collection_stores.extend(
                    candidates
                )
                continue

            ids = ",".join(
                store.magic_store_id
                for store in candidates
            )

            pre_items[key] = (
                ProductWorkbookBatchItem(
                    store_name=(
                        target.store_name
                    ),
                    status=(
                        ProductWorkbookBatchStatus
                        .ERP_AMBIGUOUS
                    ),
                    reason=(
                        "ERP_STORE_AMBIGUOUS:"
                        + ids
                    ),
                )
            )

        # Only Excel-targeted ERP stores are collected.
        batch = self.batch_collector.collect(
            stores=collection_stores,
            start_date=start_date,
            end_date=end_date,
        )

        batch_by_name = {}

        for item in batch.items:

            key = (
                normalize_product_store_name(
                    item.store_name
                )
            )

            if key in batch_by_name:
                raise ValueError(
                    "PRODUCT_BATCH_DUPLICATE_RESULT:"
                    + key
                )

            batch_by_name[
                key
            ] = item

        output_items: list[
            ProductWorkbookBatchItem
        ] = []

        for target in targets:

            key = (
                normalize_product_store_name(
                    target.store_name
                )
            )

            pre = pre_items.get(
                key
            )

            if pre is not None:
                output_items.append(
                    pre
                )
                continue

            batch_item = (
                batch_by_name.get(
                    key
                )
            )

            if batch_item is None:

                output_items.append(
                    ProductWorkbookBatchItem(
                        store_name=(
                            target.store_name
                        ),
                        status=(
                            ProductWorkbookBatchStatus
                            .COLLECTION_ERROR
                        ),
                        reason=(
                            "PRODUCT_BATCH_RESULT_MISSING"
                        ),
                    )
                )

                continue

            if (
                batch_item.status
                is ProductDetailBatchStatus.EMPTY
            ):

                output_items.append(
                    ProductWorkbookBatchItem(
                        store_name=(
                            target.store_name
                        ),
                        status=(
                            ProductWorkbookBatchStatus
                            .EMPTY
                        ),
                        erp_store_id=(
                            batch_item.store_id
                        ),
                        erp_store_name=(
                            batch_item.store_name
                        ),
                        reason=(
                            "PRODUCT_DETAIL_EMPTY"
                        ),
                    )
                )

                continue

            if (
                batch_item.status
                is ProductDetailBatchStatus.ERROR
            ):

                output_items.append(
                    ProductWorkbookBatchItem(
                        store_name=(
                            target.store_name
                        ),
                        status=(
                            ProductWorkbookBatchStatus
                            .COLLECTION_ERROR
                        ),
                        erp_store_id=(
                            batch_item.store_id
                        ),
                        erp_store_name=(
                            batch_item.store_name
                        ),
                        reason=(
                            batch_item.reason
                            or batch_item.error_type
                            or "PRODUCT_DETAIL_ERROR"
                        ),
                    )
                )

                continue

            if (
                batch_item.status
                is not ProductDetailBatchStatus.SUCCESS
                or batch_item.result is None
            ):

                output_items.append(
                    ProductWorkbookBatchItem(
                        store_name=(
                            target.store_name
                        ),
                        status=(
                            ProductWorkbookBatchStatus
                            .COLLECTION_ERROR
                        ),
                        erp_store_id=(
                            batch_item.store_id
                        ),
                        erp_store_name=(
                            batch_item.store_name
                        ),
                        reason=(
                            "INVALID_PRODUCT_BATCH_RESULT"
                        ),
                    )
                )

                continue

            result = batch_item.result

            projection = (
                project_product_detail(
                    result.records
                )
            )

            try:

                plan = (
                    build_product_workbook_plan(
                        worksheet=worksheet,
                        store_name=(
                            target.store_name
                        ),
                        projection=projection,
                        source_total_sales=(
                            result.source_total_sales
                        ),
                        source_total_quantity=(
                            result.source_total_quantity
                        ),
                    )
                )

                dry_run = (
                    dry_run_product_workbook_plan(
                        worksheet=worksheet,
                        plan=plan,
                    )
                )

            except Exception as exc:

                output_items.append(
                    ProductWorkbookBatchItem(
                        store_name=(
                            target.store_name
                        ),
                        status=(
                            ProductWorkbookBatchStatus
                            .BLOCKED
                        ),
                        erp_store_id=(
                            batch_item.store_id
                        ),
                        erp_store_name=(
                            batch_item.store_name
                        ),
                        reason=(
                            f"{type(exc).__name__}:{exc}"
                        ),
                    )
                )

                continue

            if not dry_run.can_write:

                reason = (
                    "WORKBOOK_DRY_RUN_BLOCKED:"
                    f"CONFLICT={dry_run.conflict_count},"
                    f"BLOCKED={dry_run.blocked_count}"
                )

                output_items.append(
                    ProductWorkbookBatchItem(
                        store_name=(
                            target.store_name
                        ),
                        status=(
                            ProductWorkbookBatchStatus
                            .BLOCKED
                        ),
                        erp_store_id=(
                            batch_item.store_id
                        ),
                        erp_store_name=(
                            batch_item.store_name
                        ),
                        plan=plan,
                        dry_run=dry_run,
                        reason=reason,
                    )
                )

                continue

            status = (
                ProductWorkbookBatchStatus
                .READY
                if dry_run.ready_count > 0
                else ProductWorkbookBatchStatus
                .SAME_VALUE
            )

            output_items.append(
                ProductWorkbookBatchItem(
                    store_name=(
                        target.store_name
                    ),
                    status=status,
                    erp_store_id=(
                        batch_item.store_id
                    ),
                    erp_store_name=(
                        batch_item.store_name
                    ),
                    plan=plan,
                    dry_run=dry_run,
                )
            )

        return ProductWorkbookBatchResult(
            period_start=start_date,
            period_end=end_date,
            excel_store_count=len(
                targets
            ),
            erp_store_count=len(
                erp_stores
            ),
            items=tuple(
                output_items
            ),
        )
