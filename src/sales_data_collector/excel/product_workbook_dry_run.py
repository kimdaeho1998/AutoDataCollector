from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from numbers import Real
from typing import Protocol

from .product_workbook_plan import (
    ProductWorkbookPlan,
)


class WorksheetLike(Protocol):
    def __getitem__(self, key: str):
        ...


class ProductWorkbookDryRunStatus(
    str,
    Enum,
):
    READY = "READY"
    SAME_VALUE = "SAME_VALUE"
    CONFLICT = "CONFLICT"
    BLOCKED = "BLOCKED"


@dataclass(frozen=True)
class ProductWorkbookDryRunCell:
    coordinate: str
    kind: str

    canonical_key: str | None

    current_value: object
    proposed_value: int | None

    status: ProductWorkbookDryRunStatus
    reason: str | None = None


@dataclass(frozen=True)
class ProductWorkbookDryRunResult:
    store_name: str

    sales_row: int
    quantity_row: int
    ratio_row: int

    cells: tuple[
        ProductWorkbookDryRunCell,
        ...
    ]

    warnings: tuple[str, ...]

    @property
    def ready_count(self) -> int:
        return sum(
            item.status
            == ProductWorkbookDryRunStatus.READY
            for item in self.cells
        )

    @property
    def same_value_count(self) -> int:
        return sum(
            item.status
            == ProductWorkbookDryRunStatus.SAME_VALUE
            for item in self.cells
        )

    @property
    def conflict_count(self) -> int:
        return sum(
            item.status
            == ProductWorkbookDryRunStatus.CONFLICT
            for item in self.cells
        )

    @property
    def blocked_count(self) -> int:
        return sum(
            item.status
            == ProductWorkbookDryRunStatus.BLOCKED
            for item in self.cells
        )

    @property
    def can_write(self) -> bool:
        return (
            self.conflict_count == 0
            and self.blocked_count == 0
        )


def _is_blank(
    value: object,
) -> bool:
    return (
        value is None
        or (
            isinstance(value, str)
            and not value.strip()
        )
    )


def _is_formula(
    value: object,
) -> bool:
    return (
        isinstance(value, str)
        and value.startswith("=")
    )


def _is_numeric(
    value: object,
) -> bool:
    # bool is technically int in Python,
    # but it is not valid Product workbook data.
    return (
        isinstance(value, Real)
        and not isinstance(value, bool)
    )


def _compare_numeric_cell(
    *,
    coordinate: str,
    kind: str,
    canonical_key: str | None,
    current_value: object,
    proposed_value: int | None,
) -> ProductWorkbookDryRunCell:

    if proposed_value is None:
        return ProductWorkbookDryRunCell(
            coordinate=coordinate,
            kind=kind,
            canonical_key=canonical_key,
            current_value=current_value,
            proposed_value=None,
            status=(
                ProductWorkbookDryRunStatus.BLOCKED
            ),
            reason="PROPOSED_VALUE_MISSING",
        )

    proposed = int(
        proposed_value
    )

    if _is_blank(current_value):
        return ProductWorkbookDryRunCell(
            coordinate=coordinate,
            kind=kind,
            canonical_key=canonical_key,
            current_value=current_value,
            proposed_value=proposed,
            status=(
                ProductWorkbookDryRunStatus.READY
            ),
            reason="EMPTY_TARGET",
        )

    if _is_formula(current_value):
        return ProductWorkbookDryRunCell(
            coordinate=coordinate,
            kind=kind,
            canonical_key=canonical_key,
            current_value=current_value,
            proposed_value=proposed,
            status=(
                ProductWorkbookDryRunStatus.BLOCKED
            ),
            reason="FORMULA_IN_NUMERIC_TARGET",
        )

    # R15-D3 business rule:
    #
    # The Product Workbook uses "-" as an operational zero
    # placeholder. When a valid numeric proposal exists, the
    # textual placeholder must be physically replaced by that
    # numeric value in the output workbook.
    #
    # Examples:
    #   "-" + 0     -> READY -> write numeric 0
    #   "-" + 15000 -> READY -> write 15000
    #   "-" + -3000 -> READY -> write -3000
    #
    # proposed_value=None has already been blocked above as
    # PROPOSED_VALUE_MISSING and is never coerced to zero.
    if (
        isinstance(current_value, str)
        and current_value.strip() == "-"
    ):
        return ProductWorkbookDryRunCell(
            coordinate=coordinate,
            kind=kind,
            canonical_key=canonical_key,
            current_value=current_value,
            proposed_value=proposed,
            status=(
                ProductWorkbookDryRunStatus.READY
            ),
            reason="DASH_ZERO_PLACEHOLDER",
        )

    if not _is_numeric(current_value):
        return ProductWorkbookDryRunCell(
            coordinate=coordinate,
            kind=kind,
            canonical_key=canonical_key,
            current_value=current_value,
            proposed_value=proposed,
            status=(
                ProductWorkbookDryRunStatus.BLOCKED
            ),
            reason="NON_NUMERIC_TARGET",
        )

    current = int(
        current_value
    )

    if current == proposed:
        return ProductWorkbookDryRunCell(
            coordinate=coordinate,
            kind=kind,
            canonical_key=canonical_key,
            current_value=current_value,
            proposed_value=proposed,
            status=(
                ProductWorkbookDryRunStatus.SAME_VALUE
            ),
            reason=None,
        )

    # W7-compatible placeholder rule:
    #
    # Existing numeric zero is treated as an unfilled
    # placeholder only when the proposed value is non-zero.
    #
    # Negative proposed values are valid and therefore also
    # qualify for READY here.
    if (
        current == 0
        and proposed != 0
    ):
        return ProductWorkbookDryRunCell(
            coordinate=coordinate,
            kind=kind,
            canonical_key=canonical_key,
            current_value=current_value,
            proposed_value=proposed,
            status=(
                ProductWorkbookDryRunStatus.READY
            ),
            reason="ZERO_PLACEHOLDER",
        )

    return ProductWorkbookDryRunCell(
        coordinate=coordinate,
        kind=kind,
        canonical_key=canonical_key,
        current_value=current_value,
        proposed_value=proposed,
        status=(
            ProductWorkbookDryRunStatus.CONFLICT
        ),
        reason="EXISTING_VALUE_DIFFERS",
    )


def dry_run_product_workbook_plan(
    *,
    worksheet: WorksheetLike,
    plan: ProductWorkbookPlan,
) -> ProductWorkbookDryRunResult:
    """
    Compare ProductWorkbookPlan against current workbook values.

    No workbook mutations are performed.

    Checked:
      - sparse canonical sales cells E:Y
      - sparse canonical quantity cells E:Y
      - AA sales source total
      - AA quantity source total

    Not checked/written:
      - Z residual formulas
      - ratio row formulas
      - other workbook formulas
    """

    result_cells: list[
        ProductWorkbookDryRunCell
    ] = []

    warnings = list(
        plan.warnings
    )

    for cell in plan.cells:

        sales_current = (
            worksheet[
                cell.sales_cell
            ].value
        )

        result_cells.append(
            _compare_numeric_cell(
                coordinate=cell.sales_cell,
                kind="SALES",
                canonical_key=(
                    cell.canonical_key
                ),
                current_value=(
                    sales_current
                ),
                proposed_value=(
                    cell.proposed_sales
                ),
            )
        )

        quantity_current = (
            worksheet[
                cell.quantity_cell
            ].value
        )

        result_cells.append(
            _compare_numeric_cell(
                coordinate=(
                    cell.quantity_cell
                ),
                kind="QUANTITY",
                canonical_key=(
                    cell.canonical_key
                ),
                current_value=(
                    quantity_current
                ),
                proposed_value=(
                    cell.proposed_quantity
                ),
            )
        )

    # Raw Product source totals are required because Z uses:
    #
    #   AA - SUM(E:Y)
    #
    # Therefore AA must be part of the write/dry-run contract.
    result_cells.append(
        _compare_numeric_cell(
            coordinate=(
                plan.sales_total_cell
            ),
            kind="SOURCE_TOTAL_SALES",
            canonical_key=None,
            current_value=(
                worksheet[
                    plan.sales_total_cell
                ].value
            ),
            proposed_value=(
                plan.source_total_sales
            ),
        )
    )

    result_cells.append(
        _compare_numeric_cell(
            coordinate=(
                plan.quantity_total_cell
            ),
            kind="SOURCE_TOTAL_QUANTITY",
            canonical_key=None,
            current_value=(
                worksheet[
                    plan.quantity_total_cell
                ].value
            ),
            proposed_value=(
                plan.source_total_quantity
            ),
        )
    )

    blocked = [
        item
        for item in result_cells
        if item.status
        == ProductWorkbookDryRunStatus.BLOCKED
    ]

    conflicts = [
        item
        for item in result_cells
        if item.status
        == ProductWorkbookDryRunStatus.CONFLICT
    ]

    if blocked:
        warnings.append(
            "BLOCKED_PRODUCT_WORKBOOK_CELLS:"
            f"{len(blocked)}"
        )

    if conflicts:
        warnings.append(
            "CONFLICT_PRODUCT_WORKBOOK_CELLS:"
            f"{len(conflicts)}"
        )

    return ProductWorkbookDryRunResult(
        store_name=plan.store_name,
        sales_row=plan.sales_row,
        quantity_row=plan.quantity_row,
        ratio_row=plan.ratio_row,
        cells=tuple(result_cells),
        warnings=tuple(warnings),
    )
