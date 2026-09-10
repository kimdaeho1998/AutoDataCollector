from __future__ import annotations

from dataclasses import dataclass
from typing import Protocol, Sequence

from .product_projection import ProductProjectionResult


PRODUCT_FIRST_COLUMN = "E"
PRODUCT_LAST_COLUMN = "Y"
PRODUCT_TOTAL_COLUMN = "AA"

SALES_LABEL = "매출"
QUANTITY_LABEL = "건수"
RATIO_LABEL = "비율"

BRAND_PREFIX = "선비꼬마김밥"


class WorksheetLike(Protocol):
    max_row: int

    def cell(
        self,
        *,
        row: int,
        column: int,
    ):
        ...


@dataclass(frozen=True)
class ProductWorkbookRows:
    store_name: str
    sales_row: int
    quantity_row: int
    ratio_row: int


@dataclass(frozen=True)
class ProductWorkbookCellPlan:
    canonical_key: str
    column_letter: str

    sales_cell: str
    quantity_cell: str

    proposed_sales: int
    proposed_quantity: int | None

    source_names: tuple[str, ...]
    record_count: int
    quantity_missing_count: int


@dataclass(frozen=True)
class ProductWorkbookUnmappedPlan:
    product_name: str
    classification_name: str | None
    sales_amount: int
    sales_quantity: int | None


@dataclass(frozen=True)
class ProductWorkbookPlan:
    store_name: str

    sales_row: int
    quantity_row: int
    ratio_row: int

    cells: tuple[ProductWorkbookCellPlan, ...]

    source_total_sales: int
    source_total_quantity: int

    mapped_sales_total: int
    mapped_quantity_total: int | None

    residual_sales: int
    residual_quantity: int | None

    sales_total_cell: str
    quantity_total_cell: str

    unmapped: tuple[ProductWorkbookUnmappedPlan, ...]
    warnings: tuple[str, ...]

    @property
    def mapped_cell_count(self) -> int:
        return len(self.cells)

    @property
    def unmapped_record_count(self) -> int:
        return len(self.unmapped)

    @property
    def has_missing_quantity(self) -> bool:
        return any(
            cell.proposed_quantity is None
            for cell in self.cells
        )


def normalize_product_store_name(
    value: object,
) -> str:
    """
    Product workbook store-name normalization.

    Contract:
    - Strip only the known brand prefix.
    - Normalize whitespace.
    - Do NOT remove suffixes such as:
      (신), (전), 직영, 본점, 1호점.
    """

    if value is None:
        return ""

    text = " ".join(
        str(value).strip().split()
    )

    if text.startswith(BRAND_PREFIX):
        text = text[len(BRAND_PREFIX):].strip()

    return " ".join(
        text.split()
    )


def _cell_text(
    worksheet: WorksheetLike,
    row: int,
    column: int,
) -> str:

    value = worksheet.cell(
        row=row,
        column=column,
    ).value

    if value is None:
        return ""

    return str(value).strip()


def resolve_product_workbook_rows(
    worksheet: WorksheetLike,
    *,
    store_name: str,
) -> ProductWorkbookRows:
    """
    Resolve one existing 3-row Product workbook block.

    Expected structure:

        row N     : A=store, D=매출
        row N + 1 : D=건수
        row N + 2 : D=비율

    No row insertion is performed.
    """

    target = normalize_product_store_name(
        store_name
    )

    if not target:
        raise ValueError(
            "PRODUCT_WORKBOOK_STORE_NAME_EMPTY"
        )

    candidates: list[int] = []

    for row in range(
        1,
        worksheet.max_row + 1,
    ):
        current = normalize_product_store_name(
            worksheet.cell(
                row=row,
                column=1,
            ).value
        )

        if current == target:
            candidates.append(row)

    if not candidates:
        raise ValueError(
            "PRODUCT_WORKBOOK_STORE_NOT_FOUND:"
            f"{target}"
        )

    if len(candidates) > 1:
        raise ValueError(
            "PRODUCT_WORKBOOK_STORE_AMBIGUOUS:"
            f"{target}:"
            f"ROWS={candidates}"
        )

    sales_row = candidates[0]
    quantity_row = sales_row + 1
    ratio_row = sales_row + 2

    if ratio_row > worksheet.max_row:
        raise ValueError(
            "PRODUCT_WORKBOOK_BLOCK_OUT_OF_RANGE:"
            f"{target}:"
            f"SALES_ROW={sales_row}"
        )

    sales_label = _cell_text(
        worksheet,
        sales_row,
        4,
    )

    quantity_label = _cell_text(
        worksheet,
        quantity_row,
        4,
    )

    ratio_label = _cell_text(
        worksheet,
        ratio_row,
        4,
    )

    if sales_label != SALES_LABEL:
        raise ValueError(
            "PRODUCT_WORKBOOK_SALES_LABEL_INVALID:"
            f"D{sales_row}={sales_label!r}"
        )

    if quantity_label != QUANTITY_LABEL:
        raise ValueError(
            "PRODUCT_WORKBOOK_QUANTITY_LABEL_INVALID:"
            f"D{quantity_row}={quantity_label!r}"
        )

    if ratio_label != RATIO_LABEL:
        raise ValueError(
            "PRODUCT_WORKBOOK_RATIO_LABEL_INVALID:"
            f"D{ratio_row}={ratio_label!r}"
        )

    return ProductWorkbookRows(
        store_name=target,
        sales_row=sales_row,
        quantity_row=quantity_row,
        ratio_row=ratio_row,
    )


def build_product_workbook_plan(
    *,
    worksheet: WorksheetLike,
    store_name: str,
    projection: ProductProjectionResult,
    source_total_sales: int,
    source_total_quantity: int,
) -> ProductWorkbookPlan:
    """
    Build a pure write plan.

    This function does NOT modify the worksheet.

    E:Y
        Direct canonical Product values.

    AA
        Raw Product source total.

    Z
        Not written here.
        Existing workbook formula remains authoritative:
        AA - SUM(E:Y)

    Ratio row
        Not written here.
        Existing formulas remain untouched.
    """

    rows = resolve_product_workbook_rows(
        worksheet,
        store_name=store_name,
    )

    cells: list[ProductWorkbookCellPlan] = []

    mapped_sales_total = 0
    mapped_quantity_total = 0

    mapped_quantity_complete = True

    warnings: list[str] = []

    seen_columns: set[str] = set()

    for item in projection.mapped:

        column = str(
            item.column_letter
        ).strip().upper()

        if not (
            PRODUCT_FIRST_COLUMN
            <= column
            <= PRODUCT_LAST_COLUMN
        ):
            raise ValueError(
                "PRODUCT_WORKBOOK_COLUMN_OUT_OF_RANGE:"
                f"{item.canonical_key}:"
                f"{column}"
            )

        if column in seen_columns:
            raise ValueError(
                "PRODUCT_WORKBOOK_DUPLICATE_COLUMN:"
                f"{column}"
            )

        seen_columns.add(column)

        proposed_sales = int(
            item.sales_amount
        )

        proposed_quantity = (
            None
            if item.sales_quantity is None
            else int(item.sales_quantity)
        )

        mapped_sales_total += proposed_sales

        if proposed_quantity is None:
            mapped_quantity_complete = False

            warnings.append(
                "MISSING_MAPPED_PRODUCT_QUANTITY:"
                f"{item.canonical_key}"
            )
        else:
            mapped_quantity_total += (
                proposed_quantity
            )

        cells.append(
            ProductWorkbookCellPlan(
                canonical_key=(
                    item.canonical_key
                ),
                column_letter=column,
                sales_cell=(
                    f"{column}{rows.sales_row}"
                ),
                quantity_cell=(
                    f"{column}{rows.quantity_row}"
                ),
                proposed_sales=proposed_sales,
                proposed_quantity=proposed_quantity,
                source_names=tuple(
                    item.source_names
                ),
                record_count=int(
                    item.record_count
                ),
                quantity_missing_count=int(
                    item.quantity_missing_count
                ),
            )
        )

    cells.sort(
        key=lambda item: item.column_letter
    )

    unmapped = tuple(
        ProductWorkbookUnmappedPlan(
            product_name=(
                item.record.product_name
            ),
            classification_name=(
                item.record.classification_name
            ),
            sales_amount=int(
                item.record.sales_amount
            ),
            sales_quantity=(
                None
                if item.record.sales_quantity
                is None
                else int(
                    item.record.sales_quantity
                )
            ),
        )
        for item in projection.unmapped
    )

    if unmapped:
        warnings.append(
            "UNMAPPED_PRODUCT_RECORDS:"
            f"{len(unmapped)}"
        )

    mapped_quantity_value = (
        mapped_quantity_total
        if mapped_quantity_complete
        else None
    )

    residual_quantity = (
        int(source_total_quantity)
        - mapped_quantity_total
        if mapped_quantity_complete
        else None
    )

    # Signed source values are valid.
    # Therefore residual sales/quantity are NOT rejected merely
    # because they are negative.
    residual_sales = (
        int(source_total_sales)
        - mapped_sales_total
    )

    return ProductWorkbookPlan(
        store_name=rows.store_name,
        sales_row=rows.sales_row,
        quantity_row=rows.quantity_row,
        ratio_row=rows.ratio_row,
        cells=tuple(cells),
        source_total_sales=int(
            source_total_sales
        ),
        source_total_quantity=int(
            source_total_quantity
        ),
        mapped_sales_total=(
            mapped_sales_total
        ),
        mapped_quantity_total=(
            mapped_quantity_value
        ),
        residual_sales=residual_sales,
        residual_quantity=(
            residual_quantity
        ),
        sales_total_cell=(
            f"{PRODUCT_TOTAL_COLUMN}"
            f"{rows.sales_row}"
        ),
        quantity_total_cell=(
            f"{PRODUCT_TOTAL_COLUMN}"
            f"{rows.quantity_row}"
        ),
        unmapped=unmapped,
        warnings=tuple(warnings),
    )
