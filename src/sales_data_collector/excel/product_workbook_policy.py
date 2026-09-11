from __future__ import annotations

from dataclasses import dataclass
from enum import Enum

from sales_data_collector.excel.product_projection import (
    ProductUnmappedRecord,
)


class ProductWorkbookDisposition(str, Enum):
    """
    Product Workbook-specific disposition.

    This is intentionally separate from Menu Excel disposition.
    """

    RESIDUAL = "RESIDUAL"
    EXCLUDED = "EXCLUDED"
    REVIEW_REQUIRED = "REVIEW_REQUIRED"


@dataclass(frozen=True)
class ProductWorkbookUnmappedDisposition:
    source: ProductUnmappedRecord
    disposition: ProductWorkbookDisposition
    reason: str


# Confirmed from the existing Product workbook contract.
#
# These rows exist in raw product.asp data but are not part of the
# Product Workbook AA/Z aggregation.
_RESIDUAL_CLASSIFICATIONS = frozenset(
    {
        # Existing confirmed Product Workbook residual.
        "메뉴",

        # Ordinary product classifications observed in W7 ERP.
        # These are eligible for Workbook AA/Z aggregation when
        # the product itself has no E:Y canonical mapping.
        "사이드",
        "사이드메뉴",
        "사이드류",
        "사이드/음료",
        "사이드 및 음료",
        "음료",
        "음료류",
        "음료수",
        "음료 외",
        "소스/음료",
        "밀키트",
        "밀키트류",
        "밀키트/음료",
        "분식류",
        "김밥/소스",
        "김밥/소스류",
        "김밥 / 사이드 / 소스류",
        "김밥 및 소스",
        "김밥 및 소스류",
        "김밥 & 식사",
        "메뉴.",
    }
)


_EXCLUDED_CLASSIFICATIONS = frozenset(
    {
        # Existing confirmed exclusions.
        "1인세트",
        "인기세트",
        "하부메뉴",

        # Confirmed set/submenu variants.
        "1인세트(포장안됨)",
        "1인세트(매장용)",
        "1인 세트메뉴 매장용",
        "할인세트",
        "식혜할인세트",
        "떡볶이세트",
        "어묵탕세트",
        "쫄면세트",
        "하부메뉴.",

        # Delivery-platform / delivery-only classifications.
        # Raw Product Detail remains preserved; these are excluded
        # only from Product Workbook AA/Z aggregation.
        "배달의민족",
        "요기요",
        "쿠팡",
        "기타배달",
        "배달비",
        "배달/기타",
        "대구로",
        "땡겨요",
        "만원의 행복",
        "마트세트",
        "단체 주문",
        "포장",
    }
)


def normalize_product_classification(
    value: str | None,
) -> str:
    if value is None:
        return ""

    return " ".join(
        str(value).strip().split()
    )


def classify_product_workbook_unmapped(
    item: ProductUnmappedRecord,
) -> ProductWorkbookUnmappedDisposition:

    classification = normalize_product_classification(
        item.record.classification_name
    )

    # Ordinary menu rows which do not have an E:Y canonical target
    # belong to the workbook residual represented by Z.
    if classification in _RESIDUAL_CLASSIFICATIONS:
        return ProductWorkbookUnmappedDisposition(
            source=item,
            disposition=ProductWorkbookDisposition.RESIDUAL,
            reason="UNMAPPED_GENERAL_MENU",
        )

    # Sets/submenu rows are raw Product Detail data but are outside
    # the Product Workbook aggregation scope.
    if classification in _EXCLUDED_CLASSIFICATIONS:
        return ProductWorkbookUnmappedDisposition(
            source=item,
            disposition=ProductWorkbookDisposition.EXCLUDED,
            reason=f"EXCLUDED_CLASSIFICATION:{classification}",
        )

    # Fail closed for a classification not yet proven by the workbook
    # contract. Never silently place a new category into Z.
    return ProductWorkbookUnmappedDisposition(
        source=item,
        disposition=ProductWorkbookDisposition.REVIEW_REQUIRED,
        reason=(
            "UNKNOWN_PRODUCT_WORKBOOK_CLASSIFICATION:"
            f"{classification or '<EMPTY>'}"
        ),
    )


def classify_product_workbook_unmapped_records(
    records: tuple[ProductUnmappedRecord, ...],
) -> tuple[ProductWorkbookUnmappedDisposition, ...]:

    return tuple(
        classify_product_workbook_unmapped(item)
        for item in records
    )


def workbook_residual_sales(
    items: tuple[ProductWorkbookUnmappedDisposition, ...],
) -> int:

    return sum(
        int(item.source.record.sales_amount)
        for item in items
        if item.disposition
        == ProductWorkbookDisposition.RESIDUAL
    )


def workbook_residual_quantity(
    items: tuple[ProductWorkbookUnmappedDisposition, ...],
) -> int | None:

    residual = [
        item
        for item in items
        if item.disposition
        == ProductWorkbookDisposition.RESIDUAL
    ]

    quantities = [
        item.source.record.sales_quantity
        for item in residual
    ]

    # Missing source quantity must never become zero.
    if any(value is None for value in quantities):
        return None

    return sum(
        int(value)
        for value in quantities
        if value is not None
    )


def workbook_excluded_sales(
    items: tuple[ProductWorkbookUnmappedDisposition, ...],
) -> int:

    return sum(
        int(item.source.record.sales_amount)
        for item in items
        if item.disposition
        == ProductWorkbookDisposition.EXCLUDED
    )


def workbook_excluded_quantity(
    items: tuple[ProductWorkbookUnmappedDisposition, ...],
) -> int | None:

    excluded = [
        item
        for item in items
        if item.disposition
        == ProductWorkbookDisposition.EXCLUDED
    ]

    quantities = [
        item.source.record.sales_quantity
        for item in excluded
    ]

    if any(value is None for value in quantities):
        return None

    return sum(
        int(value)
        for value in quantities
        if value is not None
    )


def review_required_count(
    items: tuple[ProductWorkbookUnmappedDisposition, ...],
) -> int:

    return sum(
        1
        for item in items
        if item.disposition
        == ProductWorkbookDisposition.REVIEW_REQUIRED
    )
