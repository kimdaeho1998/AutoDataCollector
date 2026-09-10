from __future__ import annotations

from dataclasses import dataclass
from typing import Sequence

from ..models import ProductDetailSalesRecord
from .menu_alias_registry import canonical_key_for_menu


# ---------------------------------------------------------------------
# Product workbook physical contract.
#
# IMPORTANT:
# - This mapping belongs to the Product workbook.
# - It does NOT redefine the existing 21 canonical aliases.
# - Z is intentionally absent.
# - Z remains the workbook residual/formula area.
# ---------------------------------------------------------------------

PRODUCT_CANONICAL_COLUMNS: dict[str, str] = {
    "KIMBAP_5": "E",
    "KIMBAP_10": "F",
    "KIMBAP_1": "G",
    "WASABI_CRAB_4": "H",
    "SPICY_JINMI_4": "I",
    "YUBU_4": "J",
    "BUL_EOMUK_4": "K",
    "FISH_CAKE_SOUP": "L",
    "TTEOKBOKKI_MILD": "M",
    "TTEOKBOKKI_SPICY": "N",
    "JJOLMYEON_MILD": "O",
    "JJOLMYEON_SPICY": "P",
    "SEONBI_UDON": "Q",
    "SEONBI_KIMCHI_UDON": "R",
    "UDON": "S",
    "RAMEN": "T",
    "SRIRACHA": "U",
    "CHEONGYANG": "V",
    "MAKHANI_CURRY": "W",
    "CHEESE": "X",
    "SIKHYE": "Y",
}


@dataclass(frozen=True)
class ProductCanonicalProjection:
    """
    Aggregated Product Detail values for one canonical Product column.

    sales_quantity is None when at least one contributing source row
    has no parseable quantity.

    We intentionally do not coerce missing quantity to zero.
    """

    canonical_key: str
    column_letter: str
    sales_amount: int
    sales_quantity: int | None
    source_names: tuple[str, ...]
    record_count: int
    quantity_missing_count: int


@dataclass(frozen=True)
class ProductUnmappedRecord:
    """
    Raw Product Detail row that could not be mapped to one of the
    approved 21 canonical product groups.
    """

    record: ProductDetailSalesRecord
    reason: str = "NO_CANONICAL_ALIAS"


@dataclass(frozen=True)
class ProductProjectionResult:
    """
    Sparse projection result.

    Only canonical groups actually present in the source are emitted.
    Missing canonical groups are NOT automatically converted to zero.
    """

    mapped: tuple[ProductCanonicalProjection, ...]
    unmapped: tuple[ProductUnmappedRecord, ...]

    @property
    def mapped_record_count(self) -> int:
        return sum(
            item.record_count
            for item in self.mapped
        )

    @property
    def unmapped_record_count(self) -> int:
        return len(self.unmapped)

    @property
    def mapped_sales_total(self) -> int:
        return sum(
            item.sales_amount
            for item in self.mapped
        )

    def by_canonical_key(
        self,
    ) -> dict[str, ProductCanonicalProjection]:
        return {
            item.canonical_key: item
            for item in self.mapped
        }


@dataclass
class _MutableBucket:
    sales_amount: int = 0
    sales_quantity: int = 0
    record_count: int = 0
    quantity_missing_count: int = 0
    source_names: set[str] | None = None

    def __post_init__(self) -> None:
        if self.source_names is None:
            self.source_names = set()


def project_product_detail(
    records: Sequence[ProductDetailSalesRecord],
) -> ProductProjectionResult:
    """
    Project raw Product Detail rows into the existing 21 canonical
    product groups.

    Contract:
    - raw Product Detail collection remains classification-independent.
    - existing Menu Alias Registry is reused only for product-name
      canonicalization.
    - Menu category whitelist is NOT applied here.
    - sales values preserve their sign.
    - quantity values preserve their sign.
    - None quantity is never coerced to zero.
    - unmapped product rows remain explicit.
    - raw '기타' is never routed to Excel Z.
    """

    buckets: dict[str, _MutableBucket] = {}
    unmapped: list[ProductUnmappedRecord] = []

    for record in records:
        canonical_key = canonical_key_for_menu(
            record.product_name
        )

        # Alias may theoretically resolve to a registry key that is not
        # part of this workbook's physical 21-column contract.
        if (
            canonical_key is None
            or canonical_key not in PRODUCT_CANONICAL_COLUMNS
        ):
            unmapped.append(
                ProductUnmappedRecord(
                    record=record,
                )
            )
            continue

        bucket = buckets.setdefault(
            canonical_key,
            _MutableBucket(),
        )

        bucket.sales_amount += record.sales_amount
        bucket.record_count += 1

        assert bucket.source_names is not None
        bucket.source_names.add(
            record.product_name
        )

        if record.sales_quantity is None:
            bucket.quantity_missing_count += 1
        else:
            bucket.sales_quantity += (
                record.sales_quantity
            )

    mapped: list[ProductCanonicalProjection] = []

    # Output follows workbook physical column order E:Y rather than
    # raw source order.
    for canonical_key, column_letter in (
        PRODUCT_CANONICAL_COLUMNS.items()
    ):
        bucket = buckets.get(
            canonical_key
        )

        if bucket is None:
            continue

        quantity: int | None

        if bucket.quantity_missing_count:
            quantity = None
        else:
            quantity = bucket.sales_quantity

        assert bucket.source_names is not None

        mapped.append(
            ProductCanonicalProjection(
                canonical_key=canonical_key,
                column_letter=column_letter,
                sales_amount=bucket.sales_amount,
                sales_quantity=quantity,
                source_names=tuple(
                    sorted(bucket.source_names)
                ),
                record_count=bucket.record_count,
                quantity_missing_count=(
                    bucket.quantity_missing_count
                ),
            )
        )

    return ProductProjectionResult(
        mapped=tuple(mapped),
        unmapped=tuple(unmapped),
    )
