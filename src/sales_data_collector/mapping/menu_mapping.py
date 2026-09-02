from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from enum import Enum
from typing import Iterable

from ..models import MenuMonthlySalesResult, MenuSalesRecord


class MenuRowType(str, Enum):
    MENU = "MENU"
    OPTION = "OPTION"


class MenuMappingStatus(str, Enum):
    MAPPED = "MAPPED"
    AMBIGUOUS = "AMBIGUOUS"
    UNMAPPED = "UNMAPPED"
    NOT_APPLICABLE = "NOT_APPLICABLE"


@dataclass(frozen=True)
class MenuMappingResult:
    record: MenuSalesRecord
    normalized_name: str
    row_type: MenuRowType
    status: MenuMappingStatus
    canonical_code: str | None
    reason: str
    parent_raw_menu: str | None = None


@dataclass(frozen=True)
class CanonicalMenuAggregate:
    canonical_code: str
    aliases: tuple[str, ...]
    quantity: int
    sales_amount: int


@dataclass(frozen=True)
class MenuMappingPreview:
    source: MenuMonthlySalesResult
    mappings: tuple[MenuMappingResult, ...]
    aggregates: tuple[CanonicalMenuAggregate, ...]

    @property
    def raw_row_count(self) -> int:
        return len(self.mappings)

    @property
    def menu_row_count(self) -> int:
        return sum(1 for item in self.mappings if item.row_type == MenuRowType.MENU)

    @property
    def option_row_count(self) -> int:
        return sum(1 for item in self.mappings if item.row_type == MenuRowType.OPTION)

    def count_by_status(self, status: MenuMappingStatus) -> int:
        return sum(1 for item in self.mappings if item.status == status)

    def sales_by_status(self, status: MenuMappingStatus) -> int:
        return sum(item.record.sales_amount for item in self.mappings if item.status == status)

    @property
    def total_classified_sales(self) -> int:
        return sum(item.record.sales_amount for item in self.mappings)


ALIASES: dict[str, str] = {
    "꼬마김밥 5줄": "KIMBAP_5",
    "선비꼬마김밥 5줄": "KIMBAP_5",
    "꼬마김밥 10줄": "KIMBAP_10",
    "어묵탕": "FISH_CAKE_SOUP",
    "선비우동": "UDON",
    "라면": "RAMEN",
    "국물떡볶이(순한맛)": "TTEOKBOKKI_MILD",
    "국물떡볶이(매운맛)": "TTEOKBOKKI_SPICY",
    "청양고추소스": "SAUCE_CHEONGYANG",
    "청양고추소스(20g)": "SAUCE_CHEONGYANG",
    "스리라차마요소스": "SAUCE_SRIRACHA_MAYO",
    "스리라차 마요소스(20g)": "SAUCE_SRIRACHA_MAYO",
}


AMBIGUOUS: dict[str, tuple[str, ...]] = {
    "생수": ("DRINK",),
    "식혜": ("DRINK", "SIKHYE"),
    "코카콜라": ("DRINK",),
    "제로 펩시 라임 (355ml)": ("DRINK",),
    "칠성사이다 (355ml)": ("DRINK",),
    "(HACCP 인증) 선비식혜 150ml": ("DRINK", "SIKHYE_150ML"),
    "국물떡볶이": ("TTEOKBOKKI_MILD", "TTEOKBOKKI_SPICY", "TTEOKBOKKI"),
    "국물떡볶이 (순한맛,매운맛)": ("TTEOKBOKKI_MILD", "TTEOKBOKKI_SPICY", "TTEOKBOKKI"),
    "국물떡볶이(순한맛)-밀키트": ("TTEOKBOKKI_MILKIT_MILD",),
    "선비국물떡볶이 밀키트 (냉동보관)": ("TTEOKBOKKI_MILKIT",),
    "선비우동 밀키트(냉동보관)": ("UDON_MILKIT",),
}


def normalize_menu_name(value: str) -> str:
    normalized = unicodedata.normalize("NFKC", value)
    return re.sub(r"\s+", " ", normalized.replace("\u00a0", " ")).strip()


def classify_menu_record(record: MenuSalesRecord) -> MenuMappingResult:
    normalized = normalize_menu_name(record.menu_name)
    if normalized.startswith("->"):
        return MenuMappingResult(
            record=record,
            normalized_name=normalized,
            row_type=MenuRowType.OPTION,
            status=MenuMappingStatus.NOT_APPLICABLE,
            canonical_code=None,
            reason="OPTION_ROW",
        )
    if normalized in ALIASES:
        return MenuMappingResult(
            record=record,
            normalized_name=normalized,
            row_type=MenuRowType.MENU,
            status=MenuMappingStatus.MAPPED,
            canonical_code=ALIASES[normalized],
            reason="EXPLICIT_ALIAS",
        )
    if normalized in AMBIGUOUS:
        return MenuMappingResult(
            record=record,
            normalized_name=normalized,
            row_type=MenuRowType.MENU,
            status=MenuMappingStatus.AMBIGUOUS,
            canonical_code=None,
            reason="BUSINESS_RULE_REQUIRED:" + ",".join(AMBIGUOUS[normalized]),
        )
    return MenuMappingResult(
        record=record,
        normalized_name=normalized,
        row_type=MenuRowType.MENU,
        status=MenuMappingStatus.UNMAPPED,
        canonical_code=None,
        reason="NO_EXPLICIT_ALIAS",
    )


def build_menu_mapping_preview(source: MenuMonthlySalesResult) -> MenuMappingPreview:
    mappings = tuple(classify_menu_record(record) for record in source.records)
    return MenuMappingPreview(source=source, mappings=mappings, aggregates=_aggregate_mapped(mappings))


def _aggregate_mapped(mappings: Iterable[MenuMappingResult]) -> tuple[CanonicalMenuAggregate, ...]:
    by_code: dict[str, dict[str, object]] = {}
    for item in mappings:
        if item.status != MenuMappingStatus.MAPPED or not item.canonical_code:
            continue
        bucket = by_code.setdefault(item.canonical_code, {"aliases": set(), "quantity": 0, "sales_amount": 0})
        aliases = bucket["aliases"]
        if isinstance(aliases, set):
            aliases.add(item.normalized_name)
        bucket["quantity"] = int(bucket["quantity"]) + (item.record.sales_quantity or 0)
        bucket["sales_amount"] = int(bucket["sales_amount"]) + item.record.sales_amount
    return tuple(
        CanonicalMenuAggregate(
            canonical_code=code,
            aliases=tuple(sorted(data["aliases"])) if isinstance(data["aliases"], set) else (),
            quantity=int(data["quantity"]),
            sales_amount=int(data["sales_amount"]),
        )
        for code, data in sorted(by_code.items())
    )
