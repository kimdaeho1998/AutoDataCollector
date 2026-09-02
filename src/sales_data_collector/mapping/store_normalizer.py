from __future__ import annotations

import os
import unicodedata
from dataclasses import dataclass
from enum import Enum


class StoreMatchStatus(str, Enum):
    MATCHED = "MATCHED"
    UNMATCHED = "UNMATCHED"
    AMBIGUOUS = "AMBIGUOUS"


def _normalize_text(value: str) -> str:
    normalized = unicodedata.normalize("NFKC", value or "")
    return " ".join(normalized.split()).strip()


def _configured_brand_prefix(explicit: str | None = None) -> str:
    if explicit is not None:
        return _normalize_text(explicit)

    configured = (
        os.environ.get("COLLECTOR_BRAND_PREFIX", "").strip()
        or os.environ.get("COLLECTOR_BRAND_NAME", "").strip()
    )

    return _normalize_text(configured)


def normalize_store_name(
    name: str,
    brand_prefix: str | None = None,
) -> str:
    value = _normalize_text(name)
    prefix = _configured_brand_prefix(brand_prefix)

    if prefix and value.startswith(prefix):
        remainder = value[len(prefix):].strip()

        # Do not normalize the brand itself to an empty key.
        if remainder:
            return remainder

    return value


@dataclass(frozen=True)
class StoreResolution:
    magic_store_id: str
    magic_store_name: str
    normalized_name: str
    excel_store_name: str | None
    sheet_name: str | None
    row: int | None
    status: StoreMatchStatus
    method: str | None
