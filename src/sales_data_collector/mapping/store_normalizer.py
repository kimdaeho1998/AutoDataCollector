from __future__ import annotations

import os
from dataclasses import dataclass
from enum import Enum


class StoreMatchStatus(str, Enum):
    MATCHED = "MATCHED"
    UNMATCHED = "UNMATCHED"
    AMBIGUOUS = "AMBIGUOUS"


def normalize_store_name(name: str, brand_prefix: str | None = None) -> str:
    brand_prefix = brand_prefix if brand_prefix is not None else os.environ.get("COLLECTOR_BRAND_PREFIX", "")
    value = " ".join(name.split())
    return value[len(brand_prefix):].strip() if value.startswith(brand_prefix) else value


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
