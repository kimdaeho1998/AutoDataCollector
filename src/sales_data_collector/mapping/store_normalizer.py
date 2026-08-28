from __future__ import annotations

from dataclasses import dataclass
from enum import Enum


class StoreMatchStatus(str, Enum):
    MATCHED = "MATCHED"
    UNMATCHED = "UNMATCHED"
    AMBIGUOUS = "AMBIGUOUS"


def normalize_store_name(name: str, brand_prefix: str = "\uc120\ube44\uaf2c\ub9c8\uae40\ubc25") -> str:
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
