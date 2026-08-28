from __future__ import annotations

from datetime import date, datetime, timedelta
from typing import Iterable


def parse_ymd(value: str) -> date:
    return datetime.strptime(value, "%Y-%m-%d").date()


def format_ymd(value: date) -> str:
    return value.strftime("%Y-%m-%d")


def date_range(start: date, end: date) -> list[date]:
    if end < start:
        raise ValueError("end date must be on or after start date")
    days: list[date] = []
    current = start
    while current <= end:
        days.append(current)
        current += timedelta(days=1)
    return days


def clean_int(value: str | int | None) -> int:
    if value is None:
        return 0
    if isinstance(value, int):
        return value
    digits = "".join(ch for ch in str(value) if ch.isdigit() or ch == "-")
    if digits in {"", "-"}:
        return 0
    return int(digits)


def first_nonempty(values: Iterable[str | None]) -> str | None:
    for value in values:
        if value is not None and str(value).strip():
            return str(value).strip()
    return None


