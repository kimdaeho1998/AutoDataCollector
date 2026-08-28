from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from datetime import date
from typing import Optional


class SalesStatus(str, Enum):
    SUCCESS_ZERO = "SUCCESS_ZERO"
    SUCCESS_DATA = "SUCCESS_DATA"
    NO_DATA = "NO_DATA"
    PARSE_ERROR = "PARSE_ERROR"
    HTTP_ERROR = "HTTP_ERROR"


class CollectionMode(str, Enum):
    ADMIN = "admin"
    DEVELOPER = "developer"


class ExportFormat(str, Enum):
    SUMMARY = "summary"
    SUMMARY_AND_TIMES = "summary_and_times"


@dataclass(frozen=True)
class Store:
    magic_store_id: str
    store_name: str


@dataclass(frozen=True)
class SalesRecord:
    business_date: date
    magic_store_id: str
    store_name: str
    order_count: int
    sales_amount: int
    time_slot: Optional[str] = None
    source_status: SalesStatus = SalesStatus.SUCCESS_DATA


@dataclass
class SalesResult:
    business_date: date
    magic_store_id: str
    store_name: str
    order_count: int = 0
    sales_amount: int = 0
    time_slots: list["TimeSlotRecord"] = field(default_factory=list)
    source_status: SalesStatus = SalesStatus.NO_DATA


@dataclass(frozen=True)
class TimeSlotRecord:
    business_date: date
    magic_store_id: str
    store_name: str
    time_slot: str
    order_count: int
    sales_amount: int


@dataclass(frozen=True)
class DailySalesRecord:
    business_date: date
    receipt_count: int | None
    sales_amount: int | None
    cash_amount: int | None = None
    card_amount: int | None = None
    discount_amount: int | None = None
    gross_sales_amount: int | None = None
    status: SalesStatus = SalesStatus.SUCCESS_DATA


@dataclass(frozen=True)
class PeriodSalesResult:
    receipt_count: int
    sales_amount: int
    cash_amount: int | None = None
    card_amount: int | None = None
    hall_sales_amount: int | None = None
    delivery_sales_amount: int | None = None


@dataclass(frozen=True)
class TodayStoreSalesResult:
    receipt_count: int | str
    gross_sales_amount: int | str


@dataclass(frozen=True)
class SalesAdminDailyRecord:
    """The two source-of-truth values written to the sales-admin workbook."""

    business_date: date
    store_id: str
    store_name: str
    receipt_count: int | str
    gross_sales_amount: int | str


@dataclass(frozen=True)
class MonthlySalesRecord:
    year: int
    month: int
    receipt_count: int
    sales_amount: int


@dataclass(frozen=True)
class ProductSalesResult:
    product_count: int
    sales_amount: int


@dataclass(frozen=True)
class DeliveryChannelRecord:
    channel_name: str
    order_count: int | None
    sales_amount: int


@dataclass(frozen=True)
class DeliverySalesResult:
    store_id: str
    start_date: date
    end_date: date
    total_count: int
    total_sales: int
    channels: list[DeliveryChannelRecord]
