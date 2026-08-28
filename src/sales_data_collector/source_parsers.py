from __future__ import annotations

import re
from datetime import date
from dataclasses import dataclass
from enum import Enum

from bs4 import BeautifulSoup

from .exceptions import ParseError
from .models import (
    DailySalesRecord,
    DeliveryChannelRecord,
    DeliverySalesResult,
    MonthlySalesRecord,
    PeriodSalesResult,
    ProductSalesResult,
    TodayStoreSalesResult,
)
from .utils import clean_int


_DATE = re.compile(r"(\d{2})/(\d{2})\[")
_MONTH = re.compile(r"(\d{4})/(\d{2})")
_NUMBER = re.compile(r"-?[\d,]+")


def _text(node) -> str:
    return node.get_text(" ", strip=True) if node else ""


def _numbers(text: str) -> list[int]:
    return [clean_int(value) for value in _NUMBER.findall(text)]


def _summary_value(soup: BeautifulSoup, label: str) -> int | None:
    for item in soup.select(".detail_list li, .detail_list2 li, .detail_list3 li"):
        value = _text(item)
        if label in value:
            numbers = _numbers(value)
            return (numbers[0] if label == "\uac74\uc218/\uac1d\uc218" else numbers[-1]) if numbers else None
    return None


class DailySalesParser:
    def parse(self, html: str, *, year: int, month: int) -> list[DailySalesRecord]:
        soup = BeautifulSoup(html, "html.parser")
        rows: list[DailySalesRecord] = []
        for block in soup.select(".detail li"):
            label = _text(block.select_one(":scope > .detail_title"))
            matched = _DATE.search(label)
            detail = block.select_one(":scope > .detail_1")
            if not matched or detail is None:
                continue
            day = int(matched.group(2))
            text = _text(detail)
            values = _numbers(text)
            if text == "0":
                rows.append(DailySalesRecord(date(year, month, day), 0, 0, status="SUCCESS_ZERO"))
                continue
            counts = _numbers(label)
            if not counts or not values:
                raise ParseError(f"daily row is incomplete: {label}")
            card = self._labelled_amount(text, "\uce74\ub4dc")
            cash = self._labelled_amount(text, "\ud604\uae08")
            rows.append(DailySalesRecord(date(year, month, day), counts[-1], values[-1], cash, card))
        if not rows:
            raise ParseError("daily response contains no date rows")
        return rows

    @staticmethod
    def _labelled_amount(text: str, label: str) -> int | None:
        matched = re.search(label + r"\s*([\d,]+)", text)
        return clean_int(matched.group(1)) if matched else None


class PeriodSalesParser:
    def parse(self, html: str) -> PeriodSalesResult:
        soup = BeautifulSoup(html, "html.parser")
        count = _summary_value(soup, "건수/객수")
        sales = _summary_value(soup, "\ub9e4\ucd9c\ud569\uacc4")
        if count is None or sales is None:
            labels = [
                _text(item.select_one(".tit"))
                for item in soup.select(".detail_list li, .detail_list2 li, .detail_list3 li")
            ]
            found = ", ".join(label for label in labels if label)[:300]
            raise ParseError(f"period summary is missing (labels={found or 'none'})")
        return PeriodSalesResult(
            count,
            sales,
            _summary_value(soup, "\ud604\uae08"),
            _summary_value(soup, "\uce74\ub4dc"),
            _summary_value(soup, "\ud640\ub9e4\ucd9c"),
            _summary_value(soup, "\ubc30\ub2ec\ub9e4\ucd9c"),
        )


class MonthlySalesParser:
    def parse(self, html: str) -> list[MonthlySalesRecord]:
        soup = BeautifulSoup(html, "html.parser")
        result: list[MonthlySalesRecord] = []
        for block in soup.select(".detail li"):
            title = block.select_one(":scope > .detail_title")
            date_label = _text(title.select_one("dt") if title else None)
            matched = _MONTH.fullmatch(date_label)
            detail = block.select_one(":scope > .detail_1")
            if not matched or detail is None:
                continue
            counts = _numbers(_text(title))
            amounts = _numbers(_text(detail))
            if counts and amounts:
                result.append(MonthlySalesRecord(int(matched.group(1)), int(matched.group(2)), counts[-1], amounts[-1]))
        if not result:
            raise ParseError("monthly response contains no month rows")
        return result


class ProductSalesParser:
    def parse(self, html: str) -> ProductSalesResult:
        soup = BeautifulSoup(html, "html.parser")
        count = _summary_value(soup, "\uac74\uc218")
        sales = _summary_value(soup, "\ub9e4\ucd9c\ud569\uacc4")
        if count is None or sales is None:
            raise ParseError("product summary is missing")
        return ProductSalesResult(product_count=count, sales_amount=sales)


class TodayStoreSalesParser:
    """Parse the authoritative gross-sales field from the configured detail page."""

    def parse(self, html: str) -> TodayStoreSalesResult:
        soup = BeautifulSoup(html, "html.parser")
        gross = self._summary_value(soup, "총 매출")
        count = self._summary_value(soup, "매출건수")
        if gross is None or count is None:
            raise ParseError("today-store gross sales or receipt count labels are missing")
        return TodayStoreSalesResult(receipt_count=count, gross_sales_amount=gross)

    @staticmethod
    def _summary_value(soup: BeautifulSoup, label: str) -> int | str | None:
        for item in soup.select(".detail_list li, .detail_list2 li, .detail_list3 li"):
            title = _text(item.select_one(".tit")).replace(" ", "")
            if title != label.replace(" ", ""):
                continue
            values = _numbers(_text(item.select_one(".txt")))
            return values[-1] if values else "-"
        return None

    def parse_gross_sales(self, html: str) -> int:
        return self.parse(html).gross_sales_amount


class DeliverySalesParser:
    def parse(self, html: str, *, store_id: str, start_date: date, end_date: date) -> DeliverySalesResult:
        soup = BeautifulSoup(html, "html.parser")
        orders = soup.select(".order")
        if not orders:
            raise ParseError("delivery response contains no channel rows")
        total = _numbers(_text(orders[0]))
        if len(total) < 2:
            raise ParseError("delivery summary is incomplete")
        channels: list[DeliveryChannelRecord] = []
        for order in orders[1:]:
            values = _numbers(_text(order))
            label = _text(order.select_one(".tit"))
            if label and len(values) >= 2:
                channels.append(DeliveryChannelRecord(label, values[0], values[-1]))
        return DeliverySalesResult(store_id, start_date, end_date, total[0], total[-1], channels)


class CancelDiscountParser:
    """Parse the cancellation/discount page without treating cancellations as sales."""

    def parse_discount_amount(self, html: str) -> int:
        soup = BeautifulSoup(html, "html.parser")
        text = _text(soup.select_one(".detail_list2"))
        matched = re.search("\ud560\uc778\\s+([\\d,]+)\\s+([\\d,]+)", text)
        if matched:
            return clean_int(matched.group(2))
        raise ParseError("discount summary is missing")

    def parse_daily_discount_amounts(self, html: str) -> dict[date, int]:
        """Return only discount entries; cancellation amounts are deliberately excluded."""
        soup = BeautifulSoup(html, "html.parser")
        text = _text(soup.select_one(".detail_list2"))
        summaries = list(re.finditer("\ud560\uc778\\s+\\d+\\s+[\\d,]+", text))
        if not summaries:
            raise ParseError("discount section is missing")
        section = text[summaries[-1].end():]
        totals: dict[date, int] = {}
        for year, month, day, amount in re.findall(r"(\d{4})/(\d{2})/(\d{2})\s+\d+\s+([\d,]+)", section):
            key = date(int(year), int(month), int(day))
            totals[key] = totals.get(key, 0) + clean_int(amount)
        return totals


class DiscountStatus(str, Enum):
    VALUE = "DISCOUNT_VALUE"
    ZERO = "DISCOUNT_ZERO"
    MISSING = "DISCOUNT_MISSING"
    PARSE_ERROR = "DISCOUNT_PARSE_ERROR"


@dataclass(frozen=True)
class DiscountResult:
    amount: int | None
    status: DiscountStatus


def calculate_gross_sales(net_sales_amount: int | None, discount: DiscountResult) -> int | None:
    if net_sales_amount is None or discount.status not in {DiscountStatus.VALUE, DiscountStatus.ZERO}:
        return None
    return net_sales_amount + (discount.amount or 0)


def apply_discount_to_daily(record: DailySalesRecord, discount_amount: int) -> DailySalesRecord:
    if record.sales_amount is None:
        raise ParseError("daily sales amount is missing")
    return DailySalesRecord(
        business_date=record.business_date,
        receipt_count=record.receipt_count,
        sales_amount=record.sales_amount,
        cash_amount=record.cash_amount,
        card_amount=record.card_amount,
        discount_amount=discount_amount,
        gross_sales_amount=record.sales_amount + discount_amount,
        status=record.status,
    )
