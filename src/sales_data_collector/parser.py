from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from typing import Iterable

from bs4 import BeautifulSoup

from .exceptions import ParseError
from .models import SalesResult, SalesStatus, TimeSlotRecord
from .utils import clean_int, first_nonempty


_DATE_LABEL_PATTERNS = [
    re.compile(r"(?:order|orders|주문|건수)\D{0,20}(\d[\d,]*)", re.I),
    re.compile(r"(?:sales|amount|매출|금액)\D{0,20}([-\d,]+)", re.I),
]


class ServiceSalesParser:
    """Parse service HTML into canonical sales models."""

    def parse_sales_page(
        self,
        html: str,
        *,
        business_date: date,
        magic_store_id: str,
        store_name: str,
    ) -> SalesResult:
        soup = BeautifulSoup(html, "html.parser")
        summary = self._parse_summary(soup)
        time_slots = self._parse_time_slots(soup, business_date, magic_store_id, store_name)

        if summary is None and not time_slots:
            raise ParseError("could not find sales summary or time-slot table in response")

        order_count = summary[0] if summary else sum(slot.order_count for slot in time_slots)
        sales_amount = summary[1] if summary else sum(slot.sales_amount for slot in time_slots)
        status = SalesStatus.SUCCESS_ZERO if order_count == 0 and sales_amount == 0 else SalesStatus.SUCCESS_DATA

        return SalesResult(
            business_date=business_date,
            magic_store_id=magic_store_id,
            store_name=store_name,
            order_count=order_count,
            sales_amount=sales_amount,
            time_slots=time_slots,
            source_status=status,
        )

    def _parse_summary(self, soup: BeautifulSoup) -> tuple[int, int] | None:
        text = soup.get_text("\n", strip=True)
        order_count = self._extract_labelled_int(text, ("order", "orders", "주문", "건수", "총건수", "총 건수"))
        sales_amount = self._extract_labelled_int(text, ("sales", "amount", "매출", "금액", "총매출", "총 매출"))

        if order_count is None and sales_amount is None:
            return None

        return (order_count or 0, sales_amount or 0)

    def _extract_labelled_int(self, text: str, labels: Iterable[str]) -> int | None:
        lowered = text.lower()
        for label in labels:
            idx = lowered.find(label.lower())
            if idx == -1:
                continue
            window = text[idx : idx + 80]
            match = re.search(r"(-?[\d,]+)", window)
            if match:
                return clean_int(match.group(1))
        return None

    def _parse_time_slots(
        self,
        soup: BeautifulSoup,
        business_date: date,
        magic_store_id: str,
        store_name: str,
    ) -> list[TimeSlotRecord]:
        rows: list[TimeSlotRecord] = []
        tables = soup.find_all("table")
        for table in tables:
            headers = [self._cell_text(th) for th in table.find_all("th")]
            if not headers:
                continue

            normalized = [header.lower() for header in headers]
            if not any("시간" in h or "time" in h for h in normalized):
                continue

            time_idx = self._find_header_index(headers, ("time", "시간", "시각"))
            order_idx = self._find_header_index(headers, ("order", "orders", "건수", "주문"))
            sales_idx = self._find_header_index(headers, ("sales", "amount", "매출", "금액"))

            body_rows = table.find_all("tr")
            for tr in body_rows[1:]:
                cells = [self._cell_text(td) for td in tr.find_all(["td", "th"])]
                if not cells:
                    continue
                time_slot = self._safe_cell(cells, time_idx)
                if not time_slot:
                    continue
                order_count = clean_int(self._safe_cell(cells, order_idx))
                sales_amount = clean_int(self._safe_cell(cells, sales_idx))
                rows.append(
                    TimeSlotRecord(
                        business_date=business_date,
                        magic_store_id=magic_store_id,
                        store_name=store_name,
                        time_slot=time_slot,
                        order_count=order_count,
                        sales_amount=sales_amount,
                    )
                )
            if rows:
                break
        return rows

    @staticmethod
    def _cell_text(cell) -> str:
        return cell.get_text(" ", strip=True) if cell else ""

    @staticmethod
    def _find_header_index(headers: list[str], keywords: tuple[str, ...]) -> int | None:
        for idx, header in enumerate(headers):
            lowered = header.lower()
            if any(keyword.lower() in lowered for keyword in keywords):
                return idx
        return None

    @staticmethod
    def _safe_cell(cells: list[str], idx: int | None) -> str:
        if idx is None or idx < 0 or idx >= len(cells):
            return ""
        return cells[idx]
