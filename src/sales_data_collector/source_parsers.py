from __future__ import annotations

import re
from datetime import date
from dataclasses import dataclass
from enum import Enum

from bs4 import BeautifulSoup

from .exceptions import ParseError
from .mapping.menu_category_policy import (
    MenuCategoryDecision,
    resolve_menu_category,
)
from .models import (
    DailySalesRecord,
    DeliveryChannelRecord,
    DeliverySalesResult,
    MenuMonthlySalesResult,
    MenuSalesRecord,
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


class MenuSalesParser:

    """Parse raw menu sales rows without applying canonical menu mapping."""

    menu_keywords = ("메뉴", "상품")
    quantity_keywords = ("수량", "건수", "개수")
    amount_keywords = ("매출", "금액", "합계")
    empty_keywords = ("조회된 데이터가 없습니다", "데이터가 없습니다", "검색 결과가 없습니다")
    detail_row_selectors = (".detail > li", ".detail2 > li", ".detail3 > li")
    detail_group_selector = ".detail_title"

    def parse(self, html: str, *, store_id: str, store_name: str, period_start: date, period_end: date) -> MenuMonthlySalesResult:
        soup = BeautifulSoup(html, "html.parser")
        records = self._parse_detail_rows(soup, store_id, store_name, period_start, period_end)
        if not records:
            records = self._parse_tables(soup, store_id, store_name, period_start, period_end)
        # MENU_MONTHLY totals are intentionally recomputed from
        # category-filtered records.
        #
        # Never use the page-level product summary here because
        # that summary may contain delivery-platform sales.
        source_total_sales = sum(
            record.sales_amount
            for record in records
        )
        source_total_quantity = sum(
            (record.sales_quantity or 0)
            for record in records
        )
        if not records and self._is_empty_result(soup):
            return MenuMonthlySalesResult(store_id, store_name, period_start, period_end, [], source_total_sales, source_total_quantity)
        if not records:
            raise ParseError("menu sales response contains no parseable menu rows")
        return MenuMonthlySalesResult(store_id, store_name, period_start, period_end, records, source_total_sales, source_total_quantity)

    def _parse_detail_rows(self, soup: BeautifulSoup, store_id: str, store_name: str, period_start: date, period_end: date) -> list[MenuSalesRecord]:
        records = self._parse_detail_title_children(soup, store_id, store_name, period_start, period_end)
        if records:
            return records

        records: list[MenuSalesRecord] = []
        for selector in self.detail_row_selectors:
            for block in soup.select(selector):
                title = self._detail_title(block)
                body = self._detail_body(block)
                if not title or self._looks_like_total(title) or body is None:
                    continue
                numbers = _numbers(_text(body))
                if not numbers:
                    continue
                unit_price = numbers[-3] if len(numbers) >= 3 else None
                sales_amount = numbers[-1]
                sales_quantity = numbers[-2] if len(numbers) >= 2 else None
                records.append(
                    MenuSalesRecord(
                        store_id=store_id,
                        store_name=store_name,
                        period_start=period_start,
                        period_end=period_end,
                        menu_name=title,
                        sales_quantity=sales_quantity,
                        sales_amount=sales_amount,
                        unit_price=unit_price,
                    )
                )
            if records:
                break
        return records

    def _parse_detail_title_children(self, soup: BeautifulSoup, store_id: str, store_name: str, period_start: date, period_end: date) -> list[MenuSalesRecord]:
        records: list[MenuSalesRecord] = []

        # MagicERP product.asp authoritative DOM contract:
        #
        #   <li>
        #       <div class="detail_t">
        #           ... first <font> = classification/category ...
        #       </div>
        #
        #       <div class="detail_title detail_title2">
        #           <div>
        #               <p class="menu_nm">...</p>
        #               <div class="info">unit qty sales</div>
        #           </div>
        #           ...
        #       </div>
        #   </li>
        #
        # Category belongs to the GROUP, not to menu_name.

        for group in soup.select(
            ".detail_title"
        ):

            # ----------------------------------------------------------------------------------------
            # CATEGORY HEADER
            # ----------------------------------------------------------------------------------------

            header = group.find_previous_sibling(
                "div",
                class_="detail_t",
            )

            if header is None:
                # Fail closed:
                # never place an unclassified source group
                # into MENU_MONTHLY.
                continue

            category_node = header.find(
                "font"
            )

            if category_node is None:
                continue

            classification_name = _text(
                category_node
            ).strip()

            if not classification_name:
                continue

            resolution = resolve_menu_category(
                classification_name
            )

            if (
                resolution.decision
                != MenuCategoryDecision.INCLUDE
            ):
                # SOURCE_EXCLUDED:
                # delivery/platform/set/unknown/etc.
                continue


            # ----------------------------------------------------------------------------------------
            # PRODUCT ROWS
            # ----------------------------------------------------------------------------------------

            for item in group.find_all(
                "div",
                recursive=False,
            ):

                menu_node = item.select_one(
                    "p.menu_nm"
                )

                info_node = item.select_one(
                    "div.info"
                )

                if (
                    menu_node is None
                    or info_node is None
                ):
                    continue


                menu_name = _text(
                    menu_node
                ).strip()

                info_text = _text(
                    info_node
                ).strip()


                if (
                    not menu_name
                    or not info_text
                ):
                    continue


                if self._looks_like_total(
                    menu_name
                ):
                    continue


                # info contract:
                #
                #   unit_price
                #   sales_quantity
                #   actual_sales_amount
                #
                # Important:
                # actual sales is authoritative.
                # Do NOT calculate unit_price * quantity.

                numeric_matches = list(
                    _NUMBER.finditer(
                        info_text
                    )
                )

                if len(
                    numeric_matches
                ) < 3:
                    continue


                unit_price_match = (
                    numeric_matches[-3]
                )

                quantity_match = (
                    numeric_matches[-2]
                )

                sales_match = (
                    numeric_matches[-1]
                )


                def _number_to_int(
                    match,
                ) -> int:

                    return int(
                        match.group(0)
                        .replace(",", "")
                    )


                unit_price = _number_to_int(
                    unit_price_match
                )

                sales_quantity = _number_to_int(
                    quantity_match
                )

                sales_amount = _number_to_int(
                    sales_match
                )


                records.append(
                    MenuSalesRecord(
                        store_id=store_id,
                        store_name=store_name,
                        period_start=period_start,
                        period_end=period_end,
                        menu_name=menu_name,
                        sales_amount=sales_amount,
                        sales_quantity=sales_quantity,
                        unit_price=unit_price,
                        classification_name=classification_name,
                    )
                )

        return records
    @staticmethod
    def _split_classification_and_menu_name(
        prefix: str,
    ) -> tuple[str, str] | None:
        """Split flattened MagicERP row prefix using the F8 category policy."""

        compact = " ".join(
            str(prefix).split()
        ).strip()

        if not compact:
            return None

        boundaries = [
            match.start()
            for match in re.finditer(
                r"\s+",
                compact,
            )
        ]

        # Longest valid category prefix wins.
        for boundary in reversed(
            boundaries
        ):

            classification = (
                compact[:boundary]
                .strip()
            )

            menu_name = (
                compact[boundary:]
                .strip()
            )

            if (
                not classification
                or not menu_name
            ):
                continue

            resolution = (
                resolve_menu_category(
                    classification
                )
            )

            if (
                resolution.decision
                == MenuCategoryDecision.INCLUDE
            ):
                return (
                    classification,
                    menu_name,
                )

        return None


    def _record_from_menu_item_text(
        self,
        text: str,
        *,
        store_id: str,
        store_name: str,
        period_start: date,
        period_end: date,
    ) -> MenuSalesRecord | None:
        text = " ".join(text.split())
        if not text:
            return None
        numeric_matches = list(_NUMBER.finditer(text))
        if len(numeric_matches) < 3:
            return None
        unit_price_match, quantity_match, sales_match = numeric_matches[-3:]

        prefix = text[:unit_price_match.start()].strip()

        if not prefix:
            return None

        split_result = (
            self._split_classification_and_menu_name(
                prefix
            )
        )

        if split_result is None:
            return None

        (
            classification_name,
            menu_name,
        ) = split_result

        if (
            not menu_name
            or self._looks_like_total(
                menu_name
            )
        ):
            return None

        return MenuSalesRecord(
            store_id=store_id,
            store_name=store_name,
            period_start=period_start,
            period_end=period_end,
            menu_name=menu_name,
            sales_quantity=clean_int(quantity_match.group(0)),
            sales_amount=clean_int(sales_match.group(0)),
            unit_price=clean_int(unit_price_match.group(0)),
            classification_name=classification_name,
        )

    @staticmethod
    def _unique_nodes(nodes):
        seen: set[int] = set()
        for node in nodes:
            key = id(node)
            if key in seen:
                continue
            seen.add(key)
            yield node

    @staticmethod
    def _detail_title(block) -> str:
        node = block.select_one(":scope > .detail_title, :scope > .detail_title2, :scope > .tit")
        if node is not None:
            return _text(node)
        children = block.find_all(recursive=False)
        return _text(children[0]) if children else ""

    @staticmethod
    def _detail_body(block):
        node = block.select_one(":scope > .detail_1, :scope > .detail_2, :scope > .detail_3, :scope > .txt")
        if node is not None:
            return node
        children = block.find_all(recursive=False)
        return children[1] if len(children) > 1 else None

    def _parse_tables(self, soup: BeautifulSoup, store_id: str, store_name: str, period_start: date, period_end: date) -> list[MenuSalesRecord]:
        records: list[MenuSalesRecord] = []
        for table in soup.find_all("table"):
            rows = table.find_all("tr")
            if len(rows) < 2:
                continue
            headers = [self._cell_text(cell) for cell in rows[0].find_all(["th", "td"])]
            indexes = self._resolve_indexes(headers)
            if indexes is None:
                continue
            menu_idx, quantity_idx, unit_price_idx, amount_idx = indexes
            for row in rows[1:]:
                cells = [self._cell_text(cell) for cell in row.find_all(["th", "td"])]
                if not cells or max(menu_idx, amount_idx) >= len(cells):
                    continue
                menu_name = cells[menu_idx].strip()
                if not menu_name or self._looks_like_total(menu_name):
                    continue
                if not self._has_number(cells[amount_idx]):
                    continue
                sales_amount = clean_int(cells[amount_idx])
                sales_quantity = clean_int(cells[quantity_idx]) if quantity_idx is not None and quantity_idx < len(cells) else None
                unit_price = clean_int(cells[unit_price_idx]) if unit_price_idx is not None and unit_price_idx < len(cells) else None
                records.append(
                    MenuSalesRecord(
                        store_id=store_id,
                        store_name=store_name,
                        period_start=period_start,
                        period_end=period_end,
                        menu_name=menu_name,
                        sales_quantity=sales_quantity,
                        sales_amount=sales_amount,
                        unit_price=unit_price,
                    )
                )
        return records

    def _resolve_indexes(self, headers: list[str]) -> tuple[int, int | None, int | None, int] | None:
        normalized = [header.replace(" ", "") for header in headers]
        menu_idx = self._find_index(normalized, self.menu_keywords)
        amount_candidates = [
            idx
            for idx, header in enumerate(normalized)
            if any(keyword in header for keyword in self.amount_keywords)
        ]
        amount_idx = amount_candidates[-1] if amount_candidates else None
        quantity_idx = self._find_index(normalized, self.quantity_keywords)
        unit_price_idx = self._find_index(normalized, ("단가", "판매가", "가격"))
        if menu_idx is None or amount_idx is None:
            return None
        return menu_idx, quantity_idx, unit_price_idx, amount_idx

    @staticmethod
    def _find_index(headers: list[str], keywords: tuple[str, ...]) -> int | None:
        for idx, header in enumerate(headers):
            if any(keyword in header for keyword in keywords):
                return idx
        return None

    @staticmethod
    def _cell_text(cell) -> str:
        return cell.get_text(" ", strip=True) if cell else ""

    @classmethod
    def _is_empty_result(cls, soup: BeautifulSoup) -> bool:
        text = soup.get_text(" ", strip=True)
        return any(keyword in text for keyword in cls.empty_keywords)

    @staticmethod
    def _looks_like_total(value: str) -> bool:
        compact = value.replace(" ", "")
        return compact in {"합계", "총계", "전체", "소계", "TOTAL", "총합계", "전체합계", "메뉴합계"}

    @staticmethod
    def _has_number(value: str) -> bool:
        return bool(_NUMBER.search(value))

    @staticmethod
    def _source_total(soup: BeautifulSoup) -> int | None:
        for label in ("매출합계", "합계", "총매출", "총 매출"):
            value = _summary_value(soup, label)
            if value is not None:
                return value
        return None


    @staticmethod
    def _source_total_quantity(soup: BeautifulSoup) -> int | None:
        for label in ("건수", "嫄댁닔"):
            value = _summary_value(soup, label)
            if value is not None:
                return value
        return None


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
