from __future__ import annotations

import json
import os
from dataclasses import dataclass
from datetime import date
from typing import Any

import requests
from bs4 import BeautifulSoup

from .exceptions import AuthenticationError, HttpError, ParseError
from .models import (
    DailySalesRecord,
    MenuMonthlySalesResult,
    MonthlySalesRecord,
    PeriodSalesResult,
    ProductSalesResult,
    Store,
    TodayStoreSalesResult,
)
from .parser import ServiceSalesParser
from .source_parsers import DailySalesParser, MenuSalesParser, MonthlySalesParser, PeriodSalesParser, ProductSalesParser, TodayStoreSalesParser
from .utils import format_ymd


@dataclass
class ServiceEndpoints:
    """Endpoint paths supplied through local environment variables only."""

    login_path: str = os.environ.get("COLLECTOR_LOGIN_PATH", "")
    login_submit_path: str = os.environ.get("COLLECTOR_LOGIN_SUBMIT_PATH", "")
    login_complete_path: str = os.environ.get("COLLECTOR_LOGIN_COMPLETE_PATH", "")
    store_list_path: str = os.environ.get("COLLECTOR_STORE_LIST_PATH", "")
    sales_search_path: str = os.environ.get("COLLECTOR_SALES_SEARCH_PATH", "")
    daily_sales_path: str = os.environ.get("COLLECTOR_DAILY_SALES_PATH", "")
    period_sales_path: str = os.environ.get("COLLECTOR_PERIOD_SALES_PATH", "")
    monthly_sales_path: str = os.environ.get("COLLECTOR_MONTHLY_SALES_PATH", "")
    product_sales_path: str = os.environ.get("COLLECTOR_PRODUCT_SALES_PATH", "")
    delivery_sales_path: str = os.environ.get("COLLECTOR_DELIVERY_SALES_PATH", "")
    today_store_sales_path: str = os.environ.get("COLLECTOR_TODAY_SALES_PATH", "")
    logout_path: str = os.environ.get("COLLECTOR_LOGOUT_PATH", "")


class ServiceClient:
    def __init__(
        self,
        base_url: str,
        *,
        session: requests.Session | None = None,
        endpoints: ServiceEndpoints | None = None,
        timeout: int = 30,
        parser: ServiceSalesParser | None = None,
    ) -> None:
        self.base_url = base_url.rstrip("/")
        self.session = session or requests.Session()
        self.endpoints = endpoints or ServiceEndpoints()
        self.timeout = timeout
        self.parser = parser or ServiceSalesParser()

    def login(self, user_id: str, password: str) -> None:
        # Captured browser workflow: entry page, login form submit, then completion page.
        entry = self.session.get(self._url(self._require_path("login_path")), timeout=self.timeout)
        self._ensure_ok(entry)
        response = self.session.post(
            self._url(self._require_path("login_submit_path")),
            data={
                "url": "/",
                "viewtype": "",
                "postDataKey": "",
                "encpw": "",
                "encnm": "",
                "saveID": "0",
                "enctp": "1",
                "cPW": "",
                "sPW": "",
                "smart_level": "",
                "adv_check": "",
                "autoLogin": "N",
                "id": user_id,
                "pw": password,
            },
            timeout=self.timeout,
        )
        self._ensure_ok(response)
        if self._looks_like_login_failure(response.text):
            raise AuthenticationError("login failed")
        complete = self.session.get(self._url(self._require_path("login_complete_path")), timeout=self.timeout)
        self._ensure_ok(complete)

    def get_stores(self, brand_idx: str, store_use: str = "Y") -> list[Store]:
        response = self.session.post(
            self._url(self._require_path("store_list_path")),
            data={"cmd": "getStore", "brand_idx": brand_idx, "store_use": store_use},
            timeout=self.timeout,
        )
        self._ensure_ok(response)
        try:
            payload = self._safe_json(response.text)
        except ParseError as exc:
            if self._looks_like_login_page(response.text):
                raise AuthenticationError("login verification failed") from exc
            raise
        items = payload.get("dataList")
        if not isinstance(items, list):
            raise AuthenticationError("login verification failed")
        stores: list[Store] = []
        for item in items:
            store_id = str(item.get("value", "")).strip()
            store_name = str(item.get("text", "")).strip()
            if store_id and store_name:
                stores.append(Store(magic_store_id=store_id, store_name=store_name))
        return stores

    def get_time_sales(
        self,
        *,
        business_date: date,
        brand_idx: str,
        brand_name: str,
        store_idx: str,
        store_name: str,
        store_use: str = "Y",
    ) -> SalesResult:
        payload = {
            "cmd": "search",
            "startDate": format_ymd(business_date),
            "endDate": format_ymd(business_date),
            "brandidx": brand_idx,
            "txtbrandidx": brand_name,
            "store_use": store_use,
            "txtstore_use": "사용" if store_use == "Y" else store_use,
            "storeidx": store_idx,
            "txtstoreidx": store_name,
        }
        response = self.session.post(self._url(self._require_path("sales_search_path")), data=payload, timeout=self.timeout)
        self._ensure_ok(response)
        return self.parser.parse_sales_page(
            response.text,
            business_date=business_date,
            magic_store_id=store_idx,
            store_name=store_name,
        )

    def get_daily_sales(self, *, year: int, month: int, brand_idx: str, brand_name: str, store_idx: str, store_name: str) -> list[DailySalesRecord]:
        response = self.session.post(
            self._url(self._require_path("daily_sales_path")),
            data=self._store_payload(brand_idx, brand_name, store_idx, store_name, srchYear=str(year), txtsrchYear=str(year), srchMonth=f"{month:02d}", txtsrchMonth=f"{month:02d}", storeidx_str="", usFranOrStore="1"),
            timeout=self.timeout,
        )
        self._ensure_ok(response)
        return DailySalesParser().parse(response.text, year=year, month=month)

    def get_period_sales(self, *, start_date: date, end_date: date, brand_idx: str, brand_name: str, store_idx: str, store_name: str) -> PeriodSalesResult:
        response = self.session.post(
            self._url(self._require_path("period_sales_path")),
            data=self._store_payload(
                brand_idx,
                brand_name,
                store_idx,
                store_name,
                startDate=format_ymd(start_date),
                endDate=format_ymd(end_date),
                txtsorting="\ub9e4\ucd9c\uc0c1\uc704",
            ),
            timeout=self.timeout,
        )
        self._ensure_ok(response)
        try:
            return PeriodSalesParser().parse(response.text)
        except ParseError as exc:
            raise ParseError(f"{exc}; response={self._safe_response_shape(response.text)}") from exc

    def get_monthly_sales(self, *, year: int, brand_idx: str, brand_name: str, store_idx: str, store_name: str) -> list[MonthlySalesRecord]:
        response = self.session.post(
            self._url(self._require_path("monthly_sales_path")),
            data=self._store_payload(brand_idx, brand_name, store_idx, store_name, startDate=str(year), txtstartDate=str(year), storeidx_str="", usFranOrStore="1"),
            timeout=self.timeout,
        )
        self._ensure_ok(response)
        return MonthlySalesParser().parse(response.text)

    def get_product_sales(self, *, business_date: date, brand_idx: str, brand_name: str, store_idx: str, store_name: str) -> ProductSalesResult:
        response = self.session.post(
            self._url(self._require_path("product_sales_path")),
            data=self._store_payload(brand_idx, brand_name, store_idx, store_name, startDate=format_ymd(business_date), endDate=format_ymd(business_date), storeidx_str="", usFranOrStore="1"),
            timeout=self.timeout,
        )
        self._ensure_ok(response)
        return ProductSalesParser().parse(response.text)

    def get_menu_monthly_sales(self, *, start_date: date, end_date: date, brand_idx: str, brand_name: str, store_idx: str, store_name: str) -> MenuMonthlySalesResult:
        response = self.session.post(
            self._url(self._require_path("product_sales_path")),
            data=self._store_payload(
                brand_idx,
                brand_name,
                store_idx,
                store_name,
                startDate=format_ymd(start_date),
                endDate=format_ymd(end_date),
                storeidx_str="",
                usFranOrStore="1",
            ),
            timeout=self.timeout,
        )
        self._ensure_ok(response)
        try:
            return MenuSalesParser().parse(
                response.text,
                store_id=store_idx,
                store_name=store_name,
                period_start=start_date,
                period_end=end_date,
            )
        except ParseError as exc:
            raise ParseError(f"{exc}; response={self._safe_menu_response_shape(response.text)}") from exc

    def get_today_store_gross_sales(self, *, business_date: date, brand_idx: str, store_idx: str, store_name: str) -> int:
        return self.get_today_store_sales(
            business_date=business_date,
            brand_idx=brand_idx,
            store_idx=store_idx,
            store_name=store_name,
        ).gross_sales_amount

    def get_today_store_sales(self, *, business_date: date, brand_idx: str, store_idx: str, store_name: str) -> TodayStoreSalesResult:
        """Use the confirmed POST contract for the date-specific gross-sales field."""
        response = self.session.post(
            self._url(self._require_path("today_store_sales_path")),
            data={
                "brandidx": brand_idx,
                "storeidx": store_idx,
                "txtstoreidx": store_name,
                "start_date": format_ymd(business_date),
                "store_use": "Y",
                "txtstore_use": "\uc0ac\uc6a9",
            },
            timeout=self.timeout,
        )
        self._ensure_ok(response)
        return TodayStoreSalesParser().parse(response.text)

    @staticmethod
    def _store_payload(brand_idx: str, brand_name: str, store_idx: str, store_name: str, **extra: str) -> dict[str, str]:
        payload = {
            "brandidx": brand_idx,
            "txtbrandidx": brand_name,
            "store_use": "Y",
            "txtstore_use": "\uc0ac\uc6a9",
            "storeidx": store_idx,
            "txtstoreidx": store_name,
        }
        payload.update(extra)
        return payload

    def logout(self) -> None:
        # No server logout contract was captured. Close local resources instead
        # of guessing an endpoint and generating a misleading 404 request.
        self.session.close()

    def _url(self, path: str) -> str:
        if path.startswith("http://") or path.startswith("https://"):
            return path
        return f"{self.base_url}{path}"

    def _require_path(self, endpoint_name: str) -> str:
        value = getattr(self.endpoints, endpoint_name)
        if not value:
            raise ValueError(f"Service endpoint configuration is missing: {endpoint_name}")
        return value

    def _ensure_ok(self, response: requests.Response) -> None:
        if response.status_code >= 400:
            raise HttpError(f"unexpected HTTP status: {response.status_code} ({response.request.method} {response.url})")

    def _safe_json(self, text: str) -> dict[str, Any]:
        try:
            payload = json.loads(text)
        except json.JSONDecodeError as exc:
            raise ParseError("store list response is not valid JSON") from exc
        if not isinstance(payload, dict):
            raise ParseError("store list response payload must be an object")
        return payload

    @staticmethod
    def _safe_response_shape(html: str) -> str:
        """Expose only page structure for support; never include values or session data."""
        soup = BeautifulSoup(html, "html.parser")
        title = soup.title.get_text(" ", strip=True) if soup.title else "none"
        actions = [str(form.get("action", "")) for form in soup.select("form")]
        input_names = [str(field.get("name", "")) for field in soup.select("input[name]")]
        return f"title={title[:80]!r}, actions={actions[:5]!r}, input_names={input_names[:20]!r}"

    @staticmethod
    def _safe_menu_response_shape(html: str) -> str:
        """Expose menu page structure without dumping operating rows or credentials."""
        soup = BeautifulSoup(html, "html.parser")
        title = soup.title.get_text(" ", strip=True) if soup.title else "none"
        table_headers = []
        for table in soup.find_all("table")[:5]:
            rows = table.find_all("tr")[:3]
            table_headers.append(
                [
                    cell.get_text(" ", strip=True)[:40]
                    for row in rows
                    for cell in row.find_all(["th", "td"])[:12]
                ][:12]
            )
        input_names = [str(field.get("name", "")) for field in soup.select("input[name]")[:20]]
        classes = sorted({name for tag in soup.find_all(True) for name in (tag.get("class") or [])})[:30]
        selector_shapes = {
            selector: ServiceClient._selector_shape(soup, selector)
            for selector in (".detail", ".detail > li", ".detail2", ".detail2 > li", ".detail3", ".detail3 > li", ".detail_list2", ".detail_list2 > li", ".detail_title", ".detail_title2")
        }
        return (
            f"title={title[:80]!r}, table_count={len(soup.find_all('table'))}, "
            f"table_headers={table_headers!r}, input_names={input_names!r}, classes={classes!r}, "
            f"selector_shapes={selector_shapes!r}"
        )

    @staticmethod
    def _selector_shape(soup: BeautifulSoup, selector: str) -> dict[str, Any]:
        nodes = soup.select(selector)
        samples = []
        for node in nodes[:3]:
            children = node.find_all(recursive=False)
            samples.append(
                {
                    "tag": node.name,
                    "classes": node.get("class") or [],
                    "child_count": len(children),
                    "children": [
                        {
                            "tag": child.name,
                            "classes": child.get("class") or [],
                        }
                        for child in children[:6]
                    ],
                    "text": node.get_text(" ", strip=True)[:200],
                }
            )
        return {"count": len(nodes), "samples": samples}

    @staticmethod
    def _looks_like_login_page(text: str) -> bool:
        lowered = text.lower()
        return "login.asp" in lowered or ('name="id"' in lowered and 'name="pw"' in lowered)

    @staticmethod
    def _looks_like_login_failure(text: str) -> bool:
        lowered = text.lower()
        return any(keyword in lowered for keyword in ["invalid", "error", "fail", "실패", "오류"])
