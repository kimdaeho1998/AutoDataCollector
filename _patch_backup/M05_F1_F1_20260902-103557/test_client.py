from __future__ import annotations

import sys
import unittest
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from sales_data_collector.client import ServiceClient, ServiceEndpoints
from sales_data_collector.exceptions import AuthenticationError, HttpError


class Response:
    def __init__(self, text: str = "", status_code: int = 200) -> None:
        self.text = text
        self.status_code = status_code
        self.request = type("Request", (), {"method": "POST"})()
        self.url = "https://example.test/mock"


class ClientTests(unittest.TestCase):
    def test_login_replays_configured_three_request_workflow(self) -> None:
        class Session:
            def __init__(self):
                self.calls = []

            def get(self, url, timeout):
                self.calls.append(("GET", url, None))
                return Response()

            def post(self, url, data, timeout):
                self.calls.append(("POST", url, data))
                return Response()

        session = Session()
        endpoints = ServiceEndpoints(
            login_path="/login",
            login_submit_path="/login_submit",
            login_complete_path="/login_complete",
        )
        ServiceClient("https://example.test", session=session, endpoints=endpoints).login("test-id", "test-password")

        self.assertEqual(
            [(method, url) for method, url, _ in session.calls],
            [
                ("GET", "https://example.test/login"),
                ("POST", "https://example.test/login_submit"),
                ("GET", "https://example.test/login_complete"),
            ],
        )
        payload = session.calls[1][2]
        self.assertEqual(payload["id"], "test-id")
        self.assertEqual(payload["pw"], "test-password")
        self.assertNotIn("pwSave", payload)

    def test_store_list_login_page_is_reported_as_authentication_failure(self) -> None:
        class Session:
            def post(self, url, data, timeout):
                return Response('<form action="/login"><input name="id"><input name="pw"></form>')

        endpoints = ServiceEndpoints(store_list_path="/stores")
        with self.assertRaises(AuthenticationError):
            ServiceClient("https://example.test", session=Session(), endpoints=endpoints).get_stores("BRAND001")

    def test_menu_monthly_request_uses_configured_product_path(self) -> None:
        class Session:
            def post(self, url, data, timeout):
                self.url = url
                self.data = data
                return Response(
                    """
                    <table>
                      <tr><th>메뉴명</th><th>수량</th><th>매출금액</th></tr>
                      <tr><td>Sample Menu</td><td>3</td><td>12,000</td></tr>
                    </table>
                    """
                )

        session = Session()
        endpoints = ServiceEndpoints(product_sales_path="/product?cmd=search")
        result = ServiceClient("https://example.test", session=session, endpoints=endpoints).get_menu_monthly_sales(
            start_date=date(2026, 7, 1),
            end_date=date(2026, 7, 31),
            brand_idx="BRAND001",
            brand_name="Sample Brand",
            store_idx="STORE001",
            store_name="Sample Store",
        )

        self.assertEqual(session.url, "https://example.test/product?cmd=search")
        self.assertEqual(session.data["startDate"], "2026-07-01")
        self.assertEqual(session.data["endDate"], "2026-07-31")
        self.assertEqual(session.data["storeidx"], "STORE001")
        self.assertEqual(result.records[0].sales_amount, 12000)

    def test_http_error_is_reused_for_menu_request(self) -> None:
        class Session:
            def post(self, url, data, timeout):
                return Response("server error", status_code=500)

        endpoints = ServiceEndpoints(product_sales_path="/product?cmd=search")
        with self.assertRaises(HttpError):
            ServiceClient("https://example.test", session=Session(), endpoints=endpoints).get_menu_monthly_sales(
                start_date=date(2026, 7, 1),
                end_date=date(2026, 7, 31),
                brand_idx="BRAND001",
                brand_name="Sample Brand",
                store_idx="STORE001",
                store_name="Sample Store",
            )

    def test_missing_menu_endpoint_is_required_only_when_menu_feature_runs(self) -> None:
        # Explicitly force the menu endpoint to be missing.
        #
        # ServiceEndpoints dataclass defaults are populated from the process
        # environment.  A developer/production shell may legitimately have
        # COLLECTOR_PRODUCT_SALES_PATH configured, so ServiceEndpoints()
        # alone would make this test environment-dependent and could cause
        # a real HTTP attempt to https://example.test.
        endpoints = ServiceEndpoints(product_sales_path="")
        client = ServiceClient("https://example.test", endpoints=endpoints)

        with self.assertRaisesRegex(ValueError, "product_sales_path"):
            client.get_menu_monthly_sales(
                start_date=date(2026, 7, 1),
                end_date=date(2026, 7, 31),
                brand_idx="BRAND001",
                brand_name="Sample Brand",
                store_idx="STORE001",
                store_name="Sample Store",
            )


if __name__ == "__main__":
    unittest.main()

