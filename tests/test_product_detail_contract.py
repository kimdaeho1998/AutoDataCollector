from __future__ import annotations

import unittest
from datetime import date
from unittest.mock import Mock

from sales_data_collector.client import ServiceClient


class ProductDetailContractTests(unittest.TestCase):

    def _client(self, html: str) -> ServiceClient:
        session = Mock()

        response = Mock()
        response.status_code = 200
        response.text = html

        session.post.return_value = response

        client = ServiceClient(
            base_url="https://example.invalid",
            session=session,
        )

        client.endpoints.product_sales_path = "/product.asp"

        return client

    def test_product_detail_reads_individual_rows(self) -> None:
        client = self._client(
            """
            <ul class="detail">
              <li>
                <div class="detail_title detail_title2">
                  <div>Sample Water 1,000 6 6,000</div>
                  <div>Sample Soup 6,000 1 5,373</div>
                </div>
              </li>
            </ul>
            """
        )

        result = client.get_product_detail_sales(
            start_date=date(2026, 9, 1),
            end_date=date(2026, 9, 30),
            brand_idx="BRAND001",
            brand_name="Sample Brand",
            store_idx="STORE001",
            store_name="Sample Store",
        )

        self.assertEqual(result.store_id, "STORE001")
        self.assertEqual(result.store_name, "Sample Store")
        self.assertEqual(result.period_start, date(2026, 9, 1))
        self.assertEqual(result.period_end, date(2026, 9, 30))

        self.assertEqual(len(result.records), 2)

        first = result.records[0]

        self.assertEqual(first.product_name, "Sample Water")
        self.assertEqual(first.unit_price, 1000)
        self.assertEqual(first.sales_quantity, 6)
        self.assertEqual(first.sales_amount, 6000)

        second = result.records[1]

        self.assertEqual(second.product_name, "Sample Soup")
        self.assertEqual(second.unit_price, 6000)
        self.assertEqual(second.sales_quantity, 1)
        self.assertEqual(second.sales_amount, 5373)

    def test_product_detail_preserves_negative_values(self) -> None:
        client = self._client(
            """
            <table>
              <tr>
                <th>상품명</th>
                <th>단가</th>
                <th>수량</th>
                <th>판매금액</th>
              </tr>
              <tr>
                <td>Negative Product</td>
                <td>1,000</td>
                <td>-2</td>
                <td>-2,000</td>
              </tr>
            </table>
            """
        )

        result = client.get_product_detail_sales(
            start_date=date(2026, 9, 1),
            end_date=date(2026, 9, 30),
            brand_idx="BRAND001",
            brand_name="Sample Brand",
            store_idx="STORE_NEG",
            store_name="Negative Store",
        )

        self.assertEqual(len(result.records), 1)

        record = result.records[0]

        self.assertEqual(record.sales_quantity, -2)
        self.assertEqual(record.sales_amount, -2000)

    def test_product_detail_uses_product_endpoint(self) -> None:
        client = self._client(
            """
            <table>
              <tr>
                <th>상품명</th>
                <th>수량</th>
                <th>판매금액</th>
              </tr>
              <tr>
                <td>Sample Product</td>
                <td>1</td>
                <td>1,000</td>
              </tr>
            </table>
            """
        )

        client.get_product_detail_sales(
            start_date=date(2026, 9, 1),
            end_date=date(2026, 9, 30),
            brand_idx="BRAND001",
            brand_name="Sample Brand",
            store_idx="STORE001",
            store_name="Sample Store",
        )

        client.session.post.assert_called_once()

        _, kwargs = client.session.post.call_args

        payload = kwargs["data"]

        self.assertEqual(payload["startDate"], "2026-09-01")
        self.assertEqual(payload["endDate"], "2026-09-30")


if __name__ == "__main__":
    unittest.main()

