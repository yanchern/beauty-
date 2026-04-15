from __future__ import annotations

import csv
import tempfile
import threading
import unittest
import urllib.request
from pathlib import Path

from fr_self_tax_sales_report import generate_self_tax_report
from fr_self_tax_web import make_server, render_country_page, render_home


def metric_dict(result: dict[str, object], key: str) -> dict[str, str]:
    return {label: f"{amount:.2f}" for label, amount in result[key]}


class TaxLogicTests(unittest.TestCase):
    def write_csv(self, rows: list[dict[str, str]], tmpdir: Path, name: str) -> Path:
        fieldnames = sorted({key for row in rows for key in row})
        path = tmpdir / name
        with path.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(rows)
        return path

    def base_row(self, **overrides: str) -> dict[str, str]:
        row = {
            "ACTIVITY_PERIOD": "2026Q1",
            "ASIN": "TESTASIN",
            "BUYER_VAT_NUMBER": "",
            "INVOICE_URL": "",
            "ITEM_DESCRIPTION": "test item",
            "MARKETPLACE": "Amazon",
            "PRICE_OF_ITEMS_VAT_RATE_PERCENT": "",
            "QTY": "1",
            "SALE_ARRIVAL_COUNTRY": "",
            "SALE_DEPART_COUNTRY": "",
            "SELLER_ARRIVAL_COUNTRY_VAT_NUMBER": "",
            "SELLER_SKU": "SKU-1",
            "TAXABLE_JURISDICTION": "",
            "TAX_CALCULATION_DATE": "",
            "TAX_COLLECTION_RESPONSIBILITY": "",
            "TAX_REPORTING_SCHEME": "",
            "TOTAL_ACTIVITY_VALUE_AMT_VAT_INCL": "0",
            "TRANSACTION_CURRENCY_CODE": "EUR",
            "TRANSACTION_EVENT_ID": "event-1",
            "TRANSACTION_TYPE": "SALE",
            "VAT_CALCULATION_IMPUTATION_COUNTRY": "",
        }
        row.update(overrides)
        return row

    def test_uk_metrics(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmpdir = Path(tmp)
            csv_path = self.write_csv(
                [
                    self.base_row(
                        TAX_COLLECTION_RESPONSIBILITY="MARKETPLACE",
                        SALE_DEPART_COUNTRY="GB",
                        SALE_ARRIVAL_COUNTRY="FR",
                        TOTAL_ACTIVITY_VALUE_AMT_VAT_INCL="120",
                        TRANSACTION_EVENT_ID="uk-marketplace",
                    ),
                    self.base_row(
                        TAX_COLLECTION_RESPONSIBILITY="SELLER",
                        SALE_DEPART_COUNTRY="GB",
                        SALE_ARRIVAL_COUNTRY="FR",
                        PRICE_OF_ITEMS_VAT_RATE_PERCENT="0",
                        TOTAL_ACTIVITY_VALUE_AMT_VAT_INCL="60",
                        TRANSACTION_EVENT_ID="uk-zero",
                    ),
                    self.base_row(
                        TAX_COLLECTION_RESPONSIBILITY="SELLER",
                        SALE_DEPART_COUNTRY="FR",
                        SALE_ARRIVAL_COUNTRY="GB",
                        PRICE_OF_ITEMS_VAT_RATE_PERCENT="0.2",
                        TOTAL_ACTIVITY_VALUE_AMT_VAT_INCL="120",
                        TRANSACTION_EVENT_ID="uk-a",
                    ),
                    self.base_row(
                        TAX_COLLECTION_RESPONSIBILITY="SELLER",
                        SALE_DEPART_COUNTRY="GB",
                        SALE_ARRIVAL_COUNTRY="FR",
                        PRICE_OF_ITEMS_VAT_RATE_PERCENT="0.2",
                        TOTAL_ACTIVITY_VALUE_AMT_VAT_INCL="240",
                        TRANSACTION_EVENT_ID="uk-b",
                    ),
                ],
                tmpdir,
                "uk.csv",
            )
            result = generate_self_tax_report(csv_path=csv_path, country_code="uk")
            metrics = metric_dict(result, "summary_metrics")

            self.assertEqual("120.00", metrics["代扣代缴销售额(含税)"])
            self.assertEqual("60.00", metrics["未代扣0税率销售额(含税)"])
            self.assertEqual("120.00", metrics["未代扣缴税A部分销售额(含税)"])
            self.assertEqual("240.00", metrics["未代扣缴税B部分销售额(含税)"])
            self.assertEqual("360.00", metrics["未代扣缴税销售额合计(含税)"])
            self.assertEqual("60.00", metrics["应缴税金"])

    def test_germany_metrics_and_no_tax_formula(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmpdir = Path(tmp)
            csv_path = self.write_csv(
                [
                    self.base_row(
                        TAX_COLLECTION_RESPONSIBILITY="SELLER",
                        SALE_DEPART_COUNTRY="DE",
                        SALE_ARRIVAL_COUNTRY="FR",
                        PRICE_OF_ITEMS_VAT_RATE_PERCENT="0.19",
                        TOTAL_ACTIVITY_VALUE_AMT_VAT_INCL="100",
                        TRANSACTION_EVENT_ID="de-main",
                    ),
                    self.base_row(
                        TAX_COLLECTION_RESPONSIBILITY="SELLER",
                        SALE_DEPART_COUNTRY="DE",
                        SALE_ARRIVAL_COUNTRY="DE",
                        PRICE_OF_ITEMS_VAT_RATE_PERCENT="0",
                        TOTAL_ACTIVITY_VALUE_AMT_VAT_INCL="50",
                        TRANSACTION_EVENT_ID="de-zero",
                    ),
                    self.base_row(
                        TAX_COLLECTION_RESPONSIBILITY="SELLER",
                        SALE_DEPART_COUNTRY="PL",
                        SALE_ARRIVAL_COUNTRY="DE",
                        PRICE_OF_ITEMS_VAT_RATE_PERCENT="0.19",
                        TOTAL_ACTIVITY_VALUE_AMT_VAT_INCL="75",
                        TRANSACTION_EVENT_ID="de-missing",
                    ),
                    self.base_row(
                        TAX_COLLECTION_RESPONSIBILITY="SELLER",
                        SALE_DEPART_COUNTRY="DE",
                        SALE_ARRIVAL_COUNTRY="FR",
                        PRICE_OF_ITEMS_VAT_RATE_PERCENT="0.19",
                        TRANSACTION_TYPE="COMMINGLING_BUY",
                        TOTAL_ACTIVITY_VALUE_AMT_VAT_INCL="999",
                        TRANSACTION_EVENT_ID="de-excluded",
                    ),
                ],
                tmpdir,
                "de.csv",
            )
            result = generate_self_tax_report(csv_path=csv_path, country_code="de")
            metrics = metric_dict(result, "summary_metrics")

            self.assertEqual("100.00", metrics["德国发出欧盟销售额(含税)"])
            self.assertEqual("50.00", metrics["平台遗漏订单(德国境内,AE=0)销售额(含税)"])
            self.assertEqual("75.00", metrics["平台遗漏订单(目的国德国,AE非零)销售额(含税)"])
            self.assertEqual("225.00", metrics["Seller需申报销售额合计"])
            self.assertNotIn("应缴税金", metrics)
            self.assertEqual(3, result["row_count"])

    def test_italy_metrics(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmpdir = Path(tmp)
            csv_path = self.write_csv(
                [
                    self.base_row(
                        TAX_COLLECTION_RESPONSIBILITY="SELLER",
                        SALE_DEPART_COUNTRY="IT",
                        SALE_ARRIVAL_COUNTRY="FR",
                        BUYER_VAT_NUMBER="FR123",
                        PRICE_OF_ITEMS_VAT_RATE_PERCENT="0.22",
                        TOTAL_ACTIVITY_VALUE_AMT_VAT_INCL="122",
                        TRANSACTION_EVENT_ID="it-b1",
                    ),
                    self.base_row(
                        TAX_COLLECTION_RESPONSIBILITY="SELLER",
                        SALE_DEPART_COUNTRY="FR",
                        SALE_ARRIVAL_COUNTRY="IT",
                        PRICE_OF_ITEMS_VAT_RATE_PERCENT="0.22",
                        TOTAL_ACTIVITY_VALUE_AMT_VAT_INCL="244",
                        TRANSACTION_EVENT_ID="it-b2",
                    ),
                    self.base_row(
                        TAX_COLLECTION_RESPONSIBILITY="SELLER",
                        SALE_DEPART_COUNTRY="ES",
                        SALE_ARRIVAL_COUNTRY="IT",
                        TOTAL_ACTIVITY_VALUE_AMT_VAT_INCL="122",
                        TRANSACTION_EVENT_ID="it-f",
                    ),
                ],
                tmpdir,
                "it.csv",
            )
            result = generate_self_tax_report(csv_path=csv_path, country_code="it")
            metrics = metric_dict(result, "summary_metrics")

            self.assertEqual("122.00", metrics["B1销售额(含税)"])
            self.assertEqual("244.00", metrics["B2销售额(含税)"])
            self.assertEqual("122.00", metrics["F销售额(含税)"])
            self.assertEqual("488.00", metrics["自行缴纳订单销售额合计(含税)"])
            self.assertEqual("400.00", metrics["自行缴纳订单净销售额"])
            self.assertEqual("88.00", metrics["销项税"])

    def test_netherlands_metrics(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmpdir = Path(tmp)
            csv_path = self.write_csv(
                [
                    self.base_row(
                        TAX_COLLECTION_RESPONSIBILITY="MARKETPLACE",
                        SALE_DEPART_COUNTRY="DE",
                        SALE_ARRIVAL_COUNTRY="NL",
                        TOTAL_ACTIVITY_VALUE_AMT_VAT_INCL="180",
                        TRANSACTION_EVENT_ID="nl-main",
                    ),
                    self.base_row(
                        TAX_COLLECTION_RESPONSIBILITY="MARKETPLACE",
                        SALE_DEPART_COUNTRY="GB",
                        SALE_ARRIVAL_COUNTRY="NL",
                        TOTAL_ACTIVITY_VALUE_AMT_VAT_INCL="999",
                        TRANSACTION_EVENT_ID="nl-excluded-non-eu",
                    ),
                    self.base_row(
                        TAX_COLLECTION_RESPONSIBILITY="SELLER",
                        SALE_DEPART_COUNTRY="FR",
                        SALE_ARRIVAL_COUNTRY="NL",
                        TOTAL_ACTIVITY_VALUE_AMT_VAT_INCL="999",
                        TRANSACTION_EVENT_ID="nl-excluded-seller",
                    ),
                ],
                tmpdir,
                "nl.csv",
            )
            result = generate_self_tax_report(csv_path=csv_path, country_code="nl")
            metrics = metric_dict(result, "summary_metrics")

            self.assertEqual("180.00", metrics["代扣代缴销售额(含税)"])
            self.assertEqual("180.00", metrics["代扣代缴销售额合计(含税)"])
            self.assertEqual(1, result["row_count"])

    def test_saudi_metrics(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmpdir = Path(tmp)
            sales_path = self.write_csv(
                [
                    {
                        "fulfillment-channel": "Amazon",
                        "product sales": "100",
                        "shipping credits": "20",
                        "promotional rebates": "-5",
                        "currency": "SAR",
                    },
                    {
                        "fulfillment-channel": "Merchant",
                        "product sales": "999",
                        "shipping credits": "0",
                        "promotional rebates": "0",
                        "currency": "SAR",
                    },
                ],
                tmpdir,
                "sa_sales.csv",
            )
            result = generate_self_tax_report(
                csv_path=sales_path,
                country_code="sa",
            )
            metrics = metric_dict(result, "summary_metrics")

            self.assertEqual("115.00", metrics["季度应纳税销售额(含税)"])
            self.assertEqual("100.00", metrics["季度不含税销售额"])
            self.assertEqual(1, result["row_count"])

    def test_render_templates_and_routes(self) -> None:
        home = render_home()
        self.assertIn("/france", home)
        self.assertIn("/spain", home)
        self.assertIn("/uk", home)
        self.assertIn("/germany", home)
        self.assertIn("/italy", home)
        self.assertIn("/netherlands", home)
        self.assertIn("/saudi", home)
        self.assertIn("/emblem_uk.svg", home)
        self.assertIn("/emblem_nl.svg", home)
        self.assertIn("/emblem_sa.svg", home)

        uk_page = render_country_page("uk")
        de_page = render_country_page("germany")
        it_page = render_country_page("italy")
        nl_page = render_country_page("netherlands")
        sa_page = render_country_page("saudi")
        self.assertIn("英国税金逻辑 TXT", uk_page)
        self.assertIn(".txt,text/plain", uk_page)
        self.assertIn(".docx,.doc", de_page)
        self.assertIn("意大利税金逻辑 PDF", it_page)
        self.assertIn("荷兰税金逻辑 DOCX", nl_page)
        self.assertIn("沙特税金计算方法", sa_page)
        self.assertIn(".xlsx,.xls", sa_page)

        server = make_server("127.0.0.1", 0)
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        try:
            base_url = f"http://127.0.0.1:{server.server_port}"
            for path in ("/healthz", "/france", "/spain", "/uk", "/germany", "/italy", "/netherlands", "/saudi"):
                with urllib.request.urlopen(f"{base_url}{path}") as response:
                    self.assertEqual(200, response.status)
        finally:
            server.shutdown()
            server.server_close()


if __name__ == "__main__":
    unittest.main()
