"""Report payload checks: python -m unittest test_reports -v"""
from pathlib import Path
import unittest

import pandas as pd

from finance import Rules, calculate_accounts, load_workbook
from reports import build_monthly_report_payload


SOURCE = Path.home() / "Downloads" / "Product Finance - AI Fluency Test Sample Data.xlsx"


class ReportTests(unittest.TestCase):
    @unittest.skipUnless(SOURCE.exists(), "sample workbook is not available")
    def test_report_payload_uses_calculated_metrics(self):
        contracts, usage = load_workbook(SOURCE.read_bytes())
        accounts = calculate_accounts(contracts, usage, Rules())
        as_of = accounts.as_of_date.iloc[0]
        payload = build_monthly_report_payload(
            accounts, usage, as_of, SOURCE.name, "abc")

        self.assertEqual(payload["title"], "Monthly Usage Report")
        self.assertEqual(payload["month"], pd.Timestamp(as_of).strftime("%B %Y"))
        self.assertEqual(len(payload["top_three"]), 3)
        self.assertEqual(len(payload["chart"]["categories"]), 12)
        self.assertEqual(payload["chart"]["categories"][0], f"Jan {as_of:%y}")
        self.assertGreater(payload["portfolio"]["ytd_usage"], 0)
        self.assertEqual({item["customer_id"] for item in payload["issues"]},
                         {"CUST-02", "CUST-03"})


if __name__ == "__main__":
    unittest.main()
