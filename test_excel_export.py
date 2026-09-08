"""Formula-export checks: python -m unittest test_excel_export -v"""
import json
from pathlib import Path
import tempfile
import unittest
from zipfile import ZipFile

from finance import Rules, calculate_accounts, load_workbook
from excel_export import build_audit_workbook_payload, generate_audit_workbook


SOURCE = Path.home() / "Downloads" / "Product Finance - AI Fluency Test Sample Data.xlsx"


@unittest.skipUnless(SOURCE.exists(), "sample workbook is not available")
class ExcelExportTests(unittest.TestCase):
    def test_export_contains_sources_assumptions_and_formulas(self):
        contracts, usage = load_workbook(SOURCE.read_bytes())
        accounts = calculate_accounts(contracts, usage, Rules())
        queue = accounts.loc[accounts.recommended_owner.ne("No Action")]
        payload = build_audit_workbook_payload(
            contracts, usage, queue, Rules(), usage.date.max(), SOURCE.name, "abc")

        self.assertEqual(payload["queue_customer_ids"][0], "CUST-03")
        self.assertEqual(len(payload["contracts"]), 12)
        self.assertEqual(len(payload["usage"]), 1943)

        workbook = generate_audit_workbook(payload)
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "audit.xlsx"
            path.write_bytes(workbook)
            with ZipFile(path) as archive:
                workbook_xml = archive.read("xl/workbook.xml").decode("utf-8")
                sheets = "".join(
                    archive.read(name).decode("utf-8")
                    for name in archive.namelist()
                    if name.startswith("xl/worksheets/sheet") and name.endswith(".xml"))

        for name in ["Action Queue", "Account Metrics", "Assumptions", "Contracts", "Usage Events"]:
            self.assertIn(f'name="{name}"', workbook_xml)
        self.assertGreater(sheets.count(":f>"), 500)
        self.assertIn("SUMIFS", sheets)
        self.assertIn("Account Metrics", sheets)
        self.assertIn("Assumptions", sheets)


if __name__ == "__main__":
    unittest.main()
