import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import openpyxl
import pandas as pd

import chdn_epi_master as report


class IdpAnnualTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)
        self.source = Path(self.temp_dir.name) / "source.xlsx"
        self.output = Path(self.temp_dir.name) / "output.xlsx"

        columns = ["Period ", "Organization ", "Project Name", "indicator"]
        columns += [
            f"Q{quarter} {category}"
            for quarter in range(1, 5)
            for category in ["IDP Male ", "IDP Female", "non-IDP Male ", "non-IDP Female"]
        ]
        self.template = pd.DataFrame([
            {"Period ": 2025, "Organization ": "CHDN", "Project Name": "Project A", "indicator": "Penta1"},
            {"Period ": 2026, "Organization ": "CHDN", "Project Name": "Project B", "indicator": "Penta1"},
        ], columns=columns)
        records = [
            ("2025-01-10", "yes", "male", "a"),
            ("2025-01-11", "yes", "male", "a"),
            ("2025-04-10", "yes", "male", "a"),
            ("2025-04-11", "yes", "male", "b"),
            ("2025-07-10", "yes", "female", "c"),
            ("2025-10-10", "no", "male", "d"),
            ("2026-01-10", "no", "female", "e"),
        ]
        child = pd.DataFrame(records, columns=["reporting_month", "IDP (Y/N)", "Gender", "children_code(T/C Code)"])
        child["vaccine_dose"] = "Penta1"
        child["source"] = "CHDN"
        child["age_at_dose"] = 10
        child.to_excel(self.output, sheet_name="EPI-Child-long", index=False)
        self.template.to_excel(self.source, sheet_name="IDP", index=False)

    def run_step(self):
        with patch.object(report, "OUTPUT_FILE", self.output), patch.object(
            report, "resolve_sheet_path_and_name", return_value=(self.source, "IDP")
        ):
            report.step_add_multi_year_idp()

    def test_annual_sums_quarters_and_preserves_metadata(self):
        self.run_step()
        self.run_step()

        wb = openpyxl.load_workbook(self.output, data_only=True)
        self.addCleanup(wb.close)
        self.assertEqual(wb.sheetnames.count("IDP_annual"), 1)
        rows = list(wb["IDP_annual"].values)
        self.assertEqual(rows[0], (
            "Period", "Organization", "Project Name", "indicator",
            "IDP Male", "IDP Female", "non-IDP Male", "non-IDP Female",
        ))
        self.assertEqual(rows[1], (2025, "CHDN", "Project A", "Penta1", 3, 1, 1, 0))
        self.assertEqual(rows[2], (2026, "CHDN", "Project B", "Penta1", 0, 0, 0, 1))
        quarterly = pd.read_excel(self.output, sheet_name="IDP")
        self.assertEqual(quarterly.loc[0, "Q1 IDP Male "], 1)
        self.assertEqual(quarterly.loc[0, "Q2 IDP Male "], 2)

    def test_missing_quarter_column_is_reported(self):
        self.template.drop(columns=["Q4 IDP Female"]).to_excel(self.source, sheet_name="IDP", index=False)
        with self.assertRaisesRegex(ValueError, "Q4 IDP Female"):
            self.run_step()
        wb = openpyxl.load_workbook(self.output, read_only=True)
        self.addCleanup(wb.close)
        self.assertNotIn("IDP_annual", wb.sheetnames)


if __name__ == "__main__":
    unittest.main()
