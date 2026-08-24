"""Tests for the data-handling parts of General Ternary Plot Script.py."""

import importlib.util
from pathlib import Path
import tempfile
import unittest

import pandas as pd


SCRIPT = Path(__file__).with_name("General Ternary Plot Script.py")
SPEC = importlib.util.spec_from_file_location("general_ternary_plot", SCRIPT)
module = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(module)


class GeneralTernaryPlotTests(unittest.TestCase):
    def test_prepares_case_insensitive_headers_and_rejects_invalid_rows(self):
        source = pd.DataFrame(
            {" as ": [20, -1, 0], "S": [30, 10, 0], "FE": [50, 91, 0],
             "Label": ["valid", "negative", "zero"], "Size": [5, 5, 5]}
        )
        prepared, skipped = module.prepare_dataframe(
            source, module.PlotSettings("As", "S", "Fe")
        )
        self.assertEqual(prepared["Label"].tolist(), ["valid"])
        self.assertEqual(skipped, 2)

    def test_optional_label_and_size_columns_receive_defaults(self):
        source = pd.DataFrame({"As": [10], "S": [20], "Fe": [70]})
        prepared, skipped = module.prepare_dataframe(
            source, module.PlotSettings("As", "S", "Fe", "", "")
        )
        self.assertEqual(prepared["Label"].tolist(), ["Row 2"])
        self.assertEqual(prepared["Size"].tolist(), [1.0])
        self.assertEqual(skipped, 0)

    def test_reads_every_excel_sheet(self):
        with tempfile.TemporaryDirectory() as directory:
            workbook = Path(directory) / "input.xlsx"
            with pd.ExcelWriter(workbook) as writer:
                pd.DataFrame({"As": [1]}).to_excel(writer, sheet_name="First", index=False)
                pd.DataFrame({"As": [2]}).to_excel(writer, sheet_name="Second", index=False)
            datasets = module.read_datasets(workbook)
        self.assertEqual([name for name, _ in datasets], ["First", "Second"])

    def test_rejects_repeated_elements(self):
        with self.assertRaisesRegex(ValueError, "three different"):
            module.validate_settings(module.PlotSettings("As", "as", "Fe"))

    def test_reports_missing_configured_size_header(self):
        source = pd.DataFrame({"As": [10], "S": [20], "Fe": [70], "Label": ["A"]})
        with self.assertRaisesRegex(ValueError, "missing marker size column"):
            module.prepare_dataframe(source, module.PlotSettings("As", "S", "Fe"))


if __name__ == "__main__":
    unittest.main()
