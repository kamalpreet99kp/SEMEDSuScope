"""Requirement/boundary checks and synthetic-workbook tests; no lab data altered."""

import ast
import hashlib
import importlib.util
from pathlib import Path
import sys
import re
import runpy
import traceback
import tempfile
import unittest
from unittest.mock import patch

import pandas as pd
from openpyxl import Workbook, load_workbook


REPO = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("sb_classifier", REPO / "Ag-Liberation/Sb Mineral Classification.py")
sb = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = sb
SPEC.loader.exec_module(sb)

FREE_EXAMPLES = {
    "Gudmundite": dict(Sb=58.07, Fe=26.64, S=15.29),
    "Aurostibite": dict(Au=44.72, Sb=55.28),
    "Stibnite": dict(Sb=71.68, S=28.32),
    "Boulangerite": dict(Pb=55.23, Sb=25.97, S=18.80),
    "Cupropolybasite": dict(Ag=48.25, Cu=22.11, Sb=12.10, S=17.53),
    "Cinnabar": dict(Hg=86.22, S=13.78),
    sb.AG_TETRAHEDRITE: dict(Cu=26.3, Sb=24.7, S=23.3, Ag=11.6, Hg=8.1, Fe=2.7, Zn=1.3, O=2),
    "Tetrahedrite-(Fe)": dict(Cu=38.49, Sb=29.50, S=25.25, Fe=6.77),
    "Andorite": dict(Sb=40.7, Pb=22.6, S=21.5, Ag=10.2, O=1.9, Cu=1, Fe=0.7),
    "Zn-bearing Tetrahedrite": dict(Cu=35.5, Sb=31, S=26.2, Fe=4.3, Zn=3),
    "Bournonite": dict(Pb=44.1, Sb=24.8, S=18.7, Cu=11.4, Fe=1.1),
    "Pyrargyrite": dict(Ag=59.76, Sb=22.48, S=17.76),
}
ASSOCIATION_EXAMPLES = {
    "Gudmundite-Asso": dict(Sb=30, Fe=30, S=15, O=25),
    "Aurostibite-Asso": dict(Au=20, Sb=25, S=10, Fe=15, O=30),
    "Stibnite-Asso": dict(Sb=40, S=20, O=30, Si=10),
    "Boulangerite-Asso": dict(Pb=30, Sb=15, S=15, O=40),
    "Cupropolybasite-Asso": dict(Ag=24, Cu=11, Sb=6, S=9, O=50),
    "Cinnabar-Asso": dict(Hg=43, S=7, O=50),
    "Andorite-Asso": dict(Ag=5, Pb=11, Sb=20, S=11, O=53),
    "Tetrahedrite-(Fe)-Asso": dict(Cu=18, Sb=15, S=13, Fe=4, O=50),
    "Bournonite-Asso": dict(Pb=22, Cu=6, Sb=12, S=10, O=50),
    "Pyrargyrite-Asso": dict(Ag=30, Sb=11, S=9, O=50),
}

EXPECTED_FREE_ORDER = (
    "Aurostibite", sb.AG_TETRAHEDRITE, "Andorite", "Cupropolybasite", "Pyrargyrite",
    "Gudmundite", "Stibnite", "Tetrahedrite-(Fe)", "Zn-bearing Tetrahedrite",
    "Boulangerite", "Bournonite", "Cinnabar",
)
EXPECTED_ASSOCIATION_ORDER = (
    "Aurostibite-Asso", "Andorite-Asso", "Cupropolybasite-Asso", "Pyrargyrite-Asso",
    "Gudmundite-Asso", "Stibnite-Asso", "Tetrahedrite-(Fe)-Asso",
    "Boulangerite-Asso", "Bournonite-Asso", "Cinnabar-Asso",
)


def row(elements, area=100):
    return {sb.AREA_COLUMN: area, **{f"{name} (Wt%)": value for name, value in elements.items()}}


class RulesTests(unittest.TestCase):
    def test_all_twelve_free_examples(self):
        for name, elements in FREE_EXAMPLES.items():
            with self.subTest(name=name):
                self.assertEqual(sb.classify_mineral(row(elements)), name)

    def test_all_ten_association_examples(self):
        for name, elements in ASSOCIATION_EXAMPLES.items():
            with self.subTest(name=name):
                self.assertEqual(sb.classify_mineral(row(elements)), name)

    def test_user_primary_ranges_and_inclusive_boundaries(self):
        # Independent transcription of the user's exact numerical limits.
        expected = [
            {"Sb": (44, 65), "Fe": (17, 32), "S": (5, 20)},
            {"Au": (35, 51), "Sb": (42, 65)},
            {"Sb": (58, 78), "S": (14, 38)},
            {"Pb": (42, 64), "Sb": (16, 34), "S": (6, 28)},
            {"Ag": (30, 58), "Cu": (13, 28), "Sb": (4, 18), "S": (7, 23)},
            {"Hg": (60, 95), "S": (4, 20)},
            {"Ag": (2, 45), "Cu": (15, 45), "Sb": (15, 35), "S": (10, 31), "Hg": (0, 25)},
            {"Cu": (22, 45), "Sb": (20, 37), "S": (15, 31), "Fe": (1, 10)},
            {"Ag": (4, 18), "Pb": (14, 31), "Sb": (30, 48), "S": (12, 28)},
            {"Cu": (22, 45), "Sb": (20, 37), "S": (15, 31), "Zn": (2, 11), "Fe": (0, 8)},
            {"Pb": (30, 46), "Sb": (12, 30), "S": (7, 24), "Cu": (4, 19)},
            {"Ag": (30, 65), "Sb": (14, 29), "S": (7, 22)},
            {"Sb": (15, 60), "Fe": (7, 65), "S": (3, 45)},
            {"Au": (4, 51), "Sb": (10, 62)},
            {"Sb": (15, 77), "S": (5, 48)},
            {"Pb": (15, 65), "Sb": (5, 31), "S": (4, 40)},
            {"Ag": (4, 56), "Cu": (5, 40), "Sb": (4, 18), "S": (4, 40)},
            {"Hg": (20, 90), "S": (4, 40)},
            {"Ag": (2, 17), "Pb": (5, 50), "Sb": (10, 47), "S": (4, 40)},
            {"Cu": (8, 43), "Sb": (5, 35), "S": (5, 45), "Fe": (1, 45)},
            {"Pb": (12, 46), "Sb": (5, 30), "S": (4, 24), "Cu": (3, 19)},
            {"Ag": (4, 65), "Sb": (5, 29), "S": (4, 22)},
        ]
        self.assertEqual(len(expected), len(sb.FREE_RULES + sb.ASSOCIATION_RULES))
        for rule, limits in zip(sb.FREE_RULES + sb.ASSOCIATION_RULES, expected):
            self.assertEqual(rule.ranges, limits, rule.name)
            values = {e: 0.0 for e in sb.ELEMENTS}
            values.update({e: (lo + hi) / 2 for e, (lo, hi) in limits.items()})
            for element, (lo, hi) in limits.items():
                for value, match in ((lo, True), (hi, True), (lo - 0.0001, False), (hi + 0.0001, False)):
                    probe = {**values, element: value}
                    self.assertEqual(sb.primary_match(rule, probe), match, (rule.name, element, value))

    def test_w_la_exclusions_and_equality(self):
        for examples in (FREE_EXAMPLES, ASSOCIATION_EXAMPLES):
            for name, elements in examples.items():
                for element, cap in (("W", 10), ("La", 7)):
                    self.assertEqual(sb.classify_mineral(row({**elements, element: cap})), name)
                    result = sb.classify_feature(row({**elements, element: cap + 0.001}), 200)
                    self.assertIsNone(result.category)
                    self.assertEqual(result.status, "Excluded W/La")
        self.assertIsNone(sb.classify_mineral(row(dict(Sb=3, W=11))))

    def test_size_cutoff_eligibility_and_boundary(self):
        values = FREE_EXAMPLES["Gudmundite"]
        self.assertEqual(sb.classify_mineral(row(values, 4.999), 5), sb.SIZE_CATEGORY)
        self.assertEqual(sb.classify_mineral(row(values, 5), 5), "Gudmundite")
        self.assertEqual(sb.classify_mineral(row(values, 0), 0), "Gudmundite")
        self.assertIsNone(sb.classify_mineral(row(dict(Sb=2), 1), 5))
        self.assertEqual(sb.classify_mineral(row(FREE_EXAMPLES["Cinnabar"], 1), 5), sb.SIZE_CATEGORY)

    def test_leftovers_are_last_and_au_precedes_ag(self):
        for elements, expected in [
            (dict(Sb=3, Au=3, Ag=3), "Au-Sb Leftover"),
            (dict(Sb=3, Ag=3), "Ag-Sb Leftover"),
            (dict(Sb=3), "Sb Leftover"),
            (dict(Sb=2, Au=3), None),
            (dict(Sb=3, Au=2, Ag=2), "Sb Leftover"),
            (dict(Sb=60, Au=4), "Au-Sb Leftover"),
        ]:
            self.assertEqual(sb.classify_mineral(row(elements)), expected)
        # This free example also meets the association's broad primary range.
        self.assertEqual(sb.classify_mineral(row(FREE_EXAMPLES["Gudmundite"])), "Gudmundite")

    def test_tetrahedrite_ag_zn_thresholds_and_copper_discrimination(self):
        base = dict(Cu=35, Sb=30, S=26, Fe=5)
        for changes, expected in [
            (dict(Ag=2, Zn=2), sb.AG_TETRAHEDRITE),
            (dict(Ag=1.499, Zn=2), "Zn-bearing Tetrahedrite"),
            (dict(Ag=1.499, Zn=1.999), "Tetrahedrite-(Fe)"),
        ]:
            self.assertEqual(sb.classify_mineral(row({**base, **changes})), expected)
        self.assertEqual(sb.classify_mineral(row(dict(Pb=44, Sb=25, S=19, Cu=10))), "Bournonite")
        self.assertEqual(sb.classify_mineral(row(dict(Sb=60, Fe=20, S=15))), "Gudmundite")
        self.assertEqual(sb.classify_mineral(row(dict(Ag=40, Cu=20, Sb=16, S=20))), "Cupropolybasite")
        # A stephanite-like ratio must not be forced into free pyrargyrite.
        self.assertNotEqual(sb.classify_mineral(row(dict(Ag=64, Sb=17, S=17))), "Pyrargyrite")

    def test_precious_element_caps_apply_to_every_free_and_association_rule(self):
        examples = {**FREE_EXAMPLES, **ASSOCIATION_EXAMPLES}
        for rule in sb.FREE_RULES + sb.ASSOCIATION_RULES:
            for element in ("Au", "Ag"):
                if element in rule.ranges:
                    continue
                with self.subTest(category=rule.name, element=element):
                    self.assertEqual(rule.caps[element], 1.5)
                    baseline = examples[rule.name]
                    self.assertEqual(sb.classify_mineral(row({**baseline, element: 1.499})), rule.name)
                    for value in (1.5, 3):
                        self.assertNotEqual(sb.classify_mineral(row({**baseline, element: value})), rule.name)

    def test_other_impurities_relaxed_but_diagnostic_boundaries_kept(self):
        gudmundite = FREE_EXAMPLES["Gudmundite"]
        self.assertEqual(sb.classify_mineral(row({**gudmundite, "Cu": 7.999, "Pb": 7.999})), "Gudmundite")
        self.assertNotEqual(sb.classify_mineral(row({**gudmundite, "Cu": 8})), "Gudmundite")
        rules = {r.name: r for r in sb.FREE_RULES + sb.ASSOCIATION_RULES}
        self.assertEqual(rules["Boulangerite"].caps["Cu"], 4)
        self.assertEqual(rules["Boulangerite-Asso"].caps["Cu"], 3)
        self.assertEqual(rules["Stibnite-Asso"].caps["Fe"], 7)
        self.assertEqual(rules["Pyrargyrite-Asso"].caps["Cu"], 5)
        self.assertEqual(rules["Tetrahedrite-(Fe)"].caps["Zn"], 2)

    def test_display_order_is_separate_from_routing_order(self):
        self.assertEqual(sb.FREE_ORDER, EXPECTED_FREE_ORDER)
        self.assertEqual(sb.ASSOCIATION_ORDER, EXPECTED_ASSOCIATION_ORDER)
        self.assertEqual(tuple(r.name for r in sb.FREE_RULES), tuple(FREE_EXAMPLES))
        self.assertEqual(tuple(r.name for r in sb.ASSOCIATION_RULES), tuple(ASSOCIATION_EXAMPLES))

    def test_ambiguity_retained_for_review(self):
        rules = (sb.MineralRule("Candidate A", {"Sb": (3, 10)}),
                 sb.MineralRule("Candidate B", {"Sb": (3, 10)}))
        with patch.object(sb, "FREE_RULES", rules):
            result = sb.classify_feature(row(dict(Sb=3, Ag=3)))
        self.assertEqual(result.category, "Ag-Sb Leftover")
        self.assertEqual(result.candidates, ("Candidate A", "Candidate B"))
        self.assertIn("Ambiguous", result.note)

    def test_invalid_values_not_classified(self):
        for bad in ("bad", -1, float("inf")):
            self.assertEqual(sb.classify_feature(row(dict(Sb=bad))).status, "Invalid data")
        self.assertEqual(sb.classify_feature(row(dict(Sb=3), float("nan"))).status, "Invalid data")


class WorkbookTests(unittest.TestCase):
    def test_output_structure_integrity_uniqueness_and_input_safety(self):
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / "Sample.xlsx"
            rows = [row(v) for v in FREE_EXAMPLES.values()]
            rows += [row(v) for v in ASSOCIATION_EXAMPLES.values()]
            rows += [row(dict(Sb=3)), row(dict(Sb=3, Au=3)), row(dict(Sb=3, Ag=3)),
                     row(dict(Sb=3, W=11)), row(dict(Sb=3, La=8)),
                     row(dict(Sb=2)), row(FREE_EXAMPLES["Gudmundite"], 1), row(dict(Sb="bad"))]
            for index, values in enumerate(rows):
                values.update({"Feature": f"Feature {index}", "Stage X (mm)": index / 10,
                               "Stage Y (mm)": -index / 10})
            pd.DataFrame(rows).to_excel(source, index=False)
            before = hashlib.sha256(source.read_bytes()).hexdigest()
            output = sb.build_classified_workbook(source, 5)
            book = load_workbook(output)
            try:
                self.assertEqual(book.sheetnames[:3], ["Area", "Summary", "Raw Data"])
                self.assertIn("After Size-CutOff", book.sheetnames)
                self.assertIn("Ag-AgHg Tetrahedrite (Fe)", book.sheetnames)
                self.assertEqual(len(book.sheetnames), 30)  # 26 categories + four standard sheets.
                self.assertTrue(all(len(name) <= 31 for name in book.sheetnames))
                expected_categories = EXPECTED_FREE_ORDER + EXPECTED_ASSOCIATION_ORDER + (
                    "Au-Sb Leftover", "Ag-Sb Leftover", "Sb Leftover", "After Size-CutOff")
                self.assertEqual(book.sheetnames[3:-1], [sb.excel_safe_sheet_name(n) for n in expected_categories])
                summary = book["Summary"]
                self.assertEqual([summary.cell(r, 1).value for r in range(2, summary.max_row + 1)],
                                 list(expected_categories))
                for cells in summary.iter_rows(min_row=2):
                    self.assertIsNone(cells[1].value)
                    self.assertIsNone(cells[2].value)
                raw = book["Raw Data"]
                self.assertEqual(raw.max_row - 1, len(rows))
                self.assertIn("Stage X (mm)", [cell.value for cell in raw[1]])
                checks = dict(book["Integrity Check"].values)
                self.assertEqual(checks["Row Count Match"], "Match")
                self.assertEqual(checks["Area Match"], "Match")
                self.assertEqual(checks["All raw rows accounted for"], "Match")
                self.assertEqual(checks["Excluded W >10 or La >7 rows"], 2)
                self.assertEqual(checks["Invalid data rows"], 1)
                self.assertEqual(checks["Total rows in category sheets (including After Size-CutOff)"], 26)
            finally:
                book.close()
            second = sb.build_classified_workbook(source, 5)
            self.assertNotEqual(output, second)
            self.assertTrue(output.exists())
            self.assertEqual(hashlib.sha256(source.read_bytes()).hexdigest(), before)

    def test_empty_workbook_and_required_columns(self):
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / "Empty.xlsx"
            pd.DataFrame(columns=["Sb (Wt%)", sb.AREA_COLUMN]).to_excel(source, index=False)
            output = sb.build_classified_workbook(source)
            book = load_workbook(output)
            self.assertEqual(book.sheetnames, ["Area", "Summary", "Raw Data", "Integrity Check"])
            book.close()
            pd.DataFrame({"Ag (Wt%)": [10], sb.AREA_COLUMN: [10]}).to_excel(source, index=False)
            with self.assertRaisesRegex(ValueError, "Sb"):
                sb.build_classified_workbook(source)

    def test_empty_categories_omitted_and_present_categories_sorted(self):
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / "Subset.xlsx"
            pd.DataFrame([row(FREE_EXAMPLES["Gudmundite"]), row(FREE_EXAMPLES["Aurostibite"])]).to_excel(source, index=False)
            output = sb.build_classified_workbook(source)
            book = load_workbook(output)
            try:
                self.assertEqual(book.sheetnames,
                                 ["Area", "Summary", "Raw Data", "Aurostibite", "Gudmundite", "Integrity Check"])
                self.assertEqual(book["Summary"]["A2"].value, "Aurostibite")
                self.assertEqual(book["Summary"]["A3"].value, "Gudmundite")
            finally:
                book.close()

    def test_existing_app_runner_can_launch_standalone_script(self):
        # Exercise the app's real legacy-runner function without importing Qt.
        app = REPO / "AuAgTe-Report-Automation/AuAgTe-Report-Automation-App.py"
        tree = ast.parse(app.read_bytes())
        function = next(n for n in tree.body if isinstance(n, ast.FunctionDef)
                        and n.name == "run_external_script")
        namespace = dict(Path=Path, sys=sys, re=re, runpy=runpy, traceback=traceback,
                         PICKED_DIRECTORY_MARKER="__AUAGTE_PICKED_DIRECTORY__=")
        exec(compile(ast.Module(body=[function], type_ignores=[]), str(app), "exec"), namespace)
        argv = sys.argv[:]
        picker = sb.filedialog.askopenfilename
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / "Runner.xlsx"
            pd.DataFrame([row(FREE_EXAMPLES["Gudmundite"])]).to_excel(source, index=False)
            before = hashlib.sha256(source.read_bytes()).hexdigest()
            with patch.object(sb.tk, "Tk"), patch.object(sb.simpledialog, "askfloat", return_value=5):
                exit_code = namespace["run_external_script"](
                    REPO / "Ag-Liberation/Sb Mineral Classification.py", source, None)
            self.assertEqual(exit_code, 0)
            self.assertTrue((source.parent / "ClassifiedRunner.xlsx").exists())
            self.assertEqual(hashlib.sha256(source.read_bytes()).hexdigest(), before)
        self.assertEqual(sys.argv, argv)
        self.assertIs(sb.filedialog.askopenfilename, picker)


if __name__ == "__main__":
    unittest.main()
