"""Test raw AZtec headers, duplicate Area names, formulas and workbook safety."""

import ast
from contextlib import redirect_stdout
import hashlib
import io
import importlib.util
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
from threading import Event
import unittest
from unittest.mock import MagicMock, patch

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter


REPO = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location(
    "optional_phase_a", REPO / "AuAgTe-Report-Automation/Optional-HideColumn-CorrFormulaAddition.py")
optional = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = optional
SPEC.loader.exec_module(optional)

HEADERS = ["Feature", "Area", "Field", "Rank", "Area (sq. µm)",
           "Breadth (µm)", "Length (µm)", "Sb (Wt%)"]


def fixture(path, duplicate_area=False):
    book = Workbook()
    book.active.title = "Area"
    summary = book.create_sheet("Summary")
    summary.append(["Old summary", "Area", "Percentage"])
    summary.append(["Unwanted old row", 999999, 99])
    values = {
        "Raw Data": [["Raw1", 99999, 1, 1, 9999, 4, 25, 10]],
        "Gudmundite": [["Gu1", 9001, 1, 1, 100, 4, 25, 58],
                       ["Gu2", 9002, 1, 2, 36, 2, 18, 58]],
        "Aurostibite": [["Au1", 9003, 2, 1, 64, 8, 8, 55]],
        "Cinnabar": [["Cin1", 9004, 3, 1, 999, 9, 111, 0]],
        "Cinnabar-Asso": [["Cin2", 9005, 3, 2, 888, 8, 111, 0]],
        "After Size-CutOff": [["Small1", 9006, 3, 3, 777, 7, 111, 55]],
        "Empty category": [],
    }
    for name, rows in values.items():
        ws = book.create_sheet(name)
        headers = list(HEADERS)
        if duplicate_area:
            headers[4] = "Area"
            headers[5] = "Breadth (um)"
            headers[6] = "Length"
        ws.append(headers)
        for row in rows:
            ws.append(row)
        ws["A2"].fill = PatternFill(fill_type="solid", fgColor="FFFF00")
    book.create_sheet("Integrity Check").append(["Metric", "Value"])
    book.create_sheet("Notes").append(["User notes"])
    book.save(path)
    book.close()


class OptionalFinishTests(unittest.TestCase):
    def setUp(self):
        self.folder = tempfile.TemporaryDirectory()
        self.root = Path(self.folder.name)
        self.source = self.root / "ClassifiedSample.xlsx"
        fixture(self.source)

    def tearDown(self):
        self.folder.cleanup()

    def test_columns_summary_formulas_and_source_safety(self):
        digest = hashlib.sha256(self.source.read_bytes()).hexdigest()
        output = optional.build_optional_workbook(self.source)
        book = load_workbook(output)
        try:
            for name in ("Gudmundite", "Aurostibite", "Cinnabar", "Cinnabar-Asso",
                         "After Size-CutOff", "Raw Data", "Empty category"):
                ws = book[name]
                self.assertEqual([c.value for c in ws[1]][:8], HEADERS)
                self.assertEqual(tuple(c.value for c in ws[1])[-7:], optional.NEW_HEADERS)
                self.assertEqual(ws.freeze_panes, "A2")
            gu = book["Gudmundite"]
            self.assertEqual(gu["J2"].value, 100)
            self.assertIsNone(gu["I2"].value)
            self.assertIsNone(gu["K2"].value)
            self.assertIsNone(gu["N2"].value)
            self.assertIsNone(gu["O2"].value)
            self.assertEqual(gu["L2"].value, '=IF(OR(K2="",K2=0),F2,SQRT(J2*K2/100))')
            self.assertEqual(gu["M2"].value, '=IF(OR(K2="",K2=0),G2,SQRT(J2*K2/100))')
            self.assertEqual(gu["A2"].fill.fgColor.rgb[-6:], "FFF8D6")
            summary = book["Summary"]
            self.assertEqual([summary.cell(r, 1).value for r in range(2, 5)],
                             ["Gudmundite", "Aurostibite", "Total"])
            self.assertEqual(summary["B2"].value, "=SUM('Gudmundite'!$E$2:$E$3)")
            self.assertEqual(summary["B3"].value, "=SUM('Aurostibite'!$E$2:$E$2)")
            self.assertEqual(summary["B4"].value, "=SUM(B2:B3)")
            self.assertEqual(summary["C2"].value, "=IF($B$4=0,0,B2/$B$4)")
            self.assertEqual(summary["C4"].value, "=SUM(C2:C3)")
            self.assertEqual(summary["C2"].number_format, "0.00%")
            self.assertEqual(summary["A2"].border.left.style, "thin")
            self.assertTrue(book.calculation.fullCalcOnLoad)
            self.assertEqual(book["Notes"]["A1"].value, "User notes")
        finally:
            book.close()
        self.assertEqual(hashlib.sha256(self.source.read_bytes()).hexdigest(), digest)
        second = optional.build_optional_workbook(self.source)
        self.assertNotEqual(output, second)
        self.assertTrue(output.exists())

    def test_exact_b3_geometry_formulas(self):
        source = REPO / "Ag-Liberation/Image Links_CorrFactor.py"
        tree = ast.parse(source.read_bytes())
        function = next(n for n in tree.body if isinstance(n, ast.FunctionDef)
                        and n.name == "ensure_new_columns_and_formulas")
        namespace = {"get_column_letter": get_column_letter}
        exec(compile(ast.Module(body=[function], type_ignores=[]), str(source), "exec"), namespace)
        output = optional.build_optional_workbook(self.source)
        book = load_workbook(output)
        reference = Workbook()
        try:
            ws = reference.active
            ws.append(HEADERS + list(optional.NEW_HEADERS))
            ws.append(["Gu", 1, 1, 1, 100, 4, 25, 58])
            namespace["ensure_new_columns_and_formulas"](ws, 10, 11, 12, 13, 14, 6, 7)
            self.assertEqual(book["Gudmundite"]["L2"].value, ws["L2"].value)
            self.assertEqual(book["Gudmundite"]["M2"].value, ws["M2"].value)
        finally:
            book.close()
            reference.close()

    def test_second_bare_area_and_independent_visibility(self):
        fixture(self.source, duplicate_area=True)
        def picker(choices):
            visible = {key for key, _ in choices}
            visible.remove(("area", 1))
            visible.remove(("openimage", 1))
            self.assertIn(("area", 2), visible)
            return visible
        output = optional.build_optional_workbook(self.source, visibility_picker=picker)
        book = load_workbook(output)
        try:
            ws = book["Gudmundite"]
            self.assertTrue(ws.column_dimensions["B"].hidden)
            self.assertFalse(ws.column_dimensions["E"].hidden)
            self.assertTrue(ws.column_dimensions["I"].hidden)
            self.assertEqual(ws["J2"].value, 100)
            self.assertEqual(book["Summary"]["B2"].value, "=SUM('Gudmundite'!$E$2:$E$3)")
            self.assertFalse(book["Summary"].column_dimensions["B"].hidden)
        finally:
            book.close()

    def test_grouped_hidden_columns_can_be_shown(self):
        book = load_workbook(self.source)
        book["Gudmundite"].column_dimensions.group("B", "G", hidden=True)
        book.save(self.source)
        book.close()
        output = optional.build_optional_workbook(self.source)
        book = load_workbook(output)
        try:
            ws = book["Gudmundite"]
            self.assertTrue(all(not ws.column_dimensions[get_column_letter(i)].hidden for i in range(1, 16)))
            self.assertTrue(all(d.min == d.max for d in ws.column_dimensions.values()))
        finally:
            book.close()

    def test_rerun_reuses_columns_and_preserves_manual_entries(self):
        first = optional.build_optional_workbook(self.source)
        book = load_workbook(first)
        ws = book["Gudmundite"]
        ws["K2"] = 25
        ws["N2"] = "Pyrite"
        ws["O2"] = 12.345
        book.save(first)
        book.close()
        second = optional.build_optional_workbook(first)
        book = load_workbook(second)
        try:
            ws = book["Gudmundite"]
            self.assertEqual(ws.max_column, 15)
            self.assertEqual(ws["K2"].value, 25)
            self.assertEqual(ws["N2"].value, "Pyrite")
            self.assertEqual(ws["O2"].value, 12.345)
            self.assertEqual(len(ws.data_validations.dataValidation), 1)
        finally:
            book.close()

    def test_unsupported_or_incomplete_input_rejected_without_output(self):
        book = load_workbook(self.source)
        book["Gudmundite"]["F1"] = "Unknown measurement"
        book.save(self.source)
        book.close()
        digest = hashlib.sha256(self.source.read_bytes()).hexdigest()
        with self.assertRaisesRegex(ValueError, "Gudmundite.*Breadth"):
            optional.build_optional_workbook(self.source)
        self.assertEqual(list(self.root.glob('*optional_final*')), [])
        self.assertEqual(hashlib.sha256(self.source.read_bytes()).hexdigest(), digest)
        with self.assertRaisesRegex(ValueError, "xlsx"):
            optional.build_optional_workbook(self.root / "Legacy.xls")

    def test_zero_categories_summary_and_case_insensitive_exclusions(self):
        book = load_workbook(self.source)
        for name in ("Gudmundite", "Aurostibite"):
            book.remove(book[name])
        # openpyxl appends "1" on a direct case-only rename; use an intermediate
        # name to construct the actual lowercase/uppercase input to be tested.
        book["Summary"].title = "Temporary summary"
        book["Temporary summary"].title = "summary"
        book["Cinnabar"].title = "Temporary cinnabar"
        book["Temporary cinnabar"].title = "CINNABAR"
        book["Cinnabar-Asso"].title = "Cinnabar Asso"
        book.save(self.source)
        book.close()
        output = optional.build_optional_workbook(self.source)
        book = load_workbook(output)
        try:
            self.assertIn("Summary", book.sheetnames)
            self.assertNotIn("Summary1", book.sheetnames)
            self.assertEqual(book["Summary"]["A2"].value, "Total")
            self.assertEqual(book["Summary"]["B2"].value, "=0")
            self.assertEqual(book["Summary"]["C2"].value, "=0")
        finally:
            book.close()

    def test_main_standalone_flow(self):
        with patch.object(optional.tk, "Tk"), \
             patch.object(optional.filedialog, "askopenfilename", return_value=str(self.source)), \
             patch.object(optional, "choose_category_sheets", return_value={"Gudmundite", "Aurostibite"}) as sheets, \
             patch.object(optional, "choose_visible_columns", return_value=None), \
             patch.object(optional.messagebox, "showinfo") as success:
            optional.main()
        self.assertTrue(success.called)
        self.assertTrue(sheets.called)
        self.assertTrue((self.root / "ClassifiedSample_optional_final.xlsx").exists())

    def test_only_selected_sheets_change_and_summary_still_includes_all_categories(self):
        book = load_workbook(self.source)
        book["Aurostibite"].column_dimensions["B"].hidden = True
        book.save(self.source)
        book.close()
        digest = hashlib.sha256(self.source.read_bytes()).hexdigest()
        visible = {("feature", 1), ("newbrea", 1)}
        output = optional.build_optional_workbook(self.source, selected_sheets={"Gudmundite"},
                                                   visible_columns=visible)
        book = load_workbook(output)
        try:
            self.assertEqual(book["Gudmundite"].max_column, 15)
            self.assertTrue(book["Gudmundite"].column_dimensions["E"].hidden)
            self.assertFalse(book["Gudmundite"].column_dimensions["L"].hidden)
            self.assertEqual(book["Gudmundite"]["A2"].fill.fgColor.rgb[-6:], "FFF8D6")
            for name in ("Aurostibite", "Cinnabar", "Cinnabar-Asso", "After Size-CutOff", "Raw Data"):
                self.assertEqual(book[name].max_column, 8)
                self.assertEqual(book[name]["A2"].fill.fgColor.rgb[-6:], "FFFF00")
                self.assertIsNone(book[name].freeze_panes)
            self.assertTrue(book["Aurostibite"].column_dimensions["B"].hidden)
            self.assertEqual(book["Summary"]["A2"].value, "Gudmundite")
            self.assertEqual(book["Summary"]["A3"].value, "Aurostibite")
            self.assertEqual(book["Summary"]["B3"].value, "=SUM('Aurostibite'!$E$2:$E$2)")
        finally:
            book.close()
        self.assertEqual(hashlib.sha256(self.source.read_bytes()).hexdigest(), digest)

    def test_sheet_selection_cancel_does_not_modify_or_save(self):
        digest = hashlib.sha256(self.source.read_bytes()).hexdigest()
        picker = MagicMock(return_value=None)
        with patch.object(optional, "add_correction_columns") as edits:
            output = optional.build_optional_workbook(self.source, sheet_picker=picker)
        self.assertIsNone(output)
        self.assertIn("Gudmundite", picker.call_args[0][0])
        self.assertIn("Raw Data", picker.call_args[0][0])
        self.assertNotIn("Summary", picker.call_args[0][0])
        edits.assert_not_called()
        self.assertEqual(list(self.root.glob('*optional_final*')), [])
        self.assertEqual(hashlib.sha256(self.source.read_bytes()).hexdigest(), digest)
        for selected in (set(), {"Not a category"}):
            with self.assertRaisesRegex(ValueError, "valid feature-data sheet"):
                optional.build_optional_workbook(self.source, selected_sheets=selected)

    def test_unselected_sheet_does_not_require_geometry_for_formula_addition(self):
        book = load_workbook(self.source)
        book["Gudmundite"]["F1"] = "Unrecognized breadth"
        book.save(self.source)
        book.close()
        output = optional.build_optional_workbook(self.source, selected_sheets={"Aurostibite"})
        book = load_workbook(output)
        try:
            self.assertEqual(book["Gudmundite"].max_column, 8)
            self.assertEqual(book["Gudmundite"]["F1"].value, "Unrecognized breadth")
            self.assertEqual(book["Aurostibite"].max_column, 15)
            self.assertEqual(book["Summary"]["A2"].value, "Gudmundite")
        finally:
            book.close()

    def test_pastel_highlights_are_stable_on_rerun(self):
        book = load_workbook(self.source)
        book["Gudmundite"]["H2"].fill = PatternFill(fill_type="solid", fgColor="F44336")
        book.save(self.source)
        book.close()
        first = optional.build_optional_workbook(self.source, selected_sheets={"Gudmundite"})
        second = optional.build_optional_workbook(first, selected_sheets={"Gudmundite"})
        for output in (first, second):
            book = load_workbook(output)
            try:
                self.assertEqual(book["Gudmundite"]["H2"].fill.fgColor.rgb[-6:], "FBE4E2")
                self.assertEqual(book["Gudmundite"]["I1"].fill.fgColor.rgb[-6:], "EFF4FA")
                self.assertEqual(book["Summary"]["A1"].fill.fgColor.rgb[-6:], "E7EEF6")
            finally:
                book.close()

    def test_category_gui_defaults_and_selected_names(self):
        root = MagicMock()
        root.winfo_viewable.return_value = False
        window, box = MagicMock(), MagicMock()
        box.curselection.return_value = (1, 3)
        commands = {}

        def button(*args, **kwargs):
            commands[kwargs["text"]] = kwargs["command"]
            return MagicMock()

        root.wait_window.side_effect = lambda dialog: commands["OK"]()
        with patch.object(optional.tk, "Toplevel", return_value=window), \
             patch.object(optional.tk, "Label"), patch.object(optional.tk, "Frame"), \
             patch.object(optional.tk, "Scrollbar"), patch.object(optional.tk, "Listbox", return_value=box), \
             patch.object(optional.tk, "Button", side_effect=button):
            selected = optional.choose_category_sheets(root, ["Raw Data", "Gudmundite", "Aurostibite", "Cinnabar"])
        self.assertEqual(selected, {"Gudmundite", "Cinnabar"})
        self.assertEqual([call.args[0] for call in box.selection_set.call_args_list], [1, 2, 3])
        window.transient.assert_not_called()
        window.lift.assert_called_once()

    def test_progress_is_logged_before_loading_and_waiting_for_selection(self):
        stream = io.StringIO()
        actual_load = optional.load_workbook

        def loader(*args, **kwargs):
            self.assertIn("Loading workbook", stream.getvalue())
            return actual_load(*args, **kwargs)

        def picker(choices):
            self.assertIn("Waiting for column selection", stream.getvalue())
            self.assertIn("Loaded 11 sheets", stream.getvalue())
            return None

        with redirect_stdout(stream), patch.object(optional, "load_workbook", side_effect=loader):
            optional.build_optional_workbook(self.source, visibility_picker=picker)
        output = stream.getvalue()
        expected = ["Selected A2 workbook", "Loading workbook", "Loaded 11 sheets",
                    "Validating sheets", "Adding correction columns", "correction columns 2/2 rows",
                    "Waiting for column selection", "Applying column visibility",
                    "Rebuilding Summary", "Summary category: Gudmundite", "Saving workbook",
                    "Optional Phase A output saved", "Complete. The input workbook was not changed"]
        position = 0
        for message in expected:
            found = output.find(message, position)
            self.assertGreaterEqual(found, position, message)
            position = found + len(message)

    def test_load_failure_is_logged_before_propagating(self):
        stream = io.StringIO()
        with redirect_stdout(stream), patch.object(optional, "load_workbook", side_effect=PermissionError("File locked")):
            with self.assertRaisesRegex(PermissionError, "File locked"):
                optional.build_optional_workbook(self.source)
        self.assertIn("ERROR during Loading workbook", stream.getvalue())
        self.assertIn("File locked", stream.getvalue())

    def test_blocking_stage_produces_heartbeat(self):
        received = Event()

        class Observer(io.StringIO):
            def write(self, text):
                result = super().write(text)
                if "Still working:" in text:
                    received.set()
                return result

        stream = Observer()
        with redirect_stdout(stream), patch.object(optional, "HEARTBEAT_SECONDS", 0.01):
            with optional.progress_stage("Test long operation"):
                self.assertTrue(received.wait(2), "No heartbeat arrived while the operation was blocked")
        self.assertIn("Still working: Test long operation", stream.getvalue())
        self.assertIn("Done: Test long operation", stream.getvalue())

    def test_column_dialog_not_attached_to_hidden_root(self):
        root = MagicMock()
        root.winfo_viewable.return_value = False
        window = MagicMock()
        stream = io.StringIO()

        def wait(dialog):
            self.assertIn("Column-selection window is open", stream.getvalue())
            dialog.destroy()

        root.wait_window.side_effect = wait
        with redirect_stdout(stream), patch.object(optional.tk, "Toplevel", return_value=window), \
             patch.object(optional.tk, "Label"), patch.object(optional.tk, "Frame"), \
             patch.object(optional.tk, "Scrollbar"), patch.object(optional.tk, "Listbox"), \
             patch.object(optional.tk, "Button"):
            result = optional.choose_visible_columns(root, [(("feature", 1), "Feature")])
        self.assertIsNone(result)
        window.transient.assert_not_called()
        window.deiconify.assert_called_once()
        window.lift.assert_called_once()
        window.focus_force.assert_called_once()
        window.attributes.assert_called_with("-topmost", True)

    @unittest.skipUnless(shutil.which('libreoffice') or shutil.which('soffice'), 'LibreOffice unavailable')
    def test_real_spreadsheet_recalculation(self):
        # Formula strings alone are insufficient: verify actual cached results.
        book = load_workbook(self.source)
        book["Aurostibite"].title = "O'Brien"
        book.save(self.source)
        book.close()
        output = optional.build_optional_workbook(self.source, selected_sheets={"Gudmundite", "O'Brien"})
        book = load_workbook(output)
        book["Gudmundite"]["K3"] = 25  # sqrt(36 * 25 / 100) = 3.
        book["O'Brien"]["K2"] = 0  # Zero keeps original breadth/length.
        book.save(output)
        book.close()
        zero_source = self.root / "Zero.xlsx"
        fixture(zero_source)
        book = load_workbook(zero_source)
        for name in ("Gudmundite", "Aurostibite"):
            for row_number in range(2, book[name].max_row + 1):
                book[name].cell(row_number, 5).value = 0
        book.save(zero_source)
        book.close()
        zero_output = optional.build_optional_workbook(zero_source)
        converted = self.root / "calculated"
        converted.mkdir()
        profile = self.root / "lo-profile"
        env = dict(os.environ, SAL_USE_VCLPLUGIN="svp")
        command = [shutil.which('libreoffice') or shutil.which('soffice'),
                   '-env:UserInstallation=' + profile.as_uri(), '--headless', '--convert-to', 'xlsx',
                   '--outdir', str(converted), str(output), str(zero_output)]
        result = subprocess.run(command, capture_output=True, text=True, env=env, timeout=60)
        self.assertEqual(result.returncode, 0, result.stderr + result.stdout)
        computed = load_workbook(converted / output.name, data_only=True)
        try:
            self.assertEqual(computed["Gudmundite"]["L2"].value, 4)
            self.assertEqual(computed["Gudmundite"]["M2"].value, 25)
            self.assertEqual(computed["Gudmundite"]["L3"].value, 3)
            self.assertEqual(computed["Gudmundite"]["M3"].value, 3)
            self.assertEqual(computed["O'Brien"]["L2"].value, 8)
            self.assertEqual(computed["Summary"]["B2"].value, 136)
            self.assertEqual(computed["Summary"]["B3"].value, 64)
            self.assertEqual(computed["Summary"]["B4"].value, 200)
            self.assertAlmostEqual(computed["Summary"]["C2"].value, 0.68)
            self.assertAlmostEqual(computed["Summary"]["C3"].value, 0.32)
            self.assertAlmostEqual(computed["Summary"]["C4"].value, 1)
        finally:
            computed.close()
        zero = load_workbook(converted / zero_output.name, data_only=True)
        try:
            self.assertEqual(zero["Summary"]["B4"].value, 0)
            self.assertEqual(zero["Summary"]["C4"].value, 0)
            self.assertEqual(zero["Summary"]["C2"].value, 0)
        finally:
            zero.close()


if __name__ == '__main__':
    unittest.main()
