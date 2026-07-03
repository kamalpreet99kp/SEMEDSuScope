"""Finish an Au report after image insertion and manual cleanup.

This Windows/Excel automation script performs the next workflow steps:
1. paste normalized Au-sheet chemistry into the report columns after No.
2. create an organized side-by-side block sheet with repeated blocks.

Run this from PyCharm on Windows with Excel installed. It uses Excel COM so that
pictures/shapes already present in the workbook are preserved better than a pure
openpyxl copy workflow.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from time import sleep
from tkinter import Tk, filedialog, messagebox, simpledialog

ROW_HEIGHT = 45
COLUMN_WIDTH = 8.43
FONT_SIZE = 11
XL_EDGE_LEFT = 7
XL_EDGE_TOP = 8
XL_EDGE_BOTTOM = 9
XL_EDGE_RIGHT = 10
XL_CONTINUOUS = 1
XL_THICK = 4
XL_MOVE_AND_SIZE = 1
XL_CENTER = -4108
XL_CALCULATION_MANUAL = -4135
XL_CALCULATION_AUTOMATIC = -4105
ORGANIZED_BLOCKS_SHEET_NAME = "Organized Blocks"


@dataclass(frozen=True)
class SampleFinishLayout:
    """Chemistry columns and block size for one sample type."""

    sample_type: str
    chemistry_headers: tuple[str, ...]
    block_size: int

    @property
    def half_block_size(self) -> int:
        return self.block_size // 2


SAMPLE_FINISH_LAYOUTS: dict[str, SampleFinishLayout] = {
    "1": SampleFinishLayout("Au+Ag", ("Au", "Ag"), 26),
    "2": SampleFinishLayout("Au+Ag+Cu", ("Au", "Ag", "Cu"), 26),
    "3": SampleFinishLayout("Au+Ag+Cu+Hg", ("Au", "Ag", "Cu", "Hg"), 20),
    "4": SampleFinishLayout("Au+Ag+Hg", ("Au", "Ag", "Hg"), 26),
}


class UserCancelledError(Exception):
    """Raised when the user cancels a required prompt."""


def choose_sample_layout() -> SampleFinishLayout:
    """Ask which sample layout should be used."""
    prompt = (
        "Select sample type by number:\n\n"
        "1 = Au+Ag\n"
        "2 = Au+Ag+Cu\n"
        "3 = Au+Ag+Cu+Hg\n"
        "4 = Au+Ag+Hg"
    )
    choice = simpledialog.askstring("Au Report Sample Type", prompt)
    if choice is None:
        raise UserCancelledError("Sample type selection was cancelled.")
    choice = choice.strip()
    if choice not in SAMPLE_FINISH_LAYOUTS:
        messagebox.showerror("Invalid sample type", "Please run again and enter 1, 2, 3, or 4.")
        raise UserCancelledError(f"Invalid sample type choice: {choice}")
    return SAMPLE_FINISH_LAYOUTS[choice]


def choose_workbook(title: str) -> Path:
    """Prompt for an Excel workbook path."""
    selected = filedialog.askopenfilename(
        title=title,
        filetypes=(("Excel workbooks", "*.xlsx *.xlsm"), ("All files", "*.*")),
    )
    if not selected:
        raise UserCancelledError(f"Workbook selection was cancelled: {title}")
    return Path(selected)


def choose_output_file(default_name: str) -> Path:
    """Prompt for a new output workbook path."""
    selected = filedialog.asksaveasfilename(
        title="Save finished Au report workbook as",
        defaultextension=".xlsx",
        initialfile=default_name,
        filetypes=(("Excel workbook", "*.xlsx"),),
    )
    if not selected:
        raise UserCancelledError("Output workbook selection was cancelled.")
    return Path(selected)


def default_final_output_name(report_workbook_path: Path) -> str:
    """Build the default finished workbook name from the selected intermediate workbook."""
    stem = report_workbook_path.stem
    if stem.endswith("_Inter"):
        stem = stem[: -len("_Inter")]
    elif stem.endswith(" Inter"):
        stem = stem[: -len(" Inter")]
    return f"{stem}_Final.xlsx"


def get_excel_property(excel_app, property_name: str):
    """Read an Excel Application property, returning None when Excel blocks access."""
    try:
        return getattr(excel_app, property_name)
    except Exception:
        return None


def set_excel_property(excel_app, property_name: str, value) -> bool:
    """Set an Excel Application property, returning False when Excel rejects it."""
    try:
        setattr(excel_app, property_name, value)
        return True
    except Exception:
        return False


def configure_excel_for_speed(excel_app) -> dict:
    """Disable expensive Excel UI updates during COM automation when Excel permits it."""
    previous_settings = {
        "ScreenUpdating": get_excel_property(excel_app, "ScreenUpdating"),
        "EnableEvents": get_excel_property(excel_app, "EnableEvents"),
        "DisplayAlerts": get_excel_property(excel_app, "DisplayAlerts"),
        "Calculation": get_excel_property(excel_app, "Calculation"),
    }
    set_excel_property(excel_app, "ScreenUpdating", False)
    set_excel_property(excel_app, "EnableEvents", False)
    set_excel_property(excel_app, "DisplayAlerts", False)
    previous_settings["CalculationWasChanged"] = set_excel_property(excel_app, "Calculation", XL_CALCULATION_MANUAL)
    return previous_settings


def restore_excel_settings(excel_app, previous_settings: dict) -> None:
    """Restore Excel settings after automation, skipping settings Excel rejected."""
    for property_name in ("ScreenUpdating", "EnableEvents", "DisplayAlerts"):
        previous_value = previous_settings.get(property_name)
        if previous_value is not None:
            set_excel_property(excel_app, property_name, previous_value)
    if previous_settings.get("CalculationWasChanged") and previous_settings.get("Calculation") is not None:
        set_excel_property(excel_app, "Calculation", previous_settings["Calculation"])


def workbook_full_name(workbook) -> str:
    """Return a workbook FullName safely."""
    try:
        return str(workbook.FullName)
    except Exception:
        return ""


def workbook_name(workbook) -> str:
    """Return a workbook Name safely."""
    try:
        return str(workbook.Name)
    except Exception:
        return ""


def workbook_matches_path(workbook, workbook_path: Path) -> bool:
    """Return True when a workbook appears to match the requested path."""
    wanted_full_name = str(workbook_path.resolve()).lower()
    wanted_name = workbook_path.name.lower()
    full_name = workbook_full_name(workbook).lower()
    name = workbook_name(workbook).lower()
    return full_name == wanted_full_name or name == wanted_name


def same_workbook(left_workbook, right_workbook) -> bool:
    """Return True when two COM workbook objects refer to the same open workbook."""
    try:
        left_full_name = workbook_full_name(left_workbook).lower()
        right_full_name = workbook_full_name(right_workbook).lower()
        return bool(left_full_name and right_full_name and left_full_name == right_full_name)
    except Exception:
        return False


def iter_open_workbooks(excel_app):
    """Yield open workbooks using index access, which is more reliable with COM collections."""
    try:
        count = excel_app.Workbooks.Count
    except Exception:
        count = 0
    for index in range(1, count + 1):
        try:
            yield excel_app.Workbooks(index)
        except Exception:
            continue


def open_workbook_names(excel_app) -> list[str]:
    """Return names of currently open workbooks for diagnostics."""
    names = []
    for workbook in iter_open_workbooks(excel_app):
        full_name = workbook_full_name(workbook)
        names.append(full_name or workbook_name(workbook) or "<unknown workbook>")
    return names


def find_open_workbook_by_path(excel_app, workbook_path: Path):
    """Find an already-open workbook by absolute path or workbook name."""
    for workbook in iter_open_workbooks(excel_app):
        if workbook_matches_path(workbook, workbook_path):
            return workbook

    try:
        active_workbook = excel_app.ActiveWorkbook
    except Exception:
        active_workbook = None
    if active_workbook is not None and workbook_matches_path(active_workbook, workbook_path):
        return active_workbook

    return None


def try_open_workbook(excel_app, resolved_path: Path):
    """Try several Excel COM open styles because some Excel versions reject keyword calls."""
    open_attempts = (
        lambda: excel_app.Workbooks.Open(
            Filename=str(resolved_path),
            UpdateLinks=0,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
            Local=True,
        ),
        lambda: excel_app.Workbooks.Open(str(resolved_path), 0, False),
        lambda: excel_app.Workbooks.Open(str(resolved_path)),
    )
    last_error = None
    for open_attempt in open_attempts:
        try:
            workbook = open_attempt()
            sleep(0.5)
            if workbook is not None:
                return workbook, None
            recovered_workbook = find_open_workbook_by_path(excel_app, resolved_path)
            if recovered_workbook is not None:
                return recovered_workbook, None
        except Exception as error:
            last_error = error
    return None, last_error


def open_excel_workbook(excel_app, workbook_path: Path):
    """Open an Excel workbook and recover it if COM returns None."""
    resolved_path = workbook_path.resolve()
    if not resolved_path.exists():
        raise FileNotFoundError(f"Workbook does not exist: {resolved_path}")

    already_open = find_open_workbook_by_path(excel_app, resolved_path)
    if already_open is not None:
        return already_open

    workbook, last_error = try_open_workbook(excel_app, resolved_path)
    if workbook is not None:
        return workbook

    open_names = open_workbook_names(excel_app)
    error_detail = f" Last Excel error: {last_error}" if last_error is not None else ""
    raise RuntimeError(
        f"Excel did not return or expose an opened workbook for: {resolved_path}."
        f" Open workbooks visible to Excel: {open_names}.{error_detail}"
    )


def close_conflicting_output_workbook(excel_app, report_workbook, output_path: Path) -> None:
    """Close an already-open generated output workbook before SaveAs.

    Excel cannot SaveAs a workbook to a filename that is already open anywhere in
    that Excel instance. This happens commonly after a failed/test run leaves the
    previous `_Final.xlsx` workbook open. The selected report workbook is never
    closed here; only a separate workbook matching the chosen output path/name is
    closed without saving so the current run can write a fresh final workbook.
    """
    open_output_workbook = find_open_workbook_by_path(excel_app, output_path.resolve())
    if open_output_workbook is None or same_workbook(open_output_workbook, report_workbook):
        return
    open_output_workbook.Close(SaveChanges=False)
    sleep(0.2)


def save_finished_workbook(excel_app, report_workbook, output_path: Path) -> None:
    """Save the finished workbook, handling same-name open workbook conflicts."""
    resolved_output_path = output_path.resolve()
    if workbook_matches_path(report_workbook, resolved_output_path):
        report_workbook.Save()
        return

    close_conflicting_output_workbook(excel_app, report_workbook, resolved_output_path)
    report_workbook.SaveAs(str(resolved_output_path), FileFormat=51)


def used_last_row(worksheet) -> int:
    """Return the last used row in an Excel worksheet."""
    return worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1


def used_last_column(worksheet) -> int:
    """Return the last used column in an Excel worksheet."""
    return worksheet.UsedRange.Column + worksheet.UsedRange.Columns.Count - 1


def normalize_header_text(value) -> str:
    """Normalize plain and wrapped Excel header text for matching."""
    return " ".join(str(value).replace("\n", " ").split()).strip().lower()


def find_header_cell(worksheet, header_text: str):
    """Find a cell by exact header text, case-insensitive, allowing wrapped lines."""
    wanted = normalize_header_text(header_text)
    for row_number in range(1, used_last_row(worksheet) + 1):
        for column_number in range(1, used_last_column(worksheet) + 1):
            value = worksheet.Cells(row_number, column_number).Value
            if value is not None and normalize_header_text(value) == wanted:
                return worksheet.Cells(row_number, column_number)
    raise ValueError(f"Could not find header {header_text!r} in sheet {worksheet.Name!r}.")


def row_contains_header(worksheet, row_number: int, header_text: str) -> bool:
    """Return True when a worksheet row contains the requested header."""
    wanted = normalize_header_text(header_text)
    for column_number in range(1, used_last_column(worksheet) + 1):
        value = worksheet.Cells(row_number, column_number).Value
        if value is not None and normalize_header_text(value) == wanted:
            return True
    return False


def worksheet_has_report_headers(worksheet) -> bool:
    """Return True when a worksheet appears to be an Au report sheet."""
    return row_contains_header(worksheet, 1, "R. Light Images") and row_contains_header(worksheet, 1, "SEM Images")


def workbook_sheet_names(workbook) -> list[str]:
    """Return workbook sheet names for diagnostics."""
    names = []
    try:
        count = workbook.Worksheets.Count
    except Exception:
        count = 0
    for index in range(1, count + 1):
        try:
            names.append(str(workbook.Worksheets(index).Name))
        except Exception:
            names.append(f"<sheet {index}>")
    return names


def find_worksheet_by_name(workbook, sheet_name: str):
    """Find a worksheet by name, case-insensitive; return None when missing."""
    wanted = sheet_name.strip().lower()
    try:
        count = workbook.Worksheets.Count
    except Exception:
        count = 0
    for index in range(1, count + 1):
        worksheet = workbook.Worksheets(index)
        if str(worksheet.Name).strip().lower() == wanted:
            return worksheet
    return None


def same_worksheet(left_worksheet, right_worksheet) -> bool:
    """Return True when two COM worksheet objects refer to the same sheet."""
    try:
        return left_worksheet.Name == right_worksheet.Name and left_worksheet.Parent.FullName == right_worksheet.Parent.FullName
    except Exception:
        return False


def delete_existing_organized_blocks_sheet(report_workbook, protected_worksheet=None) -> None:
    """Remove a previous generated Organized Blocks sheet so reruns do not fail.

    The finishing script can be rerun on the same intermediate/final workbook while
    testing. Excel rejects assigning a duplicate worksheet name, so remove only the
    generated output sheet before creating a fresh one. Raw report/data sheets are
    left untouched.
    """
    existing_sheet = find_worksheet_by_name(report_workbook, ORGANIZED_BLOCKS_SHEET_NAME)
    if existing_sheet is None:
        return
    if protected_worksheet is not None and same_worksheet(existing_sheet, protected_worksheet):
        raise RuntimeError(
            f"The selected report worksheet is named {ORGANIZED_BLOCKS_SHEET_NAME!r}. "
            "Please select the intermediate workbook/report sheet, not a previously generated block sheet."
        )

    excel_app = report_workbook.Application
    previous_display_alerts = get_excel_property(excel_app, "DisplayAlerts")
    set_excel_property(excel_app, "DisplayAlerts", False)
    existing_sheet.Delete()
    if previous_display_alerts is not None:
        set_excel_property(excel_app, "DisplayAlerts", previous_display_alerts)


def find_report_worksheet(workbook):
    """Find the report worksheet containing image columns instead of assuming sheet 1."""
    try:
        count = workbook.Worksheets.Count
    except Exception:
        count = 0
    for index in range(1, count + 1):
        worksheet = workbook.Worksheets(index)
        if str(worksheet.Name).strip().lower() == ORGANIZED_BLOCKS_SHEET_NAME.lower():
            continue
        if worksheet_has_report_headers(worksheet):
            return worksheet
    raise ValueError(
        "Could not find a report worksheet with headers 'R. Light Images' and 'SEM Images'. "
        f"Available sheets: {workbook_sheet_names(workbook)}"
    )


def find_last_normalized_header_row(worksheet) -> int:
    """Find the final row that contains a Normalized header in the Au sheet."""
    normalized_row = None
    for row_number in range(1, used_last_row(worksheet) + 1):
        for column_number in range(1, used_last_column(worksheet) + 1):
            value = worksheet.Cells(row_number, column_number).Value
            if value is not None and str(value).strip().lower() == "normalized":
                normalized_row = row_number
    if normalized_row is None:
        raise ValueError("Could not find a 'Normalized' header in the Au sheet.")
    return normalized_row


def find_column_on_row(worksheet, row_number: int, header_text: str) -> int:
    """Find a column by header text on one row."""
    wanted = header_text.strip().lower()
    for column_number in range(1, used_last_column(worksheet) + 1):
        value = worksheet.Cells(row_number, column_number).Value
        if value is not None and str(value).strip().lower() == wanted:
            return column_number
    raise ValueError(f"Could not find column {header_text!r} on row {row_number} of sheet {worksheet.Name!r}.")


def wrapped_header_text(header_text: str) -> str:
    """Return the wrapped report header text used in the final block layout."""
    normalized = str(header_text).strip()
    header_map = {
        "Au (Wt%)": "Au\n(Wt%)",
        "Ag (Wt%)": "Ag\n(Wt%)",
        "Cu (Wt%)": "Cu\n(Wt%)",
        "Hg (Wt%)": "Hg\n(Wt%)",
        "R. Light Images": "R. Light\nImages",
        "SEM Images": "SEM\nImages",
    }
    return header_map.get(normalized, normalized)


def apply_wrapped_headers(worksheet, header_row: int, start_column: int, end_column: int) -> None:
    """Wrap and center headers in one copied block header area."""
    for column_number in range(start_column, end_column + 1):
        cell = worksheet.Cells(header_row, column_number)
        cell.Value = wrapped_header_text(cell.Value)
        cell.WrapText = True
        cell.Font.Bold = True
        cell.Font.Size = FONT_SIZE
        cell.HorizontalAlignment = XL_CENTER
        cell.VerticalAlignment = XL_CENTER


def format_chemistry_number_columns(worksheet, first_row: int, last_row: int) -> None:
    """Format report chemistry columns to two decimal places."""
    if last_row < first_row:
        return
    reflected_light_header = find_header_cell(worksheet, "R. Light Images")
    chemistry_last_column = reflected_light_header.Column - 1
    if chemistry_last_column < 2:
        return
    chemistry_range = worksheet.Range(worksheet.Cells(first_row, 2), worksheet.Cells(last_row, chemistry_last_column))
    chemistry_range.NumberFormat = "0.00"


def format_chemistry_number_block(worksheet, first_row: int, last_row: int, first_column: int, last_column: int) -> None:
    """Format one known chemistry range to two decimal places without searching headers."""
    if last_row < first_row or last_column < first_column:
        return
    worksheet.Range(worksheet.Cells(first_row, first_column), worksheet.Cells(last_row, last_column)).NumberFormat = "0.00"


def format_report_range(worksheet, first_row: int, last_row: int, first_column: int, last_column: int) -> None:
    """Apply standard row, column, and font formatting to one report range."""
    if last_row < first_row or last_column < first_column:
        return
    worksheet.Rows(f"{first_row}:{last_row}").RowHeight = ROW_HEIGHT
    for column_number in range(first_column, last_column + 1):
        worksheet.Columns(column_number).ColumnWidth = COLUMN_WIDTH
    report_range = worksheet.Range(worksheet.Cells(first_row, first_column), worksheet.Cells(last_row, last_column))
    report_range.Font.Size = FONT_SIZE
    report_range.HorizontalAlignment = XL_CENTER
    report_range.VerticalAlignment = XL_CENTER
    if last_row >= first_row + 1:
        worksheet.Range(worksheet.Cells(first_row + 1, first_column), worksheet.Cells(last_row, last_column)).Font.Bold = False


def format_report_sheet(worksheet, last_row: int, last_column: int) -> None:
    """Apply standard row, column, and font formatting to the main report sheet."""
    format_report_range(worksheet, 1, last_row, 1, last_column)
    worksheet.Rows(1).Font.Bold = True
    worksheet.Rows(1).WrapText = True
    worksheet.Rows(1).HorizontalAlignment = XL_CENTER
    worksheet.Rows(1).VerticalAlignment = XL_CENTER


def chemistry_row_has_data(worksheet, row_number: int, source_columns: list[int]) -> bool:
    """Return True when any requested chemistry column has a value on this row."""
    for source_column in source_columns:
        value = worksheet.Cells(row_number, source_column).Value
        if value is not None and str(value).strip() != "":
            return True
    return False


def find_chemistry_data_count(worksheet, header_row: int, source_columns: list[int]) -> int:
    """Count contiguous chemistry data rows below the normalized header row."""
    count = 0
    row_number = header_row + 1
    last_row = used_last_row(worksheet)
    while row_number <= last_row and chemistry_row_has_data(worksheet, row_number, source_columns):
        count += 1
        row_number += 1
    return count


def mark_average_from_chemistry_count(report_worksheet, final_data_row: int) -> None:
    """Replace the final chemistry row number with Average and bold cells before images."""
    reflected_light_header = find_header_cell(report_worksheet, "R. Light Images")
    report_worksheet.Cells(final_data_row, 1).Value = "Average"
    average_range = report_worksheet.Range(report_worksheet.Cells(final_data_row, 1), report_worksheet.Cells(final_data_row, reflected_light_header.Column - 1))
    average_range.Font.Bold = True
    average_range.Font.Size = FONT_SIZE
    average_range.HorizontalAlignment = XL_CENTER
    average_range.VerticalAlignment = XL_CENTER


def paste_au_chemistry(report_worksheet, data_worksheet, layout: SampleFinishLayout) -> int:
    """Paste normalized chemistry from the Au sheet into the report after No."""
    header_row = find_last_normalized_header_row(data_worksheet)
    source_columns = [find_column_on_row(data_worksheet, header_row, header) for header in layout.chemistry_headers]
    chemistry_data_count = find_chemistry_data_count(data_worksheet, header_row, source_columns)
    if chemistry_data_count == 0:
        raise ValueError("No chemistry data rows were found below the final Normalized header in the Au sheet.")

    final_data_row = chemistry_data_count + 1

    for index, header in enumerate(layout.chemistry_headers, start=2):
        report_worksheet.Cells(1, index).Value = f"{header} (Wt%)"

    for report_row in range(2, final_data_row + 1):
        source_row = header_row + report_row - 1
        for offset, source_column in enumerate(source_columns, start=2):
            value = data_worksheet.Cells(source_row, source_column).Value
            if isinstance(value, (int, float)):
                value = round(value, 2)
            report_worksheet.Cells(report_row, offset).Value = value

    mark_average_from_chemistry_count(report_worksheet, final_data_row)
    format_report_sheet(report_worksheet, max(used_last_row(report_worksheet), final_data_row), used_last_column(report_worksheet))
    mark_average_from_chemistry_count(report_worksheet, final_data_row)
    format_chemistry_number_columns(report_worksheet, 2, final_data_row)
    return final_data_row


def set_shapes_to_move_and_size(worksheet) -> None:
    """Make shapes move with copied cells where Excel supports it."""
    for shape in worksheet.Shapes:
        shape.Placement = XL_MOVE_AND_SIZE


def apply_thick_outside_border(range_object) -> None:
    """Apply thick outside borders around an Excel range."""
    for edge in (XL_EDGE_LEFT, XL_EDGE_TOP, XL_EDGE_BOTTOM, XL_EDGE_RIGHT):
        border = range_object.Borders(edge)
        border.LineStyle = XL_CONTINUOUS
        border.Weight = XL_THICK


def organize_blocks(report_workbook, source_worksheet, layout: SampleFinishLayout, source_last_row: int) -> tuple:
    """Create a new side-by-side block sheet under the source report sheet and return block ranges."""
    set_shapes_to_move_and_size(source_worksheet)
    block_ranges = []
    reflected_light_header = find_header_cell(source_worksheet, "R. Light Images")
    reflected_light_column = reflected_light_header.Column
    sem_header = find_header_cell(source_worksheet, "SEM Images")
    source_last_column = sem_header.Column
    delete_existing_organized_blocks_sheet(report_workbook, protected_worksheet=source_worksheet)
    block_sheet = report_workbook.Worksheets.Add(After=source_worksheet)
    block_sheet.Name = ORGANIZED_BLOCKS_SHEET_NAME

    half = layout.half_block_size
    block_size = layout.block_size
    output_row = 1
    source_start_row = 2

    while source_start_row <= source_last_row:
        left_count = min(half, source_last_row - source_start_row + 1)
        right_start_row = source_start_row + half
        right_count = min(half, max(0, source_last_row - right_start_row + 1))

        source_worksheet.Range(source_worksheet.Cells(1, 1), source_worksheet.Cells(1, source_last_column)).Copy(
            Destination=block_sheet.Cells(output_row, 1)
        )
        source_worksheet.Range(source_worksheet.Cells(source_start_row, 1), source_worksheet.Cells(source_start_row + left_count - 1, source_last_column)).Copy(
            Destination=block_sheet.Cells(output_row + 1, 1)
        )

        right_output_column = source_last_column + 1
        if right_count > 0:
            source_worksheet.Range(source_worksheet.Cells(1, 1), source_worksheet.Cells(1, source_last_column)).Copy(
                Destination=block_sheet.Cells(output_row, right_output_column)
            )
            source_worksheet.Range(source_worksheet.Cells(right_start_row, 1), source_worksheet.Cells(right_start_row + right_count - 1, source_last_column)).Copy(
                Destination=block_sheet.Cells(output_row + 1, right_output_column)
            )

        block_height = max(left_count, right_count) + 1
        block_width = source_last_column * (2 if right_count > 0 else 1)
        block_last_row = output_row + block_height - 1

        format_report_range(block_sheet, output_row, block_last_row, 1, block_width)
        format_chemistry_number_block(block_sheet, output_row + 1, block_last_row, 2, reflected_light_column - 1)
        if right_count > 0:
            format_chemistry_number_block(
                block_sheet,
                output_row + 1,
                block_last_row,
                source_last_column + 2,
                source_last_column + reflected_light_column - 1,
            )
        apply_wrapped_headers(block_sheet, output_row, 1, source_last_column)
        if right_count > 0:
            apply_wrapped_headers(block_sheet, output_row, right_output_column, right_output_column + source_last_column - 1)
        block_range = block_sheet.Range(block_sheet.Cells(output_row, 1), block_sheet.Cells(block_last_row, block_width))
        apply_thick_outside_border(block_range)
        block_ranges.append((output_row, 1, block_last_row, block_width))

        source_start_row += block_size
        output_row = block_last_row + 2

    return block_sheet, tuple(block_ranges)


def main() -> None:
    """Run the finishing workflow."""
    root = Tk()
    root.withdraw()

    layout = choose_sample_layout()
    report_path = choose_workbook("Select the Au report workbook created by the image script")
    data_path = choose_workbook("Select the Excel data workbook containing the Au sheet")
    output_path = choose_output_file(default_final_output_name(report_path))
    import win32com.client

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    previous_excel_settings = configure_excel_for_speed(excel)

    report_workbook = open_excel_workbook(excel, report_path)
    data_workbook = open_excel_workbook(excel, data_path)
    report_worksheet = find_report_worksheet(report_workbook)
    data_worksheet = data_workbook.Worksheets("Au")

    final_data_row = paste_au_chemistry(report_worksheet, data_worksheet, layout)
    block_sheet, block_ranges = organize_blocks(report_workbook, report_worksheet, layout, final_data_row)

    save_finished_workbook(excel, report_workbook, output_path)

    data_workbook.Close(SaveChanges=False)
    report_workbook.Close(SaveChanges=True)
    restore_excel_settings(excel, previous_excel_settings)
    excel.Quit()

    messagebox.showinfo("Au report finished", f"Finished workbook saved successfully:\n\n{output_path}")


if __name__ == "__main__":
    main()
