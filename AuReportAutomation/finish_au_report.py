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


def used_last_row(worksheet) -> int:
    """Return the last used row in an Excel worksheet."""
    return worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1


def used_last_column(worksheet) -> int:
    """Return the last used column in an Excel worksheet."""
    return worksheet.UsedRange.Column + worksheet.UsedRange.Columns.Count - 1


def find_header_cell(worksheet, header_text: str):
    """Find a cell by exact header text, case-insensitive."""
    wanted = header_text.strip().lower()
    for row_number in range(1, used_last_row(worksheet) + 1):
        for column_number in range(1, used_last_column(worksheet) + 1):
            value = worksheet.Cells(row_number, column_number).Value
            if value is not None and str(value).strip().lower() == wanted:
                return worksheet.Cells(row_number, column_number)
    raise ValueError(f"Could not find header {header_text!r} in sheet {worksheet.Name!r}.")


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


def format_report_sheet(worksheet, last_row: int, last_column: int) -> None:
    """Apply standard row, column, and font formatting."""
    worksheet.Rows(f"1:{last_row}").RowHeight = ROW_HEIGHT
    for column_number in range(1, last_column + 1):
        worksheet.Columns(column_number).ColumnWidth = COLUMN_WIDTH
    worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(last_row, last_column)).Font.Size = FONT_SIZE
    worksheet.Rows(1).Font.Bold = True
    if last_row >= 2:
        worksheet.Range(worksheet.Cells(2, 1), worksheet.Cells(last_row, last_column)).Font.Bold = False


def paste_au_chemistry(report_worksheet, data_worksheet, layout: SampleFinishLayout) -> None:
    """Paste normalized chemistry from the Au sheet into the report after No."""
    header_row = find_last_normalized_header_row(data_worksheet)
    source_columns = [find_column_on_row(data_worksheet, header_row, header) for header in layout.chemistry_headers]
    report_last_row = used_last_row(report_worksheet)

    for index, header in enumerate(layout.chemistry_headers, start=2):
        report_worksheet.Cells(1, index).Value = f"{header} (Wt%)"

    for report_row in range(2, report_last_row + 1):
        source_row = header_row + report_row - 1
        for offset, source_column in enumerate(source_columns, start=2):
            report_worksheet.Cells(report_row, offset).Value = data_worksheet.Cells(source_row, source_column).Value

    format_report_sheet(report_worksheet, report_last_row, used_last_column(report_worksheet))


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


def organize_blocks(report_workbook, source_worksheet, layout: SampleFinishLayout) -> None:
    """Create a new side-by-side block sheet under the source report sheet."""
    set_shapes_to_move_and_size(source_worksheet)
    source_last_row = used_last_row(source_worksheet)
    sem_header = find_header_cell(source_worksheet, "SEM Images")
    source_last_column = sem_header.Column
    block_sheet = report_workbook.Worksheets.Add(After=source_worksheet)
    block_sheet.Name = "Organized Blocks"

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
        apply_thick_outside_border(block_sheet.Range(block_sheet.Cells(output_row, 1), block_sheet.Cells(output_row + block_height - 1, block_width)))
        format_report_sheet(block_sheet, output_row + block_height - 1, block_width)

        source_start_row += block_size
        output_row += block_height + 1


def main() -> None:
    """Run the finishing workflow."""
    root = Tk()
    root.withdraw()

    layout = choose_sample_layout()
    report_path = choose_workbook("Select the Au report workbook created by the image script")
    data_path = choose_workbook("Select the Excel data workbook containing the Au sheet")
    output_path = choose_output_file(f"{report_path.stem}_finished.xlsx")

    import win32com.client

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    report_workbook = excel.Workbooks.Open(str(report_path.resolve()))
    data_workbook = excel.Workbooks.Open(str(data_path.resolve()))
    report_worksheet = report_workbook.Worksheets(1)
    data_worksheet = data_workbook.Worksheets("Au")

    paste_au_chemistry(report_worksheet, data_worksheet, layout)
    organize_blocks(report_workbook, report_worksheet, layout)

    report_workbook.SaveAs(str(output_path.resolve()), FileFormat=51)
    data_workbook.Close(SaveChanges=False)
    report_workbook.Close(SaveChanges=True)
    excel.Quit()

    messagebox.showinfo("Au report finished", f"Finished workbook saved successfully:\n\n{output_path}")


if __name__ == "__main__":
    main()
