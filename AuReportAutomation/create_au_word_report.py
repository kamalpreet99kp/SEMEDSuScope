"""Create the Au Word report from a finished workbook.

Run this optional Windows/Word script after `finish_au_report.py` has created the
finished Excel workbook and its `Organized Blocks` worksheet. The Excel finishing
script intentionally does not create Word files anymore, so Word-specific testing
and changes can be handled separately from the working Excel workflow.
"""

from __future__ import annotations

from pathlib import Path
from time import sleep
from tkinter import Tk, filedialog, messagebox, simpledialog

ORGANIZED_BLOCKS_SHEET_NAME = "Organized Blocks"
WD_ALIGN_PARAGRAPH_CENTER = 1
WD_PAGE_BREAK = 7
WD_PASTE_OLE_OBJECT = 0
WD_IN_LINE = 0
WD_ORIENT_PORTRAIT = 0
WD_ORIENT_LANDSCAPE = 1
MSO_TRUE = -1
WORD_PROJECT_FONT_SIZE = 18
WORD_SAMPLE_FONT_SIZE = 14
FIRST_PAGE_HEADER_RESERVED_POINTS = 55

SAMPLE_ORIENTATIONS = {
    "1": ("Au+Ag", WD_ORIENT_PORTRAIT),
    "2": ("Au+Ag+Cu", WD_ORIENT_PORTRAIT),
    "3": ("Au+Ag+Cu+Hg", WD_ORIENT_LANDSCAPE),
    "4": ("Au+Ag+Hg", WD_ORIENT_PORTRAIT),
}


class UserCancelledError(Exception):
    """Raised when the user cancels a required prompt."""


def choose_sample_orientation() -> int:
    """Ask which sample layout should be used for Word page orientation."""
    prompt = (
        "Select sample type by number for Word page orientation:\n\n"
        "1 = Au+Ag (Portrait)\n"
        "2 = Au+Ag+Cu (Portrait)\n"
        "3 = Au+Ag+Cu+Hg (Landscape)\n"
        "4 = Au+Ag+Hg (Portrait)"
    )
    choice = simpledialog.askstring("Au Word Report Sample Type", prompt)
    if choice is None:
        raise UserCancelledError("Sample type selection was cancelled.")
    choice = choice.strip()
    if choice not in SAMPLE_ORIENTATIONS:
        messagebox.showerror("Invalid sample type", "Please run again and enter 1, 2, 3, or 4.")
        raise UserCancelledError(f"Invalid sample type choice: {choice}")
    return SAMPLE_ORIENTATIONS[choice][1]


def choose_workbook(title: str) -> Path:
    """Prompt for a finished Excel workbook path."""
    selected = filedialog.askopenfilename(
        title=title,
        filetypes=(("Excel workbooks", "*.xlsx *.xlsm"), ("All files", "*.*")),
    )
    if not selected:
        raise UserCancelledError(f"Workbook selection was cancelled: {title}")
    return Path(selected)


def choose_output_file(default_name: str) -> Path:
    """Prompt for a Word output path."""
    selected = filedialog.asksaveasfilename(
        title="Save Au Word report as",
        defaultextension=".docx",
        initialfile=default_name,
        filetypes=(("Word document", "*.docx"),),
    )
    if not selected:
        raise UserCancelledError("Word output selection was cancelled.")
    return Path(selected)


def ask_project_number() -> str:
    """Ask for the Word report project number."""
    project_number = simpledialog.askstring("Project No.", "Enter Project No. for the Word report:")
    if project_number is None:
        raise UserCancelledError("Project number entry was cancelled.")
    return project_number.strip() or "XXX"


def sample_name_from_workbook(path: Path) -> str:
    """Build the Word report sample name from the workbook filename."""
    sample_name = path.stem
    for suffix in ("_Final", " Final", "_Inter", " Inter"):
        if sample_name.endswith(suffix):
            sample_name = sample_name[: -len(suffix)]
    return sample_name.replace("(Au SEM)", "").replace("(au sem)", "").strip()


def used_last_row(worksheet) -> int:
    """Return the last used row in an Excel worksheet."""
    return worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1


def used_last_column(worksheet) -> int:
    """Return the last used column in an Excel worksheet."""
    return worksheet.UsedRange.Column + worksheet.UsedRange.Columns.Count - 1


def row_has_content(worksheet, row_number: int, last_column: int) -> bool:
    """Return True when a row has any cell value in the organized block area."""
    for column_number in range(1, last_column + 1):
        if worksheet.Cells(row_number, column_number).Value not in (None, ""):
            return True
    return False


def detect_block_ranges(block_sheet) -> tuple[tuple[int, int, int, int], ...]:
    """Detect organized block ranges separated by blank rows."""
    last_row = used_last_row(block_sheet)
    last_column = used_last_column(block_sheet)
    block_ranges = []
    row_number = 1

    while row_number <= last_row:
        while row_number <= last_row and not row_has_content(block_sheet, row_number, last_column):
            row_number += 1
        if row_number > last_row:
            break

        top_row = row_number
        while row_number <= last_row and row_has_content(block_sheet, row_number, last_column):
            row_number += 1
        bottom_row = row_number - 1
        block_ranges.append((top_row, 1, bottom_row, last_column))

    if not block_ranges:
        raise ValueError(f"No report blocks were detected on {ORGANIZED_BLOCKS_SHEET_NAME!r}.")
    return tuple(block_ranges)


def ole_prog_id(word_object) -> str:
    """Return an OLE object's ProgID when available."""
    try:
        return str(word_object.OLEFormat.ProgID)
    except Exception:
        return ""


def has_known_non_excel_prog_id(word_object) -> bool:
    """Return True only when Word exposes a non-Excel ProgID."""
    prog_id = ole_prog_id(word_object)
    return bool(prog_id and "Excel" not in prog_id)


def paste_excel_block_as_editable_ole_object(selection, document, block_range):
    """Paste a copied Excel block as an editable Microsoft Excel Worksheet Object."""
    before_inline_count = document.InlineShapes.Count
    before_shape_count = document.Shapes.Count
    block_range.Copy()
    sleep(0.2)
    selection.Range.PasteSpecial(
        Link=False,
        DataType=WD_PASTE_OLE_OBJECT,
        Placement=WD_IN_LINE,
        DisplayAsIcon=False,
    )

    after_inline_count = document.InlineShapes.Count
    if after_inline_count > before_inline_count:
        inline_shape = document.InlineShapes(after_inline_count)
        if has_known_non_excel_prog_id(inline_shape):
            raise RuntimeError(f"Paste Special created an inline non-Excel object. ProgID: {ole_prog_id(inline_shape)!r}")
        inline_shape.Width = block_range.Width
        inline_shape.Height = block_range.Height
        selection.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_CENTER
        return inline_shape

    after_shape_count = document.Shapes.Count
    if after_shape_count > before_shape_count:
        shape = document.Shapes(after_shape_count)
        if has_known_non_excel_prog_id(shape):
            raise RuntimeError(f"Paste Special created a floating non-Excel object. ProgID: {ole_prog_id(shape)!r}")
        inline_shape = shape.ConvertToInlineShape()
        inline_shape.Width = block_range.Width
        inline_shape.Height = block_range.Height
        selection.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_CENTER
        return inline_shape

    raise RuntimeError("Word did not create an Excel OLE object from Paste Special.")


def fit_inline_shape_to_page(inline_shape, document, reserve_header_space: bool) -> None:
    """Scale an embedded block so it fits on the current Word page."""
    page_setup = document.PageSetup
    available_width = page_setup.PageWidth - page_setup.LeftMargin - page_setup.RightMargin
    available_height = page_setup.PageHeight - page_setup.TopMargin - page_setup.BottomMargin
    if reserve_header_space:
        available_height -= FIRST_PAGE_HEADER_RESERVED_POINTS

    width_scale = available_width / inline_shape.Width if inline_shape.Width else 1
    height_scale = available_height / inline_shape.Height if inline_shape.Height else 1
    scale = min(1, width_scale, height_scale)
    if scale < 1:
        inline_shape.LockAspectRatio = MSO_TRUE
        inline_shape.Width = inline_shape.Width * scale


def create_word_report(word_app, block_sheet, block_ranges: tuple, project_number: str, sample_name: str, output_path: Path, orientation: int) -> None:
    """Create a Word report with one organized Excel block per page."""
    document = word_app.Documents.Add()
    word_app.Visible = False
    document.PageSetup.Orientation = orientation

    for block_index, (top_row, left_column, bottom_row, right_column) in enumerate(block_ranges):
        selection = word_app.Selection
        selection.EndKey(Unit=6)
        selection.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_CENTER

        if block_index > 0:
            selection.InsertBreak(WD_PAGE_BREAK)
            selection.EndKey(Unit=6)
            selection.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_CENTER

        if block_index == 0:
            selection.Font.Size = WORD_PROJECT_FONT_SIZE
            selection.Font.Bold = True
            selection.TypeText(f"Project No.: {project_number}")
            selection.TypeParagraph()

            selection.Font.Size = WORD_SAMPLE_FONT_SIZE
            selection.Font.Bold = True
            selection.TypeText(f"Sample Name: {sample_name}")
            selection.TypeParagraph()

        block_range = block_sheet.Range(block_sheet.Cells(top_row, left_column), block_sheet.Cells(bottom_row, right_column))
        inline_shape = paste_excel_block_as_editable_ole_object(selection, document, block_range)
        fit_inline_shape_to_page(inline_shape, document, reserve_header_space=(block_index == 0))

    document.SaveAs2(str(output_path.resolve()))
    document.Close(SaveChanges=True)


def main() -> None:
    """Run the optional Word export workflow."""
    root = Tk()
    root.withdraw()

    workbook_path = choose_workbook("Select the finished Au report workbook with Organized Blocks")
    orientation = choose_sample_orientation()
    project_number = ask_project_number()
    sample_name = sample_name_from_workbook(workbook_path)
    output_path = choose_output_file(f"{sample_name}.docx")

    import win32com.client

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    workbook = excel.Workbooks.Open(str(workbook_path.resolve()))
    block_sheet = workbook.Worksheets(ORGANIZED_BLOCKS_SHEET_NAME)
    block_ranges = detect_block_ranges(block_sheet)

    word = win32com.client.Dispatch("Word.Application")
    create_word_report(word, block_sheet, block_ranges, project_number, sample_name, output_path, orientation)
    word.Quit()

    workbook.Close(SaveChanges=False)
    excel.Quit()
    messagebox.showinfo("Au Word report created", f"Word report saved successfully:\n\n{output_path}")


if __name__ == "__main__":
    main()
