"""Optional standalone finish for A2 workbooks when microscope imaging is omitted.

Preserves the input workbook and its original AZtec headers. Adds B3-style
correction columns to feature-data sheets, offers B2-style column visibility,
and replaces Summary with formulas based on physical feature area in um2.
"""

from copy import copy
from dataclasses import dataclass
from pathlib import Path
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from typing import Tuple

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter, quote_sheetname
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.properties import CalcProperties


NEW_HEADERS = ("Open Image", "Area Orig", "Corr %", "New Brea", "New Len", "Asso", "Diam")
ADMIN_SHEETS = {"area", "summary", "integritycheck", "checkreport"}
NOT_SUMMARY_CATEGORIES = {"aftersizecutoff", "cinnabar", "cinnabarasso"}


def normalize(value):
    text = str(value or "").strip().lower().replace("µ", "u").replace("μ", "u")
    return re.sub(r"[^a-z0-9]", "", text)


def find_column(headers, aliases, description, sheet_name):
    positions = [i for i, header in enumerate(headers, 1) if normalize(header) in aliases]
    if len(positions) != 1:
        raise ValueError(f"{sheet_name}: expected one {description} column; found {len(positions)}.")
    return positions[0]


def physical_area_column(headers, sheet_name):
    # A2 retains the unit-bearing column. B2 renames it to the SECOND bare Area.
    physical = [i for i, header in enumerate(headers, 1)
                if normalize(header) in {"areasqum", "areaum2", "areasquareum"}]
    if len(physical) == 1:
        return physical[0]
    if len(physical) > 1:
        raise ValueError(f"{sheet_name}: more than one unit-bearing physical Area column.")
    bare = [i for i, header in enumerate(headers, 1) if normalize(header) == "area"]
    if len(bare) == 2:
        return bare[1]
    raise ValueError(f"{sheet_name}: cannot identify Area (sq. µm) or the second Area column.")


@dataclass(frozen=True)
class FeatureTable:
    sheet_name: str
    feature_col: int
    area_col: int
    breadth_col: int
    length_col: int
    rows: Tuple[int, ...]
    block_start: int
    block_exists: bool


def inspect_tables(book):
    """Validate every feature table before making any output changes."""
    tables = []
    for ws in book:
        if normalize(ws.title) in ADMIN_SHEETS:
            continue
        headers = [cell.value for cell in ws[1]]
        if "feature" not in {normalize(h) for h in headers}:
            continue
        feature = find_column(headers, {"feature"}, "Feature", ws.title)
        area = physical_area_column(headers, ws.title)
        breadth = find_column(headers, {"breadthum", "breadth"}, "Breadth (µm)", ws.title)
        length = find_column(headers, {"lengthum", "length"}, "Length (µm)", ws.title)
        exists = (len(headers) >= len(NEW_HEADERS)
                  and tuple(normalize(h) for h in headers[-len(NEW_HEADERS):])
                  == tuple(normalize(h) for h in NEW_HEADERS))
        if not exists and any(normalize(h) in {normalize(n) for n in NEW_HEADERS[:-1]} for h in headers):
            raise ValueError(f"{ws.title}: contains a partial or displaced correction block. "
                             "Select the original A2 workbook or an output made by this script.")
        start = len(headers) - len(NEW_HEADERS) + 1 if exists else len(headers) + 1
        rows = tuple(row for row in range(2, ws.max_row + 1)
                     if ws.cell(row, feature).value is not None
                     and str(ws.cell(row, feature).value).strip() != "")
        tables.append(FeatureTable(ws.title, feature, area, breadth, length, rows, start, exists))
    if not tables:
        raise ValueError("No A2 feature-data sheets found. Select the categorized workbook.")
    return tables


def add_correction_columns(ws, table):
    start = table.block_start
    columns = {name: start + offset for offset, name in enumerate(NEW_HEADERS)}
    fill = PatternFill(fill_type="solid", fgColor="D9E5F2")
    for header, col in columns.items():
        cell = ws.cell(1, col, header)
        cell.font = Font(name="Calibri", bold=True, color="20364D")
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(col)].width = 15
    for row in table.rows:
        corr = f"{get_column_letter(columns['Corr %'])}{row}"
        area = f"{get_column_letter(columns['Area Orig'])}{row}"
        breadth = f"{get_column_letter(table.breadth_col)}{row}"
        length = f"{get_column_letter(table.length_col)}{row}"
        link = ws.cell(row, columns["Open Image"])
        link.value = None
        link.hyperlink = None
        if not table.block_exists:
            # B3 copies the original physical area into Area Orig.
            ws.cell(row, columns["Area Orig"]).value = ws.cell(row, table.area_col).value
        # Exactly the B3 correction behavior: blank/zero keeps original geometry.
        ws.cell(row, columns["New Brea"]).value = (
            f'=IF(OR({corr}="",{corr}=0),{breadth},SQRT({area}*{corr}/100))')
        ws.cell(row, columns["New Len"]).value = (
            f'=IF(OR({corr}="",{corr}=0),{length},SQRT({area}*{corr}/100))')
        # Corr %, Asso and Diam remain blank/editable; preserve them on reruns.
        for header, col in columns.items():
            cell = ws.cell(row, col)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if header not in ("Open Image", "Asso"):
                cell.number_format = "0.00" if header == "Corr %" else "0.000"
    if table.rows and not table.block_exists:
        validation = DataValidation(type="decimal", operator="between", formula1=0,
                                    formula2=100, allow_blank=True)
        validation.errorTitle = "Correction percentage"
        validation.error = "Enter a number between 0 and 100, or leave blank."
        validation.showErrorMessage = True
        validation.errorStyle = "stop"
        ws.add_data_validation(validation)
        col = get_column_letter(columns["Corr %"])
        validation.add(f"{col}2:{col}{max(table.rows)}")
    ws.freeze_panes = "A2"


def header_keys(ws):
    """Distinguish repeated headers so hiding the first Area does not hide both."""
    seen = {}
    result = []
    for col in range(1, ws.max_column + 1):
        value = ws.cell(1, col).value
        text = str(value).strip() if value is not None else f"(Unnamed column {col})"
        key = normalize(text)
        seen[key] = seen.get(key, 0) + 1
        result.append(((key, seen[key]), text))
    return result


def visibility_choices(book, tables):
    choices = {}
    for table in tables:
        for key, label in header_keys(book[table.sheet_name]):
            if key not in choices:
                choices[key] = label if key[1] == 1 else f"{label} [occurrence {key[1]}]"
    return list(choices.items())


def choose_visible_columns(root, choices):
    """B2-style selection: all visible by default; cancel means show all."""
    window = tk.Toplevel(root)
    window.title("Choose columns to SHOW — optional A2 output")
    window.geometry("760x620")
    tk.Label(window, text="Select columns to SHOW; unselected columns will be hidden.\n"
             "Default: all columns selected. Ctrl/Shift allows multiple selection.",
             justify="left", anchor="w").pack(fill="x", padx=10, pady=10)
    frame = tk.Frame(window)
    frame.pack(fill="both", expand=True, padx=10, pady=5)
    scrollbar = tk.Scrollbar(frame)
    scrollbar.pack(side="right", fill="y")
    box = tk.Listbox(frame, selectmode="extended", exportselection=False,
                     yscrollcommand=scrollbar.set)
    box.pack(side="left", fill="both", expand=True)
    scrollbar.configure(command=box.yview)
    for _, label in choices:
        box.insert("end", label)
    box.selection_set(0, "end")
    result = {"visible": None}

    def confirm():
        result["visible"] = {choices[i][0] for i in box.curselection()}
        window.destroy()

    buttons = tk.Frame(window)
    buttons.pack(fill="x", padx=10, pady=10)
    tk.Button(buttons, text="Select All", command=lambda: box.selection_set(0, "end")).pack(side="left")
    tk.Button(buttons, text="Select None", command=lambda: box.selection_clear(0, "end")).pack(side="left", padx=5)
    tk.Button(buttons, text="OK", command=confirm).pack(side="right")
    tk.Button(buttons, text="Cancel (show all)", command=window.destroy).pack(side="right", padx=5)
    window.protocol("WM_DELETE_WINDOW", window.destroy)
    window.transient(root)
    window.grab_set()
    root.wait_window(window)
    return result["visible"]


def apply_visibility(ws, visible):
    # Flatten pre-existing grouped column dimensions before applying explicit
    # per-column visibility; otherwise a hidden group can override visible cells.
    dimensions = list(ws.column_dimensions.values())
    expanded = {}
    for col, (key, _) in enumerate(header_keys(ws), 1):
        letter = get_column_letter(col)
        original = ws.column_dimensions.get(letter)
        if original is None:
            original = next((d for d in dimensions if d.min and d.max and d.min <= col <= d.max), None)
        if original is not None:
            expanded[col] = copy(original)
        else:
            expanded[col] = copy(ws.column_dimensions[letter])
        dimension = expanded[col]
        dimension.index = letter
        dimension.min = col
        dimension.max = col
        dimension.hidden = visible is not None and key not in visible
        if not dimension.hidden and (dimension.width is None or dimension.width <= 0):
            dimension.width = 12
    ws.column_dimensions.clear()
    for col, dimension in expanded.items():
        ws.column_dimensions[get_column_letter(col)] = dimension


def write_summary(book, tables):
    previous = next((ws for ws in book if normalize(ws.title) == "summary"), None)
    if previous is not None:
        index = book.sheetnames.index(previous.title)
        book.remove(previous)
    else:
        index = 1 if "Area" in book.sheetnames else 0
    ws = book.create_sheet("Summary", index)
    ws.append(("Category (Sheet name)", "Total Sum of Area (sq. µm)", "Percentage"))
    categories = [t for t in tables if t.rows and normalize(t.sheet_name) != "rawdata"
                  and normalize(t.sheet_name) not in NOT_SUMMARY_CATEGORIES]
    total_row = len(categories) + 2
    for row, table in enumerate(categories, 2):
        # Force literal text for unusual sheet names starting with an equals sign.
        label = ws.cell(row, 1, table.sheet_name)
        label.data_type = "s"
        col = get_column_letter(table.area_col)
        sheet = quote_sheetname(table.sheet_name)
        ws.cell(row, 2, f"=SUM({sheet}!${col}$2:${col}${max(table.rows)})")
        ws.cell(row, 3, f"=IF($B${total_row}=0,0,B{row}/$B${total_row})")
    ws.cell(total_row, 1, "Total")
    ws.cell(total_row, 2, f"=SUM(B2:B{total_row - 1})" if categories else "=0")
    ws.cell(total_row, 3, f"=SUM(C2:C{total_row - 1})" if categories else "=0")
    thin = Side(style="thin", color="B4C2D0")
    medium = Side(style="medium", color="52657A")
    for row in ws:
        for cell in row:
            cell.font = Font(name="Calibri", size=11)
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    for cell in ws[1]:
        cell.fill = PatternFill(fill_type="solid", fgColor="20364D")
        cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    for cell in ws[total_row]:
        cell.fill = PatternFill(fill_type="solid", fgColor="D9E5F2")
        cell.font = Font(name="Calibri", size=11, bold=True)
        cell.border = Border(left=thin, right=thin, top=medium, bottom=medium)
    for row in range(2, total_row + 1):
        ws.cell(row, 2).number_format = "0.000"
        ws.cell(row, 3).number_format = "0.00%"
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 16
    ws.row_dimensions[1].height = 32
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    if categories:
        ws.auto_filter.ref = f"A1:C{total_row - 1}"


def unique_output_path(source):
    output = source.with_name(f"{source.stem}_optional_final{source.suffix}")
    number = 2
    while output.exists() or output.resolve() == source.resolve():
        output = source.with_name(f"{source.stem}_optional_final_{number}{source.suffix}")
        number += 1
    return output


def build_optional_workbook(source, visible_columns=None, visibility_picker=None):
    """Finish a copy; the original workbook is never saved or overwritten."""
    source = Path(source)
    if source.suffix.lower() not in (".xlsx", ".xlsm"):
        raise ValueError("Select an A2 .xlsx or .xlsm workbook.")
    book = load_workbook(source, keep_vba=source.suffix.lower() == ".xlsm")
    try:
        tables = inspect_tables(book)
        for table in tables:
            add_correction_columns(book[table.sheet_name], table)
        visible = visibility_picker(visibility_choices(book, tables)) if visibility_picker else visible_columns
        for table in tables:
            apply_visibility(book[table.sheet_name], visible)
        write_summary(book, tables)
        book.calculation = CalcProperties(calcId=0, fullCalcOnLoad=True, forceFullCalc=True, calcMode="auto")
        output = unique_output_path(source)
        book.save(output)
        print(f"Optional Phase A output saved: {output}", flush=True)
        return output
    finally:
        book.close()


def main():
    root = tk.Tk()
    root.withdraw()
    try:
        source = filedialog.askopenfilename(
            parent=root, title="Select A2 categorized workbook — optional Phase A finish",
            filetypes=[("Excel workbooks", "*.xlsx *.xlsm")])
        if not source:
            return
        output = build_optional_workbook(
            source, visibility_picker=lambda choices: choose_visible_columns(root, choices))
        messagebox.showinfo("Optional Phase A finish", f"Created:\n{output}", parent=root)
    except (OSError, ValueError) as exc:
        messagebox.showerror("Optional output could not be created", str(exc), parent=root)
        raise  # Also retain the detailed error in PyCharm/app Activity output.
    finally:
        root.destroy()


if __name__ == "__main__":
    main()
