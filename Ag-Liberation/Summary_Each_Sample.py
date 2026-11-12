import re
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter

IGNORE_SHEETS = {"Area", "Summary", "Raw Data", "Integrity Check", "Check Report"}
REQUIRED_HEADERS = ["Feature", "Area Orig", "Corr %", "New Brea", "New Len", "Asso", "Diam", "New Area"]
OPTIONAL_HEADERS = {"Diam", "New Area"}
HEADER_BLANK_ROWS = 5
BLOCK_GAP_COLUMNS = 2
BLOCK_COLORS = ("D9E5F2", "E2EFDA", "FCE4D6", "E4DFEC")


class LiveLogger:
    def __init__(self, root: tk.Tk):
        self.win = tk.Toplevel(root)
        self.win.title("Summary Each Sample - Live Progress")
        self.win.geometry("980x540")
        self.text = tk.Text(self.win, wrap="word")
        self.text.pack(fill="both", expand=True)
        self.write("=== Summary Each Sample ===")

    def write(self, msg: str):
        self.text.insert("end", msg + "\n")
        self.text.see("end")
        self.win.update_idletasks()
        self.win.update()


def clean_fraction_name(filename: str) -> str:
    stem = Path(filename).stem
    stem = re.sub(r"^Complete-MASTER_Classified_Tellurides_", "", stem, flags=re.IGNORECASE)
    stem = re.sub(r"_Excel_with_quick_links", "", stem, flags=re.IGNORECASE)
    stem = re.sub(r"MASTER_Classified_?", "", stem, flags=re.IGNORECASE)
    stem = re.sub(r"\s*\(AuTe\s*Scan\)\s*$", "", stem, flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", stem).strip() or stem


def choose_sample_folder(root: tk.Tk) -> Optional[Path]:
    folder = filedialog.askdirectory(parent=root, title="Select SAMPLE folder containing size-fraction Excel files")
    if not folder:
        return None
    return Path(folder)


def find_excel_files(sample_dir: Path) -> List[Path]:
    files = [p for p in sample_dir.iterdir() if p.is_file() and p.suffix.lower() in {".xlsx", ".xlsm", ".xls"}]
    return sorted(files, key=lambda p: p.name.lower())


def choose_file_order_gui(root: tk.Tk, files: Sequence[Path]) -> List[Path]:
    win = tk.Toplevel(root)
    win.title("Select and order size-fraction files")
    win.geometry("980x620")

    tk.Label(
        win,
        text=(
            "Select files to include and set order.\n"
            "Order in this list controls left-to-right placement in each output category sheet."
        ),
        justify="left",
        anchor="w",
    ).pack(fill="x", padx=10, pady=8)

    frame = tk.Frame(win)
    frame.pack(fill="both", expand=True, padx=10, pady=8)

    listbox = tk.Listbox(frame, selectmode="extended")
    listbox.pack(side="left", fill="both", expand=True)

    sb = tk.Scrollbar(frame, orient="vertical", command=listbox.yview)
    sb.pack(side="right", fill="y")
    listbox.configure(yscrollcommand=sb.set)

    display = [f"{clean_fraction_name(f.name)}   [{f.name}]" for f in files]
    for item in display:
        listbox.insert("end", item)

    listbox.selection_set(0, "end")

    result = {"selected": list(files)}

    def move(delta: int):
        sel = list(listbox.curselection())
        if not sel:
            return
        if delta < 0 and sel[0] == 0:
            return
        if delta > 0 and sel[-1] == listbox.size() - 1:
            return

        texts = [listbox.get(i) for i in range(listbox.size())]
        selected_flags = [i in sel for i in range(len(texts))]

        for i in (sel if delta < 0 else reversed(sel)):
            j = i + delta
            texts[i], texts[j] = texts[j], texts[i]
            selected_flags[i], selected_flags[j] = selected_flags[j], selected_flags[i]

        listbox.delete(0, "end")
        for t in texts:
            listbox.insert("end", t)
        for i, is_sel in enumerate(selected_flags):
            if is_sel:
                listbox.selection_set(i)

    def ok():
        selected_idx = listbox.curselection()
        if not selected_idx:
            messagebox.showwarning("No files selected", "Please select at least one input file.", parent=win)
            return
        ordered_text = [listbox.get(i) for i in range(listbox.size())]
        text_to_file = {display[i]: files[i] for i in range(len(files))}
        result["selected"] = [text_to_file[t] for t in ordered_text if listbox.selection_includes(ordered_text.index(t))]
        win.destroy()

    def cancel():
        result["selected"] = []
        win.destroy()

    btn = tk.Frame(win)
    btn.pack(fill="x", padx=10, pady=8)
    tk.Button(btn, text="Move Up", command=lambda: move(-1)).pack(side="left")
    tk.Button(btn, text="Move Down", command=lambda: move(1)).pack(side="left", padx=5)
    tk.Button(btn, text="Select All", command=lambda: listbox.selection_set(0, "end")).pack(side="left", padx=20)
    tk.Button(btn, text="Clear", command=lambda: listbox.selection_clear(0, "end")).pack(side="left")
    tk.Button(btn, text="OK", command=ok).pack(side="right")
    tk.Button(btn, text="Cancel", command=cancel).pack(side="right", padx=5)

    win.grab_set()
    win.wait_window()
    return result["selected"]


def discover_categories(files: Sequence[Path], logger: LiveLogger) -> List[str]:
    cats = []
    seen = set()
    for p in files:
        try:
            with pd.ExcelFile(p) as xls:
                for s in xls.sheet_names:
                    if s not in IGNORE_SHEETS and s not in seen:
                        cats.append(s)
                        seen.add(s)
            logger.write(f"Read sheets from: {p.name}")
        except Exception as exc:
            logger.write(f"WARNING: Could not read {p.name}: {exc}")
    return cats


def choose_categories_gui(root: tk.Tk, categories: Sequence[str]) -> List[str]:
    win = tk.Toplevel(root)
    win.title("Select mineral categories to include")
    win.geometry("900x650")

    tk.Label(win, text="Select/deselect category sheets for the summary output.", anchor="w").pack(fill="x", padx=10, pady=8)

    canvas = tk.Canvas(win)
    frame = tk.Frame(canvas)
    sb = tk.Scrollbar(win, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
    canvas.create_window((0, 0), window=frame, anchor="nw")
    frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

    vars_ = []
    for i, cat in enumerate(categories):
        v = tk.BooleanVar(value=True)
        vars_.append(v)
        tk.Checkbutton(frame, variable=v, text=cat, anchor="w", width=80, justify="left").grid(row=i, column=0, sticky="w")

    result = {"cats": list(categories)}

    def select_all():
        for v in vars_:
            v.set(True)

    def select_none():
        for v in vars_:
            v.set(False)

    def ok():
        chosen = [cat for cat, v in zip(categories, vars_) if v.get()]
        if not chosen:
            messagebox.showwarning("No categories selected", "Please select at least one category.", parent=win)
            return
        result["cats"] = chosen
        win.destroy()

    def cancel():
        result["cats"] = []
        win.destroy()

    btn = tk.Frame(win)
    btn.pack(fill="x", padx=10, pady=8)
    tk.Button(btn, text="Select All", command=select_all).pack(side="left")
    tk.Button(btn, text="Select None", command=select_none).pack(side="left", padx=5)
    tk.Button(btn, text="OK", command=ok).pack(side="right")
    tk.Button(btn, text="Cancel", command=cancel).pack(side="right", padx=5)

    win.grab_set()
    win.wait_window()
    return result["cats"]


def extract_columns(df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
    """Keep source column order, appending missing columns as blank fields."""
    cols_map = {str(c).strip().lower(): c for c in df.columns}
    canonical = {header.lower(): header for header in REQUIRED_HEADERS}
    ordered = []
    for original in df.columns:
        header = canonical.get(str(original).strip().lower())
        if header is not None and header not in ordered:
            ordered.append(header)
    ordered.extend(header for header in REQUIRED_HEADERS if header not in ordered)
    out = pd.DataFrame(index=df.index)
    missing = []
    for wanted in ordered:
        original = cols_map.get(wanted.lower())
        if original is None:
            out[wanted] = None
            if wanted not in OPTIONAL_HEADERS:
                missing.append(wanted)
        else:
            out[wanted] = df[original].values
    return out, missing


def read_values(file_path: Path, sheet_name: str, logger: LiveLogger) -> pd.DataFrame:
    """Read cached results only; never carry source formulas into the summary."""
    if file_path.suffix.lower() == ".xls":
        return pd.read_excel(file_path, sheet_name=sheet_name)
    values_book = load_workbook(file_path, read_only=True, data_only=True)
    try:
        if sheet_name not in values_book.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not present in this file")
        rows = list(values_book[sheet_name].values)
        # Excel must have calculated/saved formula results before values-only copying.
        formulas_book = load_workbook(file_path, read_only=True, data_only=False)
        try:
            uncached = 0
            for formula_row, value_row in zip(
                formulas_book[sheet_name].iter_rows(), rows
            ):
                for source, value in zip(formula_row, value_row):
                    if source.data_type == "f" and value is None:
                        uncached += 1
            if uncached:
                logger.write(
                    f"WARNING: {file_path.name} / {sheet_name}: {uncached} formula(s) "
                    "have no saved result. These cells will be blank. "
                    "Open, recalculate and save the input in Excel, then rerun."
                )
        finally:
            formulas_book.close()
        if not rows:
            return pd.DataFrame()
        headers = rows[0]
        data = list(rows[1:])
        while data and all(value is None for value in data[-1]):
            data.pop()
        return pd.DataFrame(data, columns=headers)
    finally:
        values_book.close()


def put_value(ws, row: int, column: int, value):
    """Write data literally, including text that starts with an equals sign."""
    if value is None or pd.isna(value):
        value = None
    cell = ws.cell(row=row, column=column, value=value)
    if isinstance(value, str):
        cell.data_type = "s"
    return cell


def format_block_sheet(ws, blocks):
    """Style fraction blocks with a common bottom edge and compact columns."""
    last_row = max(ws.max_row, HEADER_BLANK_ROWS + 1)
    edge = Side(style="thick", color="52657A")
    total_edge = Side(style="thin", color="52657A")
    for index, (start, width) in enumerate(blocks):
        end = start + width - 1
        fill = PatternFill(fill_type="solid", fgColor=BLOCK_COLORS[index % len(BLOCK_COLORS)])
        for col in range(start, end + 1):
            ws.cell(1, col).fill = fill
        ws.merge_cells(start_row=1, start_column=start, end_row=1, end_column=end)
        title = ws.cell(1, start)
        title.font = Font(name="Calibri", size=11, bold=True, color="20364D")
        title.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for col in range(start, end + 1):
            header = str(ws.cell(HEADER_BLANK_ROWS + 1, col).value).strip().lower()
            is_feature = header == "feature"
            is_area = header in {"area orig", "new area"}
            for row in range(1, last_row + 1):
                cell = ws.cell(row, col)
                area_outline = is_area and row >= 3
                cell.border = Border(
                    left=edge if col == start else total_edge if area_outline else Side(),
                    right=edge if col == end else total_edge if area_outline else Side(),
                    top=edge if row == 1 else total_edge if is_area and row == 3 else Side(),
                    bottom=edge if row == last_row else total_edge if is_area and row in (4, 6) else Side(),
                )
                if is_area and row in (3, 4):
                    cell.fill = fill
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                if row > HEADER_BLANK_ROWS and not isinstance(cell, MergedCell):
                    cell.font = Font(name="Calibri", size=11, bold=row == HEADER_BLANK_ROWS + 1)
                    cell.alignment = Alignment(vertical="center", wrap_text=True)
                    if row == HEADER_BLANK_ROWS + 1:
                        cell.fill = fill
                    elif isinstance(cell.value, (int, float)):
                        cell.number_format = "General" if is_feature else "0.000"
            # Size for the displayed fixed-decimal numbers, including the total.
            # Keep text compact and wrapped, but never cap numeric widths.
            longest = 8
            numeric_width = 0
            for row in range(HEADER_BLANK_ROWS + 1, last_row + 1):
                value = ws.cell(row, col).value
                if value is not None:
                    if isinstance(value, (int, float)) and not isinstance(value, bool):
                        text = str(value) if is_feature else f"{value:.3f}"
                        numeric_width = max(numeric_width, len(text))
                    else:
                        longest = max(longest, len(str(value)))
            if is_area and ws.cell(4, col).data_type == "f":
                # This estimate is for sizing only; the saved total remains a SUM formula.
                total_value = sum(
                    ws.cell(row, col).value
                    for row in range(HEADER_BLANK_ROWS + 2, last_row + 1)
                    if isinstance(ws.cell(row, col).value, (int, float))
                    and not isinstance(ws.cell(row, col).value, bool)
                )
                numeric_width = max(numeric_width, len(f"{total_value:.3f}"))
            ws.column_dimensions[get_column_letter(col)].width = max(
                min(longest + 2, 22), numeric_width + 3
            )
        if index < len(blocks) - 1:
            for col in range(end + 1, end + BLOCK_GAP_COLUMNS + 1):
                ws.column_dimensions[get_column_letter(col)].width = 3
    ws.row_dimensions[1].height = 32
    ws.freeze_panes = "A7"
    ws.sheet_view.showGridLines = False


def write_frame(ws, frame: pd.DataFrame, start: int):
    header_row = HEADER_BLANK_ROWS + 1
    for offset, header in enumerate(frame.columns):
        put_value(ws, header_row, start + offset, str(header))
    for row, values in enumerate(frame.itertuples(index=False, name=None), start=header_row + 1):
        for offset, value in enumerate(values):
            put_value(ws, row, start + offset, value)
    # These are new summary-local formulas; all copied data remains values only.
    for offset, header in enumerate(frame.columns):
        if str(header).strip().lower() in {"area orig", "new area"}:
            col = start + offset
            letter = get_column_letter(col)
            put_value(ws, 3, col, "Total")
            total = ws.cell(4, col)
            total.value = f"=SUM({letter}{header_row + 1}:{letter}{header_row + len(frame)})" if len(frame) else "=0"
            total.font = Font(bold=True, color="20364D")
            total.number_format = "0.000"


def write_category_sheet(wb: Workbook, category: str, selected_files: Sequence[Path], logger: LiveLogger):
    ws = wb.create_sheet(title=category[:31])
    blocks = []
    current_col = 1
    for file_path in selected_files:
        blocks.append((current_col, len(REQUIRED_HEADERS)))
        put_value(ws, 1, current_col, clean_fraction_name(file_path.name))
        try:
            picked, missing = extract_columns(read_values(file_path, category, logger))
            write_frame(ws, picked, current_col)
            if missing:
                note = "Missing: " + ", ".join(missing)
                put_value(ws, 2, current_col, note)
                logger.write(f"WARNING: {file_path.name} / {category}: {note}")
            logger.write(f"{category} <- {file_path.name}: {len(picked)} rows")
        except Exception as exc:
            write_frame(ws, pd.DataFrame(columns=REQUIRED_HEADERS), current_col)
            put_value(ws, 2, current_col, str(exc))
            logger.write(f"ERROR in {category} from {file_path.name}: {exc}")
        current_col += len(REQUIRED_HEADERS) + BLOCK_GAP_COLUMNS
    format_block_sheet(ws, blocks)


def write_area_sheet(wb: Workbook, selected_files: Sequence[Path], logger: LiveLogger):
    ws = wb.create_sheet(title="Area")
    blocks = []
    current_col = 1
    for file_path in selected_files:
        put_value(ws, 1, current_col, clean_fraction_name(file_path.name))
        width = 1
        try:
            frame = read_values(file_path, "Area", logger)
            width = max(1, len(frame.columns))
            write_frame(ws, frame, current_col)
            logger.write(f"Area <- {file_path.name}: {len(frame)} rows")
        except Exception as exc:
            put_value(ws, 2, current_col, str(exc))
            logger.write(f"ERROR Area from {file_path.name}: {exc}")
        blocks.append((current_col, width))
        current_col += width + BLOCK_GAP_COLUMNS
    format_block_sheet(ws, blocks)


def sample_output_name(sample_name: str) -> str:
    name = sample_name.strip()
    if not name or re.search(r'[<>:"/\\|?*\x00-\x1f]', name) or name.endswith((".", " ")):
        raise ValueError("Enter a sample name without Windows filename characters: < > : \" / \\ | ? *")
    return f"Summary_{name}.xlsx"


def save_summary(wb: Workbook, output: Path, input_files: Sequence[Path]):
    """Prevent overwriting any source workbook, even through a linked path."""
    for source in input_files:
        if output.resolve() == source.resolve() or (
            output.exists() and source.exists() and output.samefile(source)
        ):
            raise ValueError("The summary output must not replace an input workbook. Choose a new filename.")
    wb.save(output)


def main():
    root = tk.Tk()
    root.withdraw()
    logger = LiveLogger(root)

    sample_dir = choose_sample_folder(root)
    if not sample_dir:
        logger.write("No sample folder selected. Exiting.")
        root.destroy()
        return

    logger.write(f"Sample folder: {sample_dir}")
    files = find_excel_files(sample_dir)
    if not files:
        messagebox.showerror("No Excel files", "No Excel files found in selected folder.", parent=root)
        root.destroy()
        return

    logger.write(f"Found {len(files)} Excel files")
    selected_files = choose_file_order_gui(root, files)
    if not selected_files:
        logger.write("No input files selected. Exiting.")
        root.destroy()
        return

    logger.write("Selected file order:")
    for idx, p in enumerate(selected_files, start=1):
        logger.write(f"  {idx}. {p.name}")

    categories = discover_categories(selected_files, logger)
    if not categories:
        messagebox.showerror("No category sheets", "No usable category sheets were detected.", parent=root)
        root.destroy()
        return

    chosen_categories = choose_categories_gui(root, categories)
    if not chosen_categories:
        logger.write("No categories chosen. Exiting.")
        root.destroy()
        return

    while True:
        sample_name = simpledialog.askstring("Sample name", "Enter the sample name:", parent=root)
        if sample_name is None:
            root.destroy()
            return
        try:
            out_default = sample_dir / sample_output_name(sample_name)
            break
        except ValueError as exc:
            messagebox.showwarning("Invalid sample name", str(exc), parent=root)
    out_path = filedialog.asksaveasfilename(
        parent=root,
        title="Save Summary workbook",
        initialdir=str(sample_dir),
        initialfile=out_default.name,
        defaultextension=".xlsx",
        filetypes=[("Excel Workbook", "*.xlsx")],
    )
    if not out_path:
        logger.write("No output path selected. Exiting.")
        root.destroy()
        return

    wb = Workbook()
    wb.remove(wb.active)

    for cat in chosen_categories:
        write_category_sheet(wb, cat, selected_files, logger)

    write_area_sheet(wb, selected_files, logger)

    try:
        save_summary(wb, Path(out_path), files)
    except (OSError, ValueError) as exc:
        messagebox.showerror("Summary not saved", str(exc), parent=root)
        root.destroy()
        return
    logger.write(f"Saved: {out_path}")
    messagebox.showinfo("Done", f"Summary workbook created:\n{out_path}", parent=root)

    try:
        logger.win.destroy()
    except Exception:
        pass
    root.destroy()


if __name__ == "__main__":
    main()
