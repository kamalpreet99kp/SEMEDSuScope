"""
MASTER Ag Minerals workbook builder (single XLSX with all categories as sheets)
+ embeds cropped microscope images in processed sheets
+ hides columns based on GUI selection (default: show ALL)
+ adds "Area" calculator sheet
+ adds "Check Report" sheet at end
+ robust matching across:
    - sheet names
    - folder names
    - Micro-After-Correction.csv "Class"
  including differences in spaces / underscores / hyphens / '&' / 'and' / repeated underscores

Also:
- Summary + Raw Data are copied AS-IS (exact headers preserved)
- Non-processed sheets (no matching folder/cropped) copied AS-IS
- Processed sheets:
    - Safe headers applied to first 19 columns ONLY
    - Adds Micro X column from Micro-After-Correction
    - Adds Image column, embeds matching image (filename startswith Micro X)
    - Center + middle align, compact columns, float display to 2 decimals, ints no decimals
    - Row height set so image is fully visible

Requirements:
  pip install pandas openpyxl pillow xlsxwriter numpy
"""

import re
import traceback
from pathlib import Path
from typing import List, Optional, Tuple, Dict, Any

import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

import xlsxwriter
from PIL import Image


# -----------------------------
# Settings
# -----------------------------
IMG_CELL_WIDTH_PX = 220
IMG_CELL_HEIGHT_PX = 160

MAKE_THUMBNAILS = True
THUMBNAIL_MAX_SIDE_PX = 420
THUMBNAIL_JPEG_QUALITY = 70

IMG_EXTS = (".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp")
ILLEGAL_FS_CHARS = r'\/:*?"<>|'

SAFE_HEADERS_19 = [
    "Feature",
    "Area",
    "Field",
    "Rank",
    "Area",  # duplicate name expected; we handle duplicates safely
    "Aspect Ratio",
    "Beam X (pixels)",
    "Beam Y (pixels)",
    "Breadth (um)",
    "Direction (degrees)",
    "ECD",
    "Length",
    "Perimeter",
    "Shape",
    "Mean grey",
    "Spectrum Area",
    "Stage X (mm)",
    "Stage Y (mm)",
    "Stage Z (mm)",
]

AUTOSIZE_SAMPLE_ROWS = 300
MIN_COL_WIDTH = 8
MAX_COL_WIDTH = 26


# -----------------------------
# Live log window
# -----------------------------
class LiveLogger:
    def __init__(self, title="Script Output"):
        self.root = tk.Tk()
        self.root.title(title)
        self.root.geometry("900x540")

        self.text = tk.Text(self.root, wrap="word")
        self.text.pack(fill="both", expand=True)

        self.text.insert("end", "=== Live Output ===\n")
        self.text.see("end")
        self.root.update()

    def write(self, msg: str):
        self.text.insert("end", msg + "\n")
        self.text.see("end")
        self.root.update()


# -----------------------------
# Canonical matching
# -----------------------------
def canonical_key(s: str) -> str:
    """
    Canonicalize names for matching between:
      - sheet names
      - folder names
      - Micro-After-Correction Class values

    Handles:
      - spaces / underscores / hyphens
      - '&' replaced safely
      - multiple underscores
      - punctuation
      - case differences
    """
    s = str(s or "").strip().lower()

    # Replace & directly with underscore (prevents triple ___ issue)
    s = s.replace("&", "_")

    # Replace spaces and hyphens with underscore
    s = s.replace("-", "_")
    s = re.sub(r"\s+", "_", s)

    # Replace everything not alphanumeric with underscore
    s = re.sub(r"[^a-z0-9_]", "_", s)

    # Collapse multiple underscores into one
    s = re.sub(r"_+", "_", s)

    # Remove leading/trailing underscores
    s = s.strip("_")

    return s

def normalize_spaces(s: str) -> str:
    s = str(s).strip()
    s = re.sub(r"\s+", " ", s)
    return s


def safe_filename(name: str) -> str:
    cleaned = "".join(c for c in str(name) if c not in ILLEGAL_FS_CHARS).strip()
    cleaned = cleaned.rstrip(". ").strip()
    return cleaned if cleaned else "Unnamed"


def set_sheet_name_safe(name: str) -> str:
    # Excel illegal sheet chars + 31 char limit
    s = str(name)
    s = s.replace(":", "-").replace("\\", "-").replace("/", "-").replace("?", "").replace("*", "")
    s = s.replace("[", "(").replace("]", ")")
    s = s.strip()
    if len(s) > 31:
        s = s[:31]
    return s if s else "Sheet"


# -----------------------------
# Column visibility GUI
# -----------------------------
def choose_visible_columns_listbox(headers: List[str]) -> Optional[set]:
    """
    Multi-select list: choose columns to SHOW (visible). Everything else will be hidden.
    Returns set of visible headers, or None -> show all.
    Default is: SHOW ALL columns.
    """
    win = tk.Toplevel()
    win.title("Choose columns to SHOW (others hidden)")
    win.geometry("760x620")

    tk.Label(
        win,
        text="Select columns to SHOW (Ctrl/Shift to multi-select).\n"
             "Default is ALL columns selected (visible). You can deselect to hide.",
        anchor="w",
        justify="left"
    ).pack(fill="x", padx=10, pady=(10, 5))

    frame = tk.Frame(win)
    frame.pack(fill="both", expand=True, padx=10, pady=10)

    scrollbar = tk.Scrollbar(frame)
    scrollbar.pack(side="right", fill="y")

    listbox = tk.Listbox(frame, selectmode="extended", yscrollcommand=scrollbar.set)
    for h in headers:
        listbox.insert("end", h)
    listbox.pack(side="left", fill="both", expand=True)
    scrollbar.config(command=listbox.yview)

    listbox.selection_set(0, "end")  # default show all

    result = {"value": None}

    def select_all():
        listbox.selection_set(0, "end")

    def select_none():
        listbox.selection_clear(0, "end")

    def ok():
        sel = listbox.curselection()
        result["value"] = {headers[i] for i in sel}
        win.destroy()

    def cancel():
        result["value"] = None
        win.destroy()

    btn = tk.Frame(win)
    btn.pack(fill="x", padx=10, pady=(0, 10))
    tk.Button(btn, text="Select All", command=select_all).pack(side="left")
    tk.Button(btn, text="Select None", command=select_none).pack(side="left", padx=5)
    tk.Button(btn, text="OK", command=ok).pack(side="right")
    tk.Button(btn, text="Cancel (show all)", command=cancel).pack(side="right", padx=5)

    win.grab_set()
    win.wait_window()
    return result["value"]


def apply_column_hiding(ws, headers: List[str], visible_set: Optional[set], widths: Dict[int, float]):
    if visible_set is None:
        for c in range(len(headers)):
            ws.set_column(c, c, widths.get(c, 12))
        return

    for c, h in enumerate(headers):
        if h in visible_set:
            ws.set_column(c, c, widths.get(c, 12))
        else:
            ws.set_column(c, c, 0.1, None, {"hidden": True})


# -----------------------------
# Area sheet
# -----------------------------
def write_area_calculator_sheet(workbook: xlsxwriter.Workbook):
    ws = workbook.add_worksheet("Area")

    title_fmt = workbook.add_format({"bold": True, "font_size": 14})
    note_fmt = workbook.add_format({"italic": True, "font_color": "#444444"})
    header_fmt = workbook.add_format({"bold": True})
    yellow_fmt = workbook.add_format({"bold": True, "bg_color": "#FFF200"})
    input_fmt = workbook.add_format({"bg_color": "#E7F3FF"})

    ws.set_column("A:A", 34)
    ws.set_column("B:B", 18)
    ws.set_column("C:C", 10)
    ws.set_column("D:D", 42)

    ws.merge_range("A1:D1", "📘 EDX Scanned Area Calculator (Random vs Sequence)", title_fmt)
    ws.merge_range("A2:D2", "Enter your scan parameters below. Units can be mm or µm. Final result is always in mm².", note_fmt)

    ws.write("A4", "Scan Type", header_fmt)
    ws.write("B4", "Random", input_fmt)
    ws.write("D4", "(Select: Random or Sequence)")

    ws.write("A5", "Field Width", header_fmt)
    ws.write("B5", "", input_fmt)
    ws.write("C5", "µm", input_fmt)
    ws.write("D5", "(e.g., 2.00)")

    ws.write("A6", "Field Height", header_fmt)
    ws.write("B6", "", input_fmt)
    ws.write("C6", "µm", input_fmt)
    ws.write("D6", "(e.g., 1.50)")

    ws.write("A7", "Overlap (%)", header_fmt)
    ws.write("B7", 0, input_fmt)
    ws.write("D7", "(Only applies if Sequence)")

    ws.write("A8", "Number of Fields", header_fmt)
    ws.write("B8", "", input_fmt)
    ws.write("D8", "(Enter total fields scanned)")

    ws.write("A10", "💡 Final Total Area Scanned (mm²):", yellow_fmt)

    ws.data_validation("B4", {"validate": "list", "source": ["Random", "Sequence"]})
    ws.data_validation("C5", {"validate": "list", "source": ["µm", "mm"]})
    ws.data_validation("C6", {"validate": "list", "source": ["µm", "mm"]})

    ws.write_formula(
        "B10",
        '=IF(B4="Random",'
        '(IF(C5="µm",B5/1000,B5)*IF(C6="µm",B6/1000,B6)*B8),'
        '((IF(C5="µm",B5/1000,B5)*(1-B7/100))*(IF(C6="µm",B6/1000,B6)*(1-B7/100))*B8)'
        ')',
        yellow_fmt
    )


# -----------------------------
# Images + folder helpers
# -----------------------------
def rename_cropped_images(cropped_dir: Path) -> int:
    renamed = 0
    for p in cropped_dir.iterdir():
        if not p.is_file():
            continue
        if p.suffix.lower() not in IMG_EXTS:
            continue
        if "_" not in p.stem:
            continue
        new_stem = p.stem.split("_", 1)[1]
        new_path = cropped_dir / f"{new_stem}{p.suffix}"

        if new_path.exists():
            i = 1
            while True:
                cand = cropped_dir / f"{new_stem}_{i}{p.suffix}"
                if not cand.exists():
                    new_path = cand
                    break
                i += 1

        if new_path != p:
            p.rename(new_path)
            renamed += 1
    return renamed


def list_images_in_folder(cropped_dir: Path) -> List[Path]:
    return [p for p in cropped_dir.iterdir() if p.is_file() and p.suffix.lower() in IMG_EXTS]


def build_microx_image_index(cropped_dir: Path) -> Dict[str, int]:
    """
    counts per MicroX prefix from filenames:
      67608.550_20759.470_F13279.17.jpg -> prefix=67608.550
    """
    counts: Dict[str, int] = {}
    for p in list_images_in_folder(cropped_dir):
        stem = p.stem
        prefix = stem.split("_", 1)[0].strip() if "_" in stem else stem.strip()
        if prefix:
            counts[prefix] = counts.get(prefix, 0) + 1
    return counts


def find_matching_image(cropped_dir: Path, micro_x_value: str) -> Optional[Path]:
    mx = str(micro_x_value).strip()
    if mx == "" or mx.lower() in ("nan", "none"):
        return None
    for p in cropped_dir.iterdir():
        if p.is_file() and p.suffix.lower() in IMG_EXTS and p.name.startswith(mx):
            return p
    return None


def make_thumbnail(img_path: Path, thumb_dir: Path, logger: LiveLogger) -> Optional[Path]:
    try:
        thumb_dir.mkdir(exist_ok=True)
        thumb_path = thumb_dir / (img_path.stem + ".jpg")
        if thumb_path.exists():
            return thumb_path

        with Image.open(img_path) as im:
            im = im.convert("RGB")
            w, h = im.size
            scale = THUMBNAIL_MAX_SIDE_PX / max(w, h)
            if scale < 1.0:
                im = im.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
            im.save(thumb_path, format="JPEG", quality=THUMBNAIL_JPEG_QUALITY, optimize=True)

        return thumb_path
    except Exception as e:
        logger.write(f"   ⚠️ Thumbnail failed for {img_path.name}: {e}")
        return None


# -----------------------------
# GUI headers from widest sheet
# -----------------------------
def build_gui_headers_from_widest_sheet(excel_path: Path, sheet_names: List[str]) -> List[str]:
    best_cols = []
    max_cols = -1
    for nm in sheet_names:
        try:
            tmp = pd.read_excel(excel_path, sheet_name=nm, engine="openpyxl", nrows=1)
            if tmp.shape[1] > max_cols:
                max_cols = tmp.shape[1]
                best_cols = list(tmp.columns)
        except Exception:
            continue

    if not best_cols:
        best_cols = SAFE_HEADERS_19.copy()
    else:
        best_cols = best_cols.copy()
        if len(best_cols) >= 19:
            best_cols[:19] = SAFE_HEADERS_19

    if "Micro X" not in best_cols:
        best_cols.append("Micro X")
    if "Image" not in best_cols:
        best_cols.append("Image")

    return best_cols


# -----------------------------
# Formatting helpers
# -----------------------------
def is_int_like(x: Any) -> bool:
    try:
        if x is None:
            return False
        if isinstance(x, (float, np.floating)) and np.isnan(x):
            return False
        if isinstance(x, (int,)) and not isinstance(x, bool):
            return True
        if isinstance(x, float):
            return float(x).is_integer()
        return False
    except Exception:
        return False


def is_float_like(x: Any) -> bool:
    try:
        if x is None:
            return False
        if isinstance(x, (float, np.floating)) and np.isnan(x):
            return False
        return isinstance(x, float) and not float(x).is_integer()
    except Exception:
        return False


def compute_column_widths(df: pd.DataFrame, headers: List[str]) -> Dict[int, float]:
    """
    Robust autosize even with duplicate column names and odd cell types.
    """
    widths: Dict[int, float] = {}
    sample = df.head(AUTOSIZE_SAMPLE_ROWS)

    def value_len(x) -> int:
        if x is None:
            return 0
        try:
            if isinstance(x, (float, np.floating)) and np.isnan(x):
                return 0
        except Exception:
            pass
        try:
            return len(str(x))
        except Exception:
            return 0

    for idx, col in enumerate(headers):
        if col == "Image":
            widths[idx] = 34
            continue
        if col == "Micro X":
            widths[idx] = 14
            continue

        max_len = len(str(col))

        if col in sample.columns:
            block = sample[col]
            if isinstance(block, pd.DataFrame):
                for subcol in block.columns:
                    vals = block[subcol].to_numpy(dtype=object, copy=False)
                    for v in vals:
                        max_len = max(max_len, value_len(v))
            else:
                vals = block.to_numpy(dtype=object, copy=False)
                for v in vals:
                    max_len = max(max_len, value_len(v))

        w = max(MIN_COL_WIDTH, min(MAX_COL_WIDTH, max_len + 1))
        widths[idx] = w

    return widths


# -----------------------------
# Check Report
# -----------------------------
def write_check_report_sheet(workbook: xlsxwriter.Workbook, rows: List[Dict[str, Any]]):
    ws = workbook.add_worksheet("Check Report")

    header_fmt = workbook.add_format({"bold": True, "align": "center", "valign": "vcenter"})
    center_fmt = workbook.add_format({"align": "center", "valign": "vcenter"})
    percent_fmt = workbook.add_format({"num_format": "0.0%", "align": "center", "valign": "vcenter"})

    columns = [
        "Sheet",
        "Has Cropped Folder",
        "Rows",
        "Micro X Filled",
        "Duplicate Micro X Rows",
        "Images in Cropped Folder",
        "Rows with 0 Image Match",
        "Rows with >1 Image Match",
        "Images Inserted",
        "Images Missing",
        "Coverage",
    ]

    ws.set_column(0, 0, 30)
    ws.set_column(1, 1, 18)
    ws.set_column(2, 10, 22)

    for c, col in enumerate(columns):
        ws.write(0, c, col, header_fmt)

    for r, item in enumerate(rows, start=1):
        for c, col in enumerate(columns):
            val = item.get(col, "")
            if col == "Coverage" and isinstance(val, (float, int)):
                ws.write_number(r, c, float(val), percent_fmt)
            else:
                ws.write(r, c, val, center_fmt)


# -----------------------------
# Main
# -----------------------------
def main():
    logger = LiveLogger("MASTER Builder - Live Output")

    picker_root = tk.Tk()
    picker_root.withdraw()

    excel_file = filedialog.askopenfilename(
        title="Select MAIN Ag-minerals Excel file",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if not excel_file:
        logger.write("No file selected. Exiting.")
        return

    excel_path = Path(excel_file)
    base_dir = excel_path.parent
    micro_csv = base_dir / "Micro-After-Correction.csv"

    if not micro_csv.exists():
        messagebox.showerror("Missing file", f"Could not find Micro-After-Correction.csv in:\n{base_dir}")
        logger.write(f"ERROR: Missing Micro-After-Correction.csv in {base_dir}")
        return

    xls = pd.ExcelFile(excel_path)
    sheet_names = xls.sheet_names

    # Build folder map once using canonical keys
    folder_map: Dict[str, Path] = {}
    for p in base_dir.iterdir():
        if p.is_dir():
            folder_map[canonical_key(p.name)] = p

    gui_headers = build_gui_headers_from_widest_sheet(excel_path, sheet_names)
    visible_set = choose_visible_columns_listbox(gui_headers)

    out_master_path = base_dir / f"MASTER_{excel_path.stem}.xlsx"
    log_path = base_dir / "master_debug_log.txt"

    audit_rows: List[Dict[str, Any]] = []

    with open(log_path, "w", encoding="utf-8") as log_fp:

        def log(msg: str):
            logger.write(msg)
            print(msg, flush=True)
            log_fp.write(msg + "\n")
            log_fp.flush()

        log("==============================")
        log("MASTER BUILD STARTED")
        log(f"Input Excel: {excel_path}")
        log(f"Micro CSV  : {micro_csv}")
        log(f"Output XLSX: {out_master_path}")
        log("==============================\n")

        micro_df = pd.read_csv(micro_csv)
        if not {"X", "Y", "Class"}.issubset(set(micro_df.columns)):
            msg = f"Micro-After-Correction.csv must contain X, Y, Class. Found: {list(micro_df.columns)}"
            log(f"ERROR: {msg}")
            messagebox.showerror("Bad CSV", msg)
            return

        # Build Micro X lists by canonical class key
        micro_by_class: Dict[str, List] = {}
        for _, row in micro_df.iterrows():
            cls_key = canonical_key(row["Class"])
            micro_by_class.setdefault(cls_key, []).append(row["X"])

        workbook = xlsxwriter.Workbook(str(out_master_path), {"constant_memory": True})

        header_fmt = workbook.add_format({"bold": True, "align": "center", "valign": "vcenter"})
        text_center_fmt = workbook.add_format({"align": "center", "valign": "vcenter"})
        int_center_fmt = workbook.add_format({"num_format": "0", "align": "center", "valign": "vcenter"})
        float2_center_fmt = workbook.add_format({"num_format": "0.00", "align": "center", "valign": "vcenter"})

        write_area_calculator_sheet(workbook)

        used_sheetnames = {"Area"}
        total_inserted = 0
        total_missing = 0

        for original_sheet_name in sheet_names:
            if canonical_key(original_sheet_name) == "area":
                log("(Skipping original 'Area' sheet; calculator already created.)")
                continue

            log("----------------------------------")
            log(f"➡️ Sheet: {original_sheet_name}")

            safe_sheet = set_sheet_name_safe(original_sheet_name)
            base_name = safe_sheet
            i = 1
            while safe_sheet in used_sheetnames:
                suffix = f"_{i}"
                safe_sheet = (base_name[: max(0, 31 - len(suffix))] + suffix)
                i += 1
            used_sheetnames.add(safe_sheet)

            ws = workbook.add_worksheet(safe_sheet)

            try:
                df = pd.read_excel(excel_path, sheet_name=original_sheet_name, engine="openpyxl")
                log(f"   ✅ Read rows={len(df)} cols={df.shape[1]}")
            except Exception:
                log("   ❌ Failed reading sheet:")
                log(traceback.format_exc())
                ws.write(0, 0, f"Failed to read sheet: {original_sheet_name}", header_fmt)
                continue

            # Summary + Raw Data copied AS-IS
            if canonical_key(original_sheet_name) in ("summary", "raw_data"):
                log("   ✅ Copying AS-IS (kept exactly like input).")
                headers = list(df.columns)
                for c, col in enumerate(headers):
                    ws.write(0, c, col, header_fmt)
                for r in range(len(df)):
                    for c in range(len(headers)):
                        v = df.iloc[r, c]
                        ws.write(r + 1, c, "" if pd.isna(v) else v, text_center_fmt)
                for c in range(len(headers)):
                    ws.set_column(c, c, 15)

                audit_rows.append({
                    "Sheet": original_sheet_name,
                    "Has Cropped Folder": "No",
                    "Rows": len(df),
                    "Micro X Filled": "",
                    "Duplicate Micro X Rows": "",
                    "Images in Cropped Folder": "",
                    "Rows with 0 Image Match": "",
                    "Rows with >1 Image Match": "",
                    "Images Inserted": "",
                    "Images Missing": "",
                    "Coverage": "",
                })
                continue

            # Determine matching folder by canonical key
            sheet_key = canonical_key(original_sheet_name)
            cat_dir = folder_map.get(sheet_key)

            cropped_dir = None
            if cat_dir is not None:
                possible = cat_dir / "cropped"
                if possible.exists() and possible.is_dir():
                    cropped_dir = possible
                    log(f"   ✅ Folder match (canonical): {cat_dir.name}")
                else:
                    log(f"   ⚠️ Folder found but missing 'cropped': {possible}")
            else:
                log(f"   ⚠️ No matching folder for canonical key: {sheet_key}")

            # No cropped -> copy AS-IS
            if cropped_dir is None:
                log("   ✅ No cropped folder -> copying AS-IS.")
                headers = list(df.columns)
                for c, col in enumerate(headers):
                    ws.write(0, c, col, header_fmt)
                for r in range(len(df)):
                    for c in range(len(headers)):
                        v = df.iloc[r, c]
                        ws.write(r + 1, c, "" if pd.isna(v) else v, text_center_fmt)
                for c in range(len(headers)):
                    ws.set_column(c, c, 15)

                audit_rows.append({
                    "Sheet": original_sheet_name,
                    "Has Cropped Folder": "No",
                    "Rows": len(df),
                    "Micro X Filled": "",
                    "Duplicate Micro X Rows": "",
                    "Images in Cropped Folder": "",
                    "Rows with 0 Image Match": "",
                    "Rows with >1 Image Match": "",
                    "Images Inserted": "",
                    "Images Missing": "",
                    "Coverage": "",
                })
                continue

            # Processed sheet
            log("   ✅ Processed -> safe headers + Micro X + Image + embed + audit.")
            if df.shape[1] >= 19:
                cols = list(df.columns)
                cols[:19] = SAFE_HEADERS_19
                df.columns = cols

            # rename images
            renamed = rename_cropped_images(cropped_dir)
            log(f"   ✅ Renamed images in cropped: {renamed}")

            # Build cropped image index
            microx_img_counts = build_microx_image_index(cropped_dir)
            total_imgs_in_folder = sum(microx_img_counts.values())

            # Add Micro X based on canonical class key
            micro_x_vals = micro_by_class.get(sheet_key, [])
            n = len(df)
            if len(micro_x_vals) < n:
                micro_x_vals = micro_x_vals + [None] * (n - len(micro_x_vals))
            else:
                micro_x_vals = micro_x_vals[:n]

            df["Micro X"] = micro_x_vals
            df.insert(df.columns.get_loc("Micro X") + 1, "Image", "")

            headers = list(df.columns)
            micro_idx = headers.index("Micro X")
            img_idx = headers.index("Image")

            for c, col in enumerate(headers):
                ws.write(0, c, col, header_fmt)

            widths = compute_column_widths(df, headers)
            apply_column_hiding(ws, headers, visible_set, widths)

            thumb_dir = cropped_dir / "_thumbnails"

            inserted = 0
            missing = 0

            # ---- Audit row-based checks ----
            micro_series = df["Micro X"].astype(str)
            micro_clean = micro_series.map(lambda x: x.strip()).replace({"nan": "", "None": ""})
            micro_filled = int((micro_clean != "").sum())
            dup_rows = int(micro_clean[micro_clean != ""].duplicated(keep=False).sum())

            zero_match_rows = 0
            multi_match_rows = 0

            for r in range(len(df)):
                excel_row = r + 1
                ws.set_row(excel_row, IMG_CELL_HEIGHT_PX)

                # Write values (centered) with numeric formatting
                for c, col in enumerate(headers):
                    if col == "Image":
                        continue
                    v = df.iloc[r, c]
                    if pd.isna(v):
                        ws.write(excel_row, c, "", text_center_fmt)
                    else:
                        if is_int_like(v):
                            ws.write_number(excel_row, c, float(v), int_center_fmt)
                        elif is_float_like(v):
                            ws.write_number(excel_row, c, float(v), float2_center_fmt)
                        else:
                            ws.write(excel_row, c, v, text_center_fmt)

                mx = str(df.iloc[r, micro_idx]).strip()
                if mx.lower() in ("nan", "none"):
                    mx = ""

                if mx:
                    cnt = microx_img_counts.get(mx, 0)
                    if cnt == 0:
                        zero_match_rows += 1
                    elif cnt > 1:
                        multi_match_rows += 1

                img_path = find_matching_image(cropped_dir, mx)
                if img_path is None:
                    missing += 1
                    continue

                embed_path = img_path
                if MAKE_THUMBNAILS:
                    thumb = make_thumbnail(img_path, thumb_dir, logger)
                    if thumb:
                        embed_path = thumb

                try:
                    with Image.open(embed_path) as im:
                        w, h = im.size
                    x_scale = IMG_CELL_WIDTH_PX / w
                    y_scale = IMG_CELL_HEIGHT_PX / h
                    ws.insert_image(
                        excel_row, img_idx, str(embed_path),
                        {"x_scale": x_scale, "y_scale": y_scale, "x_offset": 2, "y_offset": 2}
                    )
                    inserted += 1
                except Exception as e:
                    log(f"   ⚠️ Insert image failed (Micro X={mx}): {e}")
                    missing += 1

            coverage = (inserted / micro_filled) if micro_filled else 0.0

            audit_rows.append({
                "Sheet": original_sheet_name,
                "Has Cropped Folder": "Yes",
                "Rows": len(df),
                "Micro X Filled": micro_filled,
                "Duplicate Micro X Rows": dup_rows,
                "Images in Cropped Folder": int(total_imgs_in_folder),
                "Rows with 0 Image Match": int(zero_match_rows),
                "Rows with >1 Image Match": int(multi_match_rows),
                "Images Inserted": int(inserted),
                "Images Missing": int(missing),
                "Coverage": float(coverage),
            })

            total_inserted += inserted
            total_missing += missing
            log(f"   📷 Images inserted: {inserted}, missing: {missing}")

        # TOTAL row
        def sum_int(key: str) -> int:
            s = 0
            for r in audit_rows:
                v = r.get(key, "")
                if isinstance(v, (int, np.integer)):
                    s += int(v)
            return s

        audit_rows.append({
            "Sheet": "TOTAL",
            "Has Cropped Folder": "",
            "Rows": sum_int("Rows"),
            "Micro X Filled": sum_int("Micro X Filled"),
            "Duplicate Micro X Rows": sum_int("Duplicate Micro X Rows"),
            "Images in Cropped Folder": sum_int("Images in Cropped Folder"),
            "Rows with 0 Image Match": sum_int("Rows with 0 Image Match"),
            "Rows with >1 Image Match": sum_int("Rows with >1 Image Match"),
            "Images Inserted": sum_int("Images Inserted"),
            "Images Missing": sum_int("Images Missing"),
            "Coverage": "",
        })

        # Check Report LAST
        write_check_report_sheet(workbook, audit_rows)

        workbook.close()

        log("\n==============================")
        log("MASTER BUILD FINISHED")
        log(f"Total images inserted: {total_inserted}")
        log(f"Total images missing : {total_missing}")
        log(f"Output: {out_master_path}")
        log(f"Log   : {log_path}")
        log("==============================")

    messagebox.showinfo(
        "Done",
        f"Master workbook created!\n\n{out_master_path}\n\n"
        f"A 'Check Report' sheet was added at the end.\n\n"
        f"Debug log:\n{log_path}"
    )


if __name__ == "__main__":
    main()