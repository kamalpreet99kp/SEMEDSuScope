"""
Add per-row quick-open hyperlinks to existing Ag liberation Excel reports.

Use case:
- Existing workbooks already contain embedded thumbnails in an "Image" column.
- Hidden "Micro X" column contains image identifiers.
- Cropped images are stored in "cropped" folders somewhere under a chosen root folder.

What this script does:
1) Lets user pick a parent folder.
2) Recursively finds Excel files (.xlsx/.xlsm) under that folder.
3) Recursively indexes all image files under folders named "cropped".
4) For each sheet that has both "Image" and "Micro X" headers, adds/updates a new column
   "Open Image" with clickable hyperlink to the best-matching cropped image.
5) Saves a NEW workbook next to original with suffix "_with_quick_links".

Notes:
- Existing embedded images are not modified.
- Original workbook is not overwritten.
- Hyperlinks keep workbook size low vs embedding additional full-size images.
"""

from __future__ import annotations

import re
from pathlib import Path
from collections import defaultdict
import tkinter as tk
from tkinter import filedialog, messagebox

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

IMG_EXTS = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp"}
TARGET_IMAGE_HEADER = "Image"
TARGET_MICROX_HEADER = "Micro X"
LINK_HEADER = "Open Image"


def normalize_text(value: str) -> str:
    s = str(value or "").strip().lower()
    s = s.replace(" ", "")
    return s


def clean_key(value: str) -> str:
    """Normalize Micro X / filename for robust matching."""
    s = str(value or "").strip().lower()
    s = re.sub(r"\.[a-z0-9]{2,5}$", "", s)  # drop extension
    s = s.replace("&", "and")
    s = re.sub(r"[^a-z0-9]+", "", s)
    return s


def choose_root_folder() -> Path:
    root = tk.Tk()
    root.withdraw()
    selected = filedialog.askdirectory(title="Select parent folder containing reports + cropped images")
    if not selected:
        raise RuntimeError("No folder selected")
    return Path(selected)


def find_workbooks(root_dir: Path):
    files = []
    for p in root_dir.rglob("*"):
        if p.is_file() and p.suffix.lower() in {".xlsx", ".xlsm"}:
            if p.name.startswith("~$"):
                continue
            if p.stem.endswith("_with_quick_links"):
                continue
            files.append(p)
    return sorted(files)


def build_cropped_image_index(root_dir: Path):
    exact = {}
    by_clean = defaultdict(list)

    for p in root_dir.rglob("*"):
        if not p.is_file() or p.suffix.lower() not in IMG_EXTS:
            continue
        if p.parent.name.lower() != "cropped":
            continue

        stem = p.stem
        exact[normalize_text(stem)] = p
        by_clean[clean_key(stem)].append(p)

    return exact, by_clean


def find_best_image(micro_x_value: str, exact, by_clean):
    if micro_x_value is None:
        return None

    raw = str(micro_x_value).strip()
    if not raw:
        return None

    # 1) exact normalized match
    key_exact = normalize_text(Path(raw).stem)
    if key_exact in exact:
        return exact[key_exact]

    # 2) cleaned exact
    key_clean = clean_key(raw)
    candidates = by_clean.get(key_clean, [])
    if candidates:
        return sorted(candidates, key=lambda p: len(str(p)))[0]

    # 3) startswith/contains fallback (handles prefixes/suffixes)
    for ck, vals in by_clean.items():
        if ck.startswith(key_clean) or key_clean.startswith(ck) or key_clean in ck:
            return sorted(vals, key=lambda p: len(str(p)))[0]

    return None


def get_header_map(ws):
    header_map = {}
    for cell in ws[1]:
        if cell.value is None:
            continue
        header_map[str(cell.value).strip()] = cell.column
    return header_map


def add_links_to_sheet(ws, exact_index, clean_index):
    headers = get_header_map(ws)
    if TARGET_IMAGE_HEADER not in headers or TARGET_MICROX_HEADER not in headers:
        return 0, 0

    image_col = headers[TARGET_IMAGE_HEADER]
    microx_col = headers[TARGET_MICROX_HEADER]

    # Reuse existing link column if present, else append after Image.
    if LINK_HEADER in headers:
        link_col = headers[LINK_HEADER]
    else:
        link_col = image_col + 1
        ws.insert_cols(link_col)
        ws.cell(row=1, column=link_col).value = LINK_HEADER

    ws.cell(row=1, column=link_col).font = Font(bold=True)
    ws.column_dimensions[ws.cell(row=1, column=link_col).column_letter].width = 18

    linked = 0
    missing = 0

    for r in range(2, ws.max_row + 1):
        micro_x = ws.cell(row=r, column=microx_col).value
        cell = ws.cell(row=r, column=link_col)

        img_path = find_best_image(micro_x, exact_index, clean_index)
        if img_path is None:
            cell.value = "Missing"
            missing += 1
            continue

        cell.value = "Open Image"
        cell.hyperlink = img_path.resolve().as_uri()
        cell.font = Font(color="0563C1", underline="single")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        linked += 1

    return linked, missing


def process_workbook(wb_path: Path, exact_index, clean_index):
    wb = load_workbook(wb_path)
    total_linked = 0
    total_missing = 0
    touched_sheets = 0

    for ws in wb.worksheets:
        linked, missing = add_links_to_sheet(ws, exact_index, clean_index)
        if linked or missing:
            touched_sheets += 1
            total_linked += linked
            total_missing += missing

    if touched_sheets == 0:
        return None, 0, 0, 0

    out = wb_path.with_name(f"{wb_path.stem}_with_quick_links{wb_path.suffix}")
    wb.save(out)
    return out, touched_sheets, total_linked, total_missing


def main():
    try:
        root_dir = choose_root_folder()
    except Exception as exc:
        print(f"Cancelled: {exc}")
        return

    print(f"Selected folder: {root_dir}")

    exact_index, clean_index = build_cropped_image_index(root_dir)
    print(f"Indexed images in cropped folders: {len(exact_index)}")

    workbooks = find_workbooks(root_dir)
    if not workbooks:
        messagebox.showwarning("No workbooks", "No .xlsx/.xlsm workbooks found.")
        return

    print(f"Found workbooks: {len(workbooks)}")

    processed = 0
    created_files = 0

    for wb_path in workbooks:
        print(f"\nProcessing: {wb_path}")
        try:
            out, sheet_count, linked, missing = process_workbook(wb_path, exact_index, clean_index)
            processed += 1
            if out is None:
                print("  No sheets with both 'Image' and 'Micro X' found. Skipped.")
                continue
            created_files += 1
            print(f"  Saved: {out}")
            print(f"  Sheets updated: {sheet_count}, links added: {linked}, missing: {missing}")
        except Exception as exc:
            print(f"  ERROR: {exc}")

    messagebox.showinfo(
        "Done",
        f"Processed workbooks: {processed}\n"
        f"Output files created: {created_files}\n"
        f"Root folder: {root_dir}",
    )


if __name__ == "__main__":
    main()
