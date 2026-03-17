import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox


# ----------------------------
# Helper: parse sheet selection
# ----------------------------
def parse_sheet_input(input_str: str, max_sheets: int):
    """
    Parses user input like '1,3,5-7' into a sorted list of 0-based sheet indices.
    Accepts:
      - comma lists: 1,3,6
      - ranges: 4-8
      - mixed: 1,3,5-7
    Sheet numbers entered by user are 1-based (Excel-style).
    """
    result = set()
    parts = input_str.split(",")

    for part in parts:
        part = part.strip()
        if not part:
            continue

        if "-" in part:
            start_s, end_s = part.split("-", 1)
            start = int(start_s.strip())
            end = int(end_s.strip())
            if end < start:
                start, end = end, start
            for n in range(start, end + 1):
                idx = n - 1
                if 0 <= idx < max_sheets:
                    result.add(idx)
        else:
            n = int(part)
            idx = n - 1
            if 0 <= idx < max_sheets:
                result.add(idx)

    return sorted(result)


# ----------------------------
# Main
# ----------------------------
def main():
    root = tk.Tk()
    root.withdraw()

    # Step 1: Select Excel file
    excel_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if not excel_path:
        raise Exception("No file selected.")

    # Step 2: Read sheet names
    xls = pd.ExcelFile(excel_path)
    sheet_names = xls.sheet_names
    num_sheets = len(sheet_names)

    # Step 3: Select sheets
    sheet_input = simpledialog.askstring(
        title="Select Sheets",
        prompt=(
            f"Total sheets: {num_sheets}\n\n"
            f"Enter sheet numbers to process (examples: 4-6 or 1,3,5-7):"
        )
    )
    if not sheet_input:
        raise Exception("No sheet input provided.")

    selected_indices = parse_sheet_input(sheet_input, num_sheets)
    if not selected_indices:
        raise Exception("No valid sheets selected. Check your numbers/ranges.")

    # Step 4: Max rows prompt (keep as requested)
    max_rows = simpledialog.askinteger(
        title="Max Rows",
        prompt="Enter maximum number of rows to extract per sheet (e.g., 50):",
        minvalue=1
    )
    if max_rows is None:
        raise Exception("Max row input cancelled.")

    # Output will be saved next to the input file
    out_dir = os.path.dirname(excel_path)

    # ----------------------------
    # Build Output File 1: SEM-pos-um.xlsx
    # ----------------------------
    combined_rows = []

    for idx in selected_indices:
        sheet_name = sheet_names[idx]
        class_name = sheet_name.strip()  # keep clean for Class

        try:
            df = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl")
        except Exception as e:
            print(f"❌ Could not read sheet '{sheet_name}': {e}")
            continue

        if "Stage X (mm)" not in df.columns or "Stage Y (mm)" not in df.columns:
            print(f"⚠️ Sheet '{sheet_name}' missing 'Stage X (mm)' or 'Stage Y (mm)'. Skipping.")
            continue

        sub = df[["Stage X (mm)", "Stage Y (mm)"]].copy()
        sub = sub.head(max_rows)  # limit rows

        # mm → µm : ×1000 (corrected)
        sub["Stage X"] = pd.to_numeric(sub["Stage X (mm)"], errors="coerce") * 1000
        sub["Stage Y"] = pd.to_numeric(sub["Stage Y (mm)"], errors="coerce") * 1000
        sub["Class"] = class_name

        out = sub[["Stage X", "Stage Y", "Class"]]
        combined_rows.append(out)

    if not combined_rows:
        raise Exception("No sheets were processed successfully (missing columns or read errors).")

    sem_pos_um = pd.concat(combined_rows, ignore_index=True)

    file1_path = os.path.join(out_dir, "SEM-pos-um.xlsx")
    sem_pos_um.to_excel(file1_path, index=False)

    # ----------------------------
    # Build Output File 2: Micro-After-Correction.csv
    # ----------------------------
    # Row 2 Class = "mark"
    # Row 3 Class = "mark"
    # Row 4+ Class = File1 Class (in order)
    classes = ["mark", "mark"] + sem_pos_um["Class"].tolist()

    micro_after_corr = pd.DataFrame({
        "X": [""] * len(classes),
        "Y": [""] * len(classes),
        "Class": classes
    })

    file2_path = os.path.join(out_dir, "Micro-After-Correction.csv")
    micro_after_corr.to_csv(file2_path, index=False)

    messagebox.showinfo(
        "Done",
        "Processing complete!\n\n"
        f"Saved:\n1) {file1_path}\n2) {file2_path}"
    )


if __name__ == "__main__":
    main()
