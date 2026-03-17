import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
150
# --- Helper Functions ---
def parse_sheet_input(input_str, max_index):
    """Parses sheet input like '1,3,5-7' into list of 0-based indices."""
    result = set()
    parts = input_str.split(',')
    for part in parts:
        part = part.strip()
        if '-' in part:
            start, end = part.split('-')
            result.update(range(int(start) - 1, int(end)))  # convert to 0-based
        else:
            result.add(int(part) - 1)
    # Filter out-of-bounds indices
    return sorted([i for i in result if 0 <= i < max_index])

# --- GUI Setup ---
root = tk.Tk()
root.withdraw()

# Step 1: File selection
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

# Step 3: Ask which sheets to process
sheet_input = simpledialog.askstring(
    title="Select Sheets",
    prompt=f"Enter sheet numbers to process (e.g., 1,3,5-7). Total sheets: {num_sheets}"
)
if not sheet_input:
    raise Exception("No sheet input provided.")
50
selected_indices = parse_sheet_input(sheet_input, num_sheets)
if not selected_indices:
    raise Exception("No valid sheets selected.")

# Step 4: Ask max rows to extract
max_rows = simpledialog.askinteger(
    title="Max Rows",
    prompt="Enter maximum number of rows to extract per sheet (e.g., 50):",
    minvalue=1
)
if max_rows is None:
    raise Exception("Max row input cancelled.")

# Step 5: Create base output folder
base_name = os.path.splitext(os.path.basename(excel_path))[0]
output_dir = os.path.join(os.path.dirname(excel_path), base_name)
os.makedirs(output_dir, exist_ok=True)

# Step 6: Process each selected sheet
for idx in selected_indices:
    sheet_name = sheet_names[idx]
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name, engine='openpyxl')
    except Exception as e:
        print(f"❌ Could not read sheet '{sheet_name}': {e}")
        continue

    if "Stage X (mm)" not in df.columns or "Stage Y (mm)" not in df.columns:
        print(f"⚠️ Sheet '{sheet_name}' is missing required columns. Skipping.")
        continue

    # Extract and convert
    sub_df = df[["Stage X (mm)", "Stage Y (mm)"]].copy()
    sub_df = sub_df.head(max_rows)  # Limit rows
    sub_df["Stage X (mm)"] *= 1000
    sub_df["Stage Y (mm)"] *= 1000

    # Rename columns to reflect µm
    sub_df.rename(columns={"Stage X (mm)": "Stage X (µm)", "Stage Y (mm)": "Stage Y (µm)"}, inplace=True)

    # Create subfolder
    subfolder = os.path.join(output_dir, sheet_name)
    os.makedirs(subfolder, exist_ok=True)

    # Write list.xlsx
    list_path = os.path.join(subfolder, "list.xlsx")
    sub_df.to_excel(list_path, index=False)

    # Write pos_corr.csv with headers only
    pos_corr_path = os.path.join(subfolder, "pos_corr.csv")
    pd.DataFrame(columns=["X", "Y"]).to_csv(pos_corr_path, index=False)

    print(f"✅ Processed: {sheet_name} → {subfolder}")

# Done
messagebox.showinfo("Done", f"Processing complete.\nOutput folder:\n{output_dir}")
