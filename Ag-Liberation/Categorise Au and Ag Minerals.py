import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment, Font, Border, Side


# ============================================================
#  AU + AG MINERAL CLASSIFICATION
#  - Au minerals are evaluated first, then Ag minerals
#  - Uses practical EDS wt% windows based on stoichiometry + Ag script style
#  - Applies "association caps" for non-member elements to reduce confusion
#  - Writes mineral sheets only when rows exist
# ============================================================


def v(row, col):
    x = row.get(col, 0)
    return 0 if pd.isna(x) else x


TRACE = 5.0
MINOR = 12.0
AUAG_CUTOFF = 2.0

def classify_mineral(row):
    ag = v(row, "Ag (Wt%)")
    s = v(row, "S (Wt%)")
    hg = v(row, "Hg (Wt%)")
    te = v(row, "Te (Wt%)")
    se = v(row, "Se (Wt%)")
    sb = v(row, "Sb (Wt%)")
    ars = v(row, "As (Wt%)")
    fe = v(row, "Fe (Wt%)")
    cu = v(row, "Cu (Wt%)")
    bi = v(row, "Bi (Wt%)")
    au = v(row, "Au (Wt%)")
    o = v(row, "O (Wt%)")

    # ========================================================

    # Global gate for Au+Ag workflow
    if ag <= AUAG_CUTOFF and au <= AUAG_CUTOFF:
        return None

    # AU MINERALS FIRST
    # ========================================================

    # Calaverite (AuTe2): ideal ~Au 43.6, Te 56.4
    if (20 <= au <= 65 and 25 <= te <= 75 and ag < MINOR and
            s < TRACE and se < TRACE and hg < TRACE and sb < TRACE and
            ars < TRACE and cu < TRACE and fe < TRACE and bi < TRACE):
        return "Calaverite"

    # Sylvanite ((Au,Ag)Te2): ideal around Au~34, Ag~6, Te~59 (Au/Ag variable)
    if (au >= 12 and ag >= 2 and te >= 25 and
            20 <= te <= 80 and au <= 65 and ag <= 35 and
            s < TRACE and se < TRACE and hg < TRACE and sb < TRACE and
            ars < TRACE and cu < TRACE and fe < TRACE and bi < MINOR):
        return "Sylvanite"

    # Petzite (Ag3AuTe2): Ag-rich Au telluride
    if (te >= 6 and au >= 5 and s < TRACE and se < TRACE and
            10 <= ag <= 75 and 5 <= au <= 85 and 8 <= te <= 90 and
            bi < MINOR and hg < TRACE and sb < TRACE and ars < TRACE and
            cu < TRACE and fe < TRACE):
        return "Petzite"

    # Au-Ag-Hg alloy (amalgamic Au-Ag phase)
    if (au >= 15 and ag >= 5 and hg >= 5 and
            te < TRACE and s < TRACE and se < TRACE and
            sb < TRACE and ars < TRACE and bi < TRACE and
            cu < TRACE and fe < TRACE):
        return "Au-Ag-Hg"

    # Electrum-like Au-Ag alloy
    if (au >= 20 and ag >= 5 and
            te < TRACE and s < TRACE and se < TRACE and
            sb < TRACE and ars < TRACE and hg < TRACE and bi < TRACE and
            cu < TRACE and fe < TRACE):
        return "Au-Ag (Electrum)"

    # Native Au (allow small Ag as common substitution)
    if (au >= 75 and ag < 25 and o <= 20 and
            s < TRACE and te < TRACE and se < TRACE and
            sb < TRACE and ars < TRACE and hg < TRACE and bi < TRACE and
            cu < TRACE and fe < TRACE):
        return "Native Au"

    # ========================================================
    # AG MINERALS (existing script logic + ordering)
    # ========================================================

    if (se >= 6 and bi >= 15 and te < TRACE and s < TRACE and
            3 <= ag <= 60 and 15 <= bi <= 80 and 8 <= se <= 85 and
            au < TRACE and hg < TRACE and sb < TRACE and ars < TRACE and
            cu < TRACE and fe < TRACE):
        return "Bohdanowiczite"

    if (te >= 6 and bi >= 15 and se < TRACE and s < TRACE and
            3 <= ag <= 60 and 15 <= bi <= 80 and 8 <= te <= 90 and
            au < TRACE and hg < TRACE and sb < TRACE and ars < TRACE and
            cu < TRACE and fe < TRACE):
        return "Volynskite"

    if (ag > 5 and te >= 6 and
            s < TRACE and se < TRACE and au < TRACE and
            hg < TRACE and sb < TRACE and ars < TRACE and
            cu < MINOR and fe < MINOR and bi < MINOR):
        return "Hessite & Associations"

    if (ag >= 30 and sb >= 6 and
            s < TRACE and te < TRACE and se < TRACE and
            ars < TRACE and hg < TRACE and bi < TRACE and
            au < TRACE and cu < TRACE and fe < TRACE):
        return "Dyscrasite"

    if (ag >= 85 and o <= 20 and
            s < TRACE and te < TRACE and se < TRACE and
            sb < TRACE and ars < TRACE and hg < TRACE and bi < TRACE and
            au < MINOR and cu < TRACE and fe < TRACE):
        return "Native Ag"

    if ag > 10 and s > 7 and hg < 7 and te < 6 and se < 6 and bi < 40 and sb < 8 and ars < 8 and cu < 25 and fe < 26:
        return "Acanthite & Associations"

    if 5 < ag < 75 and 5 < s < 31 and 6.99 < hg < 53 and te < 6 and sb < 15 and ars < 15 and bi < 40 and cu < 25 and fe < 26 and se < 6:
        return "Imiterite & Associations"

    if 5 < ag < 75 and s > 7.0 and hg < 12 and fe < 21 and te < 6 and ars > 2 and bi < 40 and cu < 25 and se < 6:
        return "Sulfosalt (Sb & or As ) & Asso"

    if 5 < ag < 75 and s > 4 and hg < 12 and te < 6 and sb > 6 and ars < 2 and bi < 40 and cu < 25 and fe < 21 and se < 6:
        return "Sulfosalt (Sb) & Asso"

    if ag > 1.0 and s > 15 and hg < 6 and te < 6 and ars > 10 and cu > 17 and bi < 40 and se < 6:
        return "Ag Associations with Enargite"

    if ag > 1.4 and hg > 3 and fe < 20 and te < 6 and bi < 40 and cu < 20 and se < 6:
        return "Other Ag-Hg Associations"

    if ag > 2 and s > 2.5 and hg < 4 and te < 6 and fe < 28 and bi < 40 and cu < 28 and se >= 6:
        return "Aguilarite & Associations"

    if ag > 1.4 and (fe > 10 or cu > 10) and bi < 40:
        return "Ag Associations with Sulphides"

    if ag > 1.4:
        return "All Others"

    return None


tk.Tk().withdraw()
file_path = filedialog.askopenfilename(
    title="Select Full Analysis Excel File",
    filetypes=[("Excel files", "*.xlsx *.xls")]
)

if not file_path:
    print("No file selected.")
    raise SystemExit

book = pd.read_excel(file_path, sheet_name=None)
raw_sheet_name = list(book.keys())[0]
df = book[raw_sheet_name].copy()

df["Mineral Type"] = df.apply(classify_mineral, axis=1)

area_column = "Area (sq. µm)"

all_categories_in_order = [
    "Calaverite",
    "Sylvanite",
    "Petzite",
    "Au-Ag-Hg",
    "Au-Ag (Electrum)",
    "Native Au",
    "Bohdanowiczite",
    "Volynskite",
    "Hessite & Associations",
    "Dyscrasite",
    "Native Ag",
    "Acanthite & Associations",
    "Imiterite & Associations",
    "Sulfosalt (Sb & or As ) & Asso",
    "Sulfosalt (Sb) & Asso",
    "Ag Associations with Enargite",
    "Other Ag-Hg Associations",
    "Aguilarite & Associations",
    "Ag Associations with Sulphides",
    "All Others",
]

classified_sheets = {}
summary_rows = []

for cat in all_categories_in_order:
    df_cat = df[df["Mineral Type"] == cat]
    if len(df_cat) == 0:
        continue
    classified_sheets[cat] = df_cat.drop(columns=["Mineral Type"])
    summary_rows.append((cat, df_cat[area_column].sum()))

summary_df = pd.DataFrame(summary_rows, columns=["Type of Mineral", "Total Sum of Area (sq. µm)"])
total_area = summary_df["Total Sum of Area (sq. µm)"].sum()
summary_df["Percentage"] = (summary_df["Total Sum of Area (sq. µm)"] / total_area * 100) if total_area else 0

output_path = os.path.join(os.path.dirname(file_path), f"Classified_AuAg_{os.path.basename(file_path)}")

with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    pd.DataFrame().to_excel(writer, sheet_name="Area", index=False)
    summary_df.to_excel(writer, sheet_name="Summary", index=False)
    df.drop(columns=["Mineral Type"]).to_excel(writer, sheet_name="Raw Data", index=False)

    for cat in all_categories_in_order:
        if cat in classified_sheets:
            classified_sheets[cat].to_excel(writer, sheet_name=cat, index=False)

    df_auag = df[(df["Ag (Wt%)"] > AUAG_CUTOFF) | (df["Au (Wt%)"] > AUAG_CUTOFF)]
    raw_auag_count = len(df_auag)
    raw_auag_area = df_auag[area_column].sum()

    classified_total_count = sum(len(classified_sheets[k]) for k in classified_sheets.keys())
    classified_total_area = sum(classified_sheets[k][area_column].sum() for k in classified_sheets.keys())

    integrity_data = {
        "Metric": [
            f"Rows with Ag or Au > {AUAG_CUTOFF}% in Raw Data",
            "Total rows in classified sheets (non-empty only)",
            f"Area in Raw Data (Ag or Au > {AUAG_CUTOFF}%)",
            "Area in classified sheets (non-empty only)",
            "Area Match",
            "Row Count Match"
        ],
        "Value": [
            raw_auag_count,
            classified_total_count,
            round(raw_auag_area, 4),
            round(classified_total_area, 4),
            "✅ Match" if abs(raw_auag_area - classified_total_area) < 0.01 else "❌ Mismatch",
            "✅ Match" if raw_auag_count == classified_total_count else "❌ Mismatch"
        ]
    }
    pd.DataFrame(integrity_data).to_excel(writer, sheet_name="Integrity Check", index=False)

print(f"\n✅ Classification complete.\nOutput saved to:\n{output_path}")

highlight_colors = {
    "Feature": "00BFCF",
    "Area (sq. µm)": "FFEB3B",
    "S (Wt%)": "FFA07A",
    "Fe (Wt%)": "ADD8E6",
    "Cu (Wt%)": "90EE90",
    "As (Wt%)": "D87093",
    "Ag (Wt%)": "F44336",
    "Sb (Wt%)": "A9A9A9",
    "Hg (Wt%)": "9370DB",
    "Se (Wt%)": "FFA09A",
    "Te (Wt%)": "D87099",
    "Au (Wt%)": "FFF2CC",
    "Bi (Wt%)": "E2EFDA",
    "O (Wt%)": "D9E1F2",
}

thin_border = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin")
)

wb = load_workbook(output_path)

for sheet in wb.sheetnames:
    if sheet in ["Area", "Integrity Check"]:
        continue
    ws = wb[sheet]
    if ws.max_row < 2:
        continue

    headers = {cell.value: cell.column for cell in ws[1]}
    for col_name, color in highlight_colors.items():
        if col_name in headers:
            col_idx = headers[col_name]
            fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
            for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                for cell in row:
                    cell.fill = fill

if "Summary" in wb.sheetnames:
    ws = wb["Summary"]

    for col in ws.columns:
        max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + 4

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border

    for cell in ws["A"]:
        cell.font = Font(bold=True)

wb.save(output_path)
