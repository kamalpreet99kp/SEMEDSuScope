import os
import re
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


AUAG_CUTOFF = 2.0
TELLURIDE_LIKE_TE = 3.0
AREA_COLUMN = "Area (sq. µm)"


def v(row, column):
    value = row.get(column, 0)
    return 0 if pd.isna(value) else value


def excel_safe_sheet_name(name):
    """Return a valid Excel worksheet name."""
    cleaned = re.sub(r'[\\/*?:\[\]]', "-", str(name)).strip()
    return cleaned[:31] if cleaned else "Sheet"


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
    pb = v(row, "Pb (Wt%)")
    o = v(row, "O (Wt%)")
    w = v(row, "W (Wt%)")

    # Route W-rich false positives first. All other categories effectively use W < 12.
    if w >= 12:
        return "W-Rich Features"

    # ---------- HIGH-SPECIFICITY TELLURIDE/AG-AU RULES FIRST ----------
    if (30 <= au <= 100 and ag < 1 and te < 4 and bi < 5 and pb < 5 and hg < 6 and sb < 6 and
            s < 12 and se < 12 and ars < 12 and cu < 12 and fe < 12):
        return "Native Au"

    if (ag >= 85 and o <= 40 and au < 3 and s < 10 and se < 10 and ars < 10 and cu < 10 and fe < 10 and
            te < 6 and sb < 6 and hg < 6 and bi < 6 and pb < 6):
        return "Native Ag"

    if (5 <= au <= 98 and 2 <= ag <= 98 and te < 4 and bi < 7 and pb < 7 and hg < 6 and sb < 7 and
            s < 12 and se < 12 and ars < 12 and cu < 12 and fe < 12):
        return "Au+Ag (Electrum)"

    if (35 <= au <= 70 and 18 <= bi <= 70 and ag < 6 and te < 4 and pb < 6 and
            s < 10 and se < 10 and hg < 10 and sb < 10 and ars < 10 and cu < 10 and fe < 10):
        return "Maldonite"

    if (51 <= te <= 68 and 18 <= au <= 40 and 3.5 <= ag <= 15.5 and bi < 10 and pb < 6 and hg < 6 and
            sb < 7 and s < 10 and se < 10 and ars < 10 and cu < 10 and fe < 10):
        return "Sylvanite/Krennerite"

    if (35 <= au <= 60 and 38 <= te <= 68 and ag < 4 and bi < 10 and pb < 6 and hg < 6 and sb < 7 and
            s < 10 and se < 10 and ars < 10 and cu < 10 and fe < 10):
        return "Calaverite"

    if (26 <= ag <= 52 and 22 <= te <= 42 and 12 <= au <= 34 and bi < 10 and pb < 6 and hg < 6 and
            sb < 7 and s < 10 and se < 10 and ars < 10 and cu < 10 and fe < 10):
        return "Petzite"

    if (14 <= au <= 34 and 45 <= te <= 74 and 2.0 <= cu <= 16 and ag < 6 and hg < 6 and sb < 6 and
            bi < 6 and pb < 6 and s < 10 and se < 10 and ars < 10 and fe < 10):
        return "Kostovite"

    if (22 <= au <= 40 and 15.6 <= ag <= 24 and 34 <= te <= 50 and bi < 10 and pb < 6 and hg < 6 and
            sb < 7 and s < 10 and se < 10 and ars < 10 and cu < 10 and fe < 10):
        return "Muthmannite"

    if (4 <= au <= 10 and 10 <= te <= 18 and 48 <= pb <= 64 and 3 <= sb <= 8 and 5 <= s <= 14 and
            ag < 5 and bi < 6 and hg < 5 and se < 10 and ars < 10 and cu < 10 and fe < 10):
        return "Nagyagite"

    if (30 <= ag <= 75 and 12 <= te <= 50 and au < 7 and hg < 6 and sb < 6 and bi < 6 and pb < 6 and
            se < 10 and ars < 10 and s < 10 and fe < 10 and cu < 10):
        return "Hessite"

    if (60 <= ag <= 76 and 20 <= sb <= 30 and s < 5 and bi < 10 and pb < 6 and hg < 6 and
            se < 10 and ars < 10 and cu < 10 and fe < 10 and te < 6 and au < 5):
        return "Dyscrasite"

    if (50 <= ag <= 70 and 12 <= sb <= 25 and 6 <= s <= 20 and bi < 10 and pb < 6 and hg < 6 and
            se < 10 and ars < 10 and cu < 10 and fe < 10 and te < 6 and au < 5):
        return "Pyrostilpnite/Stephanite"

    if (ag > 75 and s > 7 and hg < 7 and te < 5 and se < 5 and bi < 12 and sb < 8 and ars < 8 and
            cu < 12 and fe < 12 and au < 4):
        return "Acanthite"

    if (25 <= ag <= 49 and 8 <= s <= 19 and 6 <= hg <= 53 and te < 6 and sb < 6 and ars < 15 and
            bi < 12 and cu < 12 and fe < 15 and se < 12):
        return "Imiterite"

    if (40 <= pb <= 70 and 20 <= te <= 45 and au < 7 and ag < 8 and bi < 12 and
            s < 12 and se < 12 and hg < 12 and sb < 12 and ars < 12 and cu < 12 and fe < 12):
        return "Altaite"

    if (35 <= hg <= 80 and 20 <= te <= 45 and au < 6 and ag < 8 and bi < 12 and pb < 12 and
            s < 12 and se < 12 and sb < 12 and ars < 12 and cu < 12 and fe < 12):
        return "Coloradoite"

    if (25 <= te <= 60 and 25 <= bi <= 65 and au < 6 and ag < 6 and pb < 6 and
            s < 12 and se < 12 and hg < 12 and sb < 12 and ars < 12 and cu < 12 and fe < 12):
        return "Tellurobismuthite"

    if (11.5 <= fe <= 18.5 and 70 <= te <= 83.5 and au < 5 and ag < 5 and hg < 5 and pb < 5 and
            bi < 5 and sb < 5 and s < 10 and se < 10 and ars < 10 and cu < 10):
        return "Frohbergite"

    if (te <= 3 and hg > 4 and au < 5 and ag < 5 and pb < 5 and bi < 5 and sb < 5 and
            s < 10 and se < 10 and ars < 10 and cu < 10 and fe < 10):
        return "Hg Minerals -No Te"

    # ---------- BROADER ASSOCIATION/FALLBACK BUCKETS ----------
    if au > AUAG_CUTOFF and te < 2 and ag < 2 and (s > 10 or fe > 10 or cu > 10):
        return "Au Asso w Sulp (No Te)"

    if ag > AUAG_CUTOFF and te < 2 and au < 2 and (s > 10 or fe > 10 or cu > 10):
        return "Ag Asso w Sulp (No Te)"

    if au > AUAG_CUTOFF and te > TELLURIDE_LIKE_TE:
        return "All Other (Au-Te)"

    if ag > AUAG_CUTOFF and te > TELLURIDE_LIKE_TE:
        return "All Others (Ag-Te)"

    if pb > 5 and te > TELLURIDE_LIKE_TE:
        return "All Others (Pb-Te)"

    if hg > 5 and te > TELLURIDE_LIKE_TE:
        return "All Others (Hg-Te)"

    if bi > 5 and te > TELLURIDE_LIKE_TE:
        return "All Others (Bi-Te)"

    if te > TELLURIDE_LIKE_TE:
        return "Unidentified"

    if au > AUAG_CUTOFF or ag > AUAG_CUTOFF:
        return "Others (Unidentified)"

    return None


def choose_input_file(parent):
    selected = filedialog.askopenfilename(
        title="Select Full Analysis Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")],
        parent=parent,
    )
    if not selected:
        raise SystemExit("No file selected.")
    return selected


def choose_area_cutoff(parent):
    """Ask for the minimum feature area; zero disables size filtering."""
    cutoff = simpledialog.askfloat(
        "Area Size Cutoff",
        "Enter Area (sq. µm) cutoff for this sample.\n"
        "Features smaller than this value will go to the After Size CutOff sheet.\n"
        "Enter 0 to disable size filtering.",
        parent=parent,
        initialvalue=0.0,
        minvalue=0.0,
    )
    if cutoff is None:
        raise SystemExit("No area cutoff entered.")
    return float(cutoff)


def validate_and_prepare_columns(df):
    required_columns = [AREA_COLUMN, "Te (Wt%)", "Ag (Wt%)", "Au (Wt%)"]
    missing = [column for column in required_columns if column not in df.columns]
    if missing:
        raise ValueError("Missing required column(s): " + ", ".join(missing))

    numeric_columns = [column for column in df.columns if column.endswith("(Wt%)")]
    numeric_columns.append(AREA_COLUMN)
    for column in numeric_columns:
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0)


ALL_CATEGORIES_IN_ORDER = [
    "Calaverite",
    "Sylvanite/Krennerite",
    "Petzite",
    "Kostovite",
    "Muthmannite",
    "Nagyagite",
    "Hessite",
    "Altaite",
    "Coloradoite",
    "Tellurobismuthite",
    "Frohbergite",
    "Hg Minerals -No Te",
    "All Other (Au-Te)",
    "All Others (Ag-Te)",
    "All Others (Pb-Te)",
    "All Others (Hg-Te)",
    "All Others (Bi-Te)",
    "Unidentified",
    "Others (Unidentified)",
    "Native Au",
    "Native Ag",
    "Au+Ag (Electrum)",
    "Maldonite",
    "Dyscrasite",
    "Pyrostilpnite/Stephanite",
    "Acanthite",
    "Imiterite",
    "Au Asso w Sulp (No Te)",
    "Ag Asso w Sulp (No Te)",
    "W-Rich Features",
]

# Every category that can produce a normal worksheet is eligible for recovery
# below the area cutoff. Ordinary raw-data grains that classify as None (for
# example, background pyrite or chalcopyrite) remain excluded.
SIZE_CUTOFF_ELIGIBLE_CATEGORIES = set(ALL_CATEGORIES_IN_ORDER)


def main():
    root = tk.Tk()
    root.withdraw()

    try:
        file_path = sys.argv[1] if len(sys.argv) > 1 else choose_input_file(root)
        area_cutoff = float(sys.argv[2]) if len(sys.argv) > 2 else choose_area_cutoff(root)
        if area_cutoff < 0:
            raise ValueError("Area cutoff cannot be negative.")

        book = pd.read_excel(file_path, sheet_name=None)
        if not book:
            raise ValueError("The selected workbook has no worksheets.")

        raw_sheet_name = list(book.keys())[0]
        df = book[raw_sheet_name].copy()
        validate_and_prepare_columns(df)
        df["Mineral Type"] = df.apply(classify_mineral, axis=1)

        meets_size_cutoff = df[AREA_COLUMN] >= area_cutoff
        classified_sheets = {}
        summary_rows = []

        for category in ALL_CATEGORIES_IN_ORDER:
            category_rows = df[meets_size_cutoff & (df["Mineral Type"] == category)]
            if category_rows.empty:
                continue
            classified_sheets[category] = category_rows.drop(columns=["Mineral Type"])
            if category != "W-Rich Features":
                summary_rows.append((category, category_rows[AREA_COLUMN].sum()))

        after_size_cutoff = df[
            (~meets_size_cutoff)
            & df["Mineral Type"].isin(SIZE_CUTOFF_ELIGIBLE_CATEGORIES)
        ].copy()
        if not after_size_cutoff.empty:
            after_size_cutoff.insert(
                len(after_size_cutoff.columns),
                "Size CutOff Reason",
                f"Area < {area_cutoff:g} sq. µm",
            )

        summary_df = pd.DataFrame(
            summary_rows,
            columns=["Type of Mineral", "Total Sum of Area (sq. µm)"],
        )
        total_area = summary_df["Total Sum of Area (sq. µm)"].sum()
        summary_df["Percentage"] = (
            summary_df["Total Sum of Area (sq. µm)"] / total_area * 100
            if total_area else 0
        )

        output_path = os.path.join(
            os.path.dirname(file_path),
            f"Classified_Tellurides_{os.path.basename(file_path)}",
        )

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            pd.DataFrame().to_excel(writer, sheet_name="Area", index=False)
            summary_df.to_excel(writer, sheet_name="Summary", index=False)
            df.drop(columns=["Mineral Type"]).to_excel(writer, sheet_name="Raw Data", index=False)

            for category in ALL_CATEGORIES_IN_ORDER:
                if category in classified_sheets:
                    classified_sheets[category].to_excel(
                        writer,
                        sheet_name=excel_safe_sheet_name(category),
                        index=False,
                    )

            if not after_size_cutoff.empty:
                after_size_cutoff.to_excel(
                    writer,
                    sheet_name="After Size CutOff",
                    index=False,
                )

            integrity_focus = (
                (df["Te (Wt%)"] > TELLURIDE_LIKE_TE)
                & meets_size_cutoff
            )
            df_focus = df[integrity_focus]
            classified_rows = df_focus[df_focus["Mineral Type"].isin(ALL_CATEGORIES_IN_ORDER)]

            raw_count = len(df_focus)
            raw_area = df_focus[AREA_COLUMN].sum()
            classified_count = len(classified_rows)
            classified_area = classified_rows[AREA_COLUMN].sum()

            integrity_data = {
                "Metric": [
                    f"Rows with Te > {TELLURIDE_LIKE_TE:g} and Area >= {area_cutoff:g} sq. µm",
                    "Total rows represented in classified sheets",
                    "Area in focused Raw Data",
                    "Area represented in classified sheets",
                    "Area Match",
                    "Row Count Match",
                    "Entered area cutoff (sq. µm)",
                    "Target rows moved to After Size CutOff",
                    "Target area moved to After Size CutOff",
                ],
                "Value": [
                    raw_count,
                    classified_count,
                    round(raw_area, 4),
                    round(classified_area, 4),
                    "✅ Match" if abs(raw_area - classified_area) < 0.01 else "❌ Mismatch",
                    "✅ Match" if raw_count == classified_count else "❌ Mismatch",
                    area_cutoff,
                    len(after_size_cutoff),
                    round(after_size_cutoff[AREA_COLUMN].sum(), 4),
                ],
            }
            pd.DataFrame(integrity_data).to_excel(
                writer,
                sheet_name="Integrity Check",
                index=False,
            )

        format_output_workbook(output_path)
        messagebox.showinfo(
            "Classification complete",
            f"Output saved to:\n{output_path}",
            parent=root,
        )
        print(f"\n✅ Classification complete.\nOutput saved to:\n{output_path}")
    except (ValueError, TypeError, FileNotFoundError, PermissionError) as error:
        messagebox.showerror("Classification error", str(error), parent=root)
        raise
    finally:
        root.destroy()


def format_output_workbook(output_path):
    highlight_colors = {
        "Feature": "00BFCF",
        AREA_COLUMN: "FFEB3B",
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
        "Pb (Wt%)": "F4B183",
        "O (Wt%)": "D9E1F2",
    }
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    workbook = load_workbook(output_path)
    for sheet_name in workbook.sheetnames:
        if sheet_name in ["Area", "Integrity Check"]:
            continue
        worksheet = workbook[sheet_name]
        if worksheet.max_row < 2:
            continue
        headers = {cell.value: cell.column for cell in worksheet[1]}
        for column_name, color in highlight_colors.items():
            if column_name not in headers:
                continue
            column_index = headers[column_name]
            fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
            for row in worksheet.iter_rows(
                min_row=2,
                min_col=column_index,
                max_col=column_index,
            ):
                row[0].fill = fill

    if "Summary" in workbook.sheetnames:
        worksheet = workbook["Summary"]
        for column in worksheet.columns:
            max_length = max(len(str(cell.value)) if cell.value else 0 for cell in column)
            worksheet.column_dimensions[column[0].column_letter].width = max_length + 4
        for row in worksheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border
        for cell in worksheet["A"]:
            cell.font = Font(bold=True)

    workbook.save(output_path)


if __name__ == "__main__":
    main()
