"""Standalone Sb classification for AZtec Full Analysis workbooks.

Primary wt.% ranges were supplied by the user. Additional, explicitly listed
caps and ratio tests separate competing minerals. Association labels are EDS
screening assignments; ambiguous matches are retained for review in Raw Data.
Workbook/area-cutoff behavior follows Updated-AllAgMinerals-41-19.py (419 lines).
"""

from dataclasses import dataclass, field
from math import isfinite
from pathlib import Path
from typing import Optional
import re
import tkinter as tk
from tkinter import filedialog, simpledialog

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


AREA_COLUMN = "Area (sq. µm)"
SB_CUTOFF = 2.0
ELEMENTS = ("Sb", "Au", "Ag", "S", "Fe", "Cu", "Pb", "Hg", "Zn", "As",
            "Bi", "Te", "Se", "O", "Si", "Al", "Ca", "Mg", "W", "La")
AG_TETRAHEDRITE = "Ag/Ag-Hg bearing Tetrahedrite (Fe)"
SIZE_CATEGORY = "After Size-CutOff"
# Workbook presentation only: classification uses FREE_RULES/ASSOCIATION_RULES below.
FREE_ORDER = (
    "Aurostibite", AG_TETRAHEDRITE, "Andorite", "Cupropolybasite", "Pyrargyrite",
    "Gudmundite", "Stibnite", "Tetrahedrite-(Fe)", "Zn-bearing Tetrahedrite",
    "Boulangerite", "Bournonite", "Cinnabar",
)
ASSOCIATION_ORDER = (
    "Aurostibite-Asso", "Andorite-Asso", "Cupropolybasite-Asso", "Pyrargyrite-Asso",
    "Gudmundite-Asso", "Stibnite-Asso", "Tetrahedrite-(Fe)-Asso",
    "Boulangerite-Asso", "Bournonite-Asso", "Cinnabar-Asso",
)
CATEGORY_ORDER = FREE_ORDER + ASSOCIATION_ORDER + (
    "Au-Sb Leftover", "Ag-Sb Leftover", "Sb Leftover", SIZE_CATEGORY,
)
ATOMIC_WEIGHTS = {"Au": 196.96657, "Ag": 107.8682, "Cu": 63.546, "Sb": 121.760, "As": 74.9216,
                  "Pb": 207.2}


@dataclass(frozen=True)
class MineralRule:
    name: str
    ranges: dict
    caps: dict = field(default_factory=dict)  # Strict upper limits: value < cap.
    checks: tuple = ()


def v(row, element):
    """Match the reference's zero handling for missing/blank optional elements."""
    value = row.get(f"{element} (Wt%)", 0)
    return 0.0 if pd.isna(value) else float(value)


def atomic_ratio(values, numerator, denominator):
    top = sum(values[e] / ATOMIC_WEIGHTS[e] for e in numerator)
    bottom = sum(values[e] / ATOMIC_WEIGHTS[e] for e in denominator)
    return top / bottom if bottom > 0 else float("inf")


def additional_checks(values, checks):
    for check in checks:
        if check == "Sb dominant over As":
            if values["Sb"] / ATOMIC_WEIGHTS["Sb"] < values["As"] / ATOMIC_WEIGHTS["As"]:
                return False
        elif check == "Cupropolybasite ratio":
            if atomic_ratio(values, ("Ag", "Cu"), ("Sb", "As")) <= 4:
                return False
        elif check == "Tetrahedrite ratio":
            if not 2 <= atomic_ratio(values, ("Ag", "Cu"), ("Sb", "As")) <= 4:
                return False
        elif check == "Boulangerite ratio":
            if not 0.9 <= atomic_ratio(values, ("Pb",), ("Sb", "As")) <= 1.8:
                return False
        elif check == "Pyrargyrite ratio":
            if not 2 <= atomic_ratio(values, ("Ag",), ("Sb", "As")) <= 4:
                return False
        elif check == "Pyrargyrite association ratio":
            if not 1.5 <= atomic_ratio(values, ("Ag",), ("Sb", "As")) <= 4.2:
                return False
        elif check == "Aurostibite association ratio":
            if not 0.1 <= atomic_ratio(values, ("Au",), ("Sb",)) <= 1:
                return False
        else:
            raise ValueError(f"Unknown classification check: {check}")
    return True


# Non-constituent Au/Ag must be <1.5 wt.% in BOTH free and association rules.
# Other previous caps below 8 are relaxed to <8 except diagnostic boundaries:
# Cu in boulangerite (free <4, Asso <3), Zn in ordinary tetrahedrite (<2),
# Fe in stibnite-Asso (<7), and Cu in pyrargyrite-Asso (<5).
# Existing wider matrix/substitution/association allowances are preserved.

FREE_RULES = (
    MineralRule('Gudmundite',
                {'Sb': (44, 65), 'Fe': (17, 32), 'S': (5, 20)},
                {'O': 15, 'Si': 8, 'Al': 8, 'Ca': 8, 'Mg': 8, 'Cu': 8, 'Pb': 8, 'Ag': 1.5, 'Au': 1.5, 'Hg': 8, 'Bi': 8, 'Te': 8, 'Se': 8, 'As': 8}),
    MineralRule('Aurostibite',
                {'Au': (35, 51), 'Sb': (42, 65)},
                {'O': 15, 'Si': 8, 'Al': 8, 'Ca': 8, 'Mg': 8, 'S': 8, 'Fe': 8, 'Cu': 8, 'Pb': 8, 'Ag': 1.5, 'Hg': 8, 'Bi': 8, 'Te': 8, 'Se': 8, 'As': 8}),
    MineralRule('Stibnite',
                {'Sb': (58, 78), 'S': (14, 38)},
                {'O': 15, 'Si': 8, 'Al': 8, 'Ca': 8, 'Mg': 8, 'Fe': 8, 'Pb': 8, 'Cu': 8, 'Ag': 1.5, 'Au': 1.5, 'Hg': 8, 'Bi': 8, 'As': 8, 'Te': 8, 'Se': 8}),
    MineralRule('Boulangerite',
                {'Pb': (42, 64), 'Sb': (16, 34), 'S': (6, 28)},
                {'O': 15, 'Si': 8, 'Al': 8, 'Ca': 8, 'Mg': 8, 'Cu': 4, 'Ag': 1.5, 'Fe': 8, 'Au': 1.5, 'Hg': 8, 'Bi': 8, 'As': 8, 'Te': 8, 'Se': 8},
                ('Boulangerite ratio',)),
    MineralRule('Cupropolybasite',
                {'Ag': (30, 58), 'Cu': (13, 28), 'Sb': (4, 18), 'S': (7, 23)},
                {'O': 15, 'Si': 8, 'Al': 8, 'Ca': 8, 'Mg': 8, 'Pb': 8, 'Fe': 8, 'Hg': 8, 'Au': 1.5, 'Bi': 8, 'Te': 8, 'Se': 8},
                ('Sb dominant over As', 'Cupropolybasite ratio')),
    MineralRule('Cinnabar',
                {'Hg': (60, 95), 'S': (4, 20)},
                {'O': 15, 'Si': 8, 'Al': 8, 'Ca': 8, 'Mg': 8, 'Ag': 1.5, 'Cu': 8, 'Fe': 8, 'Pb': 8, 'Au': 1.5, 'Sb': 8, 'Te': 8, 'Se': 8, 'Bi': 8, 'As': 8}),
    MineralRule(AG_TETRAHEDRITE,
                {'Ag': (2, 45), 'Cu': (15, 45), 'Sb': (15, 35), 'S': (10, 31), 'Hg': (0, 25)},
                {'O': 15, 'Si': 8, 'Al': 8, 'Ca': 8, 'Mg': 8, 'Pb': 8, 'Au': 1.5, 'Bi': 8, 'Te': 8, 'Se': 8, 'Fe': 12, 'Zn': 12},
                ('Sb dominant over As', 'Tetrahedrite ratio')),
    MineralRule('Tetrahedrite-(Fe)',
                {'Cu': (22, 45), 'Sb': (20, 37), 'S': (15, 31), 'Fe': (1, 10)},
                {'O': 15, 'Si': 8, 'Al': 8, 'Ca': 8, 'Mg': 8, 'Pb': 8, 'Au': 1.5, 'Bi': 8, 'Te': 8, 'Se': 8, 'Ag': 1.5, 'Hg': 8, 'Zn': 2},
                ('Sb dominant over As', 'Tetrahedrite ratio')),
    MineralRule('Andorite',
                {'Ag': (4, 18), 'Pb': (14, 31), 'Sb': (30, 48), 'S': (12, 28)},
                {'O': 15, 'Si': 8, 'Al': 8, 'Ca': 8, 'Mg': 8, 'Cu': 8, 'Fe': 8, 'Au': 1.5, 'Hg': 8, 'Bi': 8, 'As': 8, 'Te': 8, 'Se': 8}),
    MineralRule('Zn-bearing Tetrahedrite',
                {'Cu': (22, 45), 'Sb': (20, 37), 'S': (15, 31), 'Zn': (2, 11), 'Fe': (0, 8)},
                {'O': 15, 'Si': 8, 'Al': 8, 'Ca': 8, 'Mg': 8, 'Pb': 8, 'Au': 1.5, 'Bi': 8, 'Te': 8, 'Se': 8, 'Ag': 1.5, 'Hg': 8},
                ('Sb dominant over As', 'Tetrahedrite ratio')),
    MineralRule('Bournonite',
                {'Pb': (30, 46), 'Sb': (12, 30), 'S': (7, 24), 'Cu': (4, 19)},
                {'O': 15, 'Si': 8, 'Al': 8, 'Ca': 8, 'Mg': 8, 'Ag': 1.5, 'Fe': 8, 'Au': 1.5, 'Hg': 8, 'Bi': 8, 'As': 8, 'Te': 8, 'Se': 8}),
    MineralRule('Pyrargyrite',
                {'Ag': (30, 65), 'Sb': (14, 29), 'S': (7, 22)},
                {'O': 15, 'Si': 8, 'Al': 8, 'Ca': 8, 'Mg': 8, 'Cu': 8, 'Pb': 8, 'Fe': 8, 'Au': 1.5, 'Hg': 8, 'Bi': 8, 'Te': 8, 'Se': 8},
                ('Sb dominant over As', 'Pyrargyrite ratio')),
)

ASSOCIATION_RULES = (
    MineralRule('Gudmundite-Asso',
                {'Sb': (15, 60), 'Fe': (7, 65), 'S': (3, 45)},
                {'As': 12, 'Bi': 30, 'Te': 8, 'Se': 8, 'Cu': 8, 'Pb': 8, 'Ag': 1.5, 'Au': 1.5, 'Hg': 10}),
    MineralRule('Aurostibite-Asso',
                {'Au': (4, 51), 'Sb': (10, 62)},
                {'As': 12, 'Bi': 30, 'Te': 8, 'Se': 8, 'Ag': 1.5, 'Cu': 25, 'Pb': 30, 'Hg': 20, 'Fe': 65},
                ('Aurostibite association ratio',)),
    MineralRule('Stibnite-Asso',
                {'Sb': (15, 77), 'S': (5, 48)},
                {'As': 12, 'Bi': 30, 'Te': 8, 'Se': 8, 'Fe': 7, 'Cu': 8, 'Pb': 8, 'Ag': 1.5, 'Au': 1.5, 'Hg': 10}),
    MineralRule('Boulangerite-Asso',
                {'Pb': (15, 65), 'Sb': (5, 31), 'S': (4, 40)},
                {'As': 12, 'Bi': 30, 'Te': 8, 'Se': 8, 'Cu': 3, 'Ag': 1.5, 'Fe': 12, 'Au': 1.5, 'Hg': 10}),
    MineralRule('Cupropolybasite-Asso',
                {'Ag': (4, 56), 'Cu': (5, 40), 'Sb': (4, 18), 'S': (4, 40)},
                {'As': 12, 'Bi': 30, 'Te': 8, 'Se': 8, 'Pb': 8, 'Fe': 25, 'Au': 1.5, 'Hg': 10},
                ('Sb dominant over As', 'Cupropolybasite ratio')),
    MineralRule('Cinnabar-Asso',
                {'Hg': (20, 90), 'S': (4, 40)},
                {'As': 12, 'Bi': 30, 'Te': 8, 'Se': 8, 'Ag': 1.5, 'Cu': 15, 'Pb': 20, 'Fe': 35, 'Au': 1.5, 'Sb': 10}),
    MineralRule('Andorite-Asso',
                {'Ag': (2, 17), 'Pb': (5, 50), 'Sb': (10, 47), 'S': (4, 40)},
                {'As': 12, 'Bi': 30, 'Te': 8, 'Se': 8, 'Cu': 8, 'Fe': 25, 'Au': 1.5, 'Hg': 10}),
    MineralRule('Tetrahedrite-(Fe)-Asso',
                {'Cu': (8, 43), 'Sb': (5, 35), 'S': (5, 45), 'Fe': (1, 45)},
                {'As': 12, 'Bi': 30, 'Te': 8, 'Se': 8, 'Ag': 1.5, 'Pb': 8, 'Zn': 15, 'Hg': 10, 'Au': 1.5}),
    MineralRule('Bournonite-Asso',
                {'Pb': (12, 46), 'Sb': (5, 30), 'S': (4, 24), 'Cu': (3, 19)},
                {'As': 12, 'Bi': 30, 'Te': 8, 'Se': 8, 'Ag': 1.5, 'Fe': 25, 'Au': 1.5, 'Hg': 10}),
    MineralRule('Pyrargyrite-Asso',
                {'Ag': (4, 65), 'Sb': (5, 29), 'S': (4, 22)},
                {'As': 12, 'Bi': 30, 'Te': 8, 'Se': 8, 'Cu': 5, 'Pb': 8, 'Fe': 30, 'Hg': 10, 'Au': 1.5},
                ('Sb dominant over As', 'Pyrargyrite association ratio')),
)


def primary_match(rule, values):
    return all(low <= values[element] <= high for element, (low, high) in rule.ranges.items())


def matches(rule, values):
    return (primary_match(rule, values)
            and all(values[element] < limit for element, limit in rule.caps.items())
            and additional_checks(values, rule.checks))


@dataclass(frozen=True)
class Classification:
    category: Optional[str]
    status: str
    candidates: tuple = ()
    note: str = ""


def leftover(values):
    if values["Sb"] > SB_CUTOFF:
        if values["Au"] > 2:
            return "Au-Sb Leftover"
        if values["Ag"] > 2:
            return "Ag-Sb Leftover"
        return "Sb Leftover"
    return None


def classify_feature(row, area_cutoff=0):
    """Exclusions -> eligibility -> size -> free -> associations -> leftovers."""
    try:
        values = {element: v(row, element) for element in ELEMENTS}
    except (TypeError, ValueError):
        return Classification(None, "Invalid data", note="Non-numeric elemental value; review Raw Data.")
    # Apply the user's exclusions before any mineral, size or leftover assignment.
    if values["W"] > 10 or values["La"] > 7:
        return Classification(None, "Excluded W/La", note="W >10 wt.% or La >7 wt.%.")
    if any(not isfinite(value) or value < 0 for value in values.values()):
        return Classification(None, "Invalid data", note="Negative or non-finite elemental value.")
    try:
        area = float(row[AREA_COLUMN])
    except (KeyError, TypeError, ValueError):
        return Classification(None, "Invalid data", note="Missing/non-numeric feature area.")
    if not isfinite(area) or area < 0:
        return Classification(None, "Invalid data", note="Missing, negative or non-finite feature area.")
    # HgS is allowed without Sb. Only a guarded cinnabar match extends eligibility.
    cinnabar_match = any(matches(rule, values) for rule in FREE_RULES + ASSOCIATION_RULES
                         if rule.name in ("Cinnabar", "Cinnabar-Asso"))
    if values["Sb"] <= SB_CUTOFF and not cinnabar_match:
        return Classification(None, "Outside scope", note="Sb <=2 wt.% and no cinnabar match.")
    if area_cutoff > 0 and area < area_cutoff:
        return Classification(SIZE_CATEGORY, "Classified", note="Eligible feature below area cutoff.")
    for rules, description in ((FREE_RULES, "Free"), (ASSOCIATION_RULES, "Association")):
        candidates = tuple(rule.name for rule in rules if matches(rule, values))
        if len(candidates) == 1:
            note = "Association screening assignment; review with images." if description == "Association" else ""
            return Classification(candidates[0], "Classified", candidates, note)
        if len(candidates) > 1:
            return Classification(leftover(values), "Classified", candidates,
                                  f"Ambiguous {description.lower()} match; retained in leftover for review.")
    return Classification(leftover(values), "Classified", note="No named mineral passed all conditions.")


def classify_mineral(row, area_cutoff=0):
    """Convenience interface matching the reference classifier."""
    return classify_feature(row, area_cutoff).category


def excel_safe_sheet_name(name):
    if name == AG_TETRAHEDRITE:
        return "Ag-AgHg Tetrahedrite (Fe)"
    return re.sub(r'[\\/*?:\[\]]', "-", str(name)).strip()[:31] or "Sheet"


def unique_output_path(source):
    source = Path(source)
    output = source.with_name(f"Classified{source.stem}.xlsx")
    number = 2
    while output.exists() or output.resolve() == source.resolve():
        output = source.with_name(f"Classified{source.stem}_{number}.xlsx")
        number += 1
    return output


HIGHLIGHT_COLORS = {
    "Feature": "00BFCF", AREA_COLUMN: "FFEB3B", "S (Wt%)": "FFA07A",
    "Fe (Wt%)": "ADD8E6", "Cu (Wt%)": "90EE90", "As (Wt%)": "D87093",
    "Ag (Wt%)": "F44336", "Sb (Wt%)": "A9A9A9", "Hg (Wt%)": "9370DB",
    "Se (Wt%)": "FFA09A", "Te (Wt%)": "D87099", "Au (Wt%)": "FFF2CC",
    "Bi (Wt%)": "E2EFDA", "Pb (Wt%)": "F4B183", "O (Wt%)": "D9E1F2",
    "Zn (Wt%)": "DDEBF7", "W (Wt%)": "F8CBAD", "La (Wt%)": "E4DFEC",
}


def format_output(output):
    book = load_workbook(output)
    try:
        for ws in book:
            if ws.title in ("Area", "Integrity Check") or ws.max_row < 2:
                continue
            headers = {cell.value: cell.column for cell in ws[1]}
            for name, color in HIGHLIGHT_COLORS.items():
                if name in headers:
                    for row in ws.iter_rows(min_row=2, min_col=headers[name], max_col=headers[name]):
                        row[0].fill = PatternFill(fill_type="solid", fgColor=color)
        summary = book["Summary"]
        thin = Side(style="thin")
        for column in summary.columns:
            width = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column)
            summary.column_dimensions[column[0].column_letter].width = width + 4
        for row in summary:
            for cell in row:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
        for cell in summary["A"]:
            cell.font = Font(bold=True)
        book.save(output)
    finally:
        book.close()


def build_classified_workbook(source, area_cutoff=0):
    """Create a separate workbook; never save into the selected input file."""
    if not isfinite(area_cutoff) or area_cutoff < 0:
        raise ValueError("Area cutoff must be a finite nonnegative number.")
    source = Path(source)
    print(f"Loading workbook: {source}", flush=True)
    with pd.ExcelFile(source) as xls:
        frame = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
    missing = [name for name in ("Sb (Wt%)", AREA_COLUMN) if name not in frame.columns]
    if missing:
        raise ValueError("Missing required column(s): " + ", ".join(missing))
    optional_missing = [f"{element} (Wt%)" for element in ELEMENTS if f"{element} (Wt%)" not in frame.columns]
    if optional_missing:
        print("Optional columns absent (treated as zero, as in the reference): "
              + ", ".join(optional_missing), flush=True)
    results = [classify_feature(row, area_cutoff) for _, row in frame.iterrows()]
    categories = pd.Series([result.category for result in results], index=frame.index, dtype=object)
    groups = {name: frame.loc[categories.eq(name)].copy() for name in CATEGORY_ORDER}
    groups = {name: group for name, group in groups.items() if not group.empty}
    eligible = pd.Series([result.status == "Classified" for result in results], index=frame.index, dtype=bool)
    area_values = pd.to_numeric(frame[AREA_COLUMN], errors="coerce")
    raw_count = int(eligible.sum())
    raw_area = float(area_values.loc[eligible].sum())
    classified_count = sum(len(group) for group in groups.values())
    classified_area = sum(float(pd.to_numeric(group[AREA_COLUMN]).sum()) for group in groups.values())
    if classified_count != raw_count or abs(classified_area - raw_area) >= 0.01:
        raise RuntimeError("Integrity check failed before writing: output counts/areas do not reconcile.")
    sheet_names = [excel_safe_sheet_name(name) for name in groups]
    if len({name.lower() for name in sheet_names}) != len(sheet_names):
        raise RuntimeError("Duplicate output sheet names after Excel name sanitization.")
    raw = frame.copy()
    audit = {
        "Sb Classification": [r.category for r in results],
        "Sb Classification Status": [r.status for r in results],
        "Sb Classification Candidates": ["; ".join(r.candidates) for r in results],
        "Sb Classification Note": [r.note for r in results],
    }
    for name, values in audit.items():
        while name in raw.columns:
            name += " (new)"
        raw[name] = values
    # Keep the reference's intentionally blank area and percentage summary cells.
    summary = pd.DataFrame({"Type of Mineral": list(groups),
                            "Total Sum of Area (sq. µm)": [None] * len(groups),
                            "Percentage": [None] * len(groups)})
    status_counts = {status: sum(r.status == status for r in results)
                     for status in ("Classified", "Excluded W/La", "Invalid data", "Outside scope")}
    review_count = sum("Ambiguous" in r.note for r in results)
    integrity_rows = [
        ("Total input rows", len(frame)),
        ("Eligibility", "Sb >2 wt.% OR guarded cinnabar match; W/La and invalid data excluded"),
        ("Area size cutoff used (sq. µm)", area_cutoff),
        ("Eligible rows before size routing", raw_count),
        ("Total rows in category sheets (including After Size-CutOff)", classified_count),
        ("Rows below area cutoff", int(categories.eq(SIZE_CATEGORY).sum())),
        ("Excluded W >10 or La >7 rows", status_counts["Excluded W/La"]),
        ("Invalid data rows", status_counts["Invalid data"]),
        ("Rows outside scope", status_counts["Outside scope"]),
        ("Ambiguous named-mineral matches sent to leftovers", review_count),
        ("Area in eligible raw rows", round(raw_area, 4)),
        ("Area in category sheets", round(classified_area, 4)),
        ("Area Match", "Match"), ("Row Count Match", "Match"),
        ("All raw rows accounted for", "Match" if sum(status_counts.values()) == len(frame) else "Mismatch"),
        ("Missing optional columns treated as zero", ", ".join(optional_missing) or "None"),
    ]
    output = unique_output_path(source)
    print(f"Writing: {output}", flush=True)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame().to_excel(writer, sheet_name="Area", index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False)
        raw.to_excel(writer, sheet_name="Raw Data", index=False)
        for name, group in groups.items():
            sheet_name = excel_safe_sheet_name(name)
            group.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"  {name} -> {sheet_name}: {len(group)} features", flush=True)
        pd.DataFrame(integrity_rows, columns=("Metric", "Value")).to_excel(
            writer, sheet_name="Integrity Check", index=False)
    format_output(output)
    print(f"Classification complete. Output: {output}", flush=True)
    if review_count:
        print(f"Review {review_count} ambiguous match(es) in the Raw Data audit columns.", flush=True)
    return output


def main():
    root = tk.Tk()
    root.withdraw()
    try:
        source = filedialog.askopenfilename(
            parent=root, title="Select Full Analysis Excel File — Sb Mineral Classification",
            filetypes=[("Excel files", "*.xlsx *.xls")])
        if not source:
            print("No file selected.", flush=True)
            return
        cutoff = simpledialog.askfloat(
            "Area Size Cutoff",
            "Enter Area (sq. µm) cutoff for this sample.\n"
            "Features smaller than this value will go to the After Size-CutOff sheet.\n"
            "Enter 0 to disable size filtering.", parent=root, minvalue=0, initialvalue=0)
        build_classified_workbook(source, 0 if cutoff is None else cutoff)
    finally:
        root.destroy()


if __name__ == "__main__":
    main()
