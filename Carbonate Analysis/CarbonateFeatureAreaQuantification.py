import re
import sys
from pathlib import Path
from tkinter import Tk, filedialog

import pandas as pd


# ============================================================
# Configurable thresholds
# ============================================================
UNWANTED_ELEMENT_MAX_WT = 3.0
SI_CUTOFF_WT = 6.0
SI_RATIO_MAX = 0.15

# Allowed chemistry for screening workflow (all others treated as unwanted)
ALLOWED_SCREEN_ELEMENTS = {"O", "Fe", "Si", "Ca", "Mg"}

FINAL_CLASSES = ["Calcite", "Dolomite", "Siderite", "Ankerite"]


# ============================================================
# Helpers: file/sheet/column detection
# ============================================================

def choose_file() -> Path:
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    file_path = filedialog.askopenfilename(
        title="Select GrainAlyser Export File",
        filetypes=[("Excel Files", "*.xlsx *.xls")],
    )
    root.destroy()
    if not file_path:
        raise SystemExit("No file selected.")
    return Path(file_path).resolve()


def normalize_name(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(text).lower())


def find_sheet(sheet_names, contains_tokens):
    """Return first sheet that contains all token fragments in normalized form."""
    token_norm = [normalize_name(t) for t in contains_tokens]
    for name in sheet_names:
        n = normalize_name(name)
        if all(t in n for t in token_norm):
            return name
    return None


def map_columns(df: pd.DataFrame):
    """Map likely GrainAlyser grain columns to standardized names by position and fallback names."""
    # Expected structure in RAW Data (Grains): 18 columns
    # [Database, Particle Id, Particle Area, ..., Grain Id, Feature Area, ..., Feature ID, Feature ECD, ..., Phase]
    out = df.copy()

    if len(out.columns) >= 18:
        out = out.iloc[:, :18].copy()
        out.columns = [
            "Database",
            "Particle_ID",
            "Particle_Area",
            "Particle_Perimeter",
            "Particle_ECD",
            "Particle_MinFeret",
            "Particle_MaxFeret",
            "Grain_ID",
            "Feature_Area",
            "Feature_Mass",
            "Feature_AreaPercentage",
            "Feature_Perimeter",
            "Feature_FreePPer",
            "Feature_ID",
            "Feature_ECD",
            "Feature_MinFeret",
            "Feature_MaxFeret",
            "Phase_AZtec",
        ]
        return out

    # Fallback for non-standard structures
    name_map = {}
    for c in out.columns:
        n = normalize_name(c)
        if n in {"database"}:
            name_map[c] = "Database"
        elif n in {"particleid", "idparticle", "particle"}:
            name_map[c] = "Particle_ID"
        elif n in {"grainid", "idgrain"}:
            name_map[c] = "Grain_ID"
        elif n in {"featureid", "idfeature"}:
            name_map[c] = "Feature_ID"
        elif n in {"featurearea", "areafeature"}:
            name_map[c] = "Feature_Area"
        elif n in {"phase", "phaseaztec"}:
            name_map[c] = "Phase_AZtec"
    out = out.rename(columns=name_map)
    return out


def choose_output_dir(input_path: Path) -> Path:
    """Prefer sibling output folder for raw/test/processed workflows when available."""
    parent = input_path.parent
    if parent.name.lower() in {"raw", "test", "processed", "output"}:
        sibling_output = parent.parent / "output"
        sibling_output.mkdir(parents=True, exist_ok=True)
        return sibling_output
    default_output = parent / "output"
    default_output.mkdir(parents=True, exist_ok=True)
    return default_output


def safe_num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")


# ============================================================
# Classification logic
# ============================================================

def carbonate_class_from_norm(ca, mg, fe):
    """Order is important: Calcite -> Siderite -> Dolomite -> Ankerite."""
    if pd.isna(ca) or pd.isna(mg) or pd.isna(fe):
        return "unclassified_carbonate"

    # 1) Calcite
    if ca > 75.00 and mg < 9.09 and fe < 9.09:
        return "Calcite"

    # 2) Siderite
    if ca < 7.69 and mg < 11.11 and 62.50 <= fe <= 100.00:
        return "Siderite"

    # 3) Dolomite
    if 54.41 <= ca <= 70.59 and 17.65 <= mg <= 44.12 and fe < 15.00:
        return "Dolomite"

    # 4) Ankerite
    if 26.32 <= ca <= 89.47 and 4.74 <= mg <= 34.21 and 15.00 <= fe <= 65.79:
        return "Ankerite"

    return "unclassified_carbonate"


# ============================================================
# Main processing
# ============================================================

def run(input_path: Path) -> Path:
    print(f"[INFO] Reading file: {input_path}")

    xl = pd.ExcelFile(input_path)
    sheets = xl.sheet_names

    grains_sheet = find_sheet(sheets, ["raw", "grain"]) or find_sheet(sheets, ["grain"])
    xray_sheet = find_sheet(sheets, ["raw", "x", "ray"]) or find_sheet(sheets, ["x", "ray"])

    if not grains_sheet or not xray_sheet:
        raise ValueError(
            f"Could not find required grain/xray sheets. Available sheets: {sheets}"
        )

    print(f"[INFO] Using grain sheet: {grains_sheet}")
    print(f"[INFO] Using x-ray sheet: {xray_sheet}")

    grains_raw = pd.read_excel(input_path, sheet_name=grains_sheet, header=1)
    xray_raw = pd.read_excel(input_path, sheet_name=xray_sheet, header=0)

    grains = map_columns(grains_raw)

    # Standardize expected key names in X-ray sheet
    xray = xray_raw.rename(
        columns={
            "Particle ID": "Particle_ID",
            "Grain ID": "Grain_ID",
        }
    ).copy()

    for key in ["Database", "Particle_ID", "Grain_ID"]:
        if key not in grains.columns:
            raise ValueError(f"Missing required merge key in grain data: {key}")
        if key not in xray.columns:
            raise ValueError(f"Missing required merge key in x-ray data: {key}")

    # Duplicate-key validation before merge
    dup_grains = grains.duplicated(subset=["Database", "Particle_ID", "Grain_ID"]).sum()
    dup_xray = xray.duplicated(subset=["Database", "Particle_ID", "Grain_ID"]).sum()
    if dup_grains > 0:
        raise ValueError(f"Duplicate grain keys found: {dup_grains}")
    if dup_xray > 0:
        raise ValueError(f"Duplicate x-ray keys found: {dup_xray}")

    merged = grains.merge(xray, on=["Database", "Particle_ID", "Grain_ID"], how="left")

    # Build chemistry column list from x-ray sheet, excluding keys
    chem_cols = [c for c in xray.columns if c not in ["Database", "Particle_ID", "Grain_ID"]]

    # Numeric conversion for relevant columns
    numeric_candidates = [
        "Particle_ID",
        "Grain_ID",
        "Feature_ID",
        "Feature_Area",
        "Feature_ECD",
        "Particle_Area",
        "Particle_ECD",
    ] + chem_cols

    for col in numeric_candidates:
        if col in merged.columns:
            merged[col] = safe_num(merged[col])

    # Feature-level chemistry availability
    merged["Has_XRay"] = ~merged[chem_cols].fillna(0).eq(0).all(axis=1)

    # Validate area availability
    merged["Missing_Area"] = merged["Feature_Area"].isna() if "Feature_Area" in merged.columns else True
    missing_area_count = int(merged["Missing_Area"].sum())

    # Unwanted-element filter (all chemistry except allowed list)
    element_like_cols = [c for c in chem_cols if re.fullmatch(r"[A-Za-z]{1,2}", str(c).strip())]
    unwanted_cols = [c for c in element_like_cols if c not in ALLOWED_SCREEN_ELEMENTS]

    if unwanted_cols:
        merged["Max_Unwanted_wt"] = merged[unwanted_cols].max(axis=1, skipna=True)
        merged["Has_Unwanted_Above_Threshold"] = merged[unwanted_cols].fillna(0).gt(UNWANTED_ELEMENT_MAX_WT).any(axis=1)
    else:
        merged["Max_Unwanted_wt"] = 0.0
        merged["Has_Unwanted_Above_Threshold"] = False

    # Ensure key screening columns exist
    for col in ["Ca", "Mg", "Fe", "Si"]:
        if col not in merged.columns:
            merged[col] = pd.NA
        merged[col] = safe_num(merged[col])

    merged[["Ca", "Mg", "Fe", "Si"]] = merged[["Ca", "Mg", "Fe", "Si"]].fillna(0.0)

    # Cation+Si screening normalization: Ca+Mg+Fe+Si = 100
    merged["Screen_Sum"] = merged["Ca"] + merged["Mg"] + merged["Fe"] + merged["Si"]
    screen_denom = merged["Screen_Sum"].replace({0: pd.NA})
    merged["Ca_Screen_Norm"] = (merged["Ca"] / screen_denom) * 100
    merged["Mg_Screen_Norm"] = (merged["Mg"] / screen_denom) * 100
    merged["Fe_Screen_Norm"] = (merged["Fe"] / screen_denom) * 100
    merged["Si_Screen_Norm"] = (merged["Si"] / screen_denom) * 100

    # Si ratio check (raw values)
    cation_sum = merged["Ca"] + merged["Mg"] + merged["Fe"]
    merged["Si_ratio"] = merged["Si"] / cation_sum.replace({0: pd.NA})

    # Candidate screening status
    si_cutoff_pass = merged["Si"] <= SI_CUTOFF_WT
    si_ratio_pass = merged["Si_ratio"] < SI_RATIO_MAX

    merged["Screen_Status"] = "pass_carbonate_candidate"
    merged.loc[merged["Has_Unwanted_Above_Threshold"], "Screen_Status"] = "fail_unwanted_elements"
    merged.loc[(~merged["Has_Unwanted_Above_Threshold"]) & (~si_cutoff_pass), "Screen_Status"] = "fail_si_cutoff"
    merged.loc[(~merged["Has_Unwanted_Above_Threshold"]) & (si_cutoff_pass) & (~si_ratio_pass), "Screen_Status"] = "fail_si_ratio"

    merged["Is_Carbonate_Candidate"] = merged["Screen_Status"] == "pass_carbonate_candidate"

    # Final carbonate normalization: Ca+Mg+Fe = 100 (for candidates)
    carb_denom = cation_sum.replace({0: pd.NA})
    merged["Ca_Carb_Norm"] = (merged["Ca"] / carb_denom) * 100
    merged["Mg_Carb_Norm"] = (merged["Mg"] / carb_denom) * 100
    merged["Fe_Carb_Norm"] = (merged["Fe"] / carb_denom) * 100

    merged["Final_Carbonate_Class"] = "not_carbonate"

    cand_mask = merged["Is_Carbonate_Candidate"]
    merged.loc[cand_mask, "Final_Carbonate_Class"] = merged.loc[cand_mask].apply(
        lambda r: carbonate_class_from_norm(r["Ca_Carb_Norm"], r["Mg_Carb_Norm"], r["Fe_Carb_Norm"]),
        axis=1,
    )

    # Sheets
    merged_raw = merged.copy()
    cleaned_screened = merged.copy()
    carbonate_candidates = merged[merged["Is_Carbonate_Candidate"]].copy()
    final_carbonates = merged[merged["Final_Carbonate_Class"].isin(FINAL_CLASSES)].copy()

    rejected_features = merged[
        (~merged["Is_Carbonate_Candidate"]) | (merged["Final_Carbonate_Class"] == "unclassified_carbonate")
    ].copy()
    rejected_features["Rejected_Reason"] = rejected_features["Screen_Status"]
    rejected_features.loc[
        rejected_features["Final_Carbonate_Class"] == "unclassified_carbonate",
        "Rejected_Reason",
    ] = "unclassified_carbonate"

    # Optional particle-level summary
    particle_summary = None
    if {"Particle_ID", "Feature_Area"}.issubset(merged.columns):
        p_total = merged.groupby("Particle_ID", dropna=False, as_index=False)["Feature_Area"].sum().rename(
            columns={"Feature_Area": "Particle_Total_Area"}
        )
        p_carb = final_carbonates.groupby("Particle_ID", dropna=False, as_index=False)["Feature_Area"].sum().rename(
            columns={"Feature_Area": "Particle_Carbonate_Area"}
        )
        p_class = (
            final_carbonates.groupby(["Particle_ID", "Final_Carbonate_Class"], dropna=False)["Feature_Area"]
            .sum()
            .reset_index()
            .sort_values(["Particle_ID", "Feature_Area"], ascending=[True, False])
        )
        p_dom = p_class.drop_duplicates(subset=["Particle_ID"]).rename(
            columns={"Final_Carbonate_Class": "Dominant_Carbonate", "Feature_Area": "Dominant_Carbonate_Area"}
        )

        particle_summary = p_total.merge(p_carb, on="Particle_ID", how="left").merge(
            p_dom[["Particle_ID", "Dominant_Carbonate", "Dominant_Carbonate_Area"]],
            on="Particle_ID",
            how="left",
        )
        particle_summary["Particle_Carbonate_Area"] = particle_summary["Particle_Carbonate_Area"].fillna(0.0)
        particle_summary["Carbonate_Area_%"] = (
            particle_summary["Particle_Carbonate_Area"]
            / particle_summary["Particle_Total_Area"].replace({0: pd.NA})
            * 100
        )

    # Summary metrics
    total_features = len(merged)
    total_area = merged["Feature_Area"].sum(skipna=True)
    cand_area = carbonate_candidates["Feature_Area"].sum(skipna=True)
    final_area = final_carbonates["Feature_Area"].sum(skipna=True)
    missing_chem = int((~merged["Has_XRay"]).sum())

    summary_rows = [
        ("input_file", str(input_path)),
        ("grain_sheet_used", grains_sheet),
        ("xray_sheet_used", xray_sheet),
        ("merge_keys", "Database, Particle_ID, Grain_ID"),
        ("merged_record_count", total_features),
        ("missing_chemistry_rows", missing_chem),
        ("missing_area_rows", missing_area_count),
        ("total_feature_area", total_area),
        ("carbonate_candidate_area", cand_area),
        ("final_carbonate_area", final_area),
        ("final_carbonate_area_pct_of_total", (final_area / total_area * 100) if total_area else pd.NA),
    ]

    for cls in FINAL_CLASSES + ["unclassified_carbonate"]:
        cdf = merged[merged["Final_Carbonate_Class"] == cls]
        area = cdf["Feature_Area"].sum(skipna=True)
        summary_rows.append((f"{cls}_count", len(cdf)))
        summary_rows.append((f"{cls}_area", area))
        summary_rows.append((f"{cls}_area_pct_total", (area / total_area * 100) if total_area else pd.NA))

    for reason in ["fail_unwanted_elements", "fail_si_cutoff", "fail_si_ratio", "unclassified_carbonate"]:
        rdf = rejected_features[rejected_features["Rejected_Reason"] == reason]
        summary_rows.append((f"rejected_{reason}_count", len(rdf)))
        summary_rows.append((f"rejected_{reason}_area", rdf["Feature_Area"].sum(skipna=True)))

    summary = pd.DataFrame(summary_rows, columns=["Metric", "Value"])

    output_dir = choose_output_dir(input_path)
    output_file = output_dir / f"Carbonate_Processed_{input_path.stem}.xlsx"

    print(f"[INFO] Writing output workbook: {output_file}")
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        merged_raw.to_excel(writer, sheet_name="merged_raw", index=False)
        cleaned_screened.to_excel(writer, sheet_name="cleaned_screened", index=False)
        carbonate_candidates.to_excel(writer, sheet_name="carbonate_candidates", index=False)
        final_carbonates.to_excel(writer, sheet_name="final_carbonates", index=False)
        rejected_features.to_excel(writer, sheet_name="rejected_features", index=False)
        summary.to_excel(writer, sheet_name="summary", index=False)
        if particle_summary is not None:
            particle_summary.to_excel(writer, sheet_name="particle_summary", index=False)

    # Concise terminal summary
    print("\n=== Carbonate Processing Summary ===")
    print(f"Features merged: {total_features}")
    print(f"Total area: {total_area:.4f}" if pd.notna(total_area) else "Total area: NA")
    print(f"Carbonate candidate area: {cand_area:.4f}" if pd.notna(cand_area) else "Carbonate candidate area: NA")
    print(f"Final carbonate area: {final_area:.4f}" if pd.notna(final_area) else "Final carbonate area: NA")
    for cls in FINAL_CLASSES:
        cdf = merged[merged["Final_Carbonate_Class"] == cls]
        area = cdf["Feature_Area"].sum(skipna=True)
        print(f"{cls}: count={len(cdf)}, area={area:.4f}" if pd.notna(area) else f"{cls}: count={len(cdf)}, area=NA")

    print("Rejected by reason:")
    for reason in ["fail_unwanted_elements", "fail_si_cutoff", "fail_si_ratio", "unclassified_carbonate"]:
        rdf = rejected_features[rejected_features["Rejected_Reason"] == reason]
        area = rdf["Feature_Area"].sum(skipna=True)
        print(f"- {reason}: count={len(rdf)}, area={area:.4f}" if pd.notna(area) else f"- {reason}: count={len(rdf)}, area=NA")

    return output_file


if __name__ == "__main__":
    if len(sys.argv) > 1:
        input_file = Path(sys.argv[1]).resolve()
    else:
        input_file = choose_file()
    run(input_file)
