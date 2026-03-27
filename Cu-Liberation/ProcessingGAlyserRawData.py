import re
import sys
from pathlib import Path
from tkinter import Tk, filedialog

import pandas as pd


UNDERSIZED_THRESHOLDS = {
    "40-75": 3.0,
    "100-150": 3.0,
    "425-600": 5.0,
    "default": 3.0,
}

QC_SAMPLE = 50

AZTEC_LIBRARY = {
    "Chalcopyrite": "Cpy_Cubanite",
    "Cpy": "Cpy_Cubanite",
    "Cubanite": "Cpy_Cubanite",
    "Bornite": "Bornite",
    "Digenite": "Digenite_Chalcocite",
    "Chalcocite": "Digenite_Chalcocite",
    "Digenite/Chalcocite": "Digenite_Chalcocite",
    "Covellite": "Covellite",
    "Pyrite": "Py_Po",
    "Pyrrhotite": "Py_Po",
    "FeOx": "Gangue",
    "Gg": "Gangue",
    "Gg1": "Gangue",
    "Gg2": "Gangue",
    "Gangue": "Gangue",
    "Quartz": "Gangue",
    "Galena": "Gangue",
    "CaO": "Gangue",
    "Carrolitte": "Carrolite",
    "Carrolite": "Carrolite",
    "Sph": "Sphalerite",
    "Sphalerite": "Sphalerite",
    "Molybdenite": "Molybdenite",
    "Unclassified": "Unclassified",
}

GANGUE_LIKE_AZTEC = {
    "Gangue",
    "Gg",
    "Gg1",
    "Gg2",
    "FeOx",
    "Quartz",
    "Galena",
    "CaO",
    "Unclassified",
}

ACCEPTED_MATCHED_PHASES = {
    "Cpy_Cubanite",
    "Bornite",
    "Digenite_Chalcocite",
    "Covellite",
    "Py_Po",
    "Carrolite",
    "Sphalerite",
    "Molybdenite",
}


def choose_file() -> Path:
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    file_path = filedialog.askopenfilename(
        title="Select GrainAlyser Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls")],
    )
    root.destroy()

    if not file_path:
        raise SystemExit("No file selected.")

    return Path(file_path).resolve()


def detect_fraction(text: str) -> str:
    text = str(text)
    for pattern, name in [
        (r"40\s*[-–]\s*75", "40-75"),
        (r"100\s*[-–]\s*150", "100-150"),
        (r"425\s*[-–]\s*600", "425-600"),
        (r"\+\s*425", "425-600"),
        (r"425\s*\+", "425-600"),
    ]:
        if re.search(pattern, text):
            return name
    return "default"


def v(x):
    try:
        if pd.isna(x):
            return 0.0
    except Exception:
        pass
    try:
        return float(x)
    except Exception:
        return 0.0


def safe_num(series):
    return pd.to_numeric(series, errors="coerce")


def classify_row(row):
    Cu = v(row.get("Cu"))
    Fe = v(row.get("Fe"))
    S = v(row.get("S"))
    Zn = v(row.get("Zn"))
    Mo = v(row.get("Mo"))
    Co = v(row.get("Co"))
    Ni = v(row.get("Ni"))

    O = v(row.get("O"))
    Si = v(row.get("Si"))
    Al = v(row.get("Al"))
    Ca = v(row.get("Ca"))
    Mg = v(row.get("Mg"))
    Na = v(row.get("Na"))
    K = v(row.get("K"))
    Pb = v(row.get("Pb"))
    Ti = v(row.get("Ti"))
    P = v(row.get("P"))
    Ba = v(row.get("Ba"))

    gangue_sum = O + Si + Al + Ca + Mg + Na + K + Pb + Ti + P + Ba
    ratio = (Cu / S) if S > 0 else None

    if not row["Has_XRay"]:
        return "No_XRay"

    if row["Undersized"]:
        return "Undersized"

    # Special non-target sulfides / minerals
    if Zn >= 20 and S >= 15 and Cu < 5:
        return "Sphalerite"

    if Mo >= 20 and S >= 15 and Cu < 5:
        return "Molybdenite"

    if Co >= 15 and S >= 20 and Cu < 20:
        return "Carrolite"

    # Pyrite / Pyrrhotite
    if Fe >= 30 and S >= 25 and Cu <= 5:
        return "Py_Po"

    # Sure-shot gangue
    if gangue_sum >= 20 and Cu < 5 and S < 7:
        return "Gangue"

    # Direct Cu sulfide rules
    if Cu >= 5 and Cu <= 45 and Fe >= 15 and S >= 12 and (ratio is None or ratio <= 1.35):
        return "Cpy_Cubanite"

    if Cu >= 45 and Cu <= 68 and Fe >= 5 and Fe <= 17 and S >= 15 and (ratio is None or (1.55 <= ratio <= 2.70)):
        return "Bornite"

    if Cu >= 70 and Fe < 5 and S >= 12 and (ratio is None or ratio >= 2.8):
        return "Digenite_Chalcocite"

    if Cu >= 55 and Cu < 70 and Fe < 5 and S >= 15 and (ratio is None or ratio < 3.6):
        return "Covellite"

    # Tightened mixed logic
    if Cu >= 5 and S >= 12:
        # Fe-rich side
        if Fe >= 13.5:
            if ratio is not None:
                if ratio < 1.55:
                    return "Cpy_Cubanite"
                if ratio <= 2.70 and Cu >= 45:
                    return "Bornite"
            return "Mixed_Cu_Sulfide"

        if 5 <= Fe < 13.5:
            if ratio is not None:
                if ratio < 1.55:
                    return "Cpy_Cubanite"
                if ratio <= 2.70:
                    return "Bornite"
            return "Mixed_Cu_Sulfide"

        # Fe-poor side
        if Fe < 5:
            if ratio is not None:
                if Cu >= 70 and ratio >= 2.8:
                    return "Digenite_Chalcocite"
                if Cu >= 55 and ratio < 3.6:
                    return "Covellite"
            return "Mixed_Cu_Sulfide"

    # Fallback gangue for clearly non-target phases
    if Cu < 5 and S < 12 and Zn < 20 and Mo < 20 and Co < 15 and not (Fe >= 30 and S >= 25):
        return "Gangue"

    # Leftover sulfide-like review bucket
    if S >= 7 or Cu >= 3 or Fe >= 10 or Zn >= 10 or Mo >= 10 or Co >= 10 or Ni >= 10:
        return "Unclassified_SulphideLike"

    # Everything else
    return "Gangue"


def second_pass_normalized_resolve(row):
    """
    Normalize only Cu-Fe-S and re-run the same style rules.
    Used only for Mixed_Cu_Sulfide and sulphide-like mismatches.
    """
    Cu = v(row.get("Cu"))
    Fe = v(row.get("Fe"))
    S = v(row.get("S"))

    total = Cu + Fe + S
    if total <= 0:
        return None

    nCu = (Cu / total) * 100.0
    nFe = (Fe / total) * 100.0
    nS = (S / total) * 100.0
    ratio = (nCu / nS) if nS > 0 else None

    # Re-apply same rule style on normalized Cu-Fe-S
    if nFe >= 30 and nS >= 25 and nCu <= 5:
        return "Py_Po"

    if nCu >= 5 and nCu <= 45 and nFe >= 15 and nS >= 12 and (ratio is None or ratio <= 1.35):
        return "Cpy_Cubanite"

    if nCu >= 45 and nCu <= 68 and nFe >= 5 and nFe <= 17 and nS >= 15 and (ratio is None or (1.55 <= ratio <= 2.70)):
        return "Bornite"

    if nCu >= 70 and nFe < 5 and nS >= 12 and (ratio is None or ratio >= 2.8):
        return "Digenite_Chalcocite"

    if nCu >= 55 and nCu < 70 and nFe < 5 and nS >= 15 and (ratio is None or ratio < 3.6):
        return "Covellite"

    # Tightened mixed logic again
    if nCu >= 5 and nS >= 12:
        if nFe >= 13.5:
            if ratio is not None:
                if ratio < 1.55:
                    return "Cpy_Cubanite"
                if ratio <= 2.70 and nCu >= 45:
                    return "Bornite"
            return None

        if 5 <= nFe < 13.5:
            if ratio is not None:
                if ratio < 1.55:
                    return "Cpy_Cubanite"
                if ratio <= 2.70:
                    return "Bornite"
            return None

        if nFe < 5:
            if ratio is not None:
                if nCu >= 70 and ratio >= 2.8:
                    return "Digenite_Chalcocite"
                if nCu >= 55 and ratio < 3.6:
                    return "Covellite"
            return None

    return None


def final_route(row):
    """
    One feature -> one destination sheet only
    outside Master_Features.
    """
    if not row["Has_XRay"]:
        return "No_XRay"

    if row["Undersized"]:
        return "Undersized"

    corrected = str(row["Phase_Corrected"])
    aztec_raw = str(row["Phase_AZtec"])
    aztec_group = str(row["Phase_AZtec_Group"])

    # Gangue only if AZtec is already gangue-like/unclassified and corrected = Gangue
    if corrected == "Gangue" and aztec_raw in GANGUE_LIKE_AZTEC:
        return "Gangue"

    # Matched accepted mineral sheets
    if corrected in ACCEPTED_MATCHED_PHASES and corrected == aztec_group:
        return corrected

    # Resolved from Mixed
    if str(row.get("Resolution_Source", "")) == "Mixed_Cu_Sulfide" and pd.notna(row.get("Resolved_Phase")):
        return "Resolved_From_Mixed"

    # Resolved from Mismatched
    if str(row.get("Resolution_Source", "")) == "Mismatched" and pd.notna(row.get("Resolved_Phase")):
        return "Resolved_From_Mismatched"

    # Mixed bucket not resolved
    if corrected == "Mixed_Cu_Sulfide":
        return "Mixed_Cu_Sulfide"

    # Everything else review-worthy goes to Mismatched
    return "Mismatched"


def run_processing(input_path: Path):
    print("Starting script...")
    print(f"Selected file: {input_path}")

    fraction = detect_fraction(input_path.name)
    threshold = UNDERSIZED_THRESHOLDS.get(fraction, UNDERSIZED_THRESHOLDS["default"])
    print(f"Detected size fraction: {fraction}")
    print(f"Undersized threshold (Feature_ECD): {threshold}")

    print("Loading workbook sheets...")
    grains = pd.read_excel(input_path, sheet_name="RAW Data (Grains)", header=1)
    xray = pd.read_excel(input_path, sheet_name="RAW Data (X-Ray)", header=0)

    grains.columns = [
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

    xray = xray.rename(columns={"Particle ID": "Particle_ID", "Grain ID": "Grain_ID"})

    print("Running merge integrity checks...")
    dup_grains = grains.duplicated(subset=["Database", "Particle_ID", "Grain_ID"]).sum()
    dup_xray = xray.duplicated(subset=["Database", "Particle_ID", "Grain_ID"]).sum()

    if dup_grains > 0:
        raise ValueError(f"Merge stopped: duplicate keys found in RAW Data (Grains): {dup_grains}")
    if dup_xray > 0:
        raise ValueError(f"Merge stopped: duplicate keys found in RAW Data (X-Ray): {dup_xray}")

    master = grains.merge(xray, on=["Database", "Particle_ID", "Grain_ID"], how="left")

    if len(master) != len(grains):
        raise ValueError("Merge stopped: merged row count does not match RAW Data (Grains).")

    print("Merge checks passed.")
    print(f"Grain rows: {len(grains)}")
    print(f"X-Ray rows: {len(xray)}")
    print(f"Merged rows: {len(master)}")

    xray_cols = [c for c in xray.columns if c not in ["Database", "Particle_ID", "Grain_ID"]]

    numeric_cols = [
        "Particle_ID", "Grain_ID", "Particle_Area", "Particle_Perimeter", "Particle_ECD",
        "Particle_MinFeret", "Particle_MaxFeret", "Feature_Area", "Feature_Mass",
        "Feature_AreaPercentage", "Feature_Perimeter", "Feature_FreePPer", "Feature_ID",
        "Feature_ECD", "Feature_MinFeret", "Feature_MaxFeret"
    ] + xray_cols

    for col in numeric_cols:
        if col in master.columns:
            master[col] = safe_num(master[col])

    master["Has_XRay"] = ~master[xray_cols].fillna(0).eq(0).all(axis=1)
    master["Undersized"] = master["Feature_ECD"].fillna(0) < threshold
    master["Phase_AZtec_Group"] = master["Phase_AZtec"].map(lambda x: AZTEC_LIBRARY.get(str(x), str(x)))

    print("Assigning first-pass corrected phases...")
    master["Phase_Corrected"] = master.apply(classify_row, axis=1)

    # Prepare second-pass resolution columns
    master["Resolved_Phase"] = pd.NA
    master["Resolution_Source"] = pd.NA
    master["Norm_Cu"] = pd.NA
    master["Norm_Fe"] = pd.NA
    master["Norm_S"] = pd.NA
    master["Norm_Cu_S_Ratio"] = pd.NA

    print("Running second-pass normalized resolution...")

    # From Mixed_Cu_Sulfide
    mixed_mask = master["Phase_Corrected"] == "Mixed_Cu_Sulfide"
    for idx in master.index[mixed_mask]:
        row = master.loc[idx]
        Cu = v(row.get("Cu"))
        Fe = v(row.get("Fe"))
        S = v(row.get("S"))
        total = Cu + Fe + S
        if total > 0:
            nCu = (Cu / total) * 100.0
            nFe = (Fe / total) * 100.0
            nS = (S / total) * 100.0
            ratio = (nCu / nS) if nS > 0 else pd.NA
            resolved = second_pass_normalized_resolve(row)
            master.at[idx, "Norm_Cu"] = nCu
            master.at[idx, "Norm_Fe"] = nFe
            master.at[idx, "Norm_S"] = nS
            master.at[idx, "Norm_Cu_S_Ratio"] = ratio
            if resolved is not None:
                master.at[idx, "Resolved_Phase"] = resolved
                master.at[idx, "Resolution_Source"] = "Mixed_Cu_Sulfide"

    # From sulphide-like mismatches
    mismatch_sulph_mask = (
        master["Has_XRay"]
        & (~master["Undersized"])
        & (master["Phase_Corrected"] == "Unclassified_SulphideLike")
    ) | (
        master["Has_XRay"]
        & (~master["Undersized"])
        & (master["Phase_Corrected"] != master["Phase_AZtec_Group"])
        & (master["Phase_Corrected"] != "Gangue")
        & (master["Phase_Corrected"] != "Mixed_Cu_Sulfide")
    )

    for idx in master.index[mismatch_sulph_mask]:
        # skip anything already resolved from Mixed
        if pd.notna(master.at[idx, "Resolved_Phase"]):
            continue

        row = master.loc[idx]
        Cu = v(row.get("Cu"))
        Fe = v(row.get("Fe"))
        S = v(row.get("S"))
        total = Cu + Fe + S
        if total > 0:
            nCu = (Cu / total) * 100.0
            nFe = (Fe / total) * 100.0
            nS = (S / total) * 100.0
            ratio = (nCu / nS) if nS > 0 else pd.NA
            resolved = second_pass_normalized_resolve(row)
            master.at[idx, "Norm_Cu"] = nCu
            master.at[idx, "Norm_Fe"] = nFe
            master.at[idx, "Norm_S"] = nS
            master.at[idx, "Norm_Cu_S_Ratio"] = ratio
            if resolved is not None:
                master.at[idx, "Resolved_Phase"] = resolved
                master.at[idx, "Resolution_Source"] = "Mismatched"

    print("Assigning final exclusive sheet routes...")
    master["Final_Sheet"] = master.apply(final_route, axis=1)

    base_keep_cols = [
        "Database",
        "Particle_ID",
        "Particle_Area",
        "Particle_Perimeter",
        "Particle_ECD",
        "Particle_MinFeret",
        "Particle_MaxFeret",
        "Grain_ID",
        "Feature_ID",
        "Feature_Area",
        "Feature_Mass",
        "Feature_AreaPercentage",
        "Feature_Perimeter",
        "Feature_FreePPer",
        "Feature_ECD",
        "Feature_MinFeret",
        "Feature_MaxFeret",
        "Phase_AZtec",
        "Phase_AZtec_Group",
        "Phase_Corrected",
    ] + xray_cols

    resolved_keep_cols = base_keep_cols + [
        "Norm_Cu",
        "Norm_Fe",
        "Norm_S",
        "Norm_Cu_S_Ratio",
        "Resolved_Phase",
        "Resolution_Source",
    ]

    # Master sheet: all XRay rows except No_XRay and Undersized
    working = master[(master["Has_XRay"]) & (~master["Undersized"])].copy()
    master_sheet = working[base_keep_cols].copy()

    # Exclusive non-master sheets
    sheets = {}
    for sheet_name in [
        "Cpy_Cubanite",
        "Bornite",
        "Digenite_Chalcocite",
        "Covellite",
        "Gangue",
        "Carrolite",
        "Sphalerite",
        "Molybdenite",
        "Py_Po",
        "Mismatched",
        "Mixed_Cu_Sulfide",
        "Resolved_From_Mixed",
        "Resolved_From_Mismatched",
        "No_XRay",
        "Undersized",
    ]:
        if sheet_name in {"Resolved_From_Mixed", "Resolved_From_Mismatched"}:
            df = master[master["Final_Sheet"] == sheet_name][resolved_keep_cols].copy()
        else:
            df = master[master["Final_Sheet"] == sheet_name][base_keep_cols].copy()

        if len(df) > 0:
            sheets[sheet_name] = df

    # Exclusivity check
    non_master = master[master["Final_Sheet"] != ""][["Database", "Particle_ID", "Grain_ID", "Final_Sheet"]].copy()
    dup_non_master = non_master.duplicated(subset=["Database", "Particle_ID", "Grain_ID"]).sum()
    if dup_non_master > 0:
        raise ValueError(f"Exclusivity check failed: {dup_non_master} features appear in more than one non-master sheet.")

    checks = pd.DataFrame(
        {
            "Metric": [
                "Detected fraction",
                "Undersized threshold (Feature_ECD)",
                "Duplicate grain keys in Sheet2",
                "Duplicate xray keys in Sheet3",
                "Duplicate features across non-master sheets",
                "Total merged rows",
                "Rows with XRay",
                "Rows without XRay",
                "Undersized rows",
                "Master_Features rows",
                "Cpy_Cubanite rows",
                "Bornite rows",
                "Digenite_Chalcocite rows",
                "Covellite rows",
                "Gangue rows",
                "Carrolite rows",
                "Sphalerite rows",
                "Molybdenite rows",
                "Py_Po rows",
                "Mismatched rows",
                "Mixed_Cu_Sulfide rows",
                "Resolved_From_Mixed rows",
                "Resolved_From_Mismatched rows",
            ],
            "Value": [
                fraction,
                threshold,
                dup_grains,
                dup_xray,
                dup_non_master,
                len(master),
                int(master["Has_XRay"].sum()),
                int((~master["Has_XRay"]).sum()),
                int(master["Undersized"].sum()),
                len(master_sheet),
                len(sheets.get("Cpy_Cubanite", [])),
                len(sheets.get("Bornite", [])),
                len(sheets.get("Digenite_Chalcocite", [])),
                len(sheets.get("Covellite", [])),
                len(sheets.get("Gangue", [])),
                len(sheets.get("Carrolite", [])),
                len(sheets.get("Sphalerite", [])),
                len(sheets.get("Molybdenite", [])),
                len(sheets.get("Py_Po", [])),
                len(sheets.get("Mismatched", [])),
                len(sheets.get("Mixed_Cu_Sulfide", [])),
                len(sheets.get("Resolved_From_Mixed", [])),
                len(sheets.get("Resolved_From_Mismatched", [])),
            ],
        }
    )

    qc = master_sheet.sample(min(QC_SAMPLE, len(master_sheet)), random_state=42).copy()

    output_file = input_path.with_name(f"Processed_{input_path.stem}.xlsx")

    print("Writing output workbook...")
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        master_sheet.to_excel(writer, sheet_name="Master_Features", index=False)

        for sheet_name in [
            "Cpy_Cubanite",
            "Bornite",
            "Digenite_Chalcocite",
            "Covellite",
            "Gangue",
            "Carrolite",
            "Sphalerite",
            "Molybdenite",
            "Py_Po",
            "Mismatched",
            "Mixed_Cu_Sulfide",
            "Resolved_From_Mixed",
            "Resolved_From_Mismatched",
            "No_XRay",
            "Undersized",
        ]:
            if sheet_name in sheets:
                sheets[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)

        checks.to_excel(writer, sheet_name="All_Checks", index=False)
        qc.to_excel(writer, sheet_name="Random_QC_50", index=False)

    print("Done.")
    print(f"Output file: {output_file}")
    print(f"Master_Features: {len(master_sheet)}")
    for name in [
        "Cpy_Cubanite",
        "Bornite",
        "Digenite_Chalcocite",
        "Covellite",
        "Gangue",
        "Carrolite",
        "Sphalerite",
        "Molybdenite",
        "Py_Po",
        "Mismatched",
        "Mixed_Cu_Sulfide",
        "Resolved_From_Mixed",
        "Resolved_From_Mismatched",
        "No_XRay",
        "Undersized",
    ]:
        print(f"{name}: {len(sheets.get(name, []))}")

    return output_file


def main():
    if len(sys.argv) > 1:
        input_path = Path(sys.argv[1]).resolve()
    else:
        input_path = choose_file()
    run_processing(input_path)


if __name__ == "__main__":
    main()