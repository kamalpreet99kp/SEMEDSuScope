import pandas as pd
import numpy as np
from tkinter import Tk, filedialog


# ============================================================
#  FILE SELECTION POP-UP
# ============================================================

Tk().withdraw()
input_file = filedialog.askopenfilename(
    title="Select GrainAlyser RAW Data Excel File (Sheet 2 – Grains)",
    filetypes=[("Excel Files", "*.xlsx *.xls")]
)

if not input_file:
    raise ValueError("No file selected.")

# Auto-generate output filename
if input_file.lower().endswith(".xlsx"):
    output_file = input_file.replace(".xlsx", "_LiberationOutput.xlsx")
elif input_file.lower().endswith(".xls"):
    output_file = input_file.replace(".xls", "_LiberationOutput.xlsx")
else:
    output_file = input_file + "_LiberationOutput.xlsx"

print("\nSelected file:", input_file)
print("Output will be saved as:", output_file)
print("--------------------------------------------------")


# ============================================================
#  STEP 1 — LOAD SHEET 2 WITH TWO-ROW HEADER
# ============================================================

xls = pd.ExcelFile(input_file)

# Detect "RAW Data (Grains)" sheet
sheet_grains = None
for s in xls.sheet_names:
    if "grain" in s.lower():
        sheet_grains = s

if sheet_grains is None:
    raise ValueError("Could not find sheet containing 'Grain'.")

# Read sheet with 2 header rows
df = pd.read_excel(xls, sheet_grains, header=[0, 1])

# Flatten multiheader
df.columns = [f"{a}_{b}".strip("_") for a, b in df.columns]


# ============================================================
#  STEP 2 — EXTRACT NECESSARY COLUMNS
# ============================================================

# Required fields
col_pid = "Particle Data_Id"
col_area = "Grain Data_AreaPercentage"
col_phase = "Grain Data_Phase"

# Verify columns exist
for col in [col_pid, col_area, col_phase]:
    if col not in df.columns:
        raise ValueError(f"Required column missing: {col}")

# Clean dataset
df = df[[col_pid, col_area, col_phase]].copy()


# ============================================================
#  STEP 3 — PHASE NORMALIZATION
# ============================================================

def normalize_phase(phase):
    if not isinstance(phase, str):
        return "Gg"

    p = phase.strip()

    # Cu-bearing groups
    if p in ["Chalcopyrite", "Cpy", "Cu Inclusion with FeO"]:
        return "Cpy"
    if p in ["Cubanite", "Cub"]:
        return "Cub"
    if p in ["Bornite", "Bo"]:
        return "Bo"
    if p in ["Covellite", "Cov"]:
        return "Cov"
    if p in ["Chalcocite", "Cc"]:
        return "Cc"
    if p in ["Native Cu", "Cu"]:
        return "Cu"
    if p == "Carrolite":
        return "Carrolite"

    # Py + Po → treat same
    if p in ["Pyrite", "Py", "Pyrrhotite", "Po"]:
        return "Py"

    # Everything else = gangue
    return "Gg"


df["NormPhase"] = df[col_phase].apply(normalize_phase)


# ============================================================
#  STEP 4 — BUILD PARTICLE-WISE COMPOSITION
# ============================================================

particles = {}

for _, row in df.iterrows():
    pid = row[col_pid]
    phase = row["NormPhase"]
    area_pct = row[col_area]

    if pid not in particles:
        particles[pid] = {
            "Cpy": 0, "Cub": 0, "Bo": 0, "Cov": 0, "Cc": 0,
            "Cu": 0, "Carrolite": 0,
            "Py": 0, "Gg": 0
        }

    particles[pid][phase] += area_pct


# ============================================================
#  STEP 5 — CLASSIFY PARTICLES
# ============================================================

results = []

for pid, comp in particles.items():
    CpyCub = comp["Cpy"] + comp["Cub"]
    Bo = comp["Bo"]
    CovCc = comp["Cov"] + comp["Cc"]
    NativeCu = comp["Cu"]
    Carrolite = comp["Carrolite"]
    Py = comp["Py"]
    Gg = comp["Gg"]

    Cu_total = CpyCub + Bo + CovCc + NativeCu + Carrolite
    if Cu_total <= 0:
        continue

    # Determine association (Py vs Gg)
    if Py == 0 and Gg == 0:
        association = "Free"
    elif Py > Gg:
        association = "Py"
    else:
        association = "Gg"

    # Dominant Cu mineral group
    groups = {
        "Cpy+Cub": CpyCub,
        "Bo": Bo,
        "Cov+Cc": CovCc,
        "Native Cu": NativeCu,
        "Carrolite": Carrolite
    }
    dominant_group = max(groups, key=groups.get)

    # Flotation category
    if Cu_total >= 80 and association == "Free":
        category = "Free"
    elif 20 <= Cu_total < 80:
        category = "Fl w/py" if association == "Py" else "Fl w/gg"
    else:  # Cu_total < 20
        category = "NF w/py" if association == "Py" else "NF w/gg"

    results.append([pid, dominant_group, Cu_total, association, category])


df_results = pd.DataFrame(
    results,
    columns=["ParticleID", "Group", "Cu%", "Assoc", "Category"]
)


# ============================================================
#  STEP 6 — BUILD FINAL SUMMARY TABLE (Row 8 & Row 9)
# ============================================================

groups = ["Cpy+Cub", "Bo", "Cov+Cc", "Native Cu", "Carrolite"]
categories = ["Free", "Fl w/py", "NF w/py", "Fl w/gg", "NF w/gg"]

final = pd.DataFrame()

for g in groups:
    for c in categories:
        final[f"{g} - {c}"] = (
            (df_results["Group"] == g) &
            (df_results["Category"] == c)
        ).astype(int)

counts = final.sum()
total_particles = counts.sum()
perc = (counts / total_particles * 100) if total_particles > 0 else counts*0


# ============================================================
#  STEP 7 — EXPORT OUTPUT EXCEL
# ============================================================

with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    summary = pd.DataFrame(
        [perc, counts],
        index=["Row8_Percent", "Row9_Count"]
    )
    summary.to_excel(writer, sheet_name="Output", startrow=0)

    df_results.to_excel(writer, sheet_name="Output", startrow=10, index=False)

print("\n===================================================")
print("✔ FINAL LIBERATION OUTPUT CREATED SUCCESSFULLY!")
print(f"Saved as: {output_file}")
print("===================================================")
