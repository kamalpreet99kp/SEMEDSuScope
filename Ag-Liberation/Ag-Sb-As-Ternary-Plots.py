"""Create one combined Ag-Sb-As ternary plot from a folder of CSV files.

The input CSV files must contain Label, Ag, Sb, As, and Size columns. Ag, Sb,
and As are expected to be wt%. The script leaves the input files unchanged and
writes one HTML, PNG, and SVG plot in the selected folder.
"""

import colorsys
import hashlib
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd
import plotly.express as px


REQUIRED_COLUMNS = ("Label", "Ag", "Sb", "As", "Size")
FIXED_COLORS = {
    "sulfosalts": "green",
    "cupropearceite-std": "red",
    "stephanite-std": "blue",
    "proustite-std": "brown",
    "xanthoconite-std": "brown",
    "proustite or xanthoconite-std": "brown",
    "proustite-xanthoconite-std": "brown",
}
MAX_MARKER_SIZE = 20
MARKER_SIZE_SCALING_FACTOR = 0.01


def make_axis(title: str, tick_angle: int) -> dict:
    """Return consistent formatting for one ternary axis."""
    return {
        "title": title,
        "min": 0,
        "tickangle": tick_angle,
        "dtick": 10,
        "tickfont": {"size": 7},
        "tickcolor": "rgba(0,0,0,1)",
        "ticklen": 5,
        "showline": True,
        "linecolor": "rgba(0,0,0,1)",
        "linewidth": 1,
        "showgrid": True,
        "gridcolor": "rgba(0,0,0,0.5)",
        "layer": "below traces",
        "ticksuffix": " %",
    }


def random_color_for_name(name: str) -> str:
    """Return a bright, repeatable color for an unlisted CSV name."""
    digest = hashlib.sha256(name.casefold().encode("utf-8")).digest()
    hue = int.from_bytes(digest[:2], "big") / 65535
    red, green, blue = colorsys.hsv_to_rgb(hue, 0.70, 0.80)
    return f"rgb({round(red * 255)}, {round(green * 255)}, {round(blue * 255)})"


def color_for_file(csv_file: Path) -> str:
    """Use a requested fixed color when its mineral name occurs in a filename."""
    normalized_name = csv_file.stem.casefold().replace(" ", "")
    for mineral_name, color in FIXED_COLORS.items():
        if mineral_name.casefold().replace(" ", "") in normalized_name:
            return color
    return random_color_for_name(csv_file.stem)


def read_and_prepare_csv(csv_file: Path) -> tuple[pd.DataFrame, int]:
    """Read a CSV and validate the raw Ag-Sb-As values for ternary plotting."""
    dataframe = pd.read_csv(csv_file, encoding="utf-8-sig")
    actual_columns = {str(column).strip().casefold(): column for column in dataframe.columns}
    missing = [column for column in REQUIRED_COLUMNS if column.casefold() not in actual_columns]
    if missing:
        raise ValueError("missing required column(s): " + ", ".join(missing))

    dataframe = dataframe[
        [actual_columns[column.casefold()] for column in REQUIRED_COLUMNS]
    ].copy()
    dataframe.columns = list(REQUIRED_COLUMNS)

    for column in ("Ag", "Sb", "As", "Size"):
        dataframe[column] = pd.to_numeric(dataframe[column], errors="coerce")

    original_row_count = len(dataframe)
    component_sum = dataframe[["Ag", "Sb", "As"]].sum(axis=1, min_count=3)
    valid_rows = (
        dataframe[["Ag", "Sb", "As", "Size"]].notna().all(axis=1)
        & (dataframe[["Ag", "Sb", "As"]] >= 0).all(axis=1)
        & (component_sum > 0)
        & (dataframe["Size"] > 0)
    )
    dataframe = dataframe.loc[valid_rows].copy()
    dataframe["Label"] = dataframe["Label"].fillna("").astype(str)
    return dataframe, original_row_count - len(dataframe)


def create_ternary_plot(
    dataframes: list[pd.DataFrame], csv_files: list[Path], output_folder: Path
) -> list[str]:
    """Combine all valid CSV data and create one plot with a layer per file."""
    prepared_dataframes = []
    color_map = {}
    for dataframe, csv_file in zip(dataframes, csv_files):
        dataframe = dataframe.copy()
        dataframe["Layer"] = csv_file.stem
        prepared_dataframes.append(dataframe)
        color_map[csv_file.stem] = color_for_file(csv_file)

    combined = pd.concat(prepared_dataframes, ignore_index=True)

    figure = px.scatter_ternary(
        combined,
        a="Ag",
        b="Sb",
        c="As",
        color="Layer",
        size="Size",
        hover_name="Label",
        custom_data=["Ag", "Sb", "As", "Size"],
        size_max=MAX_MARKER_SIZE,
        color_discrete_map=color_map,
    )
    figure.update_layout(
        title=None,
        showlegend=True,
        ternary={
            "sum": 100,
            "aaxis": make_axis("Ag", 0),
            "baxis": make_axis("Sb", 45),
            "caxis": make_axis("As", -45),
        },
    )

    for trace in figure.data:
        trace.cliponaxis = False
        trace.marker.sizeref = MARKER_SIZE_SCALING_FACTOR
        trace.hovertemplate = (
            "<b>%{hovertext}</b><br>"
            "Ag: %{customdata[0]:.2f} wt%<br>"
            "Sb: %{customdata[1]:.2f} wt%<br>"
            "As: %{customdata[2]:.2f} wt%<br>"
            "Size: %{customdata[3]:.2f}<extra>%{fullData.name}</extra>"
        )

    output_stem = output_folder / "Combined (Ternary Plot)"
    figure.write_html(str(output_stem) + ".html")

    warnings = []
    for extension in (".png", ".svg"):
        try:
            figure.write_image(str(output_stem) + extension)
        except Exception as error:
            warnings.append(f"Could not create {output_stem.name}{extension}: {error}")
    return warnings


def main() -> None:
    root = tk.Tk()
    root.withdraw()
    selected_folder = filedialog.askdirectory(
        parent=root,
        title="Select folder containing Ag-Sb-As CSV files",
    )
    if not selected_folder:
        root.destroy()
        return

    folder = Path(selected_folder)
    csv_files = sorted(folder.glob("*.csv"), key=lambda path: path.name.casefold())
    if not csv_files:
        messagebox.showerror("No CSV files", "No CSV files were found in the selected folder.")
        root.destroy()
        return

    valid_dataframes = []
    valid_csv_files = []
    completed = []
    skipped = []
    warnings = []
    for csv_file in csv_files:
        try:
            dataframe, skipped_rows = read_and_prepare_csv(csv_file)
            if dataframe.empty:
                skipped.append(f"{csv_file.name}: no valid data rows")
                continue
            valid_dataframes.append(dataframe)
            valid_csv_files.append(csv_file)
            completed.append(f"{csv_file.name} ({len(dataframe)} rows; {skipped_rows} skipped)")
        except Exception as error:
            skipped.append(f"{csv_file.name}: {error}")

    if valid_dataframes:
        warnings.extend(create_ternary_plot(valid_dataframes, valid_csv_files, folder))

    summary = [f"Created one combined plot from {len(completed)} CSV file(s)."]
    if completed:
        summary.extend(["", "Processed:", *completed])
    if skipped:
        summary.extend(["", "Not processed:", *skipped])
    if warnings:
        summary.extend(["", "Static-image warnings:", *warnings])

    if completed:
        messagebox.showinfo("Ag-Sb-As ternary plots", "\n".join(summary))
    else:
        messagebox.showerror("No plots created", "\n".join(summary))
    root.destroy()


if __name__ == "__main__":
    main()
