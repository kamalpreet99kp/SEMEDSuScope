"""Create configurable ternary plots from CSV or Excel data files.

Run this script in PyCharm (or with Python), choose the three elements and their
triangle positions, then select one or more input files.  Inputs are never
changed.  An interactive HTML plot and, when Kaleido is available, PNG and SVG
copies are written beside each input file.
"""

import colorsys
import hashlib
import re
from dataclasses import dataclass
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd
import plotly.express as px


SUPPORTED_EXTENSIONS = {".csv", ".xlsx", ".xlsm"}
MAX_MARKER_SIZE = 20
MARKER_SIZE_SCALING_FACTOR = 0.01


@dataclass(frozen=True)
class PlotSettings:
    """User-selected columns and display settings."""

    top: str
    left: str
    right: str
    label: str = "Label"
    size: str = "Size"

    @property
    def elements(self) -> tuple[str, str, str]:
        return self.top, self.left, self.right


def normalized_header(value: object) -> str:
    """Normalize a header so capitalization and surrounding spaces do not matter."""
    return str(value).strip().casefold()


def validate_settings(settings: PlotSettings) -> None:
    """Raise a useful error before files are processed."""
    if any(not element.strip() for element in settings.elements):
        raise ValueError("Top, left, and right element headers are all required.")
    if len({normalized_header(element) for element in settings.elements}) != 3:
        raise ValueError("Choose three different element headers.")


def make_axis(title: str, tick_angle: int) -> dict:
    """Return the formatting shared by all three ternary axes."""
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


def color_for_name(name: str) -> str:
    """Return a bright, repeatable color for a file or worksheet name."""
    digest = hashlib.sha256(name.casefold().encode("utf-8")).digest()
    hue = int.from_bytes(digest[:2], "big") / 65535
    red, green, blue = colorsys.hsv_to_rgb(hue, 0.70, 0.80)
    return f"rgb({round(red * 255)}, {round(green * 255)}, {round(blue * 255)})"


def prepare_dataframe(
    dataframe: pd.DataFrame, settings: PlotSettings
) -> tuple[pd.DataFrame, int]:
    """Match selected headers and remove rows that cannot be plotted."""
    header_lookup = {normalized_header(column): column for column in dataframe.columns}
    missing = [
        element for element in settings.elements if normalized_header(element) not in header_lookup
    ]
    if missing:
        raise ValueError("missing selected element column(s): " + ", ".join(missing))

    prepared = pd.DataFrame()
    for position, element in zip(("Top", "Left", "Right"), settings.elements):
        prepared[position] = pd.to_numeric(
            dataframe[header_lookup[normalized_header(element)]], errors="coerce"
        )

    label_column = header_lookup.get(normalized_header(settings.label)) if settings.label.strip() else None
    if settings.label.strip() and label_column is None:
        raise ValueError(
            f"missing label column: {settings.label} (clear the label setting to use row numbers)"
        )
    if label_column is None:
        prepared["Label"] = [f"Row {index + 2}" for index in range(len(dataframe))]
    else:
        prepared["Label"] = dataframe[label_column].fillna("").astype(str)

    size_column = header_lookup.get(normalized_header(settings.size)) if settings.size.strip() else None
    if settings.size.strip() and size_column is None:
        raise ValueError(
            f"missing marker size column: {settings.size} "
            "(clear the size setting to use equal-sized markers)"
        )
    if size_column is None:
        prepared["Size"] = 1.0
    else:
        prepared["Size"] = pd.to_numeric(dataframe[size_column], errors="coerce")

    original_count = len(prepared)
    component_sum = prepared[["Top", "Left", "Right"]].sum(axis=1, min_count=3)
    valid = (
        prepared[["Top", "Left", "Right", "Size"]].notna().all(axis=1)
        & (prepared[["Top", "Left", "Right"]] >= 0).all(axis=1)
        & (component_sum > 0)
        & (prepared["Size"] > 0)
    )
    prepared = prepared.loc[valid].copy()
    return prepared, original_count - len(prepared)


def read_datasets(input_file: Path) -> list[tuple[str, pd.DataFrame]]:
    """Read a CSV or every worksheet in an Excel workbook."""
    if input_file.suffix.casefold() == ".csv":
        return [(input_file.stem, pd.read_csv(input_file, encoding="utf-8-sig"))]
    if input_file.suffix.casefold() in {".xlsx", ".xlsm"}:
        sheets = pd.read_excel(input_file, sheet_name=None)
        return [(str(name), frame) for name, frame in sheets.items()]
    raise ValueError(f"unsupported file type: {input_file.suffix}")


def safe_filename_part(value: str) -> str:
    """Make worksheet names safe for Windows output filenames."""
    cleaned = re.sub(r'[<>:"/\\|?*]', "_", value).strip().rstrip(".")
    return cleaned or "Sheet"


def create_ternary_plot(
    dataframe: pd.DataFrame,
    settings: PlotSettings,
    layer_name: str,
    output_stem: Path,
) -> list[str]:
    """Write an interactive plot and attempt both static image formats."""
    plot_data = dataframe.copy()
    plot_data["Layer"] = layer_name
    figure = px.scatter_ternary(
        plot_data,
        a="Top",
        b="Left",
        c="Right",
        color="Layer",
        size="Size",
        hover_name="Label",
        custom_data=["Top", "Left", "Right", "Size"],
        size_max=MAX_MARKER_SIZE,
        color_discrete_map={layer_name: color_for_name(layer_name)},
    )
    figure.update_layout(
        title=None,
        showlegend=True,
        ternary={
            "sum": 100,
            "aaxis": make_axis(settings.top, 0),
            "baxis": make_axis(settings.left, 45),
            "caxis": make_axis(settings.right, -45),
        },
    )
    for trace in figure.data:
        trace.cliponaxis = False
        trace.marker.sizeref = MARKER_SIZE_SCALING_FACTOR
        trace.hovertemplate = (
            "<b>%{hovertext}</b><br>"
            f"{settings.top}: %{{customdata[0]:.2f}} wt%<br>"
            f"{settings.left}: %{{customdata[1]:.2f}} wt%<br>"
            f"{settings.right}: %{{customdata[2]:.2f}} wt%<br>"
            "Size: %{customdata[3]:.2f}<extra>%{fullData.name}</extra>"
        )

    figure.write_html(str(output_stem) + ".html")
    warnings = []
    for extension in (".png", ".svg"):
        try:
            figure.write_image(str(output_stem) + extension)
        except Exception as error:
            warnings.append(f"Could not create {output_stem.name}{extension}: {error}")
    return warnings


class SettingsDialog(tk.Toplevel):
    """Small modal dialog for choosing headers and triangle positions."""

    def __init__(self, parent: tk.Tk) -> None:
        super().__init__(parent)
        self.title("General ternary plot settings")
        self.resizable(False, False)
        self.result: PlotSettings | None = None
        defaults = (("Top element header", "As"), ("Left element header", "S"),
                    ("Right element header", "Fe"), ("Label header (optional)", "Label"),
                    ("Marker size header (optional)", "Size"))
        self.entries = []
        for row, (label, default) in enumerate(defaults):
            ttk.Label(self, text=label).grid(row=row, column=0, padx=10, pady=5, sticky="w")
            entry = ttk.Entry(self, width=28)
            entry.insert(0, default)
            entry.grid(row=row, column=1, padx=10, pady=5)
            self.entries.append(entry)
        ttk.Label(
            self,
            text="Headers are matched without regard to capitalization or outer spaces.",
        ).grid(row=5, column=0, columnspan=2, padx=10, pady=(5, 10))
        buttons = ttk.Frame(self)
        buttons.grid(row=6, column=0, columnspan=2, pady=(0, 10))
        ttk.Button(buttons, text="Continue", command=self.accept).pack(side="left", padx=5)
        ttk.Button(buttons, text="Cancel", command=self.destroy).pack(side="left", padx=5)
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.transient(parent)
        self.grab_set()
        self.entries[0].focus_set()

    def accept(self) -> None:
        values = [entry.get().strip() for entry in self.entries]
        settings = PlotSettings(*values)
        try:
            validate_settings(settings)
        except ValueError as error:
            messagebox.showerror("Invalid settings", str(error), parent=self)
            return
        self.result = settings
        self.destroy()


def output_stem_for(input_file: Path, sheet_name: str, multiple_sheets: bool) -> Path:
    """Build a non-destructive and unambiguous output name."""
    sheet_suffix = f" - {safe_filename_part(sheet_name)}" if multiple_sheets else ""
    return input_file.with_name(f"{input_file.stem}{sheet_suffix} (Ternary Plot)")


def main() -> None:
    root = tk.Tk()
    root.withdraw()
    dialog = SettingsDialog(root)
    root.wait_window(dialog)
    if dialog.result is None:
        root.destroy()
        return

    selected_files = filedialog.askopenfilenames(
        parent=root,
        title="Select CSV or Excel data files",
        filetypes=[
            ("CSV and Excel files", "*.csv *.xlsx *.xlsm"),
            ("CSV files", "*.csv"),
            ("Excel workbooks", "*.xlsx *.xlsm"),
        ],
    )
    if not selected_files:
        root.destroy()
        return

    completed, skipped, warnings = [], [], []
    for filename in selected_files:
        input_file = Path(filename)
        if input_file.suffix.casefold() not in SUPPORTED_EXTENSIONS:
            skipped.append(f"{input_file.name}: unsupported file type")
            continue
        try:
            datasets = read_datasets(input_file)
        except Exception as error:
            skipped.append(f"{input_file.name}: could not read file ({error})")
            continue
        multiple_sheets = len(datasets) > 1
        for sheet_name, raw_data in datasets:
            dataset_name = input_file.name if not multiple_sheets else f"{input_file.name} [{sheet_name}]"
            try:
                data, skipped_rows = prepare_dataframe(raw_data, dialog.result)
                if data.empty:
                    skipped.append(f"{dataset_name}: no valid data rows")
                    continue
                output_stem = output_stem_for(input_file, sheet_name, multiple_sheets)
                warnings.extend(
                    create_ternary_plot(data, dialog.result, dataset_name, output_stem)
                )
                completed.append(f"{dataset_name} ({len(data)} rows; {skipped_rows} skipped)")
            except Exception as error:
                skipped.append(f"{dataset_name}: {error}")

    summary = [f"Created plots for {len(completed)} dataset(s)."]
    if completed:
        summary.extend(["", "Processed:", *completed])
    if skipped:
        summary.extend(["", "Not processed:", *skipped])
    if warnings:
        summary.extend(["", "Static-image warnings:", *warnings])
    message = "\n".join(summary)
    if completed:
        messagebox.showinfo("General ternary plots", message, parent=root)
    else:
        messagebox.showerror("No plots created", message, parent=root)
    root.destroy()


if __name__ == "__main__":
    main()
