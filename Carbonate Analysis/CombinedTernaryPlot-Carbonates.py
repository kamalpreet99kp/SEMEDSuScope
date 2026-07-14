import copy
import csv
import os
from glob import glob
from pathlib import Path
import tkinter.filedialog

import pandas as pd
import plotly.express as px


def has_header(file):
    with open(file, newline="") as f:
        return csv.Sniffer().has_header(f.read(2048))


def makeAxis(title, tickangle, prefix, mini):
    return {
        "title": title.split(",")[0],
        "min": mini,
        "tickangle": tickangle,
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
        "ticksuffix": " " + prefix,
    }


def format_hover_template(data):
    replacements = {
        "{customdata[0]:.2f}": "{customdata[0]:.2f}%",
        "{customdata[1]:.2f}": "{customdata[1]:.2f}%",
        "{customdata[2]:.2f}": "{customdata[2]:.2f}%",
        "{customdata[3]:.2f}": "{customdata[3]:.2f}g/t",
    }
    result = copy.copy(data)
    for old_text, new_text in replacements.items():
        result = result.replace(old_text, new_text)
    return result


def ternary(data, files, col_names, dir_path):
    # Define the color mapping for each specific layer.
    color_map = {
        "Ankerite": "green",
        "Ankerite STD": "orange",
        "Calcite": "skyblue",
        "Calcite STD": "yellow",
        "Dolomite": "blue",
        "Dolomite STD": "brown",
        "Siderite": "red",
        "Siderite STD": "black",
    }

    max_marker_size = 20
    marker_size_scaling_factor = 0.01

    combined = pd.concat(data, ignore_index=True)
    new_col = []
    for i in range(len(files)):
        new_col.extend([os.path.basename(files[i][:-4])] * len(data[i]))

    combined["Layer"] = new_col

    # Create hover names from the Label column.
    hover_name = combined["Label"]

    hover_data = {
        combined.columns[0]: ":.2f",
        combined.columns[1]: ":.2f",
        combined.columns[2]: ":.2f",
        combined.columns[3]: ":.2f",
    }

    fig = px.scatter_ternary(
        combined,
        a=combined.columns[0],
        b=combined.columns[1],
        c=combined.columns[2],
        color="Layer",
        size=combined.columns[3],
        hover_name=hover_name,
        hover_data=hover_data,
        size_max=max_marker_size,
        color_discrete_map=color_map,
    )

    min_a = 0.0
    min_b = 0.0
    min_c = 0.0

    fig.update_layout(
        title=None,
        showlegend=True,
        ternary={
            "sum": 100,
            "aaxis": makeAxis(col_names[0], 0, "%", min_a),
            "baxis": makeAxis("<br>" + col_names[1], 45, "%", min_b),
            "caxis": makeAxis("<br>" + col_names[2], -45, "%", min_c),
        },
    )

    for d in fig["data"]:
        d["cliponaxis"] = False
        d["marker"]["sizeref"] = marker_size_scaling_factor
        d["hovertemplate"] = format_hover_template(d["hovertemplate"])

    output_stem = Path(dir_path) / "Combined (Ternary Plot)"
    fig.write_html(str(output_stem) + ".html")
    fig.write_image(str(output_stem) + ".png")
    fig.write_image(str(output_stem) + ".svg")


if __name__ == "__main__":
    dir_path = tkinter.filedialog.askdirectory(title="CropoutDir", initialdir="")
    if not dir_path:
        raise SystemExit("No folder selected.")

    files = glob(os.path.join(dir_path, "*.csv"))
    data = []
    col_names = None

    for f in files:
        df = pd.read_csv(f)
        if df.empty:
            continue

        original_col_names = list(df.columns)
        col_names = [
            original_col_names[1],
            original_col_names[2],
            original_col_names[3],
            original_col_names[4],
            original_col_names[0],
        ]

        header = 0 if has_header(f) else None
        df = pd.read_csv(f, header=header, names=original_col_names)
        df = df[col_names]

        if "Labels" in df.columns:
            df.drop(["Labels"], axis=0, inplace=True)

        data.append(df)

    if not data or col_names is None:
        raise SystemExit("No non-empty CSV files were found in the selected folder.")

    ternary(data, files, col_names, dir_path)
