import csv
import random
import copy
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import tkinter.filedialog
import os
from glob import glob

def has_header(file):
    with open(file) as f:
        return csv.Sniffer().has_header(f.read(2048))

def makeAxis(title, tickangle, prefix, mini):
    return {
        'title': title.split(",")[0],
        'min': mini,
        'tickangle': tickangle,
        'dtick': 10,
        'tickfont': { 'size': 7 },
        'tickcolor': 'rgba(0,0,0,1)',
        'ticklen': 5,
        'showline': True,
        'linecolor': 'rgba(0,0,0,1)',  # Color of the edge lines
        'linewidth': 1,  # Width of the edge lines
        'showgrid': True,  # This ensures no grid inside the triangle
        'gridcolor': 'rgba(0,0,0,0.5)',  # Making the grid lines a bit transparent for clarity
        'layer': 'below traces', # This will place the grid lines behind the data points
        'ticksuffix': " " + prefix
    }

def format_hover_template(data):
    replacements = {
        '{customdata[0]:.2f}': '{customdata[0]:.2f}%',
        '{customdata[1]:.2f}': '{customdata[1]:.2f}%',
        '{customdata[2]:.2f}': '{customdata[2]:.2f}%',
        '{customdata[3]:.2f}': '{customdata[3]:.2f}g/t'}
    result = copy.copy(data)
    for k, v in replacements.items():
        result = data.replace(k, v)
    return result

def ternary(data, files, col_names):
    # Define the color mapping for each specific layer
    color_map = {
        "Ankerite": "green",
        "Ankerite STD": "orange",
        "Calcite": "skyblue",
        "Calcite STD": "yellow",
        "Dolomite": "blue",
        "Dolomite STD": "brown",
        "Siderite": "red",
        "Siderite STD": "black"
    }

    # The maximum market size.
    max_marker_size = 20

    # The scaling factor used for the markers.
    marker_size_scaling_factor = 0.01

    # Combined Ternary plot
    combined = pd.concat(data)
    new_col = []
    for i in range(len(files)):
        new_col.extend([os.path.basename(files[i][:-4])] * len(data[i]))  # Use basename here

    combined["Layer"] = new_col

    # Create underlined hoover names from the "Label" column.
    hoover_name = combined['Label']

    # Prepare custom hover data.
    hover_data = {
        combined.columns[0]: ':.2f',
        combined.columns[1]: ':.2f',
        combined.columns[2]: ':.2f',
        combined.columns[3]: ':.2f'}

    fig = px.scatter_ternary(
        combined,
        a=combined.columns[0],
        b=combined.columns[1],
        c=combined.columns[2],
        color="Layer",
        size=combined.columns[3],
        hover_name=hoover_name,
        hover_data=hover_data,
        size_max=max_marker_size,
        color_discrete_map=color_map)

    title="Combined"

    min_a = 0.0
    min_b = 0.0
    min_c = 0.0

    fig.update_layout({
        'ternary': {
            'sum': 100,
            'aaxis': makeAxis(col_names[0], 0, '%', min_a),
            'baxis': makeAxis('<br>'+col_names[1], 45, '%', min_b),
            'caxis': makeAxis('<br>'+col_names[2], -45, '%', min_c)
        }
    })

    for d in fig['data']:
        d['cliponaxis'] = False
        d['marker']['sizeref'] = marker_size_scaling_factor
        d['hovertemplate'] = format_hover_template(d['hovertemplate'])
    fig.write_html(dir_path+"/"+title+' (Ternary Plot).html')
    fig.write_image(dir_path+"/"+title+' (Ternary Plot).png')
    fig.write_image(dir_path+"/"+title+' (Ternary Plot).svg')
    fig.update_layout(showlegend=True)
    # fig.show()

    return

if __name__ == "__main__":
    dir_path = tkinter.filedialog.askdirectory(title='CropoutDir', initialdir='')
    files = glob(dir_path + "/*.csv")
    data = []
    for f in files:
        df = pd.read_csv(f)  # read the CSV without passing column names
        if df.empty:
            continue  # if the dataframe is empty, skip it
        original_col_names = list(df.columns)  # extract the column names
        col_names = [original_col_names[1], original_col_names[2], original_col_names[3], original_col_names[4], original_col_names[0]]  # use the original column names, or rearrange them as per your needs
        if has_header(f):
            header = 0
        df = pd.read_csv(f, header=header, names=original_col_names)
        df = df[col_names]
        if "Labels" in df.columns:
            df.drop(["Labels"], axis=0, inplace=True)
        data.append(df)

    ternary(data, files, col_names)
