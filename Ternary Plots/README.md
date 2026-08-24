# General ternary plot tool

`General Ternary Plot Script.py` is the reusable version of the Ag-Sb-As ternary plot
workflow. It creates a separate plot for each selected CSV file or each worksheet
in a selected Excel workbook without changing the source data.

## Use in PyCharm on Windows

1. Install the packages from the repository's `requirements.txt` in the project's
   Python virtual environment.
2. Run `General Ternary Plot Script.py`.
3. Enter the spreadsheet headers for the **top**, **left**, and **right** corners.
   The initial values are `As`, `S`, and `Fe`, respectively.
4. Enter the label and marker-size headers. Clear either box if the file does not
   contain that column; row numbers or equal-sized markers will then be used.
5. Select one or more `.csv`, `.xlsx`, or `.xlsm` data files.

Headers are matched case-insensitively and leading/trailing spaces are ignored.
Each component and the marker size must be numeric. Rows containing missing or
negative component values, a zero total, or a non-positive marker size are
skipped and reported in the completion message. Ternary component values are
proportional, so they do not need to total exactly 100.

## Outputs

For an input named `Sample.xlsx`, the tool writes `Sample (Ternary Plot).html`,
`.png`, and `.svg` beside the workbook. For a workbook with multiple worksheets,
the worksheet name is included in each output filename. HTML output does not
require Kaleido; if static-image export is unavailable, the HTML is still created
and the PNG/SVG problem is reported. Existing raw input files are never modified.
