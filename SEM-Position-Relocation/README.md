# SEM Position Relocation Calculator

This GUI recalculates SEM stage positions after a sample has been removed and
remounted. It uses the old and new coordinates of the same two recognizable
reference positions, A and B.

## Workflow

1. Run `SEM_Position_Relocation_Calculator.py` in the project-specific Python
   environment (for example, from PyCharm on Windows).
2. Enter the **old SEM X/Y** and newly measured **new SEM X/Y** for references
   A and B.
3. Enter positions of interest manually, add/remove rows, or import an `.xlsx`
   or `.csv` table. Recognized coordinate headings include `Old X`/`Old Y`,
   `X`/`Y`, and `Stage X`/`Stage Y`.
4. Select **Calculate / Recalculate**. All input cells remain editable, so A,
   B, or any position can be corrected and calculated again.
5. Save a new `.xlsx` or `.csv` result. The imported source is never modified.

The Excel output contains both a `Calculated Positions` sheet and a
`Reference Inputs` sheet so the result remains traceable.

## Calculation mode

The default is a rigid two-dimensional transformation: translation plus
rotation. This is normally appropriate when returning the same sample to the
same SEM stage because physical distances have not changed.

The optional scale checkbox additionally makes the measured old A–B distance
match the new A–B distance. Leave it off unless a scale correction is known to
be necessary. The status line reports rotation, A–B distance difference, and
the scale used.

## Requirements

- Python 3.9 or newer
- Tkinter (included with normal Windows Python installations)
- `openpyxl` for Excel import/export (already listed in the repository
  `requirements.txt`)

CSV entry and output use only the Python standard library.
