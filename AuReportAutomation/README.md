# Au Report Automation

This folder is reserved for the Au report automation workflow.

Planned scripts in this directory will support gold-focused SEM/EDS report automation while keeping the work separated from the existing Ag, Cu, AZtec, and SEM-to-uScope utilities.

## Notes

- Add all Au report automation scripts and supporting notes here.
- Prefer non-destructive workflows that create new output files rather than modifying raw SEM/EDS inputs.
- Keep scripts practical to run from PyCharm on Windows with the project virtual environment.

## Current scripts

- `insert_au_report_images.py` creates the first Au report workbook for one sample by prompting for the sample type, reflected light image folder, SEM image folder, optional `Area` source workbook, and output `.xlsx` path.
- The script inserts `.jpg` and `.jpeg` files in numeric filename order while keeping row height `45`, column width `8.43`, font size `11`, and the original image files unchanged, and displayed report image size just under `1.59 cm x 1.59 cm` with a small in-cell safety margin.

- `finish_au_report.py` is the Windows/Excel finishing script for pasting normalized chemistry from the `Au` sheet and creating repeated side-by-side report blocks.
