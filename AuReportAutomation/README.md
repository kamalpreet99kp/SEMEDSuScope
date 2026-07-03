# Au Report Automation

This folder is reserved for the Au report automation workflow.

Planned scripts in this directory will support gold-focused SEM/EDS report automation while keeping the work separated from the existing Ag, Cu, AZtec, and SEM-to-uScope utilities.

## Notes

- Add all Au report automation scripts and supporting notes here.
- Prefer non-destructive workflows that create new output files rather than modifying raw SEM/EDS inputs.
- Keep scripts practical to run from PyCharm on Windows with the project virtual environment.

## Current scripts

- `insert_au_report_images.py` creates the first Au report workbook for one sample by prompting for the sample type, reflected light image folder, SEM image folder, optional `Area` source workbook, and output `.xlsx` path, and uses `Inter` in the default intermediate workbook name.
- The script inserts `.jpg` and `.jpeg` files in numeric filename order while keeping row height `45`, column width `8.43`, font size `11`, and the original image files unchanged, compressing embedded copies to control workbook size, and displayed report image size `1.58 cm x 1.58 cm`.

- `finish_au_report.py` is the Windows/Excel finishing script for pasting normalized chemistry from the `Au` sheet with two-decimal formatting and creating repeated side-by-side report blocks with wrapped headers, marks `Average` from the copied Au chemistry data count, uses the selected intermediate workbook name with `_Inter` changed to `_Final` in the default workbook name, and stops after saving the finished Excel workbook. Word report creation is intentionally kept in the separate `create_au_word_report.py` script so the working Excel workflow is isolated from Word-specific Paste Special behavior.

- `create_au_word_report.py` is the optional Word-only export script. Run it after `finish_au_report.py`; it opens the finished workbook, reads the `Organized Blocks` worksheet, detects the block ranges separated by blank rows, asks for project/sample orientation details, and creates a `.docx` by copying each Excel block and using Word Paste Special OLE object.

## Speed note

`finish_au_report.py` disables Excel screen updating, events, alerts, and automatic calculation while it runs when Excel permits those settings, and formats only the current block range during block creation to reduce COM overhead.

## Workbook opening note

`finish_au_report.py` uses a guarded Excel workbook-open helper so if COM opens a workbook but returns `None`, the script tries multiple Excel COM open styles, searches Excel's open workbooks by path/name, checks ActiveWorkbook, and then fails with visible workbook diagnostics.

## Report sheet detection note

`finish_au_report.py` scans the selected intermediate workbook for the worksheet containing `R. Light Images` and `SEM Images` headers instead of assuming the report is always the first worksheet, and it skips any previously generated `Organized Blocks` sheet during that detection.

## Rerun note

`finish_au_report.py` replaces only the previously generated `Organized Blocks` worksheet before creating a fresh one, so rerunning the finishing script on the same workbook will not fail because that generated sheet name already exists.

## Final workbook save note

`finish_au_report.py` checks whether the selected `_Final.xlsx` output workbook is already open in Excel. If a separate old generated final workbook with that path/name is open, the script closes that stale output workbook without saving before writing the fresh final workbook, avoiding Excel's same-name `SaveAs` error.

## Word Paste Special note

`create_au_word_report.py` uses Word Paste Special OLE object for each Excel block. The Excel finishing script no longer creates Word files, so Word-specific troubleshooting can be done separately from the working Excel report generation.
