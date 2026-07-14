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
- Embedded image temp files include the image role (`microscope` or `sem`) and row number so matching filenames in the two image folders cannot overwrite each other or mix columns.

- `finish_au_report.py` is the Windows/Excel finishing script for pasting normalized chemistry from the `Au` sheet with two-decimal formatting and creating repeated side-by-side report blocks with wrapped headers, marks `Average` from the copied Au chemistry data count, uses the selected intermediate workbook name with `_Inter` changed to `_Final` in the default workbook name, and stops after saving the finished Excel workbook. Word report creation is intentionally kept separate through the `run_au_word_report_macro.py` launcher plus `create_au_word_report_macro.bas` Word VBA macro, so the working Excel workflow is isolated from Word-specific Paste Special behavior.

- `run_au_word_report_macro.py` is the recommended launcher for Word export. It creates a temporary macro-enabled Word host document, imports `create_au_word_report_macro.bas`, and reads the actual VBA module name assigned by Word after importing `create_au_word_report_macro.bas`, saves/activates the temporary macro host, then runs `CreateAuWordReportFromWorkbook` using several Word-supported document-qualified macro name styles and closes the temporary host even if one style fails.
- `validate_existing_au_outputs.py` is a temporary QC helper for already-created outputs. It asks for a Word `.docx` and the three-sheet data workbook, then writes a compact pass/warning/fail QC report next to the Word file.
- The app now shows compact green/yellow/red step statuses for key QC checks, including sample-type chemistry warnings, normalized totals, image folder-name similarity, and final workbook `Organized Blocks` presence.
- `create_au_word_report_macro.bas` is the native Word VBA macro used by the launcher because it runs inside Word VBA and uses Word's native Paste Special command.

- `au_report_automation_app.py` is the initial PySide-based Au Report Automation App. It follows the reference app style with a `QMainWindow`, numbered `QTabWidget` workflow tabs, grouped controls, status labels, and summary tables. The app wires sample setup, Output File 1 creation, `Others` sheet creation, microscope/SEM resize steps, image report workbook creation, Excel report finishing, and Word macro launcher startup.
- The app reads `.full export` / CSV files with common Windows and Unicode encodings (`utf-8-sig`, `cp1252`, `latin-1`, and `utf-16`) because AZtec exports can contain symbols such as `µ` that are not always valid UTF-8.
- The app converts numeric-looking raw export values to real Excel numbers so Excel can calculate formulas and show Average/Sum/Count in the status bar instead of Count only.
- The app now copies the `Area` column into the image report workbook with its own helper, so the image-report step does not depend on a specific `copy_area_column_from_workbook` helper being present in the standalone image insertion script.
- The app also creates the image report workbook with its own local workbook builder, so it does not depend on the argument signature of `create_image_workbook` inside the standalone image insertion script.
- When the app finishes the Excel report, it shows staged progress percentages, saves the combined workbook as `_Final.xlsx`, updates the app state to that final workbook, and attempts to remove the older `_Inter.xlsx` workbook so only the final combined report remains.
- When the app launches Word export, it embeds the known final workbook and selected sample type into the temporary imported Word macro module, so the macro no longer asks for those values when started from the app and avoids unreliable Word COM macro argument passing.
- The Excel finishing and Word export automation use isolated Office COM instances when available, so unrelated Excel/Word files already open for other work should not be touched by the Au workflow.
- The app applies a light professional color theme to tabs, groups, workflow buttons, tables, logs, and progress bars so the four-step process is easier to follow.
- The raw export file picker defaults to showing all files so `.full export` files with nonstandard extensions are visible immediately.
- Step 1 can be resumed from an existing data workbook: selecting the raw export only sets sample state and detects an existing workbook, and the `Select Existing Data Workbook` button can be used when the workbook was already created. Steps 2, 3, and 4 remain streamlined for the normal run-every-time workflow.
- The app includes an AMTEL-branded header that loads the exact local logo image from `AuReportAutomation/assets/amtel_logo.png` (or `.jpg` / `.gif`) when present; it also checks `AuReportAutomation/assests/` to tolerate the common folder-name typo, and shows a workflow checklist for Step 1 through Step 4 readiness.

## Run order

1. Run `insert_au_report_images.py`.
2. Run `finish_au_report.py`.
3. Run `run_au_word_report_macro.py`.

Do not run `create_au_word_report_macro.bas` directly as a normal script; it is a support macro imported automatically by `run_au_word_report_macro.py`.

## App note

The standalone run order remains available, but the app is now the intended direction for the combined workflow. Run `au_report_automation_app.py` to start the GUI once PySide6 or PySide2 is installed.

## Speed note

`finish_au_report.py` disables Excel screen updating, events, alerts, and automatic calculation while it runs when Excel permits those settings, and formats only the current block range during block creation to reduce COM overhead.

## Workbook opening note

`finish_au_report.py` uses a guarded Excel workbook-open helper so if COM opens a workbook but returns `None`, the script tries multiple Excel COM open styles, searches Excel's open workbooks by path/name, checks ActiveWorkbook, and then fails with visible workbook diagnostics.

## Report sheet detection note

`finish_au_report.py` scans the selected intermediate workbook for the worksheet containing `R. Light Images` and `SEM Images` headers instead of assuming the report is always the first worksheet, and it skips any previously generated `Organized Blocks` sheet during that detection.

## Rerun note

`finish_au_report.py` replaces only the previously generated `Organized Blocks` worksheet before creating a fresh one, so rerunning the finishing script on the same workbook will not fail because that generated sheet name already exists. It also reapplies the bold `Average` row formatting inside `Organized Blocks` after standard block formatting is applied.

## Final workbook save note

`finish_au_report.py` checks whether the selected `_Final.xlsx` output workbook is already open in Excel. If a separate old generated final workbook with that path/name is open, the script closes that stale output workbook without saving before writing the fresh final workbook, avoiding Excel's same-name `SaveAs` error.

## Word first-page note

The Word VBA macro places the centered bold `Sample Name:` text in the first-page header instead of the document body. It does not prompt for or insert `Project No.` text, leaving the body of page 1 available for the first pasted block.

## Word Paste Special note

`run_au_word_report_macro.py` is preferred for Word output because it imports and runs `create_au_word_report_macro.bas` automatically. This keeps the copy/paste inside Word VBA, closer to the manual Paste Special workflow used for the reference file, without manually importing the macro each time. The VBA macro no longer inserts manual page breaks between blocks; it adds a normal paragraph separator and lets Word continue naturally so extra blank pages from explicit page breaks are avoided.
