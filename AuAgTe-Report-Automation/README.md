# Au-Ag-Te Report Automation

`AuAgTe-Report-Automation-App.py` is a Windows desktop workflow manager for the
two-phase Au-Ag-Te SEM/EDS and reflected-light microscope report process. It
uses the current scripts in the local `SEMEDSuScope` checkout, so updated
classification rules remain outside the app.

## Workflow

### Phase A — SEM / AZtec preparation

1. Select an AZtec **Full Analysis** export. The app derives the sample name,
   remembers the directory, and writes a new `_Excel.xlsx` workbook.
2. Select a project-specific categorization script from the automatically
   populated `Ag-Liberation` dropdown and create `Classified_...xlsx`.
3. create a separate `_DUPS_ONLY.xlsx` workbook, then manually remove confirmed
   duplicates from the categorized workbook.
4. Create `SEM-pos-um.xlsx` and `Micro-After-Correction.csv` for transfer to the
   microscope computer.

### Phase B — microscope images / final report

1. Select the transferred parent folder and interactively crop images into the
   category `cropped` subfolders.
2. Select the transferred categorized workbook and create `MASTER_...xlsx` with
   microscope images and a Check Report.
3. Create the final `_with_quick_links.xlsx` workbook with image links,
   correction columns, and formulas.

Every stage can also start independently by selecting an existing input. The
app records its paths in `AuAgTe-Report-Project.json` and provides buttons to
open intermediate files and folders.

## Run in PyCharm on Windows

Use the project virtual environment and install the packages from the repository
requirements, plus PySide6 if it is not already installed:

```powershell
pip install -r requirements.txt
pip install PySide6
python ".\AuAgTe-Report-Automation\AuAgTe-Report-Automation-App.py"
```

The existing step scripts continue to show their sheet, category, crop-centre,
tolerance, and visible-column dialogs. The app only supplies a file or directory
picker automatically when that input is already known.

## Data safety

- The raw Full Analysis export is never modified.
- Each Excel stage writes a new output file.
- The app warns before rerunning steps with fixed output names.
- Script 5 may rename files inside `cropped` folders and creates `_thumbnails`;
  this is existing behavior of that local script.
