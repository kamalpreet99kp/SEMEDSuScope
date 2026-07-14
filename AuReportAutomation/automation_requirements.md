# Au Report Automation Requirements

This document captures the initial requirements for the Au report automation workflow before script or VBA generation begins.

## Initial image insertion workflow

- The user has an existing macro-enabled Excel workbook with fixed layout dimensions:
  - Row height: `45`
  - Column width: `8.43`
  - Font size: `11`, with only headers bold
  - Original image files should not be modified; embedded workbook copies may be compressed/resized to control final workbook size
  - Displayed/embedded report image size should be height `1.58 cm`, width `1.58 cm`, rotation `0`
- The workflow should prompt the user to select two image directories:
  1. Reflected light image directory
  2. SEM image directory
- Images may be `.jpg` or `.jpeg`.
- Images are named with leading numeric identifiers such as `001`, `002`, and so on, potentially up to `1000` or more.
- Images must be inserted in numeric order from the lowest identifier to the highest identifier.
- Reflected light and SEM images must be kept independent even when the two folders contain identical filenames such as `013-resized.jpg`; temporary embedded copies must include image role and row information to prevent overwrites.
- The first image row starts at row `2` because row `1` contains headers.
- Column `A` should automatically be populated with sequential numbers from `1` to the number of image rows.
- The intermediate workbook filename should use the selected Excel data file name when available and include `Inter`.

## Sample type layouts

The column order should follow this general pattern:

`No.`, `Au (Wt%)`, `Ag (Wt%)`, optional `Cu (Wt%)`, optional `Hg (Wt%)`, `R. Light Images`, `SEM Images`, later followed by `Area`.

| Sample type | Reflected light images | SEM images | Starting headers |
| --- | --- | --- | --- |
| `Au+Ag` | Column `D` | Column `E` | `A: No.`, `B: Au (Wt%)`, `C: Ag (Wt%)`, `D: R. Light Images`, `E: SEM Images` |
| `Au+Ag+Cu` | Column `E` | Column `F` | `A: No.`, `B: Au (Wt%)`, `C: Ag (Wt%)`, `D: Cu (Wt%)`, `E: R. Light Images`, `F: SEM Images` |
| `Au+Ag+Cu+Hg` | Column `F` | Column `G` | `A: No.`, `B: Au (Wt%)`, `C: Ag (Wt%)`, `D: Cu (Wt%)`, `E: Hg (Wt%)`, `F: R. Light Images`, `G: SEM Images` |
| `Au+Ag+Hg` | Column `E` | Column `F` | `A: No.`, `B: Au (Wt%)`, `C: Ag (Wt%)`, `D: Hg (Wt%)`, `E: R. Light Images`, `F: SEM Images` |

## Area column import workflow

- The workflow should prompt the user to select an Excel data file.
- From that workbook, the script or macro should open the sheet named `Others`.
- It should find/copy the first column with header `Area`.
- The copied `Area` column should be pasted into the output workbook in the column immediately after `SEM Images`.
- Blank yellow cells in the `Area` column must remain blank and preserve their yellow fill formatting in the final file.
- The existing row height, column width, and font rule should remain fixed at row height `45`, column width `8.43`, and font size `11` with only headers bold.

## Later normalized chemistry import workflow

After the user manually edits the intermediate output by deleting some rows, columns, or images:

- The workflow should return to the same selected Excel data file.
- It should open the sheet named `Au`.
- It should locate the ending data area containing columns such as `Normalized`, `Au`, `Ag`, `Cu`, and `Hg`, depending on sample type.
- It should copy the relevant chemistry columns and paste them after column `A` and before `R. Light Images` in the final report layout.
- Pasted chemistry values should be displayed with exactly two decimal places.
- The final chemistry row copied from the `Au` sheet should replace its report `No.` value with `Average`; that row should be bold from column `A` through the column immediately before `R. Light Images`.
- The finished workbook filename should use the selected intermediate workbook name and replace `_Inter` with `_Final` when present.
- The finishing script should then create an organized `Organized Blocks` worksheet with repeated side-by-side blocks.

## Final block organization workflow

- After the `Area` column workflow and manual user cleanup, the report should be reorganized into repeated side-by-side blocks.
- Standard sample types use one block size of `26` items:
  - Left side: `1` through `13`
  - Right side: `14` through `26`
- The `Au+Ag+Cu+Hg` sample type uses one block size of `20` items:
  - Left side: `1` through `10`
  - Right side: `11` through `20`
- If there are more grains than one block can hold, the script should create additional blocks underneath the first block while preserving the sequence of all data and images.
- There should be one blank row between repeated blocks, but that blank row applies only to the final organized area after the `Area` column step.
- Block headers should be wrapped, using labels such as `Au\n(Wt%)`, `Ag\n(Wt%)`, `Cu\n(Wt%)`, `Hg\n(Wt%)`, `R. Light\nImages`, and `SEM\nImages`.
- Only the final row of the final set should show `Average` in the first `No.` column instead of the last sequential number; do not add an `Average` row to every block.
- The source `Average` row created from the Au chemistry data count should be copied naturally into the final block during block organization.
- Each complete side-by-side block should receive thick outside borders.
- If a keyboard shortcut is used for this final organization macro, the requested shortcut is `Ctrl+Shift+K`.

## Word report export workflow

- After the `Organized Blocks` sheet is created by `finish_au_report.py`, `run_au_word_report_macro.py` should import/run the Word VBA macro automatically and create a Word document with the same base name as the final workbook.
- The Word export should not prompt for `Project No.` and should not add `Project No.` text, so the first block has more room on page 1.
- Only the first Word page should include centered bold `Sample Name: xxxx` text, placed in the first-page header rather than the document body so the first block has maximum body space; `xxxx` is based on the selected Excel data filename with `(Au SEM)` removed if present.
- Word orientation should be portrait for `Au+Ag`, `Au+Ag+Cu`, and `Au+Ag+Hg`; landscape for `Au+Ag+Cu+Hg`.
- Each page should copy the Excel block and paste it using Word Paste Special as an editable `Microsoft Excel Worksheet Object`; this Word-specific behavior belongs in `create_au_word_report_macro.bas` plus `run_au_word_report_macro.py` so it can be tested/changed independently from the working Excel finishing script.
- Manual page breaks should not be inserted by default because they can create blank pages with embedded Excel objects; the Word macro should let blocks continue naturally with a normal separator between pasted objects.
- `finish_au_report.py` should reduce Excel COM overhead by disabling screen updating/events/alerts/automatic calculation during processing when Excel permits those settings and by avoiding repeated formatting of already completed blocks.

## Implementation decisions confirmed

- The automation should be script-based so it can keep growing as more steps are added.
- Each run should create a new final Excel workbook for one sample instead of modifying the original input/template workbook.
- Cell dimensions should remain fixed at row height `45` and column width `8.43`.
- Original image files should not be modified; embedded workbook copies may be compressed/resized to control final workbook size because that reduces clarity; images should be inserted from the original files and displayed at height `1.58 cm`, width `1.58 cm`, rotation `0`.
- Text should use font size `11`; only header text should be bold.
- Image folder count mismatches can be fixed manually later; the script should paste all images found in the selected folders.
- Filename sorting should be numeric, so a file beginning with `001` is placed before a file beginning with `1000`.
- Implementation should proceed one step at a time, beginning with image insertion.

## Open questions before later workflow steps

1. For the `Area` import, should the script copy only the used rows matching the inserted image count, or copy the full used `Area` column from the `Others` sheet?
2. For the later `Au` sheet import, how should the script identify the correct final `Normalized` chemistry block if the sheet contains multiple tables or repeated headers?
3. After manual cleanup, should the final block organization preserve the current visible row order exactly, or should it renumber/re-sort based on the `No.` column?


## Quality-control checks

- Step status messages should stay compact: show a green pass icon when checks pass; show only the warning/failure issue in yellow/red when attention is required.
- If a lower-complexity sample type is selected, unselected `Cu (Wt%)` or `Hg (Wt%)` values greater than `2` should fail the sample-type check and suggest the higher sample type.
- Normalized totals for each data row must not be below `85`; normalized average chemistry for the selected elements must sum to `100`.
- Step 2 should clearly indicate microscope resize completion before SEM resize, then confirm SEM resize/folder-name checks.
- Image folder-name checks should compare alphabetic text in the selected image folders against the raw export name, ignoring numbers.
- Final workbook checks should confirm the `Organized Blocks` sheet exists.
- The Word export macro should clear Excel copy mode and avoid extra success popups; the app provides an Open Word Report button.

## App workflow requirements

- The combined app should be named `Au Report Automation App`.
- The app should follow the PySide reference app style: `QMainWindow`, numbered workflow tabs, grouped controls, status labels, summary tables, and button-driven manual pause points.
- The first app implementation should create Output File 1 from the `.full export` file, create the `Others` sheet after manual `Au` edits, and resize microscope/SEM images into `resizedtosmallest` folders.
- The app should write numeric-looking raw export values as real Excel numbers, then add `Average` formulas for `Normalized` and sample chemistry columns and a `Minimum` formula for `Normalized`.
- The app should run the image report and Excel finishing steps from app state without re-asking for files/folders already selected earlier.
- The app should save the finished report as a single `_Final.xlsx` combined workbook, remove the previous `_Inter.xlsx` when possible, and show staged progress percentages while Excel COM finishing is running.
- The app should pass the known final workbook and selected sample type into the Word macro flow by embedding them into the temporary macro module, so Word export does not ask again for the workbook or orientation/sample type.
- Office COM automation should prefer isolated Excel/Word instances so unrelated workbooks/documents already open on the computer are not affected by the Au workflow.
- The app should use a light professional color theme for tabs, workflow groups, and primary/manual/Word buttons.
- The raw export picker should make `.full export` files visible without requiring the user to switch to `All Files`.
- Step 1 should support resuming from an existing data workbook: selecting the raw export should detect the existing workbook when possible, and the app should provide one `Select Existing Data Workbook` button. Step 2 should stay simple with only the normal resize buttons because it is expected to be run each time.
- The app should include a professional AMTEL-branded header and workflow checklist showing Step 1 through Step 4 readiness.
