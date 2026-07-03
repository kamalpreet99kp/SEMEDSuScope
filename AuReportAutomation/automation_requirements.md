# Au Report Automation Requirements

This document captures the initial requirements for the Au report automation workflow before script or VBA generation begins.

## Initial image insertion workflow

- The user has an existing macro-enabled Excel workbook with fixed layout dimensions:
  - Row height: `45`
  - Column width: `8.43`
- The workflow should prompt the user to select two image directories:
  1. Reflected light image directory
  2. SEM image directory
- Images may be `.jpg` or `.jpeg`.
- Images are named with leading numeric identifiers such as `001`, `002`, and so on, potentially up to `1000` or more.
- Images must be inserted in numeric order from the lowest identifier to the highest identifier.
- The first image row starts at row `2` because row `1` contains headers.
- Column `A` should automatically be populated with sequential numbers from `1` to the number of image rows.

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
- The existing row height and column width should remain fixed at row height `45` and column width `8.43`.

## Later normalized chemistry import workflow

After the user manually edits the intermediate output by deleting some rows, columns, or images:

- The workflow should return to the same selected Excel data file.
- It should open the sheet named `Au`.
- It should locate the ending data area containing columns such as `Normalized`, `Au`, `Ag`, `Cu`, and `Hg`, depending on sample type.
- It should copy the relevant chemistry columns and paste them after column `A` and before `R. Light Images` in the final report layout.

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
- Each complete side-by-side block should receive thick outside borders.
- If a keyboard shortcut is used for this final organization macro, the requested shortcut is `Ctrl+Shift+K`.

## Implementation decisions confirmed

- The automation should be script-based so it can keep growing as more steps are added.
- Each run should create a new final Excel workbook for one sample instead of modifying the original input/template workbook.
- Cell dimensions should remain fixed at row height `45` and column width `8.43`.
- Image folder count mismatches can be fixed manually later; the script should paste all images found in the selected folders.
- Filename sorting should be numeric, so a file beginning with `001` is placed before a file beginning with `1000`.
- Implementation should proceed one step at a time, beginning with image insertion.

## Open questions before later workflow steps

1. For the `Area` import, should the script copy only the used rows matching the inserted image count, or copy the full used `Area` column from the `Others` sheet?
2. For the later `Au` sheet import, how should the script identify the correct final `Normalized` chemistry block if the sheet contains multiple tables or repeated headers?
3. After manual cleanup, should the final block organization preserve the current visible row order exactly, or should it renumber/re-sort based on the `No.` column?
