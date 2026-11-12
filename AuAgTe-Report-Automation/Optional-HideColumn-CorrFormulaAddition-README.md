# Optional Phase A finish — standalone testing version

Run `Optional-HideColumn-CorrFormulaAddition.py` in PyCharm or Python, select the
A2 categorized `.xlsx`/`.xlsm` workbook, then choose which columns to show.
All columns are selected initially; unselected columns are hidden, not deleted.
Canceling the visibility dialog shows all columns. No app integration yet.

## Terminal progress

Timestamped messages are flushed immediately for PyCharm and future QProcess
Activity logging. They report file selection, workbook loading, sheet validation,
per-sheet correction processing, waiting for column selection, visibility,
Summary rebuilding, saving and completion/errors. Correction processing reports
every 1,000 feature rows; validation reports every 10,000 scanned rows. Blocking
stages such as load/save print a heartbeat every 10 seconds with elapsed time.
This is activity reporting, not an estimated percentage complete.

The column-selection dialog is brought to the front and is not attached as a
transient to the withdrawn Tk root, which could keep the dialog invisible. The
terminal explicitly says when user selection is required. Completion also
reports that it is waiting for the final message-box acknowledgement.

## A2 compatibility

The script recognizes the actual AZtec headers retained by A2:
`Feature`, `Area (sq. µm)`, `Breadth (µm)`, and `Length (µm)`. It also supports
ASCII `um` headers and the B2-style second duplicate `Area` column. It does not
require Micro X, Image, microscope folders, or a B2 master workbook.

It processes mineral/feature tables, including Raw Data, Cinnabar,
Cinnabar-Asso and After Size-CutOff. Area, Summary, Integrity Check and
Check Report are administrative sheets; they do not get correction columns.
Other non-feature sheets are preserved. All target tables are validated before
changes are made. Ambiguous/missing physical area or geometry columns produce
a sheet-specific error instead of a misleading partial result.

## Appended columns, in order

`Open Image | Area Orig | Corr % | New Brea | New Len | Asso | Diam`

- Open Image is blank: no image links are created.
- Area Orig copies the physical area, matching B3.
- Corr % is blank and editable, with an Excel validation range of 0–100.
- New Brea: blank/zero Corr % keeps original Breadth; otherwise
  `SQRT(Area Orig * Corr % / 100)`.
- New Len: blank/zero Corr % keeps original Length; otherwise the same
  square-root calculation. This deliberately matches B3's existing behavior.
- Asso is blank and editable.
- Diam is blank and editable: B3 has no diameter formula and none was specified
  for this script. No diameter calculation is invented.

Original data columns/headers are not renamed or moved. Existing source styles
are preserved; new columns get readable widths, number formats and headers.
On rerunning an output made by this script, its seven-column block is reused,
and Corr %, Asso and Diam entries are preserved. A partial/displaced old block
is rejected to avoid duplicating columns or overwriting manual work.

## Summary

Rebuilt with three columns: Category (Sheet name), Total Sum of Area (sq. µm),
and Percentage. Each nonempty mineral sheet appears in workbook order except:

- After Size-CutOff
- Cinnabar
- Cinnabar-Asso

Raw Data and administrative sheets are not categories and are also omitted.
Those three excluded category sheets are retained and still receive the new
columns. Their areas are excluded from both Summary totals and the percentage
denominator.

Area totals are SUM formulas referencing the original unit-bearing physical
area column (the second Area in B2-style headers), not the first Area identifier
and not a new corrected area. Percentages are formulas based on the sum of the
included category areas, displayed with two decimal places. The Total row shows
100.00% when the included total area is positive; all-zero/no-category inputs
show 0.00% without division errors. Display rounding is separate from the full
precision used in the formulas.

All formulas remain in the output. Excel automatic/full recalculation is
requested on opening. The source is never saved or overwritten. Output:
`<input stem>_optional_final.xlsx` (or `.xlsm`), numbered if that path exists.

## Checks

`tests/test_optional_phase_a_finish.py` covers native A2 headers, duplicate
Area columns, exact B3 formulas, exclusions, visibility including grouped
columns, source hashes, preserved styles, idempotent columns/manual entries,
unique output names, standalone GUI flow with mocked dialogs, invalid inputs,
case-insensitive administrative/category names, and actual formula
recalculation using an isolated headless LibreOffice profile. Cached numeric
results verify original/zero/entered correction percentages, area totals,
percentages, escaped sheet names and zero-area handling.
Additional tests verify messages precede loading and dialog waits, load errors
are logged, heartbeats appear during blocked operations, and a mocked hidden
root is not used as the visibility window's transient master. The latter tests
window setup without claiming validation of Windows foreground-focus behavior.

