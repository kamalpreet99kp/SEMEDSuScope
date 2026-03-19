# ProcessingGAlyserRawData.py Sheet Guide

Brief guide to the workbook sheets created by `ProcessingGAlyserRawData.py`.

## Sheets

- **Master_Features**
  - All features that have X-ray data and are not undersized.
  - Includes grain/feature measurements, AZtec phase, grouped AZtec phase, corrected phase, and X-ray chemistry.

- **Cpy_Cubanite**
  - Features exclusively routed as chalcopyrite/cubanite after correction and routing.

- **Bornite**
  - Features exclusively routed as bornite after correction and routing.

- **Digenite_Chalcocite**
  - Features exclusively routed as digenite/chalcocite after correction and routing.

- **Covellite**
  - Features exclusively routed as covellite after correction and routing.

- **Gangue**
  - Features routed as gangue when the corrected phase is gangue and the AZtec phase is already gangue-like.

- **Carrolite**
  - Features exclusively routed as carrolite after correction and routing.

- **Sphalerite**
  - Features exclusively routed as sphalerite after correction and routing.

- **Molybdenite**
  - Features exclusively routed as molybdenite after correction and routing.

- **Py_Po**
  - Features exclusively routed as pyrite/pyrrhotite after correction and routing.

- **Mismatched**
  - Review sheet for features with X-ray data that do not land in an accepted matched mineral sheet and are not resolved elsewhere.

- **Mixed_Cu_Sulfide**
  - Review sheet for Cu-sulfide features that remain mixed after first-pass classification and were not resolved in the normalized second pass.

- **Resolved_From_Mixed**
  - Features first labelled as `Mixed_Cu_Sulfide` and then resolved in the normalized Cu-Fe-S second pass.
  - Includes the normalized Cu, Fe, S, Cu/S ratio, resolved phase, and resolution source columns.

- **Resolved_From_Mismatched**
  - Features sent to mismatch review and then resolved in the normalized Cu-Fe-S second pass.
  - Includes the normalized Cu, Fe, S, Cu/S ratio, resolved phase, and resolution source columns.

- **No_XRay**
  - Features with no usable X-ray chemistry values.

- **Undersized**
  - Features below the fraction-specific `Feature_ECD` undersized threshold.

- **All_Checks**
  - Summary/QC sheet with detected fraction, undersized threshold, duplicate-key checks, merged row counts, and row counts for each major output sheet.

- **Random_QC_50**
  - Random sample of up to 50 rows from `Master_Features` for manual QC.

## Notes

- Sheets are written only when they contain rows.
- Each feature is routed to one exclusive non-master destination sheet.
- `Resolved_From_Mixed` and `Resolved_From_Mismatched` include extra normalized-resolution columns beyond the standard output columns.
