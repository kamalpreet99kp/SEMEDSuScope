# Micro-to-SEM Workflow — MM4

Run `MicroToSEMWorkflow_App_MM4.py` in PyCharm using your Windows project
interpreter. Keep all three MM4 scripts in this directory. The original app and
the supplied original helpers have not been changed.

The interface, five tabs, units, rounding, mark selection, optional mineral
checks, and AutoFeature controls are retained. The app imports only the MM4
coordinate helpers:

- `CreateTSVfromPycroOutput_App_MM4.py`: X adds the ROI centre offset; Y
  subtracts it. Bounding-box Y limits use the same correction and remain ordered.
- `CorrectSEMPositionsFromTemplateV3_MM4.py`: uses the rotation equivalent of
  `Calculation (Locked)!R7 = $T$5+Q7`. The old per-axis shortcuts are removed
  because they give wrong positions under rotation. Exact fiducial B still maps
  to SEM B.

**Excel is not called:** this app performs the transformation through the Python
helper for both grains and mineral-check positions. There is no workbook to edit
for this workflow.

## Outputs and calibration

The MM4 app creates `list_mm4.csv`, `image_list_mm4/`,
`list_sem_corrected_mm4.csv`, and `SEM_mark_positions_mm4.csv` in the same
project subfolders used by the original workflow. Existing old-microscope outputs
are preserved. Rebuilding lists replaces the JPG copies in `image_list_mm4/`;
rerunning corrections replaces the MM4 CSV outputs. Raw images remain unchanged.
The correction helper's standalone CLI defaults to `<input-stem>_sem_corrected_mm4.csv`.

The default calibration is still **0.1182 µm/pixel**, as supplied. Confirm it for
the MM4 objective/camera configuration before using coordinates on hardware.
The crop helper retains its `--um-per-pixel` CLI override and
`PycroTsvSettings(um_per_pixel=...)` API.

## Dependencies and checks

The GUI requires PySide6 (or the original PySide2 fallback). Actual SEM automation
also needs the existing Windows environment, PyAutoGUI, and the custom
`PyWinAuto.read_mpo_values` / `PyWinAuto.mpo_xy_controller` modules on the Python
import path. Those custom modules were not supplied here. Optional notifications
use `requests`. No hardware movements are made by the tests.

From the repository root:

```console
python -m unittest discover -s "MicroToSEM WorkFlow/tests" -v
python "MicroToSEM WorkFlow/MicroToSEMWorkflow_App_MM4.py"
```

Coordinate tests cover ROI signs and bounds, rotation and scale, both SEM unit
choices, fiducials, axis-aligned edge cases, degenerate references, CSV output,
and preservation of old output files. Physical microscope/SEM accuracy still
requires checking known positions on the Windows instrument setup.
