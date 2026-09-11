# Au Analysis documentation

This directory contains the controlled-document draft for the analysis stage of
the Au workflow.

## Why the Word file is stored as Base64 text

The pull-request interface used for this repository does not support binary
files. The Word document is therefore stored losslessly as the text file:

- `Au_Analysis_Micro-to-SEM_Workflow_Guide_DRAFT.docx.b64`

This is the complete Word document encoded as Base64, not a shortened copy.

## Restore the Word document on Windows 10

1. Download this directory, or download these two files into the same folder:
   - `Au_Analysis_Micro-to-SEM_Workflow_Guide_DRAFT.docx.b64`
   - `Restore-Word-Document.ps1`
2. Right-click `Restore-Word-Document.ps1` and select **Run with PowerShell**.
3. If Windows blocks the script, open PowerShell in that folder and run:

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\Restore-Word-Document.ps1
   ```

4. The script creates the editable document in the same folder:

   `Au_Analysis_Micro-to-SEM_Workflow_Guide_DRAFT.docx`

5. Open the resulting `.docx` file in Microsoft Word.

The restored document is 52,554 bytes and has SHA-256 checksum:

`74a069bb61c91092e108775aff4a90279aedcfb1f609389a67eb21881769f807`

## Basis and scope

The draft records the equipment and software information supplied by the
operator: Windows 10, JEOL JSM-6010PLUS/LV, Oxford Instruments AZtec 4.4,
X-Max^N 80 detector, and the AZtec Feature workflow. It covers preparation of
micro-to-SEM positions, two-point coordinate correction, position checking, and
the AutoFeature3.0 queue. The later reporting stage is explicitly out of scope.

The document distinguishes operator instructions from the underlying technical
reference. Yellow `[COMPLETE ...]` fields identify company-specific or
operational details that require the operator's final review. The document does
not replace equipment training, laboratory safety procedures, or the Oxford
Instruments and JEOL operating manuals.

## Source information still required for final approval

- The unformatted local `MicrotoSEMWorkflow_App.py` and its exact revision.
- `CorrectSEMPositionsFromTemplateV3.py`.
- `CreateTSVfromPycroOutput_App.py`.
- The `PyWinAuto` helper modules used to read and set MPO Test stage values.
- Approved screenshots of the five application tabs.
- An anonymized known-good project directory and expected outputs.
- Company document number, owner, approver, retention location, and review date.
