# AGENTS.md

## Project purpose
This repository contains laboratory automation and data-processing tools for:
- SEM / EDS workflows
- Oxford Instruments AZtec exported Word and data files
- Ag liberation workflows
- Cu liberation workflows
- SEM to uScope file preparation
- supporting utilities for Excel, Word, images, and file handling

## User environment
Assume the primary working environment is:
- Windows
- PyCharm
- Project-specific Python virtual environment
- SEM: JEOL JSM-6010PLUS/LV
- EDS: Oxford Instruments AZtec
- Work domain: mineralogy, liberation studies, metallurgical/mineral processing support

## How to work in this repo
- Read existing scripts in the relevant folder before proposing a new script.
- Reuse working logic from older scripts whenever practical.
- Prefer improving or combining existing scripts over rewriting from scratch.
- Preserve behavior that is already known to work unless the task explicitly asks for redesign.
- When multiple scripts in a folder solve related problems, first summarize what each script does and identify reusable parts.

## Code expectations
- Write complete working Python scripts, not partial snippets.
- Prefer reliable, practical solutions over clever but fragile ones.
- Handle real-world edge cases such as:
  - missing columns
  - variable sheet names
  - user-selected files/folders
  - mixed file naming conventions
  - large numbers of files
  - Windows path issues
- Use clear variable names and comments where logic is not obvious.
- Avoid unnecessary new dependencies unless there is a strong reason.
- Keep scripts easy to run in PyCharm on Windows.

## File and data safety
- Do not modify raw input data unless explicitly requested.
- Prefer writing new output files rather than overwriting originals.
- When changing file structure or output naming, state clearly what will be created.
- Be cautious with destructive operations such as deleting or renaming files in bulk.

## Script behavior preferences
When building or editing scripts for this user:
- Prefer GUI file/folder selection when it helps usability.
- Support Excel and Word workflows commonly used in the lab.
- Assume automation should be practical for repeated day-to-day use.
- If a script depends on an interpreter or environment issue, diagnose that before changing logic.

## Domain-specific guidance

### Ag liberation
Typical tasks may include:
- categorization of Ag minerals
- category selection in Excel
- SEM data and uScope correlation
- classification based on wt% thresholds and associations
- preserving existing category logic unless asked to revise it

When classification rules do not fit perfectly, prefer the closest realistic mineralogical fit instead of forcing a wrong rigid category.

### Cu liberation
Typical tasks may include:
- processing GrainAlyser / GAlyser data
- categorization of Cu-bearing minerals
- liberation report generation
- preserving user-confirmed logic and thresholds

### Word formatting / AZtec exports
Typical tasks may include:
- formatting exported Word files
- arranging BSE, spectra, and maps
- preserving user-specified layout rules
- keeping outputs consistent and repeatable

### SEM to uScope file prep
Typical tasks may include:
- file renaming
- coordinate matching
- folder preparation
- image organization
- preserving filename-based logic when already established in earlier scripts

## Troubleshooting style
When debugging:
1. Identify likely root causes first.
2. Inspect related scripts in the same folder before changing code.
3. Prefer one internally consistent fix over many trial-and-error edits.
4. If environment or interpreter issues explain the failure, fix those before rewriting script logic.

## Output style
When proposing a change:
- briefly explain what the script currently does
- explain what will change
- keep the final code ready to run
- mention any assumptions clearly
- mention any required packages or files only if relevant

## Validation
When possible, validate by:
- syntax checking
- checking imports
- preserving expected inputs/outputs
- using existing script patterns in the repo

## Don’t
- Don’t replace tested domain logic with generic logic without explanation.
- Don’t assume Linux/macOS paths or behavior for this repo.
- Don’t introduce unnecessary architectural complexity.
- Don’t ignore older working scripts in the same folder when creating a new one.