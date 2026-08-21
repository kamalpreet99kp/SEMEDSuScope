# Directory Comparison Tool

`Compare_Directories.py` compares a manually prepared, known-correct directory
with a directory created by an app or another workflow. It never changes either
directory and writes `Directory_Comparison_Report.xlsx` to a separately selected
output directory.

## What it compares

- Complete recursive folder/file inventory
- SHA-256 hashes for exact byte matches
- Excel sheet order, cells, formulas, formatting, hidden rows/columns, merged
  cells, normalized hyperlinks, and embedded images
- Word paragraphs, tables, sections, headers/footers, and embedded media
- Decoded pixels for common image formats
- CSV rows and cells
- Normalized text, HTML, XML, SVG, JSON, logs, and source files
- Binary hashes for other file types

The report distinguishes **Exact Match** from **Content Match**. Content Match
means the meaningful content is the same even though generated metadata,
encoding, or ZIP packaging differs.

## Run in PyCharm on Windows

Run `Compare_Directories.py` and select:

1. The manually prepared reference directory
2. The directory being validated
3. A separate output directory for the report

Required packages are already represented in the repository requirements:
`openpyxl`, `Pillow`, and `python-docx`.

## Command line

```powershell
python ".\Directory-Comparison-Tool\Compare_Directories.py" `
  --reference "C:\Samples\Manual Correct" `
  --comparison "C:\Samples\Created By App" `
  --output "C:\Samples\Comparison Results"
```

Exit code `0` means every file was an Exact Match or Content Match. Exit code
`1` means the report contains an item requiring review.
