"""Compare two sample/project directories and create a concise Excel report.

The reference directory is the manually prepared, known-correct version. The
comparison directory is the version produced by an app or another workflow.
Neither directory is modified. Run without arguments for Windows folder-picker
dialogs, or supply --reference, --comparison, and optionally --output.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import re
from collections import Counter
from dataclasses import dataclass, field
from datetime import date, datetime
from io import BytesIO
from pathlib import Path
from tkinter import Tk, filedialog, messagebox
from typing import Any, Callable
from zipfile import ZipFile

from docx import Document
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from PIL import Image, ImageOps


APP_TITLE = "Directory Content Comparison"
REPORT_NAME = "Directory_Comparison_Report.xlsx"
HASH_CHUNK_SIZE = 1024 * 1024
MAX_DETAILS_PER_FILE = 100
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".webp"}
EXCEL_EXTENSIONS = {".xlsx", ".xlsm"}
TEXT_EXTENSIONS = {
    ".txt", ".log", ".md", ".json", ".xml", ".svg", ".html", ".htm",
    ".py", ".bas", ".tsv",
}
WORD_EXTENSIONS = {".docx"}
CSV_EXTENSIONS = {".csv"}

STATUS_EXACT = "Exact Match"
STATUS_CONTENT = "Content Match"
STATUS_DIFFERENT = "Different"
STATUS_REFERENCE_ONLY = "Only in Reference"
STATUS_COMPARISON_ONLY = "Only in Comparison"
STATUS_ERROR = "Comparison Error"


@dataclass
class ComparisonResult:
    relative_path: str
    status: str
    comparison_type: str
    reference_size: int | None = None
    comparison_size: int | None = None
    summary: str = ""
    details: list[str] = field(default_factory=list)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as file:
        while chunk := file.read(HASH_CHUNK_SIZE):
            digest.update(chunk)
    return digest.hexdigest()


def sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def image_bytes_signature(value: bytes) -> tuple:
    """Compare embedded images by decoded pixels rather than file metadata."""
    try:
        with Image.open(BytesIO(value)) as image:
            normalized = ImageOps.exif_transpose(image)
            if normalized.mode not in {"RGB", "RGBA", "L"}:
                normalized = normalized.convert("RGBA")
            return normalized.size, normalized.mode, sha256_bytes(normalized.tobytes())
    except Exception:
        return ("binary", sha256_bytes(value))


def normalize_relative_path(path: Path, root: Path) -> str:
    return path.relative_to(root).as_posix()


def build_file_index(root: Path) -> tuple[dict[str, Path], dict[str, str]]:
    """Index files case-insensitively while retaining the visible path."""
    paths: dict[str, Path] = {}
    visible_names: dict[str, str] = {}
    for path in root.rglob("*"):
        if not path.is_file():
            continue
        relative = normalize_relative_path(path, root)
        key = relative.casefold()
        # Windows cannot normally contain two paths differing only by case. If
        # such files are encountered, preserve the first and report the issue.
        paths.setdefault(key, path)
        visible_names.setdefault(key, relative)
    return paths, visible_names


def add_detail(details: list[str], message: str) -> None:
    if len(details) < MAX_DETAILS_PER_FILE:
        details.append(message)
    elif len(details) == MAX_DETAILS_PER_FILE:
        details.append(f"Further differences omitted after {MAX_DETAILS_PER_FILE} details.")


def normalize_root_path(value: Any, root: Path) -> Any:
    if not isinstance(value, str):
        return value
    root_variants = {
        str(root.resolve()),
        str(root.resolve()).replace("/", "\\"),
        str(root.resolve()).replace("\\", "/"),
    }
    normalized = value
    for root_value in sorted(root_variants, key=len, reverse=True):
        normalized = re.sub(re.escape(root_value), "<PROJECT_ROOT>", normalized, flags=re.IGNORECASE)
    return normalized.replace("\\", "/")


def normalize_cell_value(value: Any, root: Path) -> Any:
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, str):
        return normalize_root_path(value, root)
    return value


def side_signature(side) -> tuple:
    color = getattr(side, "color", None)
    return (
        getattr(side, "style", None),
        getattr(color, "type", None),
        getattr(color, "rgb", None),
        getattr(color, "indexed", None),
        getattr(color, "theme", None),
        getattr(color, "tint", None),
    )


def color_signature(color) -> tuple:
    if color is None:
        return ()
    return (
        getattr(color, "type", None),
        getattr(color, "rgb", None),
        getattr(color, "indexed", None),
        getattr(color, "theme", None),
        getattr(color, "tint", None),
        getattr(color, "auto", None),
    )


def cell_style_signature(cell) -> tuple:
    font = cell.font
    fill = cell.fill
    alignment = cell.alignment
    border = cell.border
    protection = cell.protection
    return (
        font.name, font.sz, font.bold, font.italic, font.underline,
        font.strike, color_signature(font.color),
        fill.fill_type, color_signature(fill.fgColor), color_signature(fill.bgColor),
        alignment.horizontal, alignment.vertical, alignment.text_rotation,
        alignment.wrap_text, alignment.shrink_to_fit, alignment.indent,
        side_signature(border.left), side_signature(border.right),
        side_signature(border.top), side_signature(border.bottom),
        side_signature(border.diagonal),
        cell.number_format, protection.locked, protection.hidden,
    )


def hyperlink_signature(cell, root: Path) -> tuple | None:
    link = cell.hyperlink
    if link is None:
        return None
    return (
        normalize_root_path(getattr(link, "target", None), root),
        getattr(link, "location", None),
        getattr(link, "tooltip", None),
        getattr(link, "display", None),
    )


def dimension_signature(dimension) -> tuple:
    return (
        getattr(dimension, "hidden", False),
        getattr(dimension, "width", None),
        getattr(dimension, "height", None),
        getattr(dimension, "outlineLevel", 0),
        getattr(dimension, "collapsed", False),
    )


def image_anchor_signature(image) -> tuple:
    anchor = getattr(image, "anchor", None)
    if isinstance(anchor, str):
        return (anchor, image.width, image.height)
    start = getattr(anchor, "_from", None)
    end = getattr(anchor, "to", None)
    start_value = (
        getattr(start, "row", None), getattr(start, "col", None),
        getattr(start, "rowOff", None), getattr(start, "colOff", None),
    )
    end_value = (
        getattr(end, "row", None), getattr(end, "col", None),
        getattr(end, "rowOff", None), getattr(end, "colOff", None),
    )
    return start_value, end_value, image.width, image.height


def worksheet_image_signatures(worksheet) -> list[tuple]:
    signatures = []
    for image in getattr(worksheet, "_images", []):
        try:
            image_hash = image_bytes_signature(image._data())
        except Exception as exc:
            image_hash = f"unreadable:{exc}"
        signatures.append((image_anchor_signature(image), image_hash))
    return signatures


def compare_excel(reference: Path, comparison: Path, reference_root: Path, comparison_root: Path) -> list[str]:
    details: list[str] = []
    keep_vba = reference.suffix.casefold() == ".xlsm" or comparison.suffix.casefold() == ".xlsm"
    reference_book = load_workbook(reference, data_only=False, keep_vba=keep_vba)
    comparison_book = load_workbook(comparison, data_only=False, keep_vba=keep_vba)
    try:
        if reference_book.sheetnames != comparison_book.sheetnames:
            add_detail(details, f"Sheet order/names differ: {reference_book.sheetnames!r} != {comparison_book.sheetnames!r}")

        all_sheet_names = list(dict.fromkeys(reference_book.sheetnames + comparison_book.sheetnames))
        for sheet_name in all_sheet_names:
            if sheet_name not in reference_book.sheetnames:
                add_detail(details, f"Worksheet exists only in comparison: {sheet_name}")
                continue
            if sheet_name not in comparison_book.sheetnames:
                add_detail(details, f"Worksheet exists only in reference: {sheet_name}")
                continue

            ref_sheet = reference_book[sheet_name]
            cmp_sheet = comparison_book[sheet_name]
            if ref_sheet.sheet_state != cmp_sheet.sheet_state:
                add_detail(details, f"{sheet_name}: sheet visibility differs")
            if str(ref_sheet.freeze_panes or "") != str(cmp_sheet.freeze_panes or ""):
                add_detail(details, f"{sheet_name}: freeze panes differ")
            if sorted(str(item) for item in ref_sheet.merged_cells.ranges) != sorted(
                str(item) for item in cmp_sheet.merged_cells.ranges
            ):
                add_detail(details, f"{sheet_name}: merged-cell ranges differ")

            max_row = max(ref_sheet.max_row, cmp_sheet.max_row)
            max_column = max(ref_sheet.max_column, cmp_sheet.max_column)
            if (ref_sheet.max_row, ref_sheet.max_column) != (cmp_sheet.max_row, cmp_sheet.max_column):
                add_detail(
                    details,
                    f"{sheet_name}: used range differs "
                    f"({ref_sheet.max_row}x{ref_sheet.max_column} != {cmp_sheet.max_row}x{cmp_sheet.max_column})",
                )

            for row in range(1, max_row + 1):
                for column in range(1, max_column + 1):
                    ref_cell = ref_sheet.cell(row=row, column=column)
                    cmp_cell = cmp_sheet.cell(row=row, column=column)
                    coordinate = f"{sheet_name}!{get_column_letter(column)}{row}"
                    ref_value = normalize_cell_value(ref_cell.value, reference_root)
                    cmp_value = normalize_cell_value(cmp_cell.value, comparison_root)
                    if ref_value != cmp_value:
                        add_detail(details, f"{coordinate}: value/formula differs: {ref_value!r} != {cmp_value!r}")
                    if ref_cell.data_type != cmp_cell.data_type:
                        add_detail(details, f"{coordinate}: data type differs")
                    if cell_style_signature(ref_cell) != cell_style_signature(cmp_cell):
                        add_detail(details, f"{coordinate}: formatting differs")
                    if hyperlink_signature(ref_cell, reference_root) != hyperlink_signature(cmp_cell, comparison_root):
                        add_detail(details, f"{coordinate}: hyperlink differs after project-root normalization")

            column_keys = set(ref_sheet.column_dimensions) | set(cmp_sheet.column_dimensions)
            for key in sorted(column_keys):
                if dimension_signature(ref_sheet.column_dimensions[key]) != dimension_signature(cmp_sheet.column_dimensions[key]):
                    add_detail(details, f"{sheet_name}: column {key} dimensions/visibility differ")
            row_keys = set(ref_sheet.row_dimensions) | set(cmp_sheet.row_dimensions)
            for key in sorted(row_keys):
                if dimension_signature(ref_sheet.row_dimensions[key]) != dimension_signature(cmp_sheet.row_dimensions[key]):
                    add_detail(details, f"{sheet_name}: row {key} dimensions/visibility differ")

            if worksheet_image_signatures(ref_sheet) != worksheet_image_signatures(cmp_sheet):
                add_detail(details, f"{sheet_name}: embedded images, image order, size, or anchors differ")
    finally:
        reference_book.close()
        comparison_book.close()
    return details


def decoded_image_signature(path: Path) -> tuple[tuple[int, int], str, str]:
    with Image.open(path) as image:
        normalized = ImageOps.exif_transpose(image)
        if normalized.mode not in {"RGB", "RGBA", "L"}:
            normalized = normalized.convert("RGBA")
        return normalized.size, normalized.mode, sha256_bytes(normalized.tobytes())


def compare_images(reference: Path, comparison: Path) -> list[str]:
    reference_signature = decoded_image_signature(reference)
    comparison_signature = decoded_image_signature(comparison)
    if reference_signature == comparison_signature:
        return []
    details = []
    if reference_signature[0] != comparison_signature[0]:
        details.append(f"Pixel dimensions differ: {reference_signature[0]} != {comparison_signature[0]}")
    if reference_signature[1] != comparison_signature[1]:
        details.append(f"Image modes differ: {reference_signature[1]} != {comparison_signature[1]}")
    if reference_signature[2] != comparison_signature[2]:
        details.append("Decoded pixel content differs")
    return details


def read_text(path: Path) -> str:
    data = path.read_bytes()
    for encoding in ("utf-8-sig", "utf-16", "cp1252"):
        try:
            return data.decode(encoding)
        except UnicodeError:
            continue
    return data.decode("utf-8", errors="replace")


def normalize_text_content(text: str, root: Path, extension: str) -> str:
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    normalized = normalize_root_path(normalized, root)
    if extension in {".html", ".htm"}:
        # Plotly and similar exporters may generate UUID element identifiers.
        normalized = re.sub(
            r"\b[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[1-5][0-9a-fA-F]{3}-[89abAB][0-9a-fA-F]{3}-[0-9a-fA-F]{12}\b",
            "<GENERATED_ID>",
            normalized,
        )
    return normalized


def compare_text(reference: Path, comparison: Path, reference_root: Path, comparison_root: Path) -> list[str]:
    extension = reference.suffix.casefold()
    ref_text = normalize_text_content(read_text(reference), reference_root, extension)
    cmp_text = normalize_text_content(read_text(comparison), comparison_root, extension)
    if ref_text == cmp_text:
        return []
    ref_lines = ref_text.splitlines()
    cmp_lines = cmp_text.splitlines()
    details = [f"Text differs ({len(ref_lines)} lines != {len(cmp_lines)} lines)"]
    for index in range(max(len(ref_lines), len(cmp_lines))):
        ref_line = ref_lines[index] if index < len(ref_lines) else "<MISSING>"
        cmp_line = cmp_lines[index] if index < len(cmp_lines) else "<MISSING>"
        if ref_line != cmp_line:
            add_detail(details, f"Line {index + 1}: {ref_line[:200]!r} != {cmp_line[:200]!r}")
    return details


def read_csv_rows(path: Path) -> list[list[str]]:
    text = read_text(path).replace("\r\n", "\n").replace("\r", "\n")
    sample = text[:8192]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel
    return [list(row) for row in csv.reader(text.splitlines(), dialect)]


def compare_csv(reference: Path, comparison: Path) -> list[str]:
    ref_rows = read_csv_rows(reference)
    cmp_rows = read_csv_rows(comparison)
    if ref_rows == cmp_rows:
        return []
    details = [f"CSV shape differs or cell content changed ({len(ref_rows)} rows != {len(cmp_rows)} rows)"]
    for row_index in range(max(len(ref_rows), len(cmp_rows))):
        ref_row = ref_rows[row_index] if row_index < len(ref_rows) else []
        cmp_row = cmp_rows[row_index] if row_index < len(cmp_rows) else []
        for column_index in range(max(len(ref_row), len(cmp_row))):
            ref_value = ref_row[column_index] if column_index < len(ref_row) else "<MISSING>"
            cmp_value = cmp_row[column_index] if column_index < len(cmp_row) else "<MISSING>"
            if ref_value != cmp_value:
                add_detail(details, f"Row {row_index + 1}, column {column_index + 1}: {ref_value!r} != {cmp_value!r}")
    return details


def document_paragraphs(document: Document) -> list[tuple[str, str, str]]:
    return [
        (paragraph.text, paragraph.style.name if paragraph.style else "", paragraph.alignment)
        for paragraph in document.paragraphs
    ]


def document_tables(document: Document) -> list[list[list[str]]]:
    return [
        [[cell.text for cell in row.cells] for row in table.rows]
        for table in document.tables
    ]


def document_sections(document: Document) -> list[tuple]:
    return [
        (
            section.page_width, section.page_height,
            section.top_margin, section.bottom_margin,
            section.left_margin, section.right_margin,
            section.orientation,
        )
        for section in document.sections
    ]


def document_headers_footers(document: Document) -> list[tuple]:
    result = []
    for section in document.sections:
        result.append(
            (
                tuple(paragraph.text for paragraph in section.header.paragraphs),
                tuple(paragraph.text for paragraph in section.footer.paragraphs),
            )
        )
    return result


def docx_media_hashes(path: Path) -> Counter:
    with ZipFile(path) as archive:
        return Counter(
            image_bytes_signature(archive.read(name))
            for name in archive.namelist()
            if name.startswith("word/media/") and not name.endswith("/")
        )


def docx_content_xml(path: Path) -> dict[str, str]:
    """Return meaningful Word XML while ignoring volatile revision IDs."""
    with ZipFile(path) as archive:
        result = {}
        for name in archive.namelist():
            is_content_part = (
                name == "word/document.xml"
                or name in {"word/styles.xml", "word/numbering.xml", "word/footnotes.xml", "word/endnotes.xml"}
                or name.startswith("word/header")
                or name.startswith("word/footer")
            )
            if not is_content_part or not name.endswith(".xml"):
                continue
            text = archive.read(name).decode("utf-8", errors="replace")
            text = re.sub(r'\s+w:rsid[A-Za-z]+="[^"]*"', "", text)
            text = re.sub(r">\s+<", "><", text).strip()
            result[name] = text
        return result


def compare_docx(reference: Path, comparison: Path) -> list[str]:
    ref_document = Document(reference)
    cmp_document = Document(comparison)
    details = []
    if document_paragraphs(ref_document) != document_paragraphs(cmp_document):
        add_detail(details, "Paragraph text, order, style, or alignment differs")
    if document_tables(ref_document) != document_tables(cmp_document):
        add_detail(details, "Table structure or cell text differs")
    if document_sections(ref_document) != document_sections(cmp_document):
        add_detail(details, "Page size, orientation, or section margins differ")
    if document_headers_footers(ref_document) != document_headers_footers(cmp_document):
        add_detail(details, "Header or footer text differs")
    if len(ref_document.inline_shapes) != len(cmp_document.inline_shapes):
        add_detail(
            details,
            f"Inline image/object count differs: {len(ref_document.inline_shapes)} != {len(cmp_document.inline_shapes)}",
        )
    else:
        ref_sizes = [(shape.width, shape.height) for shape in ref_document.inline_shapes]
        cmp_sizes = [(shape.width, shape.height) for shape in cmp_document.inline_shapes]
        if ref_sizes != cmp_sizes:
            add_detail(details, "Inline image/object dimensions or order differ")
    if docx_media_hashes(reference) != docx_media_hashes(comparison):
        add_detail(details, "Embedded image/media content or counts differ")
    if docx_content_xml(reference) != docx_content_xml(comparison):
        add_detail(details, "Word text-flow, run formatting, page breaks, styles, or numbering XML differs")
    return details


def compare_file_pair(
    relative_path: str,
    reference: Path,
    comparison: Path,
    reference_root: Path,
    comparison_root: Path,
) -> ComparisonResult:
    reference_size = reference.stat().st_size
    comparison_size = comparison.stat().st_size
    if sha256_file(reference) == sha256_file(comparison):
        return ComparisonResult(
            relative_path, STATUS_EXACT, "SHA-256", reference_size, comparison_size,
            "Files are byte-for-byte identical.",
        )

    extension = reference.suffix.casefold()
    comparator: Callable[..., list[str]] | None = None
    comparison_type = "Binary"
    arguments: tuple = (reference, comparison)

    if extension in EXCEL_EXTENSIONS:
        comparator = compare_excel
        comparison_type = "Excel content, formulas, formatting, links, and images"
        arguments = (reference, comparison, reference_root, comparison_root)
    elif extension in WORD_EXTENSIONS:
        comparator = compare_docx
        comparison_type = "Word text, structure, and embedded media"
    elif extension in IMAGE_EXTENSIONS:
        comparator = compare_images
        comparison_type = "Decoded image pixels"
    elif extension in CSV_EXTENSIONS:
        comparator = compare_csv
        comparison_type = "CSV rows and cells"
    elif extension in TEXT_EXTENSIONS:
        comparator = compare_text
        comparison_type = "Normalized text content"
        arguments = (reference, comparison, reference_root, comparison_root)

    if comparator is None:
        return ComparisonResult(
            relative_path, STATUS_DIFFERENT, comparison_type, reference_size, comparison_size,
            "Binary contents differ; no specialized content comparator is available.",
        )

    try:
        details = comparator(*arguments)
    except Exception as exc:
        return ComparisonResult(
            relative_path, STATUS_ERROR, comparison_type, reference_size, comparison_size,
            f"Specialized comparison failed: {exc}", [f"{type(exc).__name__}: {exc}"],
        )

    if not details:
        return ComparisonResult(
            relative_path, STATUS_CONTENT, comparison_type, reference_size, comparison_size,
            "Meaningful content matches; binary metadata or encoding differs.",
        )
    return ComparisonResult(
        relative_path, STATUS_DIFFERENT, comparison_type, reference_size, comparison_size,
        details[0], details,
    )


def compare_directories(
    reference_root: Path,
    comparison_root: Path,
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> list[ComparisonResult]:
    reference_index, reference_names = build_file_index(reference_root)
    comparison_index, comparison_names = build_file_index(comparison_root)
    all_keys = sorted(set(reference_index) | set(comparison_index))
    results = []

    for position, key in enumerate(all_keys, start=1):
        visible_name = reference_names.get(key) or comparison_names[key]
        if progress_callback:
            progress_callback(position, len(all_keys), visible_name)
        reference = reference_index.get(key)
        comparison = comparison_index.get(key)
        if reference is None:
            results.append(
                ComparisonResult(
                    visible_name, STATUS_COMPARISON_ONLY, "Directory inventory",
                    None, comparison.stat().st_size, "File is absent from the reference directory.",
                )
            )
            continue
        if comparison is None:
            results.append(
                ComparisonResult(
                    visible_name, STATUS_REFERENCE_ONLY, "Directory inventory",
                    reference.stat().st_size, None, "File is absent from the comparison directory.",
                )
            )
            continue
        if reference_names[key] != comparison_names[key]:
            result = compare_file_pair(
                visible_name, reference, comparison, reference_root, comparison_root
            )
            add_detail(
                result.details,
                f"Path capitalization differs: {reference_names[key]!r} != {comparison_names[key]!r}",
            )
            if result.status in {STATUS_EXACT, STATUS_CONTENT}:
                result.status = STATUS_DIFFERENT
                result.summary = "File content matches, but relative path capitalization differs."
            results.append(result)
            continue
        results.append(
            compare_file_pair(visible_name, reference, comparison, reference_root, comparison_root)
        )
    return results


def autosize_columns(worksheet, maximum_width: int = 70) -> None:
    for column_cells in worksheet.columns:
        length = max(len(str(cell.value or "")) for cell in column_cells)
        worksheet.column_dimensions[column_cells[0].column_letter].width = min(max(length + 2, 10), maximum_width)


def write_report(
    report_path: Path,
    reference_root: Path,
    comparison_root: Path,
    results: list[ComparisonResult],
) -> None:
    workbook = Workbook()
    summary_sheet = workbook.active
    summary_sheet.title = "Summary"
    differences_sheet = workbook.create_sheet("Differences")
    all_files_sheet = workbook.create_sheet("All Files")

    counts = Counter(result.status for result in results)
    meaningful_matches = counts[STATUS_EXACT] + counts[STATUS_CONTENT]
    difference_count = len(results) - meaningful_matches
    overall = "PASS" if difference_count == 0 else "REVIEW REQUIRED"

    summary_rows = [
        ("Directory Comparison Result", overall),
        ("Reference directory", str(reference_root)),
        ("Comparison directory", str(comparison_root)),
        ("Generated", datetime.now().isoformat(timespec="seconds")),
        ("Total relative files", len(results)),
        ("Exact byte matches", counts[STATUS_EXACT]),
        ("Content matches (binary metadata differs)", counts[STATUS_CONTENT]),
        ("Different files", counts[STATUS_DIFFERENT]),
        ("Only in reference", counts[STATUS_REFERENCE_ONLY]),
        ("Only in comparison", counts[STATUS_COMPARISON_ONLY]),
        ("Comparison errors", counts[STATUS_ERROR]),
        ("Items requiring review", difference_count),
    ]
    for row in summary_rows:
        summary_sheet.append(row)

    headers = [
        "Relative Path", "Status", "Comparison Type", "Reference Size",
        "Comparison Size", "Summary", "Details",
    ]
    differences_sheet.append(headers)
    all_files_sheet.append(headers)
    for result in results:
        row = [
            result.relative_path,
            result.status,
            result.comparison_type,
            result.reference_size,
            result.comparison_size,
            result.summary,
            "\n".join(result.details),
        ]
        all_files_sheet.append(row)
        if result.status not in {STATUS_EXACT, STATUS_CONTENT}:
            differences_sheet.append(row)

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    status_fills = {
        STATUS_EXACT: "C6EFCE",
        STATUS_CONTENT: "E2F0D9",
        STATUS_DIFFERENT: "FFC7CE",
        STATUS_REFERENCE_ONLY: "FCE4D6",
        STATUS_COMPARISON_ONLY: "FFF2CC",
        STATUS_ERROR: "E4DFEC",
    }
    for worksheet in (differences_sheet, all_files_sheet):
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions
        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
        for row in range(2, worksheet.max_row + 1):
            status_cell = worksheet.cell(row=row, column=2)
            color = status_fills.get(status_cell.value)
            if color:
                status_cell.fill = PatternFill("solid", fgColor=color)
            worksheet.cell(row=row, column=7).alignment = Alignment(wrap_text=True, vertical="top")
        autosize_columns(worksheet)
        worksheet.column_dimensions["G"].width = 90

    summary_sheet["A1"].font = Font(bold=True, size=14)
    summary_sheet["B1"].font = Font(bold=True, size=14, color="008000" if overall == "PASS" else "C00000")
    summary_sheet.freeze_panes = "A2"
    autosize_columns(summary_sheet)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(report_path)


def choose_directories() -> tuple[Path, Path, Path] | None:
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    reference = filedialog.askdirectory(
        parent=root,
        title="Select the manually prepared REFERENCE directory (known correct)",
    )
    if not reference:
        root.destroy()
        return None
    comparison = filedialog.askdirectory(
        parent=root,
        title="Select the directory to COMPARE against the reference",
    )
    if not comparison:
        root.destroy()
        return None
    output = filedialog.askdirectory(
        parent=root,
        title="Select where to save the comparison report",
        initialdir=str(Path(comparison).parent),
    )
    if not output:
        root.destroy()
        return None
    root.destroy()
    return Path(reference), Path(comparison), Path(output)


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--reference", type=Path, help="Known-correct reference directory")
    parser.add_argument("--comparison", type=Path, help="Directory to compare")
    parser.add_argument("--output", type=Path, help="Output directory for the Excel report")
    return parser.parse_args()


def main() -> int:
    args = parse_arguments()
    if args.reference or args.comparison:
        if not args.reference or not args.comparison:
            raise SystemExit("Both --reference and --comparison are required together.")
        reference_root = args.reference.resolve()
        comparison_root = args.comparison.resolve()
        output_directory = (args.output or Path.cwd()).resolve()
        use_messagebox = False
    else:
        selected = choose_directories()
        if selected is None:
            return 0
        reference_root, comparison_root, output_directory = selected
        reference_root = reference_root.resolve()
        comparison_root = comparison_root.resolve()
        output_directory = output_directory.resolve()
        use_messagebox = True

    if not reference_root.is_dir():
        raise SystemExit(f"Reference directory does not exist: {reference_root}")
    if not comparison_root.is_dir():
        raise SystemExit(f"Comparison directory does not exist: {comparison_root}")
    if reference_root == comparison_root:
        raise SystemExit("Reference and comparison directories must be different.")
    if reference_root.is_relative_to(comparison_root) or comparison_root.is_relative_to(reference_root):
        raise SystemExit("Reference and comparison directories cannot be inside one another.")
    if output_directory.is_relative_to(reference_root) or output_directory.is_relative_to(comparison_root):
        raise SystemExit("Choose an output directory outside both directories being compared.")

    def progress(position: int, total: int, relative_path: str) -> None:
        print(f"[{position}/{total}] {relative_path}", flush=True)

    results = compare_directories(reference_root, comparison_root, progress)
    report_path = output_directory / REPORT_NAME
    write_report(report_path, reference_root, comparison_root, results)
    review_count = sum(
        result.status not in {STATUS_EXACT, STATUS_CONTENT}
        for result in results
    )
    message = (
        f"Comparison complete.\n\nReport: {report_path}\n\n"
        f"Files checked: {len(results)}\nItems requiring review: {review_count}"
    )
    print(message, flush=True)
    if use_messagebox:
        root = Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        messagebox.showinfo(APP_TITLE, message, parent=root)
        root.destroy()
    return 0 if review_count == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
