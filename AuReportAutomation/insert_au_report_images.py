"""Create the first Au report workbook by inserting reflected light and SEM images.

This script is intentionally focused on step 1 of the Au report automation:
- ask for the sample type
- ask for reflected light and SEM image folders
- sort .jpg/.jpeg images by the first number in each filename
- create a new Excel workbook
- add the correct headers for the selected sample type
- insert reflected light and SEM images into the correct columns
- number rows automatically in column A

Run from PyCharm or a terminal with the project virtual environment active.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from tempfile import TemporaryDirectory
from tkinter import Tk, filedialog, messagebox, simpledialog

from copy import copy

from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from PIL import Image as PillowImage

ROW_HEIGHT = 45
COLUMN_WIDTH = 8.43
FONT_SIZE = 11
IMAGE_EXTENSIONS = {".jpg", ".jpeg"}
ORIGINAL_IMAGE_SIZE_CM = 1.58
DISPLAY_IMAGE_SIZE_CM = 1.59
PIXELS_PER_CM_AT_96_DPI = 96 / 2.54
ORIGINAL_IMAGE_WIDTH_PX = round(ORIGINAL_IMAGE_SIZE_CM * PIXELS_PER_CM_AT_96_DPI)
ORIGINAL_IMAGE_HEIGHT_PX = round(ORIGINAL_IMAGE_SIZE_CM * PIXELS_PER_CM_AT_96_DPI)
TARGET_IMAGE_WIDTH_PX = round(DISPLAY_IMAGE_SIZE_CM * PIXELS_PER_CM_AT_96_DPI)
TARGET_IMAGE_HEIGHT_PX = round(DISPLAY_IMAGE_SIZE_CM * PIXELS_PER_CM_AT_96_DPI)
HEADER_FILL = PatternFill(fill_type="solid", fgColor="D9EAF7")


@dataclass(frozen=True)
class SampleLayout:
    """Column layout for one Au report sample type."""

    sample_type: str
    headers: tuple[str, ...]
    reflected_light_column: int
    sem_column: int


SAMPLE_LAYOUTS: dict[str, SampleLayout] = {
    "1": SampleLayout(
        sample_type="Au+Ag",
        headers=("No.", "Au (Wt%)", "Ag (Wt%)", "R. Light Images", "SEM Images"),
        reflected_light_column=4,
        sem_column=5,
    ),
    "2": SampleLayout(
        sample_type="Au+Ag+Cu",
        headers=("No.", "Au (Wt%)", "Ag (Wt%)", "Cu (Wt%)", "R. Light Images", "SEM Images"),
        reflected_light_column=5,
        sem_column=6,
    ),
    "3": SampleLayout(
        sample_type="Au+Ag+Cu+Hg",
        headers=("No.", "Au (Wt%)", "Ag (Wt%)", "Cu (Wt%)", "Hg (Wt%)", "R. Light Images", "SEM Images"),
        reflected_light_column=6,
        sem_column=7,
    ),
    "4": SampleLayout(
        sample_type="Au+Ag+Hg",
        headers=("No.", "Au (Wt%)", "Ag (Wt%)", "Hg (Wt%)", "R. Light Images", "SEM Images"),
        reflected_light_column=5,
        sem_column=6,
    ),
}


class UserCancelledError(Exception):
    """Raised when the user cancels one of the required GUI prompts."""


def get_first_number(path: Path) -> int:
    """Return the first number found in a filename for numeric image sorting."""
    match = re.search(r"\d+", path.stem)
    if not match:
        return 10**12
    return int(match.group())


def sorted_image_files(folder: Path) -> list[Path]:
    """Return .jpg/.jpeg images sorted by their first filename number, then name."""
    image_files = [path for path in folder.iterdir() if path.is_file() and path.suffix.lower() in IMAGE_EXTENSIONS]
    return sorted(image_files, key=lambda path: (get_first_number(path), path.name.lower()))


def choose_sample_layout() -> SampleLayout:
    """Ask the user which Au sample layout to use."""
    prompt = (
        "Select sample type by number:\n\n"
        "1 = Au+Ag\n"
        "2 = Au+Ag+Cu\n"
        "3 = Au+Ag+Cu+Hg\n"
        "4 = Au+Ag+Hg"
    )
    choice = simpledialog.askstring("Au Report Sample Type", prompt)
    if choice is None:
        raise UserCancelledError("Sample type selection was cancelled.")
    choice = choice.strip()
    if choice not in SAMPLE_LAYOUTS:
        messagebox.showerror("Invalid sample type", "Please run again and enter 1, 2, 3, or 4.")
        raise UserCancelledError(f"Invalid sample type choice: {choice}")
    return SAMPLE_LAYOUTS[choice]


def choose_folder(title: str) -> Path:
    """Ask the user to select a folder."""
    selected = filedialog.askdirectory(title=title)
    if not selected:
        raise UserCancelledError(f"Folder selection was cancelled: {title}")
    return Path(selected)


def choose_output_file(default_name: str) -> Path:
    """Ask the user where to save the new final workbook."""
    selected = filedialog.asksaveasfilename(
        title="Save new Au report workbook as",
        defaultextension=".xlsx",
        initialfile=default_name,
        filetypes=(("Excel workbook", "*.xlsx"),),
    )
    if not selected:
        raise UserCancelledError("Output workbook selection was cancelled.")
    return Path(selected)


def set_standard_dimensions(worksheet, max_row: int, max_column: int) -> None:
    """Apply the requested row height and column width."""
    for row_number in range(1, max_row + 1):
        worksheet.row_dimensions[row_number].height = ROW_HEIGHT
    for column_number in range(1, max_column + 1):
        worksheet.column_dimensions[get_column_letter(column_number)].width = COLUMN_WIDTH


def prepare_embedded_image(image_path: Path, temporary_folder: Path) -> Path:
    """Create a square embedded copy representing the requested 1.58 cm original size.

    Excel/openpyxl stores image dimensions in pixels. At 96 DPI, 1.58 cm and 1.59 cm
    both round to 60 px, which keeps Excel's displayed scale effectively at 100% while
    still honoring the requested 1.59 cm report display size.
    """
    output_path = temporary_folder / f"{image_path.stem}_au_report{image_path.suffix.lower()}"
    with PillowImage.open(image_path) as source_image:
        resized_image = source_image.convert("RGB").resize(
            (ORIGINAL_IMAGE_WIDTH_PX, ORIGINAL_IMAGE_HEIGHT_PX),
            PillowImage.Resampling.LANCZOS,
        )
        resized_image.save(output_path)
    return output_path


def exact_report_image_size() -> tuple[int, int]:
    """Return the required 1.59 cm x 1.59 cm display size in Excel pixels."""
    return TARGET_IMAGE_WIDTH_PX, TARGET_IMAGE_HEIGHT_PX


def add_image_to_cell(worksheet, image_path: Path, row_number: int, column_number: int, temporary_folder: Path) -> None:
    """Insert one image into a worksheet cell."""
    embedded_image_path = prepare_embedded_image(image_path, temporary_folder)
    image = ExcelImage(str(embedded_image_path))
    image.width, image.height = exact_report_image_size()
    worksheet.add_image(image, f"{get_column_letter(column_number)}{row_number}")



def apply_standard_text_format(worksheet, max_row: int, max_column: int) -> None:
    """Apply 11-point text everywhere, with bold text only in the header row."""
    for row_number in range(1, max_row + 1):
        for column_number in range(1, max_column + 1):
            cell = worksheet.cell(row=row_number, column=column_number)
            cell.font = Font(size=FONT_SIZE, bold=(row_number == 1))
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def prepare_headers(worksheet, layout: SampleLayout) -> None:
    """Add headers for the selected sample layout."""
    for column_number, header in enumerate(layout.headers, start=1):
        cell = worksheet.cell(row=1, column=column_number, value=header)
        cell.font = Font(size=FONT_SIZE, bold=True)
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def find_header_column(worksheet, header_name: str) -> int:
    """Find the first column whose row-1 header matches header_name."""
    wanted = header_name.strip().lower()
    for cell in worksheet[1]:
        if str(cell.value).strip().lower() == wanted:
            return cell.column
    raise ValueError(f"Could not find header {header_name!r} in row 1 of sheet {worksheet.title!r}.")


def copy_area_column_from_workbook(report_worksheet, data_workbook_path: Path, destination_column: int) -> None:
    """Copy the Area column from the Others sheet into the report workbook.

    Values and cell fill formatting are copied so blank yellow cells stay blank and yellow.
    """
    data_workbook = load_workbook(data_workbook_path)
    if "Others" not in data_workbook.sheetnames:
        raise ValueError(f"The selected workbook does not contain a sheet named 'Others': {data_workbook_path}")

    source_worksheet = data_workbook["Others"]
    source_column = find_header_column(source_worksheet, "Area")

    header_cell = report_worksheet.cell(row=1, column=destination_column, value="Area")
    header_cell.font = Font(size=FONT_SIZE, bold=True)
    header_cell.fill = HEADER_FILL
    header_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    report_worksheet.column_dimensions[get_column_letter(destination_column)].width = COLUMN_WIDTH
    apply_standard_text_format(report_worksheet, max_row=source_worksheet.max_row, max_column=destination_column)

    for row_number in range(2, source_worksheet.max_row + 1):
        source_cell = source_worksheet.cell(row=row_number, column=source_column)
        destination_cell = report_worksheet.cell(row=row_number, column=destination_column, value=source_cell.value)
        destination_cell.font = Font(size=FONT_SIZE, bold=False)
        destination_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        destination_cell.fill = copy(source_cell.fill)
        report_worksheet.row_dimensions[row_number].height = ROW_HEIGHT


def choose_excel_file(title: str) -> Path | None:
    """Ask the user to select an Excel workbook, returning None if skipped."""
    selected = filedialog.askopenfilename(
        title=title,
        filetypes=(("Excel workbooks", "*.xlsx *.xlsm"), ("All files", "*.*")),
    )
    if not selected:
        return None
    return Path(selected)


def create_image_workbook(
    layout: SampleLayout,
    reflected_images: list[Path],
    sem_images: list[Path],
    temporary_folder: Path,
) -> Workbook:
    """Create the first Au report workbook with headers, numbers, and images."""
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = layout.sample_type.replace("+", "_")

    image_count = max(len(reflected_images), len(sem_images))
    max_column = len(layout.headers)
    set_standard_dimensions(worksheet, max_row=image_count + 1, max_column=max_column)
    apply_standard_text_format(worksheet, max_row=image_count + 1, max_column=max_column)
    prepare_headers(worksheet, layout)

    for index in range(image_count):
        row_number = index + 2
        number_cell = worksheet.cell(row=row_number, column=1, value=index + 1)
        number_cell.font = Font(size=FONT_SIZE, bold=False)
        number_cell.alignment = Alignment(horizontal="center", vertical="center")

        if index < len(reflected_images):
            add_image_to_cell(worksheet, reflected_images[index], row_number, layout.reflected_light_column, temporary_folder)
        if index < len(sem_images):
            add_image_to_cell(worksheet, sem_images[index], row_number, layout.sem_column, temporary_folder)

    worksheet.freeze_panes = "A2"
    return workbook


def main() -> None:
    """Run the GUI workflow."""
    root = Tk()
    root.withdraw()

    layout = choose_sample_layout()
    reflected_folder = choose_folder("Select reflected light image folder")
    sem_folder = choose_folder("Select SEM image folder")

    reflected_images = sorted_image_files(reflected_folder)
    sem_images = sorted_image_files(sem_folder)
    if not reflected_images and not sem_images:
        messagebox.showerror("No images found", "No .jpg or .jpeg images were found in either selected folder.")
        raise UserCancelledError("No images found in selected folders.")

    output_file = choose_output_file(f"Au_Report_{layout.sample_type.replace('+', '_')}.xlsx")
    with TemporaryDirectory() as temporary_directory:
        workbook = create_image_workbook(layout, reflected_images, sem_images, Path(temporary_directory))

        area_imported = False
        if messagebox.askyesno("Import Area column?", "Do you want to import the Area column from the Others sheet now?"):
            data_workbook_path = choose_excel_file("Select Excel data file containing the Others sheet")
            if data_workbook_path is not None:
                worksheet = workbook.active
                area_column = layout.sem_column + 1
                copy_area_column_from_workbook(worksheet, data_workbook_path, area_column)
                area_imported = True

        workbook.save(output_file)

    messagebox.showinfo(
        "Au report workbook created",
        "Workbook saved successfully.\n\n"
        f"Output: {output_file}\n"
        f"Reflected light images inserted: {len(reflected_images)}\n"
        f"SEM images inserted: {len(sem_images)}\n"
        f"Area column imported: {'Yes' if area_imported else 'No'}",
    )


if __name__ == "__main__":
    main()
