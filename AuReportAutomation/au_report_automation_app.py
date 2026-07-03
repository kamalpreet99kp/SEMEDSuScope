"""Au Report Automation App.

PySide-based workflow app for one Au SEM/EDS report sample. The UI style follows
`Reference file.txt` from the `reference-file-for-codex` branch: a QMainWindow
with numbered QTabWidget workflow steps, grouped controls, status labels, tables,
and a log panel.
"""

from __future__ import annotations

import csv
import os
import subprocess
import sys
import time
from dataclasses import dataclass
from glob import glob
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

try:
    from PySide6.QtCore import Qt
    from PySide6.QtWidgets import (
        QApplication,
        QAbstractItemView,
        QComboBox,
        QFileDialog,
        QGroupBox,
        QHBoxLayout,
        QHeaderView,
        QLabel,
        QLineEdit,
        QMainWindow,
        QMessageBox,
        QPlainTextEdit,
        QPushButton,
        QTableWidget,
        QTableWidgetItem,
        QTabWidget,
        QVBoxLayout,
        QWidget,
    )
except ImportError:
    try:
        from PySide2.QtCore import Qt
        from PySide2.QtWidgets import (
            QApplication,
            QAbstractItemView,
            QComboBox,
            QFileDialog,
            QGroupBox,
            QHBoxLayout,
            QHeaderView,
            QLabel,
            QLineEdit,
            QMainWindow,
            QMessageBox,
            QPlainTextEdit,
            QPushButton,
            QTableWidget,
            QTableWidgetItem,
            QTabWidget,
            QVBoxLayout,
            QWidget,
        )
    except ImportError as exc:
        raise SystemExit(
            "PySide6 is required to run the app. Install it with:\n"
            "    pip install PySide6\n"
            "or install PySide2 if PySide6 is not available."
        ) from exc


def _qt_attr(group_name: str, attr_name: str, fallback_name: str):
    group = getattr(Qt, group_name, None)
    if group is not None and hasattr(group, attr_name):
        return getattr(group, attr_name)
    return getattr(Qt, fallback_name)


def _class_enum(class_object, group_name: str, attr_name: str, fallback_name: str):
    group = getattr(class_object, group_name, None)
    if group is not None and hasattr(group, attr_name):
        return getattr(group, attr_name)
    return getattr(class_object, fallback_name)


ALIGN_CENTER = _qt_attr("AlignmentFlag", "AlignCenter", "AlignCenter")
HEADER_STRETCH = _class_enum(QHeaderView, "ResizeMode", "Stretch", "Stretch")
NO_EDIT_TRIGGERS = _class_enum(QAbstractItemView, "EditTrigger", "NoEditTriggers", "NoEditTriggers")

ELEMENTS_BY_SAMPLE_TYPE = {
    "Au+Ag": ("Au", "Ag"),
    "Au+Ag+Cu": ("Au", "Ag", "Cu"),
    "Au+Ag+Cu+Hg": ("Au", "Ag", "Cu", "Hg"),
    "Au+Ag+Hg": ("Au", "Ag", "Hg"),
}
IMAGE_EXTENSIONS = {".jpg", ".jpeg"}
RESIZED_FOLDER_NAME = "resizedtosmallest"


@dataclass
class AppState:
    sample_type: str = "Au+Ag"
    raw_export_path: Path | None = None
    sample_dir: Path | None = None
    sample_name: str = ""
    data_workbook_path: Path | None = None
    microscope_image_list_dir: Path | None = None
    microscope_resized_dir: Path | None = None
    sem_image_dir: Path | None = None
    sem_resized_dir: Path | None = None
    report_workbook_path: Path | None = None


def normalize_header(value) -> str:
    return " ".join(str(value or "").replace("\n", " ").split()).strip().lower()


def find_header_index(headers: list[str], wanted_header: str) -> int:
    wanted = normalize_header(wanted_header)
    for index, header in enumerate(headers):
        if normalize_header(header) == wanted:
            return index
    raise ValueError(f"Could not find required header {wanted_header!r}.")


def read_full_export_csv(csv_path: Path) -> list[list[str]]:
    """Read a .full export CSV-like file into rows."""
    with csv_path.open("r", encoding="utf-8-sig", newline="") as file:
        sample = file.read(4096)
        file.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample)
        except csv.Error:
            dialect = csv.excel
        return [row for row in csv.reader(file, dialect)]


def write_rows(worksheet, rows: list[list[str]]) -> None:
    for row in rows:
        worksheet.append(row)


def autosize_columns(worksheet) -> None:
    for column_cells in worksheet.columns:
        max_length = 0
        column_letter = get_column_letter(column_cells[0].column)
        for cell in column_cells:
            max_length = max(max_length, len(str(cell.value or "")))
        worksheet.column_dimensions[column_letter].width = min(max(max_length + 2, 10), 45)


def create_data_workbook(raw_export_path: Path, sample_type: str) -> Path:
    """Create Output File 1 with Raw Data and Au normalized sheets."""
    rows = read_full_export_csv(raw_export_path)
    if not rows:
        raise ValueError(f"No rows were found in {raw_export_path}.")

    headers = rows[0]
    au_index = find_header_index(headers, "Au (Wt%)")
    elements = ELEMENTS_BY_SAMPLE_TYPE[sample_type]
    element_indices = {element: find_header_index(headers, f"{element} (Wt%)") for element in elements}

    workbook = Workbook()
    raw_sheet = workbook.active
    raw_sheet.title = "Raw Data"
    write_rows(raw_sheet, rows)

    au_sheet = workbook.create_sheet("Au")
    au_sheet.append(headers)
    for row in rows[1:]:
        try:
            au_value = float(str(row[au_index]).strip() or 0)
        except (ValueError, IndexError):
            au_value = 0
        if au_value > 0:
            au_sheet.append(row)

    existing_last_column = len(headers)
    normalized_column = existing_last_column + 3
    element_start_column = normalized_column + 2
    au_sheet.cell(row=1, column=normalized_column, value="Normalized")
    for offset, element in enumerate(elements):
        au_sheet.cell(row=1, column=element_start_column + offset, value=element)

    for row_number in range(2, au_sheet.max_row + 1):
        source_refs = [f"{get_column_letter(element_indices[element] + 1)}{row_number}" for element in elements]
        normalized_cell = f"{get_column_letter(normalized_column)}{row_number}"
        au_sheet.cell(row=row_number, column=normalized_column, value=f"=SUM({','.join(source_refs)})")
        for offset, element in enumerate(elements):
            source_cell = f"{get_column_letter(element_indices[element] + 1)}{row_number}"
            au_sheet.cell(
                row=row_number,
                column=element_start_column + offset,
                value=f"={source_cell}*100/{normalized_cell}",
            )

    for sheet in (raw_sheet, au_sheet):
        for row in sheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(size=11, bold=(cell.row == 1))
        autosize_columns(sheet)
        sheet.freeze_panes = "A2"

    output_path = raw_export_path.with_suffix(".xlsx")
    workbook.save(output_path)
    return output_path


def create_others_sheet(data_workbook_path: Path) -> None:
    """Create/update Others sheet from the edited Au sheet Area column."""
    workbook = load_workbook(data_workbook_path)
    if "Au" not in workbook.sheetnames:
        raise ValueError("The data workbook does not contain an 'Au' sheet.")
    au_sheet = workbook["Au"]
    headers = [cell.value for cell in au_sheet[1]]
    area_column_index = find_header_index(headers, "Area") + 1

    if "Others" in workbook.sheetnames:
        del workbook["Others"]
    others_sheet = workbook.create_sheet("Others")
    others_sheet.cell(row=1, column=1, value="Area")
    for row_number in range(2, au_sheet.max_row + 1):
        source_cell = au_sheet.cell(row=row_number, column=area_column_index)
        target_cell = others_sheet.cell(row=row_number, column=1, value=source_cell.value)
        if source_cell.fill:
            target_cell.fill = source_cell.fill.copy()

    others_sheet.cell(row=1, column=6, value="SEM Image No.")
    others_sheet.cell(row=1, column=7, value="uScope Image No.")
    for row in others_sheet.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.font = Font(size=11, bold=(cell.row == 1))
    autosize_columns(others_sheet)
    workbook.save(data_workbook_path)


def prepare_image_names(folder: Path) -> None:
    """Apply the existing image-renaming behavior before resize."""
    for image_path in folder.iterdir():
        if image_path.is_file() and image_path.suffix.lower() == ".jpeg":
            image_path.rename(image_path.with_suffix(".jpg"))

    for image_path in folder.iterdir():
        if image_path.is_file() and image_path.name.startswith("Electron Image "):
            new_name = image_path.name.replace("Electron Image ", "")
            stem, suffix = new_name.rsplit(".", 1)
            image_path.rename(image_path.with_name(f"{stem.zfill(3)}.{suffix}"))


def resize_images_to_smallest(folder: Path) -> Path:
    """Crop all JPG images in a folder to a shared smallest size."""
    try:
        import cv2
    except ImportError as exc:
        raise RuntimeError("OpenCV is required for image resizing. Install it with: pip install opencv-python") from exc

    folder = Path(folder)
    prepare_image_names(folder)
    output_folder = folder / RESIZED_FOLDER_NAME
    output_folder.mkdir(exist_ok=True)

    image_paths = sorted(Path(path) for path in glob(str(folder / "*.jpg")))
    if not image_paths:
        raise ValueError(f"No .jpg images were found in {folder}.")

    widths = []
    heights = []
    for image_path in image_paths:
        image = cv2.imread(str(image_path))
        if image is None:
            continue
        height, width = image.shape[:2]
        widths.append(width)
        heights.append(height)

    if not widths or not heights:
        raise ValueError(f"No readable .jpg images were found in {folder}.")

    crop_width = min(widths)
    crop_height = int(0.75 * crop_width) if crop_width > min(heights) else int(crop_width / 1.33)
    (output_folder / f"NEW Max Size is {crop_width} {crop_height} .txt").write_text("")

    for image_path in image_paths:
        image = cv2.imread(str(image_path))
        if image is None:
            continue
        height, width = image.shape[:2]
        mid_x, mid_y = width // 2, height // 2
        half_width, half_height = crop_width // 2, crop_height // 2
        cropped = image[mid_y - half_height:mid_y + half_height, mid_x - half_width:mid_x + half_width]
        cv2.imwrite(str(output_folder / f"{image_path.stem}-resized.jpg"), cropped)
    return output_folder


def open_path(path: Path) -> None:
    """Open a file/folder in the OS default application."""
    if sys.platform.startswith("win"):
        os.startfile(str(path))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


class AuReportAutomationApp(QMainWindow):
    """PySide workflow app for the Au report automation."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Au Report Automation App")
        self.resize(1280, 820)
        self.state = AppState()

        self.tabs = QTabWidget()
        self.setup_tab = self._build_setup_tab()
        self.images_tab = self._build_images_tab()
        self.report_tab = self._build_report_tab()
        self.word_tab = self._build_word_tab()
        self.tabs.addTab(self.setup_tab, "1. Data Workbook")
        self.tabs.addTab(self.images_tab, "2. Resize Images")
        self.tabs.addTab(self.report_tab, "3. Excel Report")
        self.tabs.addTab(self.word_tab, "4. Word Report")
        self.setCentralWidget(self.tabs)

    def _build_setup_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)

        group = QGroupBox("Sample Setup and Output File 1")
        group_layout = QVBoxLayout(group)

        sample_layout = QHBoxLayout()
        self.sample_type_combo = QComboBox()
        self.sample_type_combo.addItems(ELEMENTS_BY_SAMPLE_TYPE.keys())
        self.raw_export_edit = QLineEdit()
        self.raw_export_edit.setPlaceholderText("Select .full export / CSV file")
        self.select_raw_button = QPushButton("Select Raw Export")
        sample_layout.addWidget(QLabel("Sample type"))
        sample_layout.addWidget(self.sample_type_combo)
        sample_layout.addWidget(self.raw_export_edit, stretch=1)
        sample_layout.addWidget(self.select_raw_button)

        buttons_layout = QHBoxLayout()
        self.create_data_button = QPushButton("Create Data Workbook")
        self.create_others_button = QPushButton("Create Others Sheet After Manual Au Edits")
        self.open_data_button = QPushButton("Open Data Workbook")
        buttons_layout.addWidget(self.create_data_button)
        buttons_layout.addWidget(self.create_others_button)
        buttons_layout.addWidget(self.open_data_button)
        buttons_layout.addStretch(1)

        self.data_status = QLabel("Select the .full export file to start. The containing folder becomes the sample folder.")
        group_layout.addLayout(sample_layout)
        group_layout.addLayout(buttons_layout)
        group_layout.addWidget(self.data_status)

        self.summary_table = QTableWidget(0, 2)
        self.summary_table.setHorizontalHeaderLabels(["Item", "Path / Value"])
        self.summary_table.horizontalHeader().setSectionResizeMode(HEADER_STRETCH)
        self.summary_table.verticalHeader().setVisible(False)
        self.summary_table.setEditTriggers(NO_EDIT_TRIGGERS)

        layout.addWidget(group)
        layout.addWidget(self.summary_table, stretch=1)
        self.app_log = QPlainTextEdit()
        self.app_log.setReadOnly(True)
        layout.addWidget(self.app_log, stretch=1)
        self._connect_setup_signals()
        return tab

    def _build_images_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        group = QGroupBox("Resize Microscope and SEM Images")
        group_layout = QVBoxLayout(group)

        self.resize_microscope_button = QPushButton("Select Microscope Parent Folder and Resize")
        self.resize_sem_button = QPushButton("Select SEM Folder and Resize")
        self.image_status = QLabel("Resize microscope images first, then SEM images.")
        group_layout.addWidget(self.resize_microscope_button)
        group_layout.addWidget(self.resize_sem_button)
        group_layout.addWidget(self.image_status)
        layout.addWidget(group)
        layout.addStretch(1)

        self.resize_microscope_button.clicked.connect(self._resize_microscope_images)
        self.resize_sem_button.clicked.connect(self._resize_sem_images)
        return tab

    def _build_report_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        group = QGroupBox("Excel Report Workflow")
        group_layout = QVBoxLayout(group)
        self.report_status = QLabel(
            "Next implementation step: wire the existing insert/finish report functions to the app state without repeated prompts."
        )
        group_layout.addWidget(self.report_status)
        layout.addWidget(group)
        layout.addStretch(1)
        return tab

    def _build_word_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        group = QGroupBox("Word Report Workflow")
        group_layout = QVBoxLayout(group)
        self.word_status = QLabel(
            "Next implementation step: run the Word macro launcher with the final workbook known from the app state."
        )
        group_layout.addWidget(self.word_status)
        layout.addWidget(group)
        layout.addStretch(1)
        return tab

    def _connect_setup_signals(self) -> None:
        self.sample_type_combo.currentTextChanged.connect(self._set_sample_type)
        self.select_raw_button.clicked.connect(self._select_raw_export)
        self.create_data_button.clicked.connect(self._create_data_workbook)
        self.create_others_button.clicked.connect(self._create_others_sheet)
        self.open_data_button.clicked.connect(self._open_data_workbook)

    def _set_sample_type(self, sample_type: str) -> None:
        self.state.sample_type = sample_type
        self._refresh_summary()

    def _select_raw_export(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(
            self,
            "Select .full export / CSV file",
            str(Path.home()),
            "EDS exports (*.full export *.csv *.txt);;All files (*.*)",
        )
        if not selected:
            return
        raw_path = Path(selected)
        self.state.raw_export_path = raw_path
        self.state.sample_dir = raw_path.parent
        self.state.sample_name = raw_path.stem
        self.raw_export_edit.setText(str(raw_path))
        self.data_status.setText(f"Selected raw export. Sample folder: {raw_path.parent}")
        self._append_log(f"Selected raw export: {raw_path}")
        self._refresh_summary()

    def _create_data_workbook(self) -> None:
        if self.state.raw_export_path is None:
            QMessageBox.warning(self, "Missing raw export", "Select the .full export file first.")
            return
        try:
            self.state.data_workbook_path = create_data_workbook(self.state.raw_export_path, self.state.sample_type)
            self.data_status.setText(f"Created data workbook: {self.state.data_workbook_path}")
            self._append_log(f"Created data workbook: {self.state.data_workbook_path}")
            self._refresh_summary()
            open_path(self.state.data_workbook_path)
        except Exception as exc:
            QMessageBox.critical(self, "Data workbook failed", str(exc))

    def _create_others_sheet(self) -> None:
        if self.state.data_workbook_path is None:
            QMessageBox.warning(self, "Missing data workbook", "Create the data workbook first.")
            return
        try:
            create_others_sheet(self.state.data_workbook_path)
            self.data_status.setText("Created/updated Others sheet after manual Au edits.")
            self._append_log(f"Created/updated Others sheet in: {self.state.data_workbook_path}")
            open_path(self.state.data_workbook_path)
        except Exception as exc:
            QMessageBox.critical(self, "Others sheet failed", str(exc))

    def _open_data_workbook(self) -> None:
        if self.state.data_workbook_path and self.state.data_workbook_path.exists():
            open_path(self.state.data_workbook_path)

    def _resize_microscope_images(self) -> None:
        selected = QFileDialog.getExistingDirectory(self, "Select microscope parent folder")
        if not selected:
            return
        image_list_dir = Path(selected) / "grains" / "image_list"
        if not image_list_dir.exists():
            QMessageBox.warning(self, "Missing image_list", f"Could not find:\n{image_list_dir}")
            return
        try:
            self.state.microscope_image_list_dir = image_list_dir
            self.state.microscope_resized_dir = resize_images_to_smallest(image_list_dir)
            self.image_status.setText(f"Microscope resized folder: {self.state.microscope_resized_dir}")
            self._append_log(f"Created microscope resized folder: {self.state.microscope_resized_dir}")
            self._refresh_summary()
        except Exception as exc:
            QMessageBox.critical(self, "Microscope resize failed", str(exc))

    def _resize_sem_images(self) -> None:
        selected = QFileDialog.getExistingDirectory(self, "Select SEM image folder")
        if not selected:
            return
        try:
            self.state.sem_image_dir = Path(selected)
            self.state.sem_resized_dir = resize_images_to_smallest(self.state.sem_image_dir)
            self.image_status.setText(f"SEM resized folder: {self.state.sem_resized_dir}")
            self._append_log(f"Created SEM resized folder: {self.state.sem_resized_dir}")
            self._refresh_summary()
        except Exception as exc:
            QMessageBox.critical(self, "SEM resize failed", str(exc))

    def _refresh_summary(self) -> None:
        rows = [
            ("Sample type", self.state.sample_type),
            ("Sample name", self.state.sample_name),
            ("Sample folder", self.state.sample_dir),
            ("Raw export", self.state.raw_export_path),
            ("Data workbook", self.state.data_workbook_path),
            ("Microscope image_list", self.state.microscope_image_list_dir),
            ("Microscope resized", self.state.microscope_resized_dir),
            ("SEM folder", self.state.sem_image_dir),
            ("SEM resized", self.state.sem_resized_dir),
        ]
        self.summary_table.setRowCount(len(rows))
        for row_number, (label, value) in enumerate(rows):
            label_item = QTableWidgetItem(str(label))
            value_item = QTableWidgetItem("" if value is None else str(value))
            label_item.setTextAlignment(ALIGN_CENTER)
            value_item.setTextAlignment(ALIGN_CENTER)
            self.summary_table.setItem(row_number, 0, label_item)
            self.summary_table.setItem(row_number, 1, value_item)

    def _append_log(self, message: str) -> None:
        timestamp = time.strftime("%H:%M:%S")
        if hasattr(self, "app_log"):
            self.app_log.appendPlainText(f"[{timestamp}] {message}")


def main() -> None:
    app = QApplication(sys.argv)
    window = AuReportAutomationApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
