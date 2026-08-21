"""Au/Ag-Te report automation workflow app.

The app coordinates the existing, locally maintained laboratory scripts without
copying their mineralogical rules into the GUI.  Each external script remains
usable on its own; when the app already knows an input path it temporarily
supplies that path to the script's Tk file picker.
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import runpy
import sys
import time
import traceback
from dataclasses import asdict, dataclass, fields
from datetime import datetime
from pathlib import Path
from typing import Callable, Optional

import pandas as pd


APP_TITLE = "Au/Ag-Te Report Automation"
PROJECT_FILE_NAME = "AuAgTe-Report-Project.json"
APP_DIR = Path(__file__).resolve().parent
REPO_ROOT = APP_DIR.parent
AG_LIBERATION_DIR = REPO_ROOT / "Ag-Liberation"
DUPLICATE_SCRIPT = REPO_ROOT / "Duplicate-Removal-in-Excel" / "HighlighDuplicatesForFullAnalysisData.py"
MICROSCOPE_PREP_SCRIPT = REPO_ROOT / "SEM-to-uScope-File-Prep" / "New-SEMtoMicroscopeFilePrep.py"
CROP_SCRIPT = REPO_ROOT / "Crop Images" / "Crop to Center_All Images at once.py"
MASTER_SCRIPT = AG_LIBERATION_DIR / "V3.Reports-SEMData-uScopeImagesCorelation-HidesColumns.py"
FINAL_SCRIPT = AG_LIBERATION_DIR / "Image Links_CorrFactor.py"
APPENDIX_SCRIPT = REPO_ROOT / "Editing-EDS-Exported-Word-Files" / "Arrange_Microscope_Images_By_Mineral.py"
TERNARY_SCRIPT = AG_LIBERATION_DIR / "Ag-Sb-As-Ternary-Plots.py"
PICKED_DIRECTORY_MARKER = "__AUAGTE_PICKED_DIRECTORY__="


def run_external_script(
    script_path: Path,
    input_path: Path | None,
    directory: Path | None,
    crop_value: float | None = None,
    suppress_messageboxes: bool = False,
) -> int:
    """Run a legacy script while supplying known answers to its first pickers."""
    import tkinter.filedialog as filedialog
    import tkinter.messagebox as messagebox

    original_open = filedialog.askopenfilename
    original_directory = filedialog.askdirectory
    original_message_functions = {
        name: getattr(messagebox, name)
        for name in ("showinfo", "showwarning", "showerror")
    }

    if input_path:
        supplied_file = str(input_path)
        filedialog.askopenfilename = lambda *args, **kwargs: supplied_file
    if directory:
        supplied_directory = str(directory)
        filedialog.askdirectory = lambda *args, **kwargs: supplied_directory
    else:
        def recorded_directory_picker(*args, **kwargs):
            selected_directory = original_directory(*args, **kwargs)
            if selected_directory:
                print(f"{PICKED_DIRECTORY_MARKER}{selected_directory}", flush=True)
            return selected_directory

        filedialog.askdirectory = recorded_directory_picker

    if suppress_messageboxes:
        def log_messagebox(title, message, **_kwargs):
            print(f"{title}: {message}", flush=True)
            return "ok"

        for function_name in original_message_functions:
            setattr(messagebox, function_name, log_messagebox)

    try:
        if crop_value is None:
            runpy.run_path(str(script_path), run_name="__main__")
        else:
            source = script_path.read_text(encoding="utf-8")
            replacement = rf"\g<1>{crop_value:.2f}\g<2>"
            source, replacement_count = re.subn(
                r"(?m)^(\s*crop_value\s*=\s*)(?:\d+(?:\.\d*)?|\.\d+)(\s*(?:#.*)?)$",
                replacement,
                source,
                count=1,
            )
            if replacement_count != 1:
                raise RuntimeError(f"Could not find crop_value in {script_path.name}.")
            script_globals = {
                "__name__": "__main__",
                "__file__": str(script_path),
                "__package__": None,
                "__cached__": None,
            }
            exec(compile(source, str(script_path), "exec"), script_globals)
        return 0
    except SystemExit as exc:
        return int(exc.code) if isinstance(exc.code, int) else (0 if exc.code is None else 1)
    except Exception:
        traceback.print_exc()
        return 1
    finally:
        filedialog.askopenfilename = original_open
        filedialog.askdirectory = original_directory
        for function_name, original_function in original_message_functions.items():
            setattr(messagebox, function_name, original_function)


def runner_main() -> int | None:
    """Handle the private subprocess mode before loading a Qt binding."""
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--run-script")
    parser.add_argument("--input")
    parser.add_argument("--directory")
    parser.add_argument("--crop-value", type=float)
    parser.add_argument("--suppress-messageboxes", action="store_true")
    args, _ = parser.parse_known_args()
    if not args.run_script:
        return None
    return run_external_script(
        Path(args.run_script),
        Path(args.input) if args.input else None,
        Path(args.directory) if args.directory else None,
        args.crop_value,
        args.suppress_messageboxes,
    )


runner_result = runner_main()
if runner_result is not None:
    raise SystemExit(runner_result)


try:
    from PySide6.QtCore import QProcess, QSettings, Qt, QUrl
    from PySide6.QtGui import QColor, QDesktopServices, QPixmap
    from PySide6.QtWidgets import (
        QApplication,
        QComboBox,
        QDoubleSpinBox,
        QFileDialog,
        QFrame,
        QGroupBox,
        QHBoxLayout,
        QLabel,
        QMainWindow,
        QMessageBox,
        QPlainTextEdit,
        QPushButton,
        QScrollArea,
        QTabWidget,
        QVBoxLayout,
        QWidget,
    )
except ImportError:
    try:
        from PySide2.QtCore import QProcess, QSettings, Qt, QUrl
        from PySide2.QtGui import QColor, QDesktopServices, QPixmap
        from PySide2.QtWidgets import (
            QApplication,
            QComboBox,
            QDoubleSpinBox,
            QFileDialog,
            QFrame,
            QGroupBox,
            QHBoxLayout,
            QLabel,
            QMainWindow,
            QMessageBox,
            QPlainTextEdit,
            QPushButton,
            QScrollArea,
            QTabWidget,
            QVBoxLayout,
            QWidget,
        )
    except ImportError as exc:
        raise SystemExit("Install PySide6 to run the app: pip install PySide6") from exc


@dataclass
class ProjectState:
    sample_name: str = ""
    phase_a_directory: str = ""
    raw_file_1: str = ""
    excel_file_2: str = ""
    category_script: str = ""
    categorized_file_3: str = ""
    duplicates_file_4: str = ""
    sem_positions_file: str = ""
    correction_csv: str = ""
    phase_b_directory: str = ""
    master_file_5: str = ""
    final_file_6: str = ""
    phase_b_crop_value: float = 0.4
    phase_c_appendix_directory: str = ""
    phase_c_crop_value: float = 0.4
    appendix_file: str = ""
    ternary_directory: str = ""
    ternary_html: str = ""
    updated_at: str = ""

    @classmethod
    def from_dict(cls, data: dict) -> "ProjectState":
        allowed = {item.name for item in fields(cls)}
        return cls(**{key: value for key, value in data.items() if key in allowed})


def qt_horizontal() -> object:
    return getattr(getattr(Qt, "Orientation", Qt), "Horizontal")


def path_exists(value: str) -> bool:
    return bool(value) and Path(value).exists()


def find_amtel_logo_path() -> Path | None:
    """Find the AMTEL image previously supplied for the Au Automation app.

    The logo is intentionally loaded from the local checkout rather than copied
    or recreated. This lets both automation apps use the same image on the lab
    computer, even when the logo asset is not tracked by Git.
    """
    supported_extensions = {".png", ".jpg", ".jpeg", ".bmp", ".webp"}
    search_directories = [APP_DIR, REPO_ROOT / "AuReportAutomation"]
    candidates: list[Path] = []

    for directory in search_directories:
        if not directory.exists():
            continue
        candidates.extend(
            path
            for path in directory.rglob("*")
            if path.is_file() and path.suffix.lower() in supported_extensions
        )

    def logo_priority(path: Path) -> tuple[int, str]:
        normalized_name = path.stem.lower().replace("_", " ").replace("-", " ")
        if "amtel" in normalized_name and "logo" in normalized_name:
            priority = 0
        elif "amtel" in normalized_name:
            priority = 1
        elif "logo" in normalized_name:
            priority = 2
        else:
            priority = 3
        return priority, str(path).lower()

    return min(candidates, key=logo_priority) if candidates else None


def unique_excel_output(source: Path) -> Path:
    candidate = source.with_name(f"{source.stem}_Excel.xlsx")
    if candidate.resolve() != source.resolve() and not candidate.exists():
        return candidate
    counter = 2
    while True:
        candidate = source.with_name(f"{source.stem}_Excel_{counter}.xlsx")
        if not candidate.exists():
            return candidate
        counter += 1


def read_delimited_file(source: Path) -> pd.DataFrame:
    for encoding in ("utf-8-sig", "utf-16", "cp1252"):
        try:
            text = source.read_text(encoding=encoding)
        except UnicodeError:
            continue
        try:
            dialect = csv.Sniffer().sniff(text[:8192], delimiters=",;\t|")
            separator = dialect.delimiter
        except csv.Error:
            separator = "\t" if "\t" in text[:8192] else ","
        return pd.read_csv(source, sep=separator, encoding=encoding)
    raise ValueError("The Full Analysis file encoding could not be read.")


def convert_full_analysis(source: Path) -> Path:
    """Create a non-destructive XLSX copy of an AZtec Full Analysis export."""
    destination = unique_excel_output(source)
    suffix = source.suffix.lower()
    if suffix in {".xlsx", ".xlsm", ".xls"}:
        sheets = pd.read_excel(source, sheet_name=None)
        with pd.ExcelWriter(destination, engine="openpyxl") as writer:
            for sheet_name, frame in sheets.items():
                frame.to_excel(writer, sheet_name=str(sheet_name)[:31], index=False)
    else:
        frame = read_delimited_file(source)
        frame.to_excel(destination, sheet_name="Raw Data", index=False)
    return destination


class StepCard(QGroupBox):
    """A consistent workflow card with status, paths, and action buttons."""

    def __init__(self, title: str, description: str = "", accent_color: str = "#236aa1"):
        super().__init__(title)
        self.setObjectName("stepCard")
        self.setStyleSheet(
            "QGroupBox#stepCard {"
            f"border-top: 4px solid {accent_color};"
            "}"
        )
        layout = QVBoxLayout(self)
        if description:
            description_label = QLabel(description)
            description_label.setWordWrap(True)
            description_label.setObjectName("description")
            layout.addWidget(description_label)
        self.status = QLabel("Not started")
        self.status.setWordWrap(True)
        self.status.setObjectName("statusPending")
        layout.addWidget(self.status)
        self.path_label = QLabel("No output selected")
        self.path_label.setWordWrap(True)
        self.path_label.setTextInteractionFlags(getattr(Qt, "TextSelectableByMouse"))
        # Paths remain available to the workflow controller and activity log,
        # but are intentionally hidden to keep each step card uncluttered.
        self.path_label.hide()
        self.controls = QHBoxLayout()
        layout.addLayout(self.controls)

    def add_button(self, text: str, callback: Callable, primary: bool = False) -> QPushButton:
        button = QPushButton(text)
        if primary:
            button.setObjectName("primaryButton")
        button.clicked.connect(callback)
        self.controls.addWidget(button)
        return button

    def set_state(self, text: str, state: str = "pending") -> None:
        self.status.setText(text)
        self.status.setObjectName({"ok": "statusOk", "error": "statusError"}.get(state, "statusPending"))
        self.status.style().unpolish(self.status)
        self.status.style().polish(self.status)


class AuAgTeReportAutomationApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.state = ProjectState()
        self.settings = QSettings("SEMEDSuScope", "AuAgTeReportAutomation")
        self.process: QProcess | None = None
        self.pending_step: str = ""
        self.pending_output: Path | None = None
        self.pending_input_path: Path | None = None
        self.pending_directory: Path | None = None
        self.pending_selected_directories: list[Path] = []
        self.pending_started_at: float = 0.0
        self.setWindowTitle(APP_TITLE)
        self.resize(1180, 860)
        self._build_ui()
        self._load_last_project()
        self.refresh_all()

    def _build_ui(self) -> None:
        central = QWidget()
        root = QVBoxLayout(central)

        heading_row = QHBoxLayout()
        brand_panel = QWidget()
        brand_panel.setObjectName("brandPanel")
        brand_layout = QHBoxLayout(brand_panel)
        brand_layout.setContentsMargins(12, 5, 14, 5)
        brand_layout.setSpacing(12)

        self.logo_label = QLabel()
        self.logo_label.setObjectName("amtelLogo")
        self.logo_label.setFixedSize(250, 68)
        self.logo_label.setAlignment(getattr(Qt, "AlignCenter"))
        self._load_amtel_logo()
        brand_layout.addWidget(self.logo_label)

        app_name = QLabel("Au/Ag-Te Report Automation")
        app_name.setObjectName("appName")
        brand_layout.addWidget(app_name)
        heading_row.addWidget(brand_panel)
        heading_row.addStretch()

        self.sample_label = QLabel("Sample: Not selected")
        self.sample_label.setObjectName("sampleName")
        self.sample_label.setToolTip("Sample name taken from the selected Full Analysis filename.")
        heading_row.addWidget(self.sample_label)
        heading_row.addSpacing(12)

        for label, callback in (("New Project", self.new_project), ("Open Project", self.open_project), ("Save Project", self.save_project)):
            button = QPushButton(label)
            button.clicked.connect(callback)
            heading_row.addWidget(button)
        root.addLayout(heading_row)

        self.tabs = QTabWidget()
        self.phase_a_tab = self._build_phase_a()
        self.phase_b_tab = self._build_phase_b()
        self.phase_c_tab = self._build_phase_c()
        self.tabs.addTab(self.phase_a_tab, "Phase A-Before uScope Acq.")
        self.tabs.addTab(self.phase_b_tab, "Phase B-After uScope Acq.")
        self.tabs.addTab(self.phase_c_tab, "Phase C-Appendices && Ternary Plots")
        self.tabs.tabBar().setTabTextColor(0, QColor("#176b5d"))
        self.tabs.tabBar().setTabTextColor(1, QColor("#7a4d9b"))
        self.tabs.tabBar().setTabTextColor(2, QColor("#9a641d"))
        root.addWidget(self.tabs, 1)

        log_group = QGroupBox("Activity log")
        log_layout = QVBoxLayout(log_group)
        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumBlockCount(3000)
        log_layout.addWidget(self.log)
        root.addWidget(log_group, 0)

        self.setCentralWidget(central)
        self.setStyleSheet("""
            QMainWindow { background: #f3f6f9; }
            QWidget#brandPanel { background: white; border: 1px solid #d3dde7; border-radius: 7px; }
            QLabel#amtelLogo { border: none; background: transparent; color: #153f66; font-size: 24px; font-weight: 800; }
            QLabel#appName { border: none; background: transparent; color: #425d75; font-size: 20px; font-weight: 700; }
            QLabel#sampleName { color: #204b70; background: #edf4fa; border: 1px solid #b9ccdc; border-radius: 6px; padding: 8px 12px; font-size: 13px; font-weight: 700; }
            QGroupBox#stepCard { background: white; border: 1px solid #cbd5df; border-radius: 8px; margin-top: 18px; padding: 15px 12px 12px 12px; font-weight: 700; font-size: 15px; }
            QGroupBox#stepCard::title { subcontrol-origin: margin; subcontrol-position: top left; left: 14px; top: 1px; padding: 2px 8px; background: white; color: #183b59; }
            QLabel#description { color: #435466; font-weight: 400; }
            QLabel#cropExplanation { color: #657587; font-size: 12px; font-weight: 400; }
            QLabel#statusPending { color: #7a5a00; font-weight: 700; }
            QLabel#statusOk { color: #176b3a; font-weight: 700; }
            QLabel#statusError { color: #a32121; font-weight: 700; }
            QPushButton { padding: 6px 11px; }
            QPushButton#primaryButton { background: #236aa1; color: white; border: 0; border-radius: 4px; padding: 8px 13px; font-weight: 700; }
            QTabBar::tab { min-width: 245px; padding: 10px 18px; font-weight: 700; background: #e5ebf1; border: 1px solid #c4ced8; }
            QTabBar::tab:selected { background: white; border-bottom-color: white; }
        """)

    def _load_amtel_logo(self) -> None:
        logo_path = find_amtel_logo_path()
        if logo_path is None:
            self.logo_label.setText("AMTEL")
            self.logo_label.setToolTip(
                "AMTEL logo image was not found in AuReportAutomation or "
                "AuAgTe-Report-Automation."
            )
            return

        pixmap = QPixmap(str(logo_path))
        if pixmap.isNull():
            self.logo_label.setText("AMTEL")
            self.logo_label.setToolTip(f"Could not load AMTEL logo: {logo_path}")
            return

        aspect_mode = getattr(getattr(Qt, "AspectRatioMode", Qt), "KeepAspectRatio")
        transform_mode = getattr(getattr(Qt, "TransformationMode", Qt), "SmoothTransformation")
        self.logo_label.setPixmap(
            pixmap.scaled(
                self.logo_label.size(),
                aspect_mode,
                transform_mode,
            )
        )
        self.logo_label.setToolTip(str(logo_path))

    def _scroll_page(self, content_layout: QVBoxLayout, background_color: str) -> QScrollArea:
        container = QWidget()
        container.setObjectName("phaseContent")
        container.setLayout(content_layout)
        area = QScrollArea()
        area.setWidgetResizable(True)
        area.setFrameShape(QFrame.NoFrame)
        area.setWidget(container)
        area.setStyleSheet(
            "QScrollArea { border: none;"
            f" background-color: {background_color};"
            "}"
            "QScrollArea > QWidget > QWidget {"
            f" background-color: {background_color};"
            "}"
            "QWidget#phaseContent {"
            f" background-color: {background_color};"
            "}"
        )
        return area

    def _add_crop_control(self, card: StepCard, phase: str) -> QDoubleSpinBox:
        row = QHBoxLayout()
        label = QLabel("Crop fraction (0.1–1.0):")
        label.setToolTip(
            "Fraction of the original image width and height retained around "
            "the selected crop centre."
        )
        row.addWidget(label)
        spin_box = QDoubleSpinBox()
        spin_box.setRange(0.1, 1.0)
        spin_box.setSingleStep(0.1)
        spin_box.setDecimals(1)
        spin_box.setValue(0.4)
        spin_box.setToolTip(label.toolTip())
        row.addWidget(spin_box)
        explanation = QLabel("1.0 keeps the complete image; smaller values crop more tightly.")
        explanation.setObjectName("cropExplanation")
        row.addWidget(explanation)
        row.addStretch()
        card.layout().insertLayout(card.layout().count() - 1, row)
        spin_box.valueChanged.connect(
            lambda value: setattr(self.state, f"phase_{phase}_crop_value", float(value))
        )
        return spin_box

    def _build_phase_a(self) -> QWidget:
        wrapper = QWidget()
        outer = QVBoxLayout(wrapper)
        content = QVBoxLayout()

        self.raw_card = StepCard("A1. Full Analysis → XLSX", accent_color="#2b8a78")
        self.raw_card.add_button("Select Full Analysis File", self.select_and_convert_raw, True)
        self.raw_card.add_button("Open XLSX File", lambda: self.open_state_path("excel_file_2"))
        content.addWidget(self.raw_card)

        self.category_card = StepCard("A2. Categorize Minerals", accent_color="#2b8a78")
        category_row = QHBoxLayout()
        category_row.addWidget(QLabel("Categorization script:"))
        self.category_combo = QComboBox()
        category_row.addWidget(self.category_combo, 1)
        refresh = QPushButton("Refresh scripts")
        refresh.clicked.connect(self.refresh_category_scripts)
        category_row.addWidget(refresh)
        self.category_card.layout().insertLayout(1, category_row)
        self.category_card.add_button("Run Categorization", self.run_categorization, True)
        self.category_card.add_button("Use Existing File", self.select_existing_file_3)
        self.category_card.add_button("Open File", lambda: self.open_state_path("categorized_file_3"))
        content.addWidget(self.category_card)

        self.duplicates_card = StepCard("A3. Find Duplicate && Review", accent_color="#2b8a78")
        self.duplicates_card.add_button("Run Duplicate Check", self.run_duplicates, True)
        self.duplicates_card.add_button("Select Existing Categorized File", self.select_existing_file_3)
        self.duplicates_card.add_button("Open File", lambda: self.open_state_path("duplicates_file_4"))
        content.addWidget(self.duplicates_card)

        self.prep_card = StepCard("A4. Prepare SEM-to-uScope Files", accent_color="#2b8a78")
        self.prep_card.add_button("Generate Coordinate Files", self.run_microscope_prep, True)
        self.prep_card.add_button("Use Existing Categorized File", self.select_existing_file_3)
        self.prep_card.add_button("Open Files", self.open_coordinate_files)
        content.addWidget(self.prep_card)
        content.addStretch()
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(self._scroll_page(content, "#eaf7f2"))
        return wrapper

    def _build_phase_b(self) -> QWidget:
        wrapper = QWidget()
        outer = QVBoxLayout(wrapper)
        content = QVBoxLayout()

        self.crop_card = StepCard("B1. Crop uScope Images", "Select parent directory with all categories/folders", "#8a5aa5")
        self.phase_b_crop_spin = self._add_crop_control(self.crop_card, "b")
        self.crop_card.add_button("Select Parent Directory", self.run_cropper, True)
        self.crop_card.add_button("Open Parent Directory", lambda: self.open_state_path("phase_b_directory"))
        content.addWidget(self.crop_card)

        self.master_card = StepCard("B2. Master Workbook: EDS Data + uScope Images", 'This step requires Original "Categorized File" & "Micro-After-Correction.csv" in the parent directory', "#8a5aa5")
        self.master_card.add_button("Build Master Workbook", self.run_master, True)
        self.master_card.add_button("Open File", lambda: self.open_state_path("master_file_5"))
        content.addWidget(self.master_card)

        self.final_card = StepCard("B3. Add Image Links && Formulas", accent_color="#8a5aa5")
        self.final_card.add_button("Create Final File", self.run_final, True)
        self.final_card.add_button("Use Existing Master Workbook", lambda: self.select_existing_output("master_file_5", "Select master workbook"))
        self.final_card.add_button("Open File", lambda: self.open_state_path("final_file_6"))
        content.addWidget(self.final_card)
        content.addStretch()
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(self._scroll_page(content, "#f4edfa"))
        return wrapper

    def _build_phase_c(self) -> QWidget:
        wrapper = QWidget()
        outer = QVBoxLayout(wrapper)
        content = QVBoxLayout()

        self.appendix_crop_card = StepCard(
            "C1. Crop Appendix Images",
            "Select parent directory with category folders containing the manually selected images",
            "#d08a2e",
        )
        self.phase_c_crop_spin = self._add_crop_control(self.appendix_crop_card, "c")
        self.appendix_crop_card.add_button("Run Cropper", self.run_appendix_cropper, True)
        content.addWidget(self.appendix_crop_card)

        self.appendix_card = StepCard("C2. Create Appendices", accent_color="#d08a2e")
        self.appendix_card.add_button("Create Appendices", self.run_appendix_script, True)
        self.appendix_card.add_button("Open File", lambda: self.open_state_path("appendix_file"))
        content.addWidget(self.appendix_card)

        self.ternary_card = StepCard("C3. Create Ternary Plots", accent_color="#d08a2e")
        self.ternary_card.add_button("Create Ternary Plots", self.run_ternary_script, True)
        self.ternary_card.add_button("Open File", lambda: self.open_state_path("ternary_html"))
        content.addWidget(self.ternary_card)

        content.addStretch()
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(self._scroll_page(content, "#fff6e5"))
        return wrapper

    def append_log(self, message: str) -> None:
        stamp = datetime.now().strftime("%H:%M:%S")
        self.log.appendPlainText(f"[{stamp}] {message}")

    def refresh_category_scripts(self) -> None:
        previous = self.state.category_script or self.category_combo.currentData()
        self.category_combo.clear()
        scripts = []
        if AG_LIBERATION_DIR.exists():
            excluded_scripts = {
                MASTER_SCRIPT.resolve(),
                FINAL_SCRIPT.resolve(),
                TERNARY_SCRIPT.resolve(),
            }
            scripts = sorted(
                (
                    path
                    for path in AG_LIBERATION_DIR.glob("*.py")
                    if path.resolve() not in excluded_scripts
                ),
                key=lambda path: path.name.lower(),
            )
        for path in scripts:
            self.category_combo.addItem(path.stem, str(path))
        if previous:
            for index in range(self.category_combo.count()):
                if self.category_combo.itemData(index) == str(previous):
                    self.category_combo.setCurrentIndex(index)
                    break

    def refresh_all(self) -> None:
        self.refresh_category_scripts()
        self.sample_label.setText(f"Sample: {self.state.sample_name or 'Not selected'}")
        self.phase_b_crop_spin.blockSignals(True)
        self.phase_b_crop_spin.setValue(float(self.state.phase_b_crop_value))
        self.phase_b_crop_spin.blockSignals(False)
        self.phase_c_crop_spin.blockSignals(True)
        self.phase_c_crop_spin.setValue(float(self.state.phase_c_crop_value))
        self.phase_c_crop_spin.blockSignals(False)
        self._refresh_card(self.raw_card, self.state.excel_file_2)
        self._refresh_card(self.category_card, self.state.categorized_file_3)
        self._refresh_card(self.duplicates_card, self.state.duplicates_file_4)
        prep_ready = path_exists(self.state.sem_positions_file) and path_exists(self.state.correction_csv)
        self.prep_card.path_label.setText(f"{self.state.sem_positions_file or 'SEM-pos-um.xlsx not selected'}\n{self.state.correction_csv or 'Micro-After-Correction.csv not selected'}")
        self.prep_card.set_state("Status: Completed" if prep_ready else "Status: Not completed", "ok" if prep_ready else "pending")
        cropped_count = self._cropped_folder_count(self.state.phase_b_directory)
        self.crop_card.path_label.setText(self.state.phase_b_directory or "No Phase B directory selected")
        self.crop_card.set_state("Status: Completed" if cropped_count else "Status: Not completed", "ok" if cropped_count else "pending")
        self._refresh_card(self.master_card, self.state.master_file_5)
        self._refresh_card(self.final_card, self.state.final_file_6)
        appendix_cropped_count = self._cropped_folder_count(self.state.phase_c_appendix_directory)
        self.appendix_crop_card.set_state(
            "Status: Completed" if appendix_cropped_count else "Status: Not completed",
            "ok" if appendix_cropped_count else "pending",
        )
        self._refresh_card(self.appendix_card, self.state.appendix_file)
        self._refresh_card(self.ternary_card, self.state.ternary_html)

    def _refresh_card(self, card: StepCard, value: str) -> None:
        card.path_label.setText(value or "No output selected")
        exists = path_exists(value)
        card.set_state("Status: Completed" if exists else "Status: Not completed", "ok" if exists else "pending")

    def _cropped_folder_count(self, directory_value: str) -> int:
        directory = Path(directory_value) if directory_value else None
        if not directory or not directory.exists():
            return 0
        return sum(1 for child in directory.iterdir() if child.is_dir() and (child / "cropped").is_dir())

    def select_and_convert_raw(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(self, "Select AZtec Full Analysis file", self.state.phase_a_directory or str(Path.home()), "Full Analysis / data files (*.*)")
        if not selected:
            return
        source = Path(selected)
        try:
            QApplication.setOverrideCursor(getattr(Qt, "WaitCursor"))
            output = convert_full_analysis(source)
        except Exception as exc:
            QMessageBox.critical(self, "Conversion failed", str(exc))
            return
        finally:
            QApplication.restoreOverrideCursor()
        self.state.raw_file_1 = str(source)
        self.state.sample_name = source.stem
        self.state.phase_a_directory = str(source.parent)
        self.state.excel_file_2 = str(output)
        self.append_log(f"Created XLSX file: {output}")
        self.auto_save()

    def select_existing_file_3(self) -> None:
        selected = self.choose_excel("Select Categorization File", self.state.phase_a_directory or self.state.phase_b_directory)
        if selected:
            self.state.categorized_file_3 = str(selected)
            self.state.duplicates_file_4 = ""
            self.state.sem_positions_file = ""
            self.state.correction_csv = ""
            if not self.state.phase_a_directory:
                self.state.phase_a_directory = str(selected.parent)
            self.auto_save()

    def select_existing_output(self, field_name: str, title: str) -> None:
        selected = self.choose_excel(title, self.state.phase_b_directory or self.state.phase_a_directory)
        if selected:
            setattr(self.state, field_name, str(selected))
            if field_name in {"master_file_5", "final_file_6"}:
                self.state.phase_b_directory = str(selected.parent)
            self.auto_save()

    def choose_excel(self, title: str, start: str = "") -> Path | None:
        selected, _ = QFileDialog.getOpenFileName(self, title, start or str(Path.home()), "Excel workbooks (*.xlsx *.xlsm *.xls)")
        return Path(selected) if selected else None

    def _require_input(self, field_name: str, title: str, phase: str = "a") -> Path | None:
        current = getattr(self.state, field_name)
        if path_exists(current):
            return Path(current)
        start = self.state.phase_a_directory if phase == "a" else self.state.phase_b_directory
        chosen = self.choose_excel(title, start)
        if chosen:
            setattr(self.state, field_name, str(chosen))
            if phase == "b":
                self.state.phase_b_directory = str(chosen.parent)
            self.auto_save()
        return chosen

    def run_categorization(self) -> None:
        input_path = self._require_input("excel_file_2", "Select Full Analysis XLSX File")
        script_data = self.category_combo.currentData()
        if not input_path or not script_data:
            QMessageBox.warning(self, "Missing input", "Select the Full Analysis XLSX file and a categorization script first.")
            return
        script = Path(script_data)
        self.state.category_script = str(script)
        output = input_path.with_name(f"Classified_{input_path.name}")
        self._run_script("categorization", script, input_path, None, output)

    def run_duplicates(self) -> None:
        input_path = self._require_input("categorized_file_3", "Select Categorization File")
        if input_path:
            output = input_path.with_name(f"{input_path.stem}_DUPS_ONLY.xlsx")
            self._run_script("duplicates", DUPLICATE_SCRIPT, input_path, None, output)

    def run_microscope_prep(self) -> None:
        input_path = self._require_input("categorized_file_3", "Select Manually Reviewed Categorization File")
        if not input_path:
            return
        outputs = [input_path.parent / "SEM-pos-um.xlsx", input_path.parent / "Micro-After-Correction.csv"]
        if any(path.exists() for path in outputs) and not self.confirm_replace(outputs):
            return
        self._run_script("microscope_prep", MICROSCOPE_PREP_SCRIPT, input_path, None, outputs[0])

    def run_cropper(self) -> None:
        selected = QFileDialog.getExistingDirectory(self, "Select Phase B parent directory", self.state.phase_b_directory or str(Path.home()))
        if not selected:
            return
        self.state.phase_b_directory = selected
        self.state.phase_b_crop_value = float(self.phase_b_crop_spin.value())
        self.auto_save()
        self._run_script(
            "crop",
            CROP_SCRIPT,
            None,
            Path(selected),
            None,
            crop_value=self.state.phase_b_crop_value,
        )

    def run_appendix_cropper(self) -> None:
        selected = QFileDialog.getExistingDirectory(
            self,
            "Select appendix parent directory",
            self.state.phase_c_appendix_directory or str(Path.home()),
        )
        if not selected:
            return
        self.state.phase_c_appendix_directory = selected
        self.state.phase_c_crop_value = float(self.phase_c_crop_spin.value())
        self.state.appendix_file = ""
        self.auto_save()
        self._run_script(
            "appendix_crop",
            CROP_SCRIPT,
            None,
            Path(selected),
            None,
            crop_value=self.state.phase_c_crop_value,
        )

    def run_appendix_script(self) -> None:
        if not self.state.phase_c_appendix_directory:
            QMessageBox.warning(
                self,
                "Appendix images not selected",
                "Run the Phase C crop step or select its parent directory first.",
            )
            return
        self._run_script("appendix", APPENDIX_SCRIPT, None, None, None)

    def run_ternary_script(self) -> None:
        selected = QFileDialog.getExistingDirectory(
            self,
            "Select directory containing ternary-plot CSV files",
            self.state.ternary_directory or str(Path.home()),
        )
        if not selected:
            return
        directory = Path(selected)
        self.state.ternary_directory = selected
        self.state.ternary_html = ""
        self.auto_save()
        expected_output = directory / "Combined (Ternary Plot).html"
        self._run_script(
            "ternary",
            TERNARY_SCRIPT,
            None,
            directory,
            expected_output,
            suppress_messageboxes=True,
        )

    def _find_phase_b_file_3(self) -> Path | None:
        known = Path(self.state.categorized_file_3) if path_exists(self.state.categorized_file_3) else None
        phase_b = Path(self.state.phase_b_directory) if self.state.phase_b_directory else None
        if known and phase_b and known.parent == phase_b:
            return known
        if known and phase_b and phase_b.is_dir():
            transferred_copy = phase_b / known.name
            if transferred_copy.exists():
                self.state.categorized_file_3 = str(transferred_copy)
                self.append_log(f"Auto-detected Categorization File: {transferred_copy}")
                self.auto_save()
                return transferred_copy
        if phase_b and phase_b.is_dir():
            candidates = [
                path
                for path in phase_b.glob("*.xlsx")
                if path.is_file()
                and ("classif" in path.name.casefold() or "categor" in path.name.casefold())
                and "dups_only" not in path.name.casefold()
                and not path.name.casefold().startswith("master_")
                and "with_quick_links" not in path.name.casefold()
            ]
            if candidates:
                detected = max(candidates, key=lambda path: path.stat().st_mtime)
                self.state.categorized_file_3 = str(detected)
                self.append_log(f"Auto-detected Categorization File: {detected}")
                self.auto_save()
                return detected
        chosen = self.choose_excel("Select the Categorization File transferred to Phase B", self.state.phase_b_directory)
        if chosen:
            self.state.categorized_file_3 = str(chosen)
            self.state.phase_b_directory = str(chosen.parent)
            self.auto_save()
        return chosen

    def run_master(self) -> None:
        input_path = self._find_phase_b_file_3()
        if not input_path:
            return
        correction = input_path.parent / "Micro-After-Correction.csv"
        if not correction.exists():
            QMessageBox.warning(self, "Missing correction file", f"Micro-After-Correction.csv was not found beside the Categorization File:\n{input_path.parent}")
            return
        output = input_path.with_name(f"MASTER_{input_path.stem}.xlsx")
        if output.exists() and not self.confirm_replace([output, input_path.parent / "master_debug_log.txt"]):
            return
        self._run_script("master", MASTER_SCRIPT, input_path, None, output)

    def run_final(self) -> None:
        input_path = self._require_input("master_file_5", "Select Master Workbook", "b")
        if not input_path:
            return
        output = input_path.with_name(f"{input_path.stem}_with_quick_links{input_path.suffix}")
        if output.exists() and not self.confirm_replace([output]):
            return
        self._run_script("final", FINAL_SCRIPT, input_path, None, output)

    def confirm_replace(self, paths: list[Path]) -> bool:
        existing = [str(path) for path in paths if path.exists()]
        if not existing:
            return True
        answer = QMessageBox.question(self, "Existing output", "The following output already exists and may be replaced:\n\n" + "\n".join(existing) + "\n\nContinue?")
        return answer == QMessageBox.Yes

    def _run_script(
        self,
        step: str,
        script: Path,
        input_path: Path | None,
        directory: Path | None,
        expected_output: Path | None,
        crop_value: float | None = None,
        suppress_messageboxes: bool = False,
    ) -> None:
        if self.process is not None:
            QMessageBox.warning(self, "Process running", "Wait for the current script to finish.")
            return
        if not script.exists():
            QMessageBox.critical(self, "Script missing", f"Could not find:\n{script}")
            return
        arguments = [str(Path(__file__).resolve()), "--run-script", str(script)]
        if input_path:
            arguments += ["--input", str(input_path)]
        if directory:
            arguments += ["--directory", str(directory)]
        if crop_value is not None:
            if not 0.1 <= crop_value <= 1.0:
                QMessageBox.critical(self, "Invalid crop fraction", "Crop fraction must be between 0.1 and 1.0.")
                return
            arguments += ["--crop-value", f"{crop_value:.1f}"]
        if suppress_messageboxes:
            arguments.append("--suppress-messageboxes")
        self.pending_step = step
        self.pending_output = expected_output
        self.pending_input_path = input_path
        self.pending_directory = directory
        self.pending_selected_directories = []
        self.pending_started_at = time.time()
        self.process = QProcess(self)
        self.process.setProcessChannelMode(QProcess.MergedChannels)
        self.process.readyReadStandardOutput.connect(self._read_process_output)
        self.process.finished.connect(self._process_finished)
        self.append_log(f"Starting {script.name}")
        self.process.start(sys.executable, arguments)

    def _read_process_output(self) -> None:
        if self.process:
            text = bytes(self.process.readAllStandardOutput()).decode(errors="replace").rstrip()
            if text:
                for line in text.splitlines():
                    if line.startswith(PICKED_DIRECTORY_MARKER):
                        selected_directory = line[len(PICKED_DIRECTORY_MARKER):].strip()
                        if selected_directory:
                            self.pending_selected_directories.append(Path(selected_directory))
                        continue
                    self.append_log(line)

    def _process_finished(self, exit_code: int, _exit_status) -> None:
        step, output = self.pending_step, self.pending_output
        self._read_process_output()
        self.append_log(f"Step {step} finished with exit code {exit_code}")
        if exit_code == 0:
            if step == "categorization":
                categorization_output = output if output and output.exists() else self._find_recent_output(
                    "*.xlsx",
                    [self.pending_input_path.parent] if self.pending_input_path else [],
                    excluded_paths=[self.pending_input_path] if self.pending_input_path else [],
                )
                if categorization_output:
                    self.state.categorized_file_3 = str(categorization_output)
                    self.append_log(f"Recorded Categorization File: {categorization_output}")
            elif step == "duplicates" and output and output.exists():
                self.state.duplicates_file_4 = str(output)
            elif step == "microscope_prep" and output:
                self.state.sem_positions_file = str(output)
                self.state.correction_csv = str(output.parent / "Micro-After-Correction.csv")
            elif step == "master":
                master_output = output if output and output.exists() else self._find_recent_output(
                    "*.xlsx",
                    [self.pending_input_path.parent] if self.pending_input_path else [],
                    excluded_paths=[self.pending_input_path] if self.pending_input_path else [],
                )
                if master_output:
                    self.state.master_file_5 = str(master_output)
                    self.state.phase_b_directory = str(master_output.parent)
                    self.append_log(f"Recorded Master Workbook: {master_output}")
            elif step == "final" and output and output.exists():
                self.state.final_file_6 = str(output)
            elif step == "appendix":
                appendix_output = self._find_recent_output("*.docx", self.pending_selected_directories)
                if appendix_output:
                    self.state.appendix_file = str(appendix_output)
            elif step == "ternary":
                ternary_output = output if output and output.exists() else self._find_recent_output(
                    "*.html",
                    [Path(self.state.ternary_directory)] if self.state.ternary_directory else [],
                )
                if ternary_output:
                    self.state.ternary_html = str(ternary_output)
            self.auto_save()
        else:
            QMessageBox.critical(self, "Script failed", f"The {step} step exited with code {exit_code}. Review the activity log.")
        self.process.deleteLater()
        self.process = None
        self.pending_step = ""
        self.pending_output = None
        self.pending_input_path = None
        self.pending_directory = None
        self.pending_selected_directories = []
        self.refresh_all()

    def _find_recent_output(
        self,
        pattern: str,
        directories: list[Path],
        excluded_paths: list[Path | None] | None = None,
    ) -> Path | None:
        excluded = {
            path.resolve()
            for path in (excluded_paths or [])
            if path is not None
        }
        candidates = []
        for directory in reversed(directories):
            if not directory.exists():
                continue
            try:
                candidates.extend(
                    path
                    for path in directory.glob(pattern)
                    if path.is_file() and path.resolve() not in excluded
                )
            except OSError:
                continue
        if not candidates:
            return None
        recent_candidates = [
            path
            for path in candidates
            if path.stat().st_mtime >= self.pending_started_at - 2
        ]
        return max(recent_candidates or candidates, key=lambda path: path.stat().st_mtime)

    def open_coordinate_files(self) -> None:
        """Open both outputs produced by the SEM-to-uScope preparation step."""
        opened = False
        missing = []
        for value in (self.state.sem_positions_file, self.state.correction_csv):
            if value and Path(value).exists():
                QDesktopServices.openUrl(QUrl.fromLocalFile(str(Path(value).resolve())))
                opened = True
            else:
                missing.append(value or "Output path not recorded")
        if missing:
            message = "The following coordinate output(s) could not be found:\n\n" + "\n".join(missing)
            QMessageBox.warning(self, "Coordinate files unavailable", message)
        elif opened:
            self.append_log("Opened SEM positions workbook and correction CSV.")

    def open_state_path(self, field_name: str) -> None:
        value = getattr(self.state, field_name)
        if value:
            self.open_path(Path(value))
        else:
            QMessageBox.information(self, "Not selected", "No path is recorded for this item.")

    def open_path(self, path: Path) -> None:
        if not path.exists():
            QMessageBox.warning(self, "Path unavailable", f"Could not find:\n{path}")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(path.resolve())))

    def new_project(self) -> None:
        if QMessageBox.question(self, "New project", "Clear the current workflow state?") == QMessageBox.Yes:
            self.state = ProjectState()
            self.log.clear()
            self.refresh_all()

    def _project_path(self) -> Path | None:
        if self.state.phase_a_directory:
            return Path(self.state.phase_a_directory) / PROJECT_FILE_NAME
        if self.state.phase_b_directory:
            return Path(self.state.phase_b_directory) / PROJECT_FILE_NAME
        if self.state.phase_c_appendix_directory:
            return Path(self.state.phase_c_appendix_directory) / PROJECT_FILE_NAME
        if self.state.ternary_directory:
            return Path(self.state.ternary_directory) / PROJECT_FILE_NAME
        return None

    def save_project(self) -> None:
        default = self._project_path() or (Path.home() / PROJECT_FILE_NAME)
        selected, _ = QFileDialog.getSaveFileName(self, "Save workflow project", str(default), "Au/Ag-Te project (*.json)")
        if selected:
            self._write_state(Path(selected))
            self.settings.setValue("last_project", selected)
            self.append_log(f"Saved project state: {selected}")

    def open_project(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(self, "Open workflow project", str(Path.home()), "Au/Ag-Te project (*.json)")
        if selected:
            self._read_state(Path(selected))

    def _write_state(self, path: Path) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        self.state.phase_b_crop_value = float(self.phase_b_crop_spin.value())
        self.state.phase_c_crop_value = float(self.phase_c_crop_spin.value())
        self.state.updated_at = datetime.now().isoformat(timespec="seconds")
        path.write_text(json.dumps(asdict(self.state), indent=2), encoding="utf-8")

    def _read_state(self, path: Path) -> None:
        try:
            self.state = ProjectState.from_dict(json.loads(path.read_text(encoding="utf-8")))
        except (OSError, ValueError, TypeError) as exc:
            QMessageBox.critical(self, "Project could not be opened", str(exc))
            return
        self.settings.setValue("last_project", str(path))
        self.append_log(f"Opened project state: {path}")
        self.refresh_all()

    def auto_save(self) -> None:
        project_path = self._project_path()
        if project_path:
            try:
                self._write_state(project_path)
                self.settings.setValue("last_project", str(project_path))
            except OSError as exc:
                self.append_log(f"Project auto-save warning: {exc}")
        self.refresh_all()

    def _load_last_project(self) -> None:
        value = self.settings.value("last_project", "")
        if value and Path(str(value)).exists():
            self._read_state(Path(str(value)))

    def closeEvent(self, event) -> None:
        if self.process is not None:
            answer = QMessageBox.question(self, "Script running", "A workflow script is still running. Close the app and stop it?")
            if answer != QMessageBox.Yes:
                event.ignore()
                return
            self.process.kill()
        project_path = self._project_path()
        if project_path:
            try:
                self._write_state(project_path)
            except OSError:
                pass
        event.accept()


def main() -> None:
    app = QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    window = AuAgTeReportAutomationApp()
    window.show()
    raise SystemExit(app.exec())


if __name__ == "__main__":
    main()
