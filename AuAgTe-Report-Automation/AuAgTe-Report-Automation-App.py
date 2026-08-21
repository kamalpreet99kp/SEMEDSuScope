"""Au-Ag-Te report automation workflow app.

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
import runpy
import sys
import traceback
from dataclasses import asdict, dataclass, fields
from datetime import datetime
from pathlib import Path
from typing import Callable, Optional

import pandas as pd


APP_TITLE = "Au-Ag-Te Report Automation"
PROJECT_FILE_NAME = "AuAgTe-Report-Project.json"
APP_DIR = Path(__file__).resolve().parent
REPO_ROOT = APP_DIR.parent
AG_LIBERATION_DIR = REPO_ROOT / "Ag-Liberation"
DUPLICATE_SCRIPT = REPO_ROOT / "Duplicate-Removal-in-Excel" / "HighlighDuplicatesForFullAnalysisData.py"
MICROSCOPE_PREP_SCRIPT = REPO_ROOT / "SEM-to-uScope-File-Prep" / "New-SEMtoMicroscopeFilePrep.py"
CROP_SCRIPT = REPO_ROOT / "Crop Images" / "Crop to Center_All Images at once.py"
MASTER_SCRIPT = AG_LIBERATION_DIR / "V3.Reports-SEMData-uScopeImagesCorelation-HidesColumns.py"
FINAL_SCRIPT = AG_LIBERATION_DIR / "Image Links_CorrFactor.py"


def run_external_script(script_path: Path, input_path: Path | None, directory: Path | None) -> int:
    """Run a legacy script while supplying known answers to its first pickers."""
    import tkinter.filedialog as filedialog

    original_open = filedialog.askopenfilename
    original_directory = filedialog.askdirectory

    if input_path:
        supplied_file = str(input_path)
        filedialog.askopenfilename = lambda *args, **kwargs: supplied_file
    if directory:
        supplied_directory = str(directory)
        filedialog.askdirectory = lambda *args, **kwargs: supplied_directory

    try:
        runpy.run_path(str(script_path), run_name="__main__")
        return 0
    except SystemExit as exc:
        return int(exc.code) if isinstance(exc.code, int) else (0 if exc.code is None else 1)
    except Exception:
        traceback.print_exc()
        return 1
    finally:
        filedialog.askopenfilename = original_open
        filedialog.askdirectory = original_directory


def runner_main() -> int | None:
    """Handle the private subprocess mode before loading a Qt binding."""
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--run-script")
    parser.add_argument("--input")
    parser.add_argument("--directory")
    args, _ = parser.parse_known_args()
    if not args.run_script:
        return None
    return run_external_script(
        Path(args.run_script),
        Path(args.input) if args.input else None,
        Path(args.directory) if args.directory else None,
    )


runner_result = runner_main()
if runner_result is not None:
    raise SystemExit(runner_result)


try:
    from PySide6.QtCore import QProcess, QSettings, Qt, QUrl
    from PySide6.QtGui import QDesktopServices
    from PySide6.QtWidgets import (
        QApplication,
        QComboBox,
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
        from PySide2.QtGui import QDesktopServices
        from PySide2.QtWidgets import (
            QApplication,
            QComboBox,
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
    updated_at: str = ""

    @classmethod
    def from_dict(cls, data: dict) -> "ProjectState":
        allowed = {item.name for item in fields(cls)}
        return cls(**{key: value for key, value in data.items() if key in allowed})


def qt_horizontal() -> object:
    return getattr(getattr(Qt, "Orientation", Qt), "Horizontal")


def path_exists(value: str) -> bool:
    return bool(value) and Path(value).exists()


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

    def __init__(self, title: str, description: str):
        super().__init__(title)
        self.setObjectName("stepCard")
        layout = QVBoxLayout(self)
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
        layout.addWidget(self.path_label)
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
        self.setWindowTitle(APP_TITLE)
        self.resize(1180, 860)
        self._build_ui()
        self._load_last_project()
        self.refresh_all()

    def _build_ui(self) -> None:
        central = QWidget()
        root = QVBoxLayout(central)

        heading_row = QHBoxLayout()
        heading = QLabel(APP_TITLE)
        heading.setObjectName("heading")
        heading_row.addWidget(heading)
        heading_row.addStretch()
        for label, callback in (("New Project", self.new_project), ("Open Project", self.open_project), ("Save Project", self.save_project)):
            button = QPushButton(label)
            button.clicked.connect(callback)
            heading_row.addWidget(button)
        root.addLayout(heading_row)

        self.sample_summary = QLabel("No sample loaded")
        self.sample_summary.setObjectName("summary")
        self.sample_summary.setWordWrap(True)
        root.addWidget(self.sample_summary)

        self.tabs = QTabWidget()
        self.phase_a_tab = self._build_phase_a()
        self.phase_b_tab = self._build_phase_b()
        self.tabs.addTab(self.phase_a_tab, "Phase A — SEM / AZtec Preparation")
        self.tabs.addTab(self.phase_b_tab, "Phase B — Microscope Images / Final Report")
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
            QLabel#heading { font-size: 24px; font-weight: 700; color: #17324d; }
            QLabel#summary { background: #e8f1fa; border: 1px solid #b7cde1; border-radius: 5px; padding: 9px; }
            QGroupBox#stepCard { background: white; border: 1px solid #cbd5df; border-radius: 7px; margin-top: 12px; padding: 12px; font-weight: 700; }
            QGroupBox#stepCard::title { subcontrol-origin: margin; left: 12px; padding: 0 5px; }
            QLabel#description { color: #435466; font-weight: 400; }
            QLabel#statusPending { color: #7a5a00; font-weight: 700; }
            QLabel#statusOk { color: #176b3a; font-weight: 700; }
            QLabel#statusError { color: #a32121; font-weight: 700; }
            QPushButton { padding: 6px 11px; }
            QPushButton#primaryButton { background: #236aa1; color: white; border: 0; border-radius: 4px; padding: 8px 13px; font-weight: 700; }
        """)

    def _scroll_page(self, content_layout: QVBoxLayout) -> QScrollArea:
        container = QWidget()
        container.setLayout(content_layout)
        area = QScrollArea()
        area.setWidgetResizable(True)
        area.setFrameShape(QFrame.NoFrame)
        area.setWidget(container)
        return area

    def _build_phase_a(self) -> QWidget:
        wrapper = QWidget()
        outer = QVBoxLayout(wrapper)
        content = QVBoxLayout()

        self.raw_card = StepCard("A1. Full Analysis → Excel (Files 1 and 2)", "Select the raw AZtec Full Analysis export. The sample name and Phase A directory are taken from it, and a new XLSX copy is created beside it.")
        self.raw_card.add_button("Select Full Analysis and Convert", self.select_and_convert_raw, True)
        self.raw_card.add_button("Use Existing File 2", self.select_existing_file_2)
        self.raw_card.add_button("Open File 2", lambda: self.open_state_path("excel_file_2"))
        self.raw_card.add_button("Open Phase A Folder", lambda: self.open_state_path("phase_a_directory"))
        content.addWidget(self.raw_card)

        self.category_card = StepCard("A2. Categorize Minerals (File 3)", "Choose the current project-specific categorization script from Ag-Liberation. It runs against File 2 and creates a separate categorized workbook.")
        category_row = QHBoxLayout()
        category_row.addWidget(QLabel("Categorization script:"))
        self.category_combo = QComboBox()
        category_row.addWidget(self.category_combo, 1)
        refresh = QPushButton("Refresh scripts")
        refresh.clicked.connect(self.refresh_category_scripts)
        category_row.addWidget(refresh)
        self.category_card.layout().insertLayout(2, category_row)
        self.category_card.add_button("Run Categorization", self.run_categorization, True)
        self.category_card.add_button("Use Existing File 3", self.select_existing_file_3)
        self.category_card.add_button("Open File 3", lambda: self.open_state_path("categorized_file_3"))
        self.category_card.add_button("Open Script", self.open_selected_category_script)
        content.addWidget(self.category_card)

        self.duplicates_card = StepCard("A3. Find and Review Duplicates (File 4)", "Run the fixed duplicate checker on File 3. It creates a DUPS_ONLY workbook; review it and manually remove confirmed duplicates from File 3.")
        self.duplicates_card.add_button("Run Duplicate Check", self.run_duplicates, True)
        self.duplicates_card.add_button("Use Existing File 4", lambda: self.select_existing_output("duplicates_file_4", "Select duplicates-only workbook"))
        self.duplicates_card.add_button("Open File 3", lambda: self.open_state_path("categorized_file_3"))
        self.duplicates_card.add_button("Open File 4", lambda: self.open_state_path("duplicates_file_4"))
        content.addWidget(self.duplicates_card)

        self.prep_card = StepCard("A4. Prepare Microscope Coordinates", "After manual duplicate cleanup, create SEM-pos-um.xlsx and Micro-After-Correction.csv beside File 3. These complete the transfer package for microscopy.")
        self.prep_card.add_button("Generate Coordinate Files", self.run_microscope_prep, True)
        self.prep_card.add_button("Use Existing Outputs", self.select_existing_prep_outputs)
        self.prep_card.add_button("Open SEM Positions", lambda: self.open_state_path("sem_positions_file"))
        self.prep_card.add_button("Open Correction CSV", lambda: self.open_state_path("correction_csv"))
        content.addWidget(self.prep_card)
        content.addStretch()
        outer.addWidget(self._scroll_page(content))
        return wrapper

    def _build_phase_b(self) -> QWidget:
        wrapper = QWidget()
        outer = QVBoxLayout(wrapper)
        content = QVBoxLayout()

        self.crop_card = StepCard("B1. Crop Microscope Images", "Select the transferred Phase B parent directory. The interactive cropper processes chosen category folders and creates cropped subfolders.")
        self.crop_card.add_button("Select Folder and Run Cropper", self.run_cropper, True)
        self.crop_card.add_button("Use Existing Phase B Folder", self.select_phase_b_directory)
        self.crop_card.add_button("Open Phase B Folder", lambda: self.open_state_path("phase_b_directory"))
        self.crop_card.add_button("Open Crop Script", lambda: self.open_path(CROP_SCRIPT))
        content.addWidget(self.crop_card)

        self.master_card = StepCard("B2. Build Master Workbook (File 5)", "Select the transferred categorized File 3 if it is not already known in Phase B. The script requires Micro-After-Correction.csv and cropped category folders beside it.")
        self.master_card.add_button("Build Master Workbook", self.run_master, True)
        self.master_card.add_button("Use Existing File 5", lambda: self.select_existing_output("master_file_5", "Select MASTER workbook"))
        self.master_card.add_button("Open File 5", lambda: self.open_state_path("master_file_5"))
        self.master_card.add_button("Open Debug Log", self.open_master_log)
        content.addWidget(self.master_card)

        self.final_card = StepCard("B3. Add Links, Correction Columns and Formulas (File 6)", "Run the final script against File 5. It adds quick image links, Corr %, Area Orig, New Brea, New Len and Asso while preserving File 5.")
        self.final_card.add_button("Create Final File", self.run_final, True)
        self.final_card.add_button("Use Existing File 6", lambda: self.select_existing_output("final_file_6", "Select final workbook"))
        self.final_card.add_button("Open File 5", lambda: self.open_state_path("master_file_5"))
        self.final_card.add_button("Open Final File 6", lambda: self.open_state_path("final_file_6"))
        content.addWidget(self.final_card)
        content.addStretch()
        outer.addWidget(self._scroll_page(content))
        return wrapper

    def append_log(self, message: str) -> None:
        stamp = datetime.now().strftime("%H:%M:%S")
        self.log.appendPlainText(f"[{stamp}] {message}")

    def refresh_category_scripts(self) -> None:
        previous = self.state.category_script or self.category_combo.currentData()
        self.category_combo.clear()
        scripts = []
        if AG_LIBERATION_DIR.exists():
            scripts = sorted(
                (path for path in AG_LIBERATION_DIR.glob("*.py") if "categor" in path.name.lower()),
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
        self.sample_summary.setText(
            f"Sample: {self.state.sample_name or 'Not selected'}    |    "
            f"Phase A: {self.state.phase_a_directory or 'Not selected'}    |    "
            f"Phase B: {self.state.phase_b_directory or 'Not selected'}"
        )
        self._refresh_card(self.raw_card, self.state.excel_file_2, "File 2 ready")
        self._refresh_card(self.category_card, self.state.categorized_file_3, "File 3 ready")
        self._refresh_card(self.duplicates_card, self.state.duplicates_file_4, "File 4 ready for manual review")
        prep_ready = path_exists(self.state.sem_positions_file) and path_exists(self.state.correction_csv)
        self.prep_card.path_label.setText(f"{self.state.sem_positions_file or 'SEM-pos-um.xlsx not selected'}\n{self.state.correction_csv or 'Micro-After-Correction.csv not selected'}")
        self.prep_card.set_state("Phase A coordinate package ready" if prep_ready else "Not completed", "ok" if prep_ready else "pending")
        cropped_count = self._cropped_folder_count()
        self.crop_card.path_label.setText(self.state.phase_b_directory or "No Phase B directory selected")
        self.crop_card.set_state(f"{cropped_count} cropped folder(s) detected" if cropped_count else "Not completed", "ok" if cropped_count else "pending")
        self._refresh_card(self.master_card, self.state.master_file_5, "File 5 ready")
        self._refresh_card(self.final_card, self.state.final_file_6, "Final File 6 ready")

    def _refresh_card(self, card: StepCard, value: str, complete_text: str) -> None:
        card.path_label.setText(value or "No output selected")
        exists = path_exists(value)
        card.set_state(complete_text if exists else "Not completed", "ok" if exists else "pending")

    def _cropped_folder_count(self) -> int:
        directory = Path(self.state.phase_b_directory) if self.state.phase_b_directory else None
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
        self.append_log(f"Created File 2: {output}")
        self.auto_save()

    def select_existing_file_2(self) -> None:
        selected = self.choose_excel("Select existing File 2", self.state.phase_a_directory)
        if selected:
            self.state.excel_file_2 = str(selected)
            self.state.phase_a_directory = str(selected.parent)
            if not self.state.sample_name:
                self.state.sample_name = selected.stem.removesuffix("_Excel")
            self.auto_save()

    def select_existing_file_3(self) -> None:
        selected = self.choose_excel("Select categorized File 3", self.state.phase_a_directory or self.state.phase_b_directory)
        if selected:
            self.state.categorized_file_3 = str(selected)
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

    def select_existing_prep_outputs(self) -> None:
        sem_file = self.choose_excel("Select SEM-pos-um.xlsx", self.state.phase_a_directory)
        if not sem_file:
            return
        csv_file, _ = QFileDialog.getOpenFileName(self, "Select Micro-After-Correction.csv", str(sem_file.parent), "CSV files (*.csv)")
        if not csv_file:
            return
        self.state.sem_positions_file = str(sem_file)
        self.state.correction_csv = csv_file
        self.state.phase_a_directory = str(sem_file.parent)
        self.auto_save()

    def select_phase_b_directory(self) -> None:
        selected = QFileDialog.getExistingDirectory(self, "Select Phase B parent directory", self.state.phase_b_directory or str(Path.home()))
        if selected:
            self.state.phase_b_directory = selected
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
        input_path = self._require_input("excel_file_2", "Select File 2")
        script_data = self.category_combo.currentData()
        if not input_path or not script_data:
            QMessageBox.warning(self, "Missing input", "Select File 2 and a categorization script first.")
            return
        script = Path(script_data)
        self.state.category_script = str(script)
        output = input_path.with_name(f"Classified_{input_path.name}")
        self._run_script("categorization", script, input_path, None, output)

    def run_duplicates(self) -> None:
        input_path = self._require_input("categorized_file_3", "Select categorized File 3")
        if input_path:
            output = input_path.with_name(f"{input_path.stem}_DUPS_ONLY.xlsx")
            self._run_script("duplicates", DUPLICATE_SCRIPT, input_path, None, output)

    def run_microscope_prep(self) -> None:
        input_path = self._require_input("categorized_file_3", "Select manually reviewed File 3")
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
        self.auto_save()
        self._run_script("crop", CROP_SCRIPT, None, Path(selected), None)

    def _find_phase_b_file_3(self) -> Path | None:
        known = Path(self.state.categorized_file_3) if path_exists(self.state.categorized_file_3) else None
        phase_b = Path(self.state.phase_b_directory) if self.state.phase_b_directory else None
        if known and phase_b and known.parent == phase_b:
            return known
        chosen = self.choose_excel("Select the categorized File 3 transferred to Phase B", self.state.phase_b_directory)
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
            QMessageBox.warning(self, "Missing correction file", f"Micro-After-Correction.csv was not found beside File 3:\n{input_path.parent}")
            return
        output = input_path.with_name(f"MASTER_{input_path.stem}.xlsx")
        if output.exists() and not self.confirm_replace([output, input_path.parent / "master_debug_log.txt"]):
            return
        self._run_script("master", MASTER_SCRIPT, input_path, None, output)

    def run_final(self) -> None:
        input_path = self._require_input("master_file_5", "Select master File 5", "b")
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

    def _run_script(self, step: str, script: Path, input_path: Path | None, directory: Path | None, expected_output: Path | None) -> None:
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
        self.pending_step = step
        self.pending_output = expected_output
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
                    self.append_log(line)

    def _process_finished(self, exit_code: int, _exit_status) -> None:
        step, output = self.pending_step, self.pending_output
        self._read_process_output()
        self.append_log(f"Step {step} finished with exit code {exit_code}")
        if exit_code == 0:
            if step == "categorization" and output and output.exists():
                self.state.categorized_file_3 = str(output)
            elif step == "duplicates" and output and output.exists():
                self.state.duplicates_file_4 = str(output)
            elif step == "microscope_prep" and output:
                self.state.sem_positions_file = str(output)
                self.state.correction_csv = str(output.parent / "Micro-After-Correction.csv")
            elif step == "master" and output and output.exists():
                self.state.master_file_5 = str(output)
                self.state.phase_b_directory = str(output.parent)
            elif step == "final" and output and output.exists():
                self.state.final_file_6 = str(output)
            self.auto_save()
        else:
            QMessageBox.critical(self, "Script failed", f"The {step} step exited with code {exit_code}. Review the activity log.")
        self.process.deleteLater()
        self.process = None
        self.pending_step = ""
        self.pending_output = None
        self.refresh_all()

    def open_selected_category_script(self) -> None:
        value = self.category_combo.currentData()
        if value:
            self.open_path(Path(value))

    def open_master_log(self) -> None:
        directory = Path(self.state.phase_b_directory) if self.state.phase_b_directory else None
        if directory:
            self.open_path(directory / "master_debug_log.txt")

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
        return None

    def save_project(self) -> None:
        default = self._project_path() or (Path.home() / PROJECT_FILE_NAME)
        selected, _ = QFileDialog.getSaveFileName(self, "Save workflow project", str(default), "Au-Ag-Te project (*.json)")
        if selected:
            self._write_state(Path(selected))
            self.settings.setValue("last_project", selected)
            self.append_log(f"Saved project state: {selected}")

    def open_project(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(self, "Open workflow project", str(Path.home()), "Au-Ag-Te project (*.json)")
        if selected:
            self._read_state(Path(selected))

    def _write_state(self, path: Path) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
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
