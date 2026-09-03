"""Desktop launcher for the locally maintained EDS Word editing scripts."""

from __future__ import annotations

import argparse
import runpy
import sys
import traceback
from datetime import datetime
from pathlib import Path


APP_TITLE = "Editing EDS Exported Word Files"
APP_DIR = Path(__file__).resolve().parent
REPO_ROOT = APP_DIR.parent
FEATURE_SCRIPT = APP_DIR / "Formatting & Text Addition In Word File.py"
MINERAL_SCRIPT = APP_DIR / "WordFormatting_With_SEM&uScopeImages.py"
GRAIN_SCRIPT = APP_DIR / "Word Formatting with Grain Numbers.py"
CROP_SCRIPT = REPO_ROOT / "Crop Images" / "CropSelect.py"


def run_external_script(script: Path, files: list[str], directory: str | None) -> int:
    """Run a local script and supply paths already selected in the app."""
    import tkinter.filedialog as filedialog

    original_open = filedialog.askopenfilename
    original_opens = filedialog.askopenfilenames
    original_directory = filedialog.askdirectory
    try:
        if files:
            filedialog.askopenfilename = lambda *args, **kwargs: files[0]
            filedialog.askopenfilenames = lambda *args, **kwargs: tuple(files)
        if directory:
            filedialog.askdirectory = lambda *args, **kwargs: directory
        runpy.run_path(str(script), run_name="__main__")
        return 0
    except SystemExit as exc:
        return int(exc.code) if isinstance(exc.code, int) else (0 if exc.code is None else 1)
    except Exception:
        traceback.print_exc()
        return 1
    finally:
        filedialog.askopenfilename = original_open
        filedialog.askopenfilenames = original_opens
        filedialog.askdirectory = original_directory


def runner_main() -> int | None:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--run-script")
    parser.add_argument("--input", action="append", default=[])
    parser.add_argument("--directory")
    args, _ = parser.parse_known_args()
    if not args.run_script:
        return None
    return run_external_script(Path(args.run_script), args.input, args.directory)


runner_result = runner_main()
if runner_result is not None:
    raise SystemExit(runner_result)


try:
    from PySide6.QtCore import QProcess, QSettings, Qt, QUrl
    from PySide6.QtGui import QDesktopServices
    from PySide6.QtWidgets import (
        QApplication, QFileDialog, QFrame, QGroupBox, QHBoxLayout, QLabel,
        QMainWindow, QMessageBox, QPlainTextEdit, QPushButton, QTabWidget,
        QVBoxLayout, QWidget,
    )
except ImportError:
    try:
        from PySide2.QtCore import QProcess, QSettings, Qt, QUrl
        from PySide2.QtGui import QDesktopServices
        from PySide2.QtWidgets import (
            QApplication, QFileDialog, QFrame, QGroupBox, QHBoxLayout, QLabel,
            QMainWindow, QMessageBox, QPlainTextEdit, QPushButton, QTabWidget,
            QVBoxLayout, QWidget,
        )
    except ImportError as exc:
        raise SystemExit("Install PySide6 to run this app: pip install PySide6") from exc


class ActionCard(QGroupBox):
    def __init__(self, title: str):
        super().__init__(title)
        self.setObjectName("actionCard")
        self.row = QHBoxLayout(self)

    def button(self, text: str, callback, primary: bool = False) -> QPushButton:
        control = QPushButton(text)
        if primary:
            control.setObjectName("primaryButton")
        control.clicked.connect(callback)
        self.row.addWidget(control)
        return control


class EditingEDSApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = QSettings("SEMEDSuScope", "EditingEDSExportedWordFiles")
        self.process: QProcess | None = None
        self.running_label = ""
        self.working_directory = str(self.settings.value("working_directory", ""))
        self.setWindowTitle(APP_TITLE)
        self.resize(1040, 720)
        self._build_ui()

    def _build_ui(self) -> None:
        central = QWidget()
        layout = QVBoxLayout(central)

        header = QHBoxLayout()
        mark = QLabel("SEMEDS")
        mark.setObjectName("brandMark")
        header.addWidget(mark)
        title = QLabel(APP_TITLE)
        title.setObjectName("heading")
        header.addWidget(title)
        header.addStretch()
        self.run_status = QLabel("READY")
        self.run_status.setObjectName("readyStatus")
        header.addWidget(self.run_status)
        layout.addLayout(header)

        divider = QFrame()
        divider.setFrameShape(QFrame.HLine)
        layout.addWidget(divider)

        tabs = QTabWidget()
        tabs.addTab(self._feature_tab(), "Exps From Feature Scan")
        tabs.addTab(self._mineral_tab(), "Mineral Check")
        tabs.addTab(self._grain_tab(), "Jarosite Scan")
        layout.addWidget(tabs, 1)

        log_box = QGroupBox("Live Output")
        log_layout = QVBoxLayout(log_box)
        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumBlockCount(3000)
        self.log.setPlaceholderText("Script progress and completion messages appear here.")
        log_layout.addWidget(self.log)
        log_controls = QHBoxLayout()
        clear = QPushButton("Clear Output")
        clear.clicked.connect(self.log.clear)
        stop = QPushButton("Stop Running Script")
        stop.clicked.connect(self.stop_process)
        log_controls.addStretch()
        log_controls.addWidget(clear)
        log_controls.addWidget(stop)
        log_layout.addLayout(log_controls)
        layout.addWidget(log_box, 1)

        self.setCentralWidget(central)
        self.setStyleSheet("""
            QMainWindow { background: #f3f6f9; }
            QLabel#brandMark { background: #17324d; color: white; padding: 9px 13px; border-radius: 4px; font-size: 16px; font-weight: 800; }
            QLabel#heading { color: #17324d; font-size: 23px; font-weight: 700; padding-left: 8px; }
            QLabel#readyStatus { color: #176b3a; font-weight: 800; padding: 7px 12px; background: #e4f4e9; border-radius: 4px; }
            QLabel#workingPath { color: #435466; background: white; border: 1px solid #cbd5df; border-radius: 4px; padding: 9px; }
            QGroupBox { color: #17324d; font-weight: 700; }
            QGroupBox#actionCard { background: white; border: 1px solid #cbd5df; border-radius: 7px; margin-top: 13px; padding: 15px; }
            QGroupBox#actionCard::title { subcontrol-origin: margin; left: 12px; padding: 0 5px; }
            QPushButton { padding: 8px 13px; }
            QPushButton#primaryButton { background: #236aa1; color: white; border: 0; border-radius: 4px; font-weight: 700; }
            QPlainTextEdit { background: #17212b; color: #dce7ef; border: 0; font-family: Consolas, monospace; padding: 6px; }
            QTabBar::tab { padding: 9px 18px; }
            QTabBar::tab:selected { color: #236aa1; font-weight: 700; }
        """)

    def _page(self) -> tuple[QWidget, QVBoxLayout]:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 24, 24, 24)
        return page, layout

    def _feature_tab(self) -> QWidget:
        page, layout = self._page()
        card = ActionCard("Feature Scan Explanations")
        card.button("Select Word + Excel and Run", self.run_feature_scan, True)
        card.button("Open Script", lambda: self.open_path(FEATURE_SCRIPT))
        layout.addWidget(card)
        layout.addStretch()
        return page

    def _mineral_tab(self) -> QWidget:
        page, layout = self._page()
        crop = ActionCard("Step 1 — Crop uScope Images")
        crop.button("Select Image Folder and Crop", self.run_cropper, True)
        crop.button("Open Crop Script", lambda: self.open_path(CROP_SCRIPT))
        layout.addWidget(crop)
        self.directory_label = QLabel(self.working_directory or "No working folder selected")
        self.directory_label.setObjectName("workingPath")
        self.directory_label.setWordWrap(True)
        layout.addWidget(self.directory_label)
        word = ActionCard("Step 2 — Format Word File")
        word.button("Select AZtec Word File and Run", self.run_mineral_formatting, True)
        word.button("Open Working Folder", self.open_working_directory)
        word.button("Open Script", lambda: self.open_path(MINERAL_SCRIPT))
        layout.addWidget(word)
        layout.addStretch()
        return page

    def _grain_tab(self) -> QWidget:
        page, layout = self._page()
        card = ActionCard("Grain Number Formatting")
        card.button("Select AZtec Word File and Run", self.run_grain_formatting, True)
        card.button("Open Script", lambda: self.open_path(GRAIN_SCRIPT))
        layout.addWidget(card)
        layout.addStretch()
        return page

    def start_directory(self) -> str:
        return self.working_directory if Path(self.working_directory).is_dir() else str(Path.home())

    def run_feature_scan(self) -> None:
        word, _ = QFileDialog.getOpenFileName(self, "Select AZtec Word file", self.start_directory(), "Word Documents (*.docx)")
        if not word:
            return
        excel, _ = QFileDialog.getOpenFileName(self, "Select explanation Excel file", str(Path(word).parent), "Excel Workbooks (*.xlsx)")
        if not excel:
            return
        self.working_directory = str(Path(word).parent)
        self._remember_directory()
        self.run_script("Feature Scan", FEATURE_SCRIPT, [word, excel])

    def run_cropper(self) -> None:
        selected = QFileDialog.getExistingDirectory(self, "Select folder containing microscope JPG images", self.start_directory())
        if not selected:
            return
        self.working_directory = selected
        self._remember_directory()
        self.run_script("Crop Images", CROP_SCRIPT, directory=selected)

    def run_mineral_formatting(self) -> None:
        word, _ = QFileDialog.getOpenFileName(self, "Select AZtec Word file", self.start_directory(), "Word Documents (*.docx)")
        if not word:
            return
        selected_parent = str(Path(word).parent)
        if selected_parent != self.working_directory:
            self.append_log(f"Working folder updated to the Word file folder: {selected_parent}")
            self.working_directory = selected_parent
            self._remember_directory()
        self.run_script("Mineral Check", MINERAL_SCRIPT, [word])

    def run_grain_formatting(self) -> None:
        word, _ = QFileDialog.getOpenFileName(self, "Select AZtec Word file", self.start_directory(), "Word Documents (*.docx)")
        if word:
            self.working_directory = str(Path(word).parent)
            self._remember_directory()
            self.run_script("Jarosite Scan", GRAIN_SCRIPT, [word])

    def run_script(self, label: str, script: Path, inputs: list[str] | None = None, directory: str | None = None) -> None:
        if self.process is not None:
            QMessageBox.warning(self, "Script running", "Wait for the current script to finish or stop it.")
            return
        if not script.exists():
            QMessageBox.critical(self, "Script not found", f"The local script was not found:\n{script}")
            return
        arguments = ["-u", str(Path(__file__).resolve()), "--run-script", str(script)]
        for input_path in inputs or []:
            arguments.extend(["--input", input_path])
        if directory:
            arguments.extend(["--directory", directory])
        self.running_label = label
        self.process = QProcess(self)
        self.process.setWorkingDirectory(str(script.parent))
        self.process.setProcessChannelMode(QProcess.MergedChannels)
        self.process.readyReadStandardOutput.connect(self.read_output)
        self.process.finished.connect(self.process_finished)
        self.set_running(True)
        self.append_log(f"Starting {label}: {script.name}")
        self.process.start(sys.executable, arguments)

    def read_output(self) -> None:
        if not self.process:
            return
        output = bytes(self.process.readAllStandardOutput()).decode(errors="replace")
        for line in output.rstrip().splitlines():
            self.append_log(line)

    def process_finished(self, exit_code: int, _status) -> None:
        self.read_output()
        label = self.running_label
        if exit_code == 0:
            self.append_log(f"Completed {label} successfully.")
        else:
            self.append_log(f"ERROR: {label} stopped with exit code {exit_code}.")
            QMessageBox.critical(self, "Script failed", f"{label} stopped with exit code {exit_code}.\nReview Live Output for details.")
        self.process.deleteLater()
        self.process = None
        self.running_label = ""
        self.set_running(False)

    def stop_process(self) -> None:
        if self.process:
            self.append_log(f"Stopping {self.running_label}...")
            self.process.kill()

    def set_running(self, running: bool) -> None:
        self.run_status.setText("RUNNING" if running else "READY")
        self.run_status.setStyleSheet("color: #9a6500; background: #fff1cc;" if running else "")

    def append_log(self, message: str) -> None:
        self.log.appendPlainText(f"[{datetime.now().strftime('%H:%M:%S')}] {message}")

    def _remember_directory(self) -> None:
        self.settings.setValue("working_directory", self.working_directory)
        self.directory_label.setText(self.working_directory)

    def open_working_directory(self) -> None:
        if self.working_directory:
            self.open_path(Path(self.working_directory))

    def open_path(self, path: Path) -> None:
        if not path.exists():
            QMessageBox.warning(self, "Path unavailable", f"Could not find:\n{path}")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(path.resolve())))

    def closeEvent(self, event) -> None:
        if self.process:
            answer = QMessageBox.question(self, "Script running", "Stop the running script and close the app?")
            if answer != QMessageBox.Yes:
                event.ignore()
                return
            self.process.kill()
        event.accept()


def main() -> None:
    app = QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    window = EditingEDSApp()
    window.show()
    raise SystemExit(app.exec())


if __name__ == "__main__":
    main()
