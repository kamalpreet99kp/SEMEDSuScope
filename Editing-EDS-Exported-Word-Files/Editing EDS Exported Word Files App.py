"""Desktop launcher for the locally maintained EDS Word editing scripts."""

from __future__ import annotations

import argparse
import runpy
import sys
import tempfile
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
    from PySide6.QtGui import QColor, QDesktopServices, QPixmap
    from PySide6.QtWidgets import (
        QApplication, QFileDialog, QFrame, QGroupBox, QHBoxLayout, QLabel,
        QDialog, QMainWindow, QMessageBox, QPlainTextEdit, QPushButton, QTabWidget,
        QVBoxLayout, QWidget,
    )
except ImportError:
    try:
        from PySide2.QtCore import QProcess, QSettings, Qt, QUrl
        from PySide2.QtGui import QColor, QDesktopServices, QPixmap
        from PySide2.QtWidgets import (
            QApplication, QFileDialog, QFrame, QGroupBox, QHBoxLayout, QLabel,
            QDialog, QMainWindow, QMessageBox, QPlainTextEdit, QPushButton, QTabWidget,
            QVBoxLayout, QWidget,
        )
    except ImportError as exc:
        raise SystemExit("Install PySide6 to run this app: pip install PySide6") from exc


def find_amtel_logo_path() -> Path | None:
    """Locate the shared AMTEL logo in the local checkout."""
    image_extensions = {".png", ".jpg", ".jpeg", ".bmp", ".webp"}
    search_directories = [
        APP_DIR,
        REPO_ROOT / "AuReportAutomation",
        REPO_ROOT / "AuAgTe-Report-Automation",
    ]
    candidates = [
        path
        for directory in search_directories
        if directory.exists()
        for path in directory.rglob("*")
        if path.is_file() and path.suffix.lower() in image_extensions
    ]

    def priority(path: Path) -> tuple[int, str]:
        name = path.stem.casefold().replace("_", " ").replace("-", " ")
        if "amtel" in name and "logo" in name:
            rank = 0
        elif "amtel" in name:
            rank = 1
        elif "logo" in name:
            rank = 2
        else:
            rank = 3
        return rank, str(path).casefold()

    return min(candidates, key=priority) if candidates else None


class ActionCard(QGroupBox):
    def __init__(self, title: str, explanation: str = "", accent: str = "#236aa1"):
        super().__init__(title)
        self.setObjectName("actionCard")
        self.setStyleSheet(f"QGroupBox#actionCard {{ border-top: 4px solid {accent}; }}")
        self.card_layout = QVBoxLayout(self)
        self.row = QHBoxLayout()
        self.card_layout.addLayout(self.row)
        if explanation:
            note = QLabel(explanation)
            note.setObjectName("explanation")
            self.card_layout.addWidget(note)

    def button(self, text: str, callback, primary: bool = False) -> QPushButton:
        control = QPushButton(text)
        if primary:
            control.setObjectName("primaryButton")
        control.clicked.connect(callback)
        self.row.addWidget(control)
        return control


class OrientationPreviewDialog(QDialog):
    """Compare corresponding cropout and BSE images before Word formatting."""

    def __init__(self, cropout_images: list[Path], bse_images: list[Path], rotate_callback, parent=None):
        super().__init__(parent)
        self.cropout_images = cropout_images
        self.bse_images = bse_images
        self.rotate_callback = rotate_callback
        self.pair_count = min(len(cropout_images), len(bse_images))
        self.current_pair = 0
        self.setWindowTitle("Check uScope Image Orientation")
        self.resize(1040, 680)

        layout = QVBoxLayout(self)
        self.pair_label = QLabel()
        self.pair_label.setObjectName("pairHeading")
        self.pair_label.setAlignment(getattr(Qt, "AlignCenter"))
        layout.addWidget(self.pair_label)

        image_row = QHBoxLayout()
        self.cropout_preview = self._image_panel("Cropped uScope Image", image_row)
        self.bse_preview = self._image_panel("BSE Image", image_row)
        layout.addLayout(image_row, 1)

        controls = QHBoxLayout()
        self.previous_button = QPushButton("← Previous Pair")
        self.previous_button.clicked.connect(self.previous_pair)
        self.next_button = QPushButton("Next Pair →")
        self.next_button.clicked.connect(self.next_pair)
        rotate_button = QPushButton("Rotate All uScope Images 90° Clockwise")
        rotate_button.setObjectName("rotateButton")
        rotate_button.clicked.connect(self.rotate_all)
        done_button = QPushButton("Orientation OK")
        done_button.setObjectName("primaryButton")
        done_button.clicked.connect(self.accept)
        controls.addWidget(self.previous_button)
        controls.addWidget(self.next_button)
        controls.addStretch()
        controls.addWidget(rotate_button)
        controls.addWidget(done_button)
        layout.addLayout(controls)

        self.setStyleSheet("""
            QDialog { background: #f4edfa; }
            QLabel#pairHeading { color: #56356d; font-size: 16px; font-weight: 700; padding: 6px; }
            QLabel#imageTitle { color: #435466; font-size: 14px; font-weight: 700; }
            QLabel#imagePreview { background: white; border: 1px solid #bcc7d1; border-radius: 5px; }
            QPushButton { padding: 8px 12px; }
            QPushButton#rotateButton { background: #8a5aa5; color: white; border: 0; border-radius: 4px; font-weight: 700; }
            QPushButton#primaryButton { background: #236aa1; color: white; border: 0; border-radius: 4px; font-weight: 700; }
        """)
        self.show_pair()

    def _image_panel(self, title: str, row: QHBoxLayout) -> QLabel:
        panel = QVBoxLayout()
        title_label = QLabel(title)
        title_label.setObjectName("imageTitle")
        title_label.setAlignment(getattr(Qt, "AlignCenter"))
        preview = QLabel()
        preview.setObjectName("imagePreview")
        preview.setMinimumSize(440, 480)
        preview.setAlignment(getattr(Qt, "AlignCenter"))
        panel.addWidget(title_label)
        panel.addWidget(preview, 1)
        row.addLayout(panel, 1)
        return preview

    def show_pair(self) -> None:
        self.pair_label.setText(f"Pair {self.current_pair + 1} of {self.pair_count}")
        self._show_image(self.cropout_preview, self.cropout_images[self.current_pair])
        self._show_image(self.bse_preview, self.bse_images[self.current_pair])
        self.previous_button.setEnabled(self.current_pair > 0)
        self.next_button.setEnabled(self.current_pair < self.pair_count - 1)

    @staticmethod
    def _show_image(label: QLabel, path: Path) -> None:
        pixmap = QPixmap(str(path))
        if pixmap.isNull():
            label.setText(f"Could not display\n{path.name}")
            return
        aspect = getattr(getattr(Qt, "AspectRatioMode", Qt), "KeepAspectRatio")
        transform = getattr(getattr(Qt, "TransformationMode", Qt), "SmoothTransformation")
        label.setPixmap(pixmap.scaled(430, 465, aspect, transform))
        label.setToolTip(str(path))

    def previous_pair(self) -> None:
        self.current_pair = max(0, self.current_pair - 1)
        self.show_pair()

    def next_pair(self) -> None:
        self.current_pair = min(self.pair_count - 1, self.current_pair + 1)
        self.show_pair()

    def rotate_all(self) -> None:
        if self.rotate_callback():
            self.show_pair()


class EditingEDSApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = QSettings("SEMEDSuScope", "EditingEDSExportedWordFiles")
        self.process: QProcess | None = None
        self.running_label = ""
        self.pending_output: Path | None = None
        self.feature_output: Path | None = None
        self.mineral_output: Path | None = None
        self.grain_output: Path | None = None
        saved_mineral_word = str(self.settings.value("mineral_word_file", ""))
        self.mineral_word_file = Path(saved_mineral_word) if saved_mineral_word else None
        self.mineral_orientation_confirmed = False
        self.working_directory = str(self.settings.value("working_directory", ""))
        self.setWindowTitle(APP_TITLE)
        self.resize(1040, 820)
        self._build_ui()

    def _build_ui(self) -> None:
        central = QWidget()
        layout = QVBoxLayout(central)

        header = QHBoxLayout()
        brand_panel = QWidget()
        brand_panel.setObjectName("brandPanel")
        brand_layout = QHBoxLayout(brand_panel)
        brand_layout.setContentsMargins(12, 5, 14, 5)
        self.logo_label = QLabel()
        self.logo_label.setObjectName("amtelLogo")
        self.logo_label.setFixedSize(250, 68)
        self.logo_label.setAlignment(getattr(Qt, "AlignCenter"))
        self._load_amtel_logo()
        brand_layout.addWidget(self.logo_label)
        title = QLabel(APP_TITLE)
        title.setObjectName("heading")
        brand_layout.addWidget(title)
        header.addWidget(brand_panel)
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
        tabs.tabBar().setTabTextColor(0, QColor("#176b5d"))
        tabs.tabBar().setTabTextColor(1, QColor("#7a4d9b"))
        tabs.tabBar().setTabTextColor(2, QColor("#9a641d"))
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
            QWidget#brandPanel { background: white; border: 1px solid #d3dde7; border-radius: 7px; }
            QLabel#amtelLogo { border: none; background: transparent; color: #153f66; font-size: 24px; font-weight: 800; }
            QLabel#heading { border: none; background: transparent; color: #425d75; font-size: 20px; font-weight: 700; }
            QLabel#readyStatus { color: #176b3a; font-weight: 800; padding: 7px 12px; background: #e4f4e9; border-radius: 4px; }
            QLabel#workingPath { color: #435466; background: white; border: 1px solid #cbd5df; border-radius: 4px; padding: 9px; }
            QGroupBox { color: #17324d; font-weight: 700; }
            QGroupBox#actionCard { background: white; border: 1px solid #cbd5df; border-radius: 7px; margin-top: 13px; padding: 15px; }
            QGroupBox#actionCard::title { subcontrol-origin: margin; left: 12px; padding: 0 5px; }
            QLabel#explanation { color: #66788a; font-size: 14px; font-style: italic; font-weight: 400; padding-top: 5px; }
            QPushButton { padding: 8px 13px; }
            QPushButton#primaryButton { background: #236aa1; color: white; border: 0; border-radius: 4px; font-weight: 700; }
            QPlainTextEdit { background: #e8ecef; color: #263746; border: 1px solid #c5ced6; font-family: Consolas, monospace; padding: 6px; }
            QTabBar::tab { min-width: 220px; padding: 10px 16px; font-size: 14px; font-weight: 700; background: #e5ebf1; border: 1px solid #c4ced8; }
            QTabBar::tab:selected { background: white; border-bottom-color: white; }
        """)

    def _load_amtel_logo(self) -> None:
        logo_path = find_amtel_logo_path()
        if logo_path is None:
            self.logo_label.setText("AMTEL")
            self.logo_label.setToolTip("AMTEL logo image was not found in the local checkout.")
            return
        pixmap = QPixmap(str(logo_path))
        if pixmap.isNull():
            self.logo_label.setText("AMTEL")
            return
        aspect = getattr(getattr(Qt, "AspectRatioMode", Qt), "KeepAspectRatio")
        transform = getattr(getattr(Qt, "TransformationMode", Qt), "SmoothTransformation")
        self.logo_label.setPixmap(pixmap.scaled(self.logo_label.size(), aspect, transform))
        self.logo_label.setToolTip(str(logo_path))

    def _page(self, color: str) -> tuple[QWidget, QVBoxLayout]:
        page = QWidget()
        page.setObjectName("workflowPage")
        page.setStyleSheet(f"QWidget#workflowPage {{ background-color: {color}; }}")
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 24, 24, 24)
        return page, layout

    def _feature_tab(self) -> QWidget:
        page, layout = self._page("#eaf7f2")
        card = ActionCard("", "Formats & adds Feature ID + Explanation.", "#2b8a78")
        card.button("Select Word+Excel File", self.run_feature_scan, True)
        card.button("Open File", lambda: self.open_output(self.feature_output))
        layout.addWidget(card)
        layout.addStretch()
        return page

    def _mineral_tab(self) -> QWidget:
        page, layout = self._page("#f4edfa")
        crop = ActionCard("Step 1 — Crop uScope Images", "Select and crop the required uScope image area.", "#8a5aa5")
        crop.button("Select Image Folder", self.run_cropper, True)
        crop.button("Open Folder", self.open_cropout_directory)
        layout.addWidget(crop)
        self.directory_label = QLabel(self.working_directory or "No working folder selected")
        self.directory_label.setObjectName("workingPath")
        self.directory_label.setWordWrap(True)
        layout.addWidget(self.directory_label)
        orientation = ActionCard(
            "Step 2 — Check Image Orientation",
            "Compare the cropped uScope images with the corresponding BSE images.",
            "#8a5aa5",
        )
        orientation.button("Select Word File && Check Orientation", self.check_mineral_orientation, True)
        layout.addWidget(orientation)
        word = ActionCard("Step 3 — Format Word File", "Formats & adds uScope images.", "#8a5aa5")
        word.button("Format Selected Word File", self.run_mineral_formatting, True)
        word.button("Open File", lambda: self.open_output(self.mineral_output))
        layout.addWidget(word)
        layout.addStretch()
        return page

    def _grain_tab(self) -> QWidget:
        page, layout = self._page("#fff6e5")
        card = ActionCard("", "Formats & adds Grain No.", "#d08a2e")
        card.button("Select Word File", self.run_grain_formatting, True)
        card.button("Open File", lambda: self.open_output(self.grain_output))
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
        self.feature_output = Path(word).with_name(f"Updated {Path(word).stem}.docx")
        self.run_script("Feature Scan", FEATURE_SCRIPT, [word, excel], expected_output=self.feature_output)

    def run_cropper(self) -> None:
        selected = QFileDialog.getExistingDirectory(self, "Select folder containing microscope JPG images", self.start_directory())
        if not selected:
            return
        self.working_directory = selected
        self.mineral_word_file = None
        self.mineral_output = None
        self.mineral_orientation_confirmed = False
        self.settings.remove("mineral_word_file")
        self._remember_directory()
        self.run_script("Crop Images", CROP_SCRIPT, directory=selected)

    def check_mineral_orientation(self) -> None:
        cropout_directory = Path(self.working_directory) / "cropout"
        cropout_images = self._image_files(cropout_directory)
        if not cropout_images:
            QMessageBox.warning(
                self,
                "Cropped images not found",
                "Complete Step 1 first. No supported images were found in:\n"
                f"{cropout_directory}",
            )
            return
        word, _ = QFileDialog.getOpenFileName(self, "Select AZtec Word file", self.start_directory(), "Word Documents (*.docx)")
        if not word:
            return
        if Path(word).parent.resolve() != Path(self.working_directory).resolve():
            QMessageBox.warning(
                self,
                "Word file is in a different folder",
                "Select the Word file located beside the cropout folder created in Step 1.\n\n"
                f"Expected folder:\n{self.working_directory}",
            )
            return
        self.mineral_word_file = Path(word)
        self.mineral_orientation_confirmed = False
        self.settings.setValue("mineral_word_file", word)
        self.mineral_output = self.mineral_word_file.with_name(
            f"Modified {self.mineral_word_file.stem}.docx"
        )

        try:
            with tempfile.TemporaryDirectory() as temporary_directory:
                helpers = runpy.run_path(str(MINERAL_SCRIPT), run_name="mineral_preview_helpers")
                ordered_images = helpers["extract_ordered_images"](word, temporary_directory)
                groups = helpers["group_by_bse"](ordered_images)
                bse_images = [Path(group["BSE"]) for group in groups if group["BSE"]]
                pair_count = min(len(cropout_images), len(bse_images))
                if pair_count == 0:
                    QMessageBox.warning(
                        self,
                        "BSE images not found",
                        "No BSE image groups could be identified in the selected Word file.",
                    )
                    return
                if len(cropout_images) != len(bse_images):
                    self.append_log(
                        f"Orientation check: {len(cropout_images)} uScope image(s), "
                        f"{len(bse_images)} BSE image(s); showing {pair_count} matched pair(s)."
                    )
                dialog = OrientationPreviewDialog(
                    cropout_images[:pair_count],
                    bse_images[:pair_count],
                    lambda: self.rotate_cropout_images(cropout_images),
                    self,
                )
                if hasattr(dialog, "exec"):
                    dialog_result = dialog.exec()
                else:
                    dialog_result = dialog.exec_()
                if not dialog_result:
                    self.append_log("Orientation check cancelled.")
                    return
        except Exception as exc:
            traceback.print_exc()
            QMessageBox.critical(self, "Orientation preview failed", str(exc))
            self.append_log(f"ERROR: Orientation preview failed: {exc}")
            return
        self.mineral_orientation_confirmed = True
        self.append_log(f"Orientation confirmed for: {self.mineral_word_file.name}")

    @staticmethod
    def _image_files(directory: Path) -> list[Path]:
        extensions = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp"}
        if not directory.is_dir():
            return []
        return sorted(
            (path for path in directory.iterdir() if path.is_file() and path.suffix.lower() in extensions),
            key=lambda path: path.name.casefold(),
        )

    def rotate_cropout_images(self, images: list[Path]) -> bool:
        answer = QMessageBox.question(
            self,
            "Rotate all cropped images",
            f"Rotate all {len(images)} images in the cropout folder 90° clockwise?",
        )
        if answer != QMessageBox.Yes:
            return False
        try:
            from PIL import Image, ImageOps

            for image_path in images:
                temporary_path = image_path.with_name(
                    f".{image_path.stem}.rotating{image_path.suffix}"
                )
                try:
                    with Image.open(image_path) as image:
                        rotated = ImageOps.exif_transpose(image).rotate(-90, expand=True)
                        if image_path.suffix.lower() in {".jpg", ".jpeg"} and rotated.mode not in ("RGB", "L"):
                            rotated = rotated.convert("RGB")
                        rotated.save(temporary_path)
                    temporary_path.replace(image_path)
                finally:
                    if temporary_path.exists():
                        temporary_path.unlink()
        except Exception as exc:
            QMessageBox.critical(self, "Rotation failed", str(exc))
            self.append_log(f"ERROR: Could not rotate all cropout images: {exc}")
            return False
        self.append_log(f"Rotated {len(images)} cropout image(s) 90° clockwise.")
        return True

    def run_mineral_formatting(self) -> None:
        if (
            not self.mineral_word_file
            or not self.mineral_word_file.is_file()
            or not self.mineral_orientation_confirmed
        ):
            QMessageBox.information(
                self,
                "Word file not selected",
                "Complete Step 2 and check the image orientation first.",
            )
            return
        cropout_directory = Path(self.working_directory) / "cropout"
        if not self._image_files(cropout_directory):
            QMessageBox.warning(self, "Cropped images not found", f"No images were found in:\n{cropout_directory}")
            return
        self.mineral_output = self.mineral_word_file.with_name(
            f"Modified {self.mineral_word_file.stem}.docx"
        )
        self.run_script(
            "Mineral Check",
            MINERAL_SCRIPT,
            [str(self.mineral_word_file)],
            expected_output=self.mineral_output,
        )

    def run_grain_formatting(self) -> None:
        word, _ = QFileDialog.getOpenFileName(self, "Select AZtec Word file", self.start_directory(), "Word Documents (*.docx)")
        if word:
            self.working_directory = str(Path(word).parent)
            self._remember_directory()
            self.grain_output = Path(word).with_name(f"Modified {Path(word).stem}.docx")
            self.run_script("Jarosite Scan", GRAIN_SCRIPT, [word], expected_output=self.grain_output)

    def run_script(self, label: str, script: Path, inputs: list[str] | None = None, directory: str | None = None, expected_output: Path | None = None) -> None:
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
        self.pending_output = expected_output
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
            if self.pending_output:
                if self.pending_output.exists():
                    self.append_log(f"Output ready: {self.pending_output}")
                else:
                    self.append_log(f"WARNING: Expected output was not found: {self.pending_output}")
        else:
            self.append_log(f"ERROR: {label} stopped with exit code {exit_code}.")
            QMessageBox.critical(self, "Script failed", f"{label} stopped with exit code {exit_code}.\nReview Live Output for details.")
        self.process.deleteLater()
        self.process = None
        self.running_label = ""
        self.pending_output = None
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

    def open_cropout_directory(self) -> None:
        if not self.working_directory:
            QMessageBox.information(self, "Folder unavailable", "Run Step 1 first.")
            return
        self.open_path(Path(self.working_directory) / "cropout")

    def open_output(self, output: Path | None) -> None:
        if output is None:
            QMessageBox.information(self, "File unavailable", "Run this step first.")
            return
        self.open_path(output)

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
