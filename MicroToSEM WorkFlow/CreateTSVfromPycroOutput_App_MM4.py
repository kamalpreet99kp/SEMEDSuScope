import argparse
import csv
import re
import shutil
from dataclasses import dataclass
from pathlib import Path


# Retained calibration; confirm this value for the MM4 objective/camera setup.
UM_PER_PIXEL = 0.1182  # MM3/MM2 = 0.1182, MM1 = 0.1665
umperpixel = UM_PER_PIXEL
IMAGE_SUFFIXES = {".jpg", ".jpeg"}
DEFAULT_PROJECT_DIRECTORY = r"\\amtelmain\amtelmain\SEM\Marked grains for SEM\current"
MINERAL_CHECK_DIRECTORY_NAME = "mineral check"


@dataclass(frozen=True)
class PycroTsvSettings:
    um_per_pixel: float = UM_PER_PIXEL


@dataclass(frozen=True)
class PositionRecord:
    x: float
    y: float
    image_paths: list[str]
    bbox: tuple[float, float, float, float]

    def as_csv_row(self):
        return {
            "X": self.x,
            "Y": self.y,
            "Image Paths": repr(self.image_paths),
        }


def _image_paths_in_directory(directory):
    return sorted(
        path
        for path in Path(directory).iterdir()
        if path.is_file() and path.suffix.lower() in IMAGE_SUFFIXES
    )


def _parse_crop_image(image_path, settings):
    image_path = Path(image_path).expanduser().resolve()
    filename = image_path.name

    roi_match = re.search(
        r"Roi(-?\d+)_(-?\d+)_(-?\d+)_(-?\d+)\.(?:jpg|jpeg)$",
        filename,
        flags=re.IGNORECASE,
    )
    if roi_match is None:
        raise ValueError(f"Could not parse ROI coordinates from {filename!r}")

    pos_match = re.search(
        r"Pos-?\d+_-?\d+_([+-]?\d+(?:\.\d+)?)_([+-]?\d+(?:\.\d+)?)_F",
        filename,
    )
    if pos_match is None:
        raise ValueError(f"Could not parse microscope coordinates from {filename!r}")

    x1, y1, x2, y2 = (int(value) for value in roi_match.groups())
    stage_x, stage_y = (float(value) for value in pos_match.groups())
    mid_x = (x1 + x2) / 2
    mid_y = (y1 + y2) / 2
    record_x = stage_x + mid_x * settings.um_per_pixel
    # MM4 image Y increases opposite to stage Y.
    record_y = stage_y - mid_y * settings.um_per_pixel
    bbox = (
        stage_x + x1 * settings.um_per_pixel,
        min(stage_y - y1 * settings.um_per_pixel,
            stage_y - y2 * settings.um_per_pixel),
        stage_x + x2 * settings.um_per_pixel,
        max(stage_y - y1 * settings.um_per_pixel,
            stage_y - y2 * settings.um_per_pixel),
    )
    return PositionRecord(
        x=record_x,
        y=record_y,
        image_paths=[str(image_path)],
        bbox=bbox,
    )


def build_position_records(directory, settings=None):
    settings = settings or PycroTsvSettings()
    directory = Path(directory).expanduser()
    if not directory.exists():
        raise FileNotFoundError(f"Directory not found: {directory}")
    if not directory.is_dir():
        raise NotADirectoryError(f"Expected a directory: {directory}")

    image_paths = _image_paths_in_directory(directory)
    if not image_paths:
        raise ValueError(f"No JPG images found in {directory}")

    return [_parse_crop_image(image_path, settings) for image_path in image_paths]


def write_position_records(records, output_path):
    output_path = Path(output_path).expanduser()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8", newline="") as file:
        writer = csv.DictWriter(
            file,
            fieldnames=["X", "Y", "Image Paths"],
            lineterminator="\n",
        )
        writer.writeheader()
        for record in records:
            writer.writerow(record.as_csv_row())
    return output_path


def generate_list_csv(directory, output_path=None, settings=None):
    directory = Path(directory).expanduser()
    output_path = Path(output_path).expanduser() if output_path else directory / "list_mm4.csv"
    records = build_position_records(directory, settings=settings)
    write_position_records(records, output_path)
    return records, output_path


def write_image_list(directory, image_paths=None):
    directory = Path(directory).expanduser()
    if image_paths is None:
        image_paths = _image_paths_in_directory(directory)
    else:
        image_paths = list(image_paths)

    image_list_directory = directory / "image_list_mm4"
    image_list_directory.mkdir(parents=True, exist_ok=True)

    for existing_path in image_list_directory.iterdir():
        if existing_path.is_file() and existing_path.suffix.lower() in IMAGE_SUFFIXES:
            existing_path.unlink()

    for index, image_path in enumerate(image_paths, start=1):
        image_path = Path(image_path)
        shutil.copyfile(
            image_path,
            image_list_directory / f"{index:03d}{image_path.suffix.lower()}",
        )

    return image_list_directory, len(image_paths)


def _first_image_paths_from_records(records):
    image_paths = []
    for index, record in enumerate(records, start=1):
        if not record.image_paths:
            raise ValueError(f"Position record {index} does not have an image path.")
        image_paths.append(record.image_paths[0])
    return image_paths


def _find_optional_subdirectory(project_directory, directory_name):
    exact_path = project_directory / directory_name
    if exact_path.exists() and exact_path.is_dir():
        return exact_path
    if not project_directory.exists() or not project_directory.is_dir():
        return exact_path

    target_name = directory_name.casefold()
    for path in project_directory.iterdir():
        if path.is_dir() and path.name.casefold() == target_name:
            return path
    return exact_path


def _generate_project_subdirectory_list(directory, settings):
    records, csv_path = generate_list_csv(
        directory,
        directory / "list_mm4.csv",
        settings=settings,
    )
    image_list, image_count = write_image_list(
        directory,
        _first_image_paths_from_records(records),
    )
    return {
        "records": records,
        "csv": csv_path,
        "directory": directory,
        "image_list": image_list,
        "image_list_count": image_count,
    }


def generate_mark_and_grain_lists(project_directory, settings=None):
    settings = settings or PycroTsvSettings()
    project_directory = Path(project_directory).expanduser()
    mark_directory = project_directory / "mark"
    grains_directory = project_directory / "grains"
    mineral_check_directory = _find_optional_subdirectory(
        project_directory, MINERAL_CHECK_DIRECTORY_NAME
    )
    missing = [
        str(path)
        for path in (mark_directory, grains_directory)
        if not path.exists() or not path.is_dir()
    ]
    if missing:
        raise FileNotFoundError(
            "Expected sub-directories named 'mark' and 'grains'. Missing: "
            + ", ".join(missing)
        )

    result = {
        "mark": _generate_project_subdirectory_list(mark_directory, settings),
        "grains": _generate_project_subdirectory_list(grains_directory, settings),
        "mineral_check": {
            "records": [],
            "csv": None,
            "directory": mineral_check_directory,
            "image_list": None,
            "image_list_count": 0,
        },
    }
    if mineral_check_directory.exists() and mineral_check_directory.is_dir():
        result["mineral_check"] = _generate_project_subdirectory_list(
            mineral_check_directory, settings
        )
    return result


def select_directory(initial_dir=DEFAULT_PROJECT_DIRECTORY):
    import tkinter as tk
    import tkinter.filedialog

    root = tk.Tk()
    root.withdraw()
    try:
        return tkinter.filedialog.askdirectory(
            initialdir=initial_dir,
            title="Select folder containing Pycro JPG crops",
        )
    finally:
        root.destroy()


def _build_arg_parser():
    parser = argparse.ArgumentParser(
        description="Create MM4 list_mm4.csv from Pycro crop image names without grouping."
    )
    parser.add_argument(
        "directory",
        nargs="?",
        help="Folder containing Pycro JPG crops. Opens a folder picker when omitted.",
    )
    parser.add_argument("--output", help="Output CSV path. Defaults to list_mm4.csv")
    parser.add_argument(
        "--um-per-pixel",
        type=float,
        default=UM_PER_PIXEL,
        help=f"Microns per pixel. Defaults to {UM_PER_PIXEL}.",
    )
    return parser


def main(argv=None):
    args = _build_arg_parser().parse_args(argv)
    directory = args.directory or select_directory()
    if not directory:
        raise SystemExit("No directory selected.")

    settings = PycroTsvSettings(um_per_pixel=args.um_per_pixel)
    records, output_path = generate_list_csv(
        directory,
        args.output,
        settings=settings,
    )
    print(f"Wrote {len(records)} positions to: {output_path}")


if __name__ == "__main__":
    main()
