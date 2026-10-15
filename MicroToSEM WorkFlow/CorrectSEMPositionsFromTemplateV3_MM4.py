import argparse
import csv
import math
import re
import sys
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path


SEM_UNIT_FACTORS = {
    "mm": 1000.0,
    "um": 1.0,
}


def excel_atan2(x_num, y_num):
    """Excel ATAN2(x_num, y_num) matches Python atan2(y_num, x_num)."""
    return math.atan2(y_num, x_num)


def parse_point(text):
    parts = re.split(r"[\s,;]+", text.strip().strip("()[]"))
    parts = [part for part in parts if part]
    if len(parts) != 2:
        raise ValueError("enter two numbers as X,Y")
    return float(parts[0]), float(parts[1])


def prompt_point(label, units):
    while True:
        raw = input(f"{label} ({units}) as X,Y: ").strip()
        try:
            return parse_point(raw)
        except ValueError as exc:
            print(f"Could not read that point: {exc}")


def prompt_sem_units():
    while True:
        raw = input("SEM position units [mm]: ").strip().lower() or "mm"
        if raw in SEM_UNIT_FACTORS:
            return raw
        print("Please enter 'mm' or 'um'.")


def prompt_decimals():
    raw = input("Decimal places for corrected SEM output [2]: ").strip()
    if not raw:
        return 2
    try:
        decimals = int(raw)
    except ValueError:
        print("Using 2 decimal places.")
        return 2
    if decimals < 0:
        print("Using 2 decimal places.")
        return 2
    return decimals


def choose_input_file():
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError:
        tk = None
        filedialog = None

    if tk is not None:
        root = tk.Tk()
        root.withdraw()
        path = filedialog.askopenfilename(
            title="Select list.csv",
            filetypes=[
                ("CSV or TSV files", "*.csv *.tsv"),
                ("All files", "*.*"),
            ],
        )
        root.destroy()
        if path:
            return Path(path)

    while True:
        raw = input("Path to list.csv: ").strip().strip('"')
        if raw:
            return Path(raw)


def sniff_csv_dialect(path):
    with path.open("r", encoding="utf-8-sig", newline="") as file:
        sample = file.read(4096)

    try:
        return csv.Sniffer().sniff(sample, delimiters=",\t;")
    except csv.Error:
        first_line = sample.splitlines()[0] if sample else ""
        dialect = csv.excel()
        if "\t" in first_line:
            dialect.delimiter = "\t"
        elif ";" in first_line:
            dialect.delimiter = ";"
        else:
            dialect.delimiter = ","
        return dialect


def normalize_column_name(name):
    return re.sub(r"[^a-z0-9]+", "", name.lower())


def find_column(fieldnames, candidates):
    normalized = {normalize_column_name(name): name for name in fieldnames}
    for candidate in candidates:
        match = normalized.get(normalize_column_name(candidate))
        if match is not None:
            return match
    raise ValueError(
        "Could not find one of these columns: " + ", ".join(candidates)
    )


def read_position_rows(path):
    dialect = sniff_csv_dialect(path)
    with path.open("r", encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file, dialect=dialect)
        if reader.fieldnames is None:
            raise ValueError("The input file does not have a header row.")
        fieldnames = list(reader.fieldnames)
        x_col = find_column(fieldnames, ["X", "Average_X", "Microscope_X"])
        y_col = find_column(fieldnames, ["Y", "Average_Y", "Microscope_Y"])
        rows = list(reader)

    if not rows:
        raise ValueError("The input file does not contain any positions.")
    return fieldnames, rows, x_col, y_col


def as_float(value, row_number, column_name):
    text = str(value).strip()
    if not text:
        raise ValueError(f"row {row_number}: blank value in {column_name}")
    try:
        return float(text)
    except ValueError as exc:
        raise ValueError(
            f"row {row_number}: could not parse {column_name}={text!r}"
        ) from exc


def validate_fiducials(micro_a, micro_b, sem_a_internal, sem_b_internal):
    micro_ba = (micro_b[0] - micro_a[0], micro_b[1] - micro_a[1])
    sem_ba = (
        sem_b_internal[0] - sem_a_internal[0],
        sem_b_internal[1] - sem_a_internal[1],
    )
    micro_len = math.hypot(micro_ba[0], micro_ba[1])
    sem_len = math.hypot(sem_ba[0], sem_ba[1])
    if micro_len == 0:
        raise ValueError("Microscope A and B positions are identical.")
    if sem_len == 0:
        raise ValueError("SEM A and B positions are identical.")


def make_transform(micro_a, micro_b, sem_a, sem_b, sem_unit_factor):
    sem_a_internal = (sem_a[0] * sem_unit_factor, sem_a[1] * sem_unit_factor)
    sem_b_internal = (sem_b[0] * sem_unit_factor, sem_b[1] * sem_unit_factor)
    validate_fiducials(micro_a, micro_b, sem_a_internal, sem_b_internal)

    micro_ba = (micro_b[0] - micro_a[0], micro_b[1] - micro_a[1])
    sem_ba = (
        sem_b_internal[0] - sem_a_internal[0],
        sem_b_internal[1] - sem_a_internal[1],
    )

    micro_ba_length = math.hypot(micro_ba[0], micro_ba[1])
    sem_ba_length = math.hypot(sem_ba[0], sem_ba[1])
    zoom_ratio = micro_ba_length / sem_ba_length
    micro_ba_angle = math.degrees(excel_atan2(micro_ba[0], micro_ba[1]))
    sem_ba_angle = math.degrees(excel_atan2(sem_ba[0], sem_ba[1]))

    def transform(point):
        point_to_b = (micro_b[0] - point[0], micro_b[1] - point[1])
        point_to_b_length = math.hypot(point_to_b[0], point_to_b[1])

        if point_to_b_length == 0:
            return sem_b

        point_to_b_angle = math.degrees(
            excel_atan2(point_to_b[0], point_to_b[1])
        )
        angle_from_ba = point_to_b_angle - micro_ba_angle
        # MM4: Template V3 R7 = $T$5+Q7 (rotation, not reflection).
        sem_angle = 360.0 + sem_ba_angle + angle_from_ba
        sem_length = point_to_b_length / zoom_ratio
        sem_x_internal = sem_b_internal[0] - sem_length * math.cos(
            math.radians(sem_angle)
        )
        sem_y_internal = sem_b_internal[1] - sem_length * math.sin(
            math.radians(sem_angle)
        )

        # Do not pin individual axes to SEM B: a rotated point can share one
        # microscope coordinate with B while differing on both SEM axes.
        # The exact point B is already handled above.

        return sem_x_internal / sem_unit_factor, sem_y_internal / sem_unit_factor

    transform.zoom_ratio = zoom_ratio
    transform.micro_ba_angle = micro_ba_angle
    transform.sem_ba_angle = sem_ba_angle
    return transform


def fixed_decimal(value, decimals):
    quantum = Decimal("1").scaleb(-decimals)
    rounded = Decimal(str(value)).quantize(quantum, rounding=ROUND_HALF_UP)
    if rounded == 0:
        rounded = abs(rounded)
    return f"{rounded:.{decimals}f}"


def default_output_path(input_path):
    return input_path.with_name(f"{input_path.stem}_sem_corrected_mm4.csv")


def write_corrected_csv(
    input_path,
    output_path,
    micro_a,
    micro_b,
    sem_a,
    sem_b,
    sem_units,
    decimals,
):
    fieldnames, rows, x_col, y_col = read_position_rows(input_path)
    transform = make_transform(
        micro_a,
        micro_b,
        sem_a,
        sem_b,
        SEM_UNIT_FACTORS[sem_units],
    )

    passthrough_fields = [
        name
        for name in fieldnames
        if name not in {x_col, y_col, "X", "Y", "Microscope_X", "Microscope_Y"}
    ]
    output_fields = ["X", "Y", "Microscope_X", "Microscope_Y"] + passthrough_fields
    output_rows = []

    for index, row in enumerate(rows, start=2):
        micro_x = as_float(row.get(x_col, ""), index, x_col)
        micro_y = as_float(row.get(y_col, ""), index, y_col)
        sem_x, sem_y = transform((micro_x, micro_y))

        output_row = {
            "X": fixed_decimal(sem_x, decimals),
            "Y": fixed_decimal(sem_y, decimals),
            "Microscope_X": row.get(x_col, ""),
            "Microscope_Y": row.get(y_col, ""),
        }
        for field in passthrough_fields:
            output_row[field] = row.get(field, "")
        output_rows.append(output_row)

    if output_path.resolve() == input_path.resolve():
        raise ValueError("Refusing to overwrite the input file.")

    with output_path.open("w", encoding="utf-8", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=output_fields, lineterminator="\n")
        writer.writeheader()
        writer.writerows(output_rows)

    return len(output_rows), transform


def build_arg_parser():
    parser = argparse.ArgumentParser(
        description=(
            "Convert microscope X/Y positions into SEM-corrected positions "
            "using the MM4 rotation correction (Template V3 R7 = $T$5+Q7)."
        )
    )
    parser.add_argument("input", nargs="?", help="Path to list.csv or list.tsv")
    parser.add_argument("--output", help="Output CSV path")
    parser.add_argument("--micro-a", help="Microscope A position as X,Y")
    parser.add_argument("--micro-b", help="Microscope B position as X,Y")
    parser.add_argument("--sem-a", help="SEM A position as X,Y")
    parser.add_argument("--sem-b", help="SEM B position as X,Y")
    parser.add_argument(
        "--sem-units",
        choices=sorted(SEM_UNIT_FACTORS),
        help="Units used for the entered SEM A/B positions and output CSV",
    )
    parser.add_argument(
        "--decimals",
        type=int,
        default=None,
        help="Decimal places for corrected SEM X/Y values",
    )
    return parser


def get_point_from_args_or_prompt(arg_value, label, units):
    if arg_value:
        return parse_point(arg_value)
    return prompt_point(label, units)


def main(argv=None):
    args = build_arg_parser().parse_args(argv)

    input_path = Path(args.input) if args.input else choose_input_file()
    input_path = input_path.expanduser()
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    sem_units = args.sem_units or prompt_sem_units()
    decimals = args.decimals if args.decimals is not None else prompt_decimals()
    if decimals < 0:
        raise ValueError("--decimals must be 0 or greater")

    micro_a = get_point_from_args_or_prompt(
        args.micro_a, "Microscope A", "same units as input X/Y"
    )
    micro_b = get_point_from_args_or_prompt(
        args.micro_b, "Microscope B", "same units as input X/Y"
    )
    sem_a = get_point_from_args_or_prompt(args.sem_a, "SEM A", sem_units)
    sem_b = get_point_from_args_or_prompt(args.sem_b, "SEM B", sem_units)

    output_path = Path(args.output) if args.output else default_output_path(input_path)
    count, transform = write_corrected_csv(
        input_path=input_path,
        output_path=output_path,
        micro_a=micro_a,
        micro_b=micro_b,
        sem_a=sem_a,
        sem_b=sem_b,
        sem_units=sem_units,
        decimals=decimals,
    )

    print(f"Wrote {count} corrected SEM positions to: {output_path}")
    print(f"Output X/Y units: {sem_units}")
    print(f"Template V3 magnitude ratio: {transform.zoom_ratio:.12g}")
    print(f"Microscope BA angle: {transform.micro_ba_angle:.6f} degrees")
    print(f"SEM BA angle: {transform.sem_ba_angle:.6f} degrees")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)
