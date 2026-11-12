from __future__ import annotations

import math
from dataclasses import dataclass
from pathlib import Path
from tkinter import (
    BOTH,
    END,
    LEFT,
    MULTIPLE,
    RIGHT,
    Y,
    Button,
    Frame,
    Label,
    Listbox,
    Scrollbar,
    Tk,
    filedialog,
)
from tkinter.simpledialog import askfloat
from typing import Dict, List, Optional, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Border, Font, PatternFill, Side


# ===========================
# SETTINGS
# ===========================
DEFAULT_TOL_UM = 20.0
REQUIRE_DIFFERENT_FIELD = True

BOLD_FONT = Font(bold=True)
HEADER_FILL = PatternFill("solid", fgColor="D9E1F2")

# Group colors cycle when there are more groups than available colors.
# Repeated colors do not cause ambiguity because each group also has:
# 1. Consecutive rows
# 2. A Duplicate Group number
# 3. A thick outline border
GROUP_COLORS = [
    "FFF59D",  # yellow
    "B3E5FC",  # light blue
    "C8E6C9",  # light green
    "FFCCBC",  # light orange
    "E1BEE7",  # light purple
    "F8BBD0",  # light pink
    "D7CCC8",  # light brown/grey
    "DCEDC8",  # pale green
    "BBDEFB",  # pale blue
    "FFE0B2",  # pale orange
]

GROUP_BORDER_SIDE = Side(style="thick", color="000000")
DUPLICATE_GROUP_HEADER = "Duplicate Group"


# ===========================
# UI helpers
# ===========================
def pick_file() -> Path:
    """Ask the user to select the classified Excel workbook."""
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    selected_path = filedialog.askopenfilename(
        title="Select the 'Classified xxxxxx' Excel file",
        filetypes=[("Excel files", "*.xlsx")],
    )

    root.destroy()

    if not selected_path:
        raise SystemExit("No file selected.")

    return Path(selected_path)


def select_sheets_gui(sheetnames: List[str]) -> List[str]:
    """Allow the user to select one or more workbook sheets."""
    root = Tk()
    root.title("Select sheets to process (duplicates only)")
    root.attributes("-topmost", True)

    Label(
        root,
        text="Select sheets to process.\nUse Ctrl-click or Shift-click to select multiple sheets.",
    ).pack(padx=10, pady=8)

    frame = Frame(root)
    frame.pack(padx=10, pady=5, fill=BOTH, expand=True)

    scrollbar = Scrollbar(frame)
    scrollbar.pack(side=RIGHT, fill=Y)

    sheet_listbox = Listbox(
        frame,
        selectmode=MULTIPLE,
        yscrollcommand=scrollbar.set,
        width=60,
        height=18,
    )

    for sheet_name in sheetnames:
        sheet_listbox.insert(END, sheet_name)

    sheet_listbox.pack(side=LEFT, fill=BOTH, expand=True)
    scrollbar.config(command=sheet_listbox.yview)

    selected_sheets: List[str] = []

    def select_all() -> None:
        sheet_listbox.select_set(0, END)

    def confirm_selection() -> None:
        for selected_index in sheet_listbox.curselection():
            selected_sheets.append(sheet_listbox.get(selected_index))
        root.destroy()

    button_frame = Frame(root)
    button_frame.pack(pady=10)

    Button(
        button_frame,
        text="Select All",
        command=select_all,
        width=12,
    ).pack(side=LEFT, padx=6)

    Button(
        button_frame,
        text="OK",
        command=confirm_selection,
        width=12,
    ).pack(side=LEFT, padx=6)

    root.mainloop()

    return selected_sheets


def ask_for_tolerance() -> float:
    """
    Ask for the maximum XY distance between potential duplicates.

    The tolerance is measured in micrometres and is based on the exported
    Stage X and Stage Y coordinates. It is not a particle-size threshold.
    """
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    tolerance = askfloat(
        "Duplicate tolerance (µm)",
        "Enter the maximum stage-coordinate distance for two features\n"
        "to be considered possible duplicates.\n\n"
        "This is an XY distance in µm, not a particle-size value.\n\n"
        "Recommended starting value: 20 µm\n"
        "Increase it if genuine duplicates are missed.\n"
        "Reduce it if unrelated features are grouped.",
        initialvalue=DEFAULT_TOL_UM,
        minvalue=0.5,
        maxvalue=500.0,
        parent=root,
    )

    root.destroy()

    if tolerance is None:
        raise SystemExit("Tolerance selection cancelled.")

    return tolerance


# ===========================
# Core helpers
# ===========================
def norm_header(value: str) -> str:
    """Normalize an Excel header for flexible column matching."""
    return (
        (value or "")
        .strip()
        .lower()
        .replace("\n", " ")
        .replace("µ", "u")
    )


def find_col_idx(
    headers: List[str],
    candidates: List[str],
) -> Optional[int]:
    """Find a one-based column index using partial header matches."""
    normalized_headers = [norm_header(header) for header in headers]

    for candidate in candidates:
        normalized_candidate = candidate.lower()

        for column_index, header in enumerate(
            normalized_headers,
            start=1,
        ):
            if normalized_candidate in header:
                return column_index

    return None


def is_um_header(header: str) -> bool:
    """Return True when a coordinate header indicates micrometres."""
    return "um" in norm_header(header)


def safe_float(value) -> Optional[float]:
    """Convert an Excel value to float without stopping on invalid data."""
    if value is None:
        return None

    if isinstance(value, str):
        value = value.strip().replace(",", ".")

    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def safe_str(value) -> str:
    """Convert an Excel value to a stripped string."""
    if value is None:
        return ""

    return str(value).strip()


@dataclass
class Pt:
    """A feature coordinate read from the source worksheet."""

    excel_row: int
    feature_id: str
    field_id: str
    x_um: float
    y_um: float


class UnionFind:
    """Combine features that belong to the same duplicate group."""

    def __init__(self, number_of_points: int):
        self.parent = list(range(number_of_points))
        self.rank = [0] * number_of_points

    def find(self, point_index: int) -> int:
        while self.parent[point_index] != point_index:
            self.parent[point_index] = self.parent[
                self.parent[point_index]
            ]
            point_index = self.parent[point_index]

        return point_index

    def union(self, first_index: int, second_index: int) -> None:
        first_root = self.find(first_index)
        second_root = self.find(second_index)

        if first_root == second_root:
            return

        if self.rank[first_root] < self.rank[second_root]:
            self.parent[first_root] = second_root
        elif self.rank[first_root] > self.rank[second_root]:
            self.parent[second_root] = first_root
        else:
            self.parent[second_root] = first_root
            self.rank[first_root] += 1


def grid_key(
    x_um: float,
    y_um: float,
    cell_size_um: float,
) -> Tuple[int, int]:
    """Return the spatial-search grid cell containing a coordinate."""
    return (
        int(math.floor(x_um / cell_size_um)),
        int(math.floor(y_um / cell_size_um)),
    )


def build_duplicate_groups(
    worksheet,
    tolerance_um: float,
) -> Tuple[Optional[str], Dict[int, List[int]], List[Pt]]:
    """
    Find duplicate feature groups in one worksheet.

    Returns:
        error_note:
            Description of a problem, or None.

        duplicate_groups:
            Mapping of group number to indexes in the points list.

        points:
            Valid coordinate records read from the worksheet.
    """
    maximum_row = worksheet.max_row
    maximum_column = worksheet.max_column

    if maximum_row < 2:
        return "Empty sheet", {}, []

    raw_headers = [
        worksheet.cell(row=1, column=column).value
        for column in range(1, maximum_column + 1)
    ]
    headers = [
        "" if header is None else str(header)
        for header in raw_headers
    ]

    stage_x_column = find_col_idx(headers, ["stage x"])
    stage_y_column = find_col_idx(headers, ["stage y"])
    feature_column = find_col_idx(
        headers,
        ["feature id", "feature"],
    )
    field_column = find_col_idx(headers, ["field"])

    if not all(
        [
            stage_x_column,
            stage_y_column,
            feature_column,
            field_column,
        ]
    ):
        return (
            "Missing one or more required columns "
            "(Stage X, Stage Y, Feature ID, Field)",
            {},
            [],
        )

    stage_x_header = headers[stage_x_column - 1]
    stage_y_header = headers[stage_y_column - 1]

    # If either coordinate header identifies µm, treat the coordinate values
    # as µm. Otherwise, assume the coordinates are in mm and convert to µm.
    if is_um_header(stage_x_header) or is_um_header(stage_y_header):
        coordinate_scale = 1.0
    else:
        coordinate_scale = 1000.0

    points: List[Pt] = []

    for excel_row in range(2, maximum_row + 1):
        stage_x = safe_float(
            worksheet.cell(
                row=excel_row,
                column=stage_x_column,
            ).value
        )
        stage_y = safe_float(
            worksheet.cell(
                row=excel_row,
                column=stage_y_column,
            ).value
        )

        # Rows without valid coordinates cannot be compared.
        if stage_x is None or stage_y is None:
            continue

        points.append(
            Pt(
                excel_row=excel_row,
                feature_id=safe_str(
                    worksheet.cell(
                        row=excel_row,
                        column=feature_column,
                    ).value
                ),
                field_id=safe_str(
                    worksheet.cell(
                        row=excel_row,
                        column=field_column,
                    ).value
                ),
                x_um=stage_x * coordinate_scale,
                y_um=stage_y * coordinate_scale,
            )
        )

    number_of_points = len(points)

    if number_of_points < 2:
        return "Not enough points", {}, points

    # Put points into spatial buckets. This avoids comparing every point
    # against every other point in large worksheets.
    buckets: Dict[Tuple[int, int], List[int]] = {}

    for point_index, point in enumerate(points):
        point_grid_key = grid_key(
            point.x_um,
            point.y_um,
            tolerance_um,
        )
        buckets.setdefault(point_grid_key, []).append(point_index)

    union_find = UnionFind(number_of_points)

    # Compare each point against points in its grid cell and the eight
    # neighbouring grid cells.
    for first_index, first_point in enumerate(points):
        grid_x, grid_y = grid_key(
            first_point.x_um,
            first_point.y_um,
            tolerance_um,
        )

        for x_offset in (-1, 0, 1):
            for y_offset in (-1, 0, 1):
                nearby_indexes = buckets.get(
                    (
                        grid_x + x_offset,
                        grid_y + y_offset,
                    ),
                    [],
                )

                for second_index in nearby_indexes:
                    # Avoid comparing a point with itself and avoid testing
                    # the same pair twice.
                    if second_index <= first_index:
                        continue

                    second_point = points[second_index]

                    if (
                        REQUIRE_DIFFERENT_FIELD
                        and first_point.field_id == second_point.field_id
                    ):
                        continue

                    distance_um = math.hypot(
                        first_point.x_um - second_point.x_um,
                        first_point.y_um - second_point.y_um,
                    )

                    if distance_um <= tolerance_um:
                        union_find.union(
                            first_index,
                            second_index,
                        )

    all_groups: Dict[int, List[int]] = {}

    for point_index in range(number_of_points):
        group_root = union_find.find(point_index)
        all_groups.setdefault(group_root, []).append(point_index)

    duplicate_point_groups = [
        group
        for group in all_groups.values()
        if len(group) > 1
    ]

    if not duplicate_point_groups:
        return None, {}, points

    duplicate_groups: Dict[int, List[int]] = {}

    for group_id, group in enumerate(
        duplicate_point_groups,
        start=1,
    ):
        duplicate_groups[group_id] = group

    return None, duplicate_groups, points


# ===========================
# Excel output helpers
# ===========================
def copy_row_values(
    source_worksheet,
    destination_worksheet,
    source_row: int,
    destination_row: int,
    maximum_column: int,
) -> None:
    """Copy the values from one worksheet row to another worksheet."""
    for column in range(1, maximum_column + 1):
        source_value = source_worksheet.cell(
            row=source_row,
            column=column,
        ).value

        destination_worksheet.cell(
            row=destination_row,
            column=column,
            value=source_value,
        )


def copy_headers(
    source_worksheet,
    destination_worksheet,
    maximum_column: int,
) -> int:
    """
    Copy the source headers and add the Duplicate Group column.

    Returns the last output column number.
    """
    for column in range(1, maximum_column + 1):
        destination_worksheet.cell(
            row=1,
            column=column,
            value=source_worksheet.cell(
                row=1,
                column=column,
            ).value,
        )

    duplicate_group_column = maximum_column + 1

    group_header_cell = destination_worksheet.cell(
        row=1,
        column=duplicate_group_column,
        value=DUPLICATE_GROUP_HEADER,
    )
    group_header_cell.fill = HEADER_FILL
    group_header_cell.font = BOLD_FONT

    return duplicate_group_column


def apply_group_fill(
    worksheet,
    first_row: int,
    last_row: int,
    last_column: int,
    color: str,
) -> None:
    """Apply one background color to all cells in a duplicate group."""
    group_fill = PatternFill(
        fill_type="solid",
        fgColor=color,
    )

    for row in range(first_row, last_row + 1):
        for column in range(1, last_column + 1):
            worksheet.cell(
                row=row,
                column=column,
            ).fill = group_fill


def outline_group(
    worksheet,
    first_row: int,
    last_row: int,
    last_column: int,
) -> None:
    """
    Draw one thick outline around a complete duplicate group.

    The border encloses the group as one block rather than drawing a thick
    border around every individual row.
    """
    for row in range(first_row, last_row + 1):
        for column in range(1, last_column + 1):
            worksheet.cell(
                row=row,
                column=column,
            ).border = Border(
                left=(
                    GROUP_BORDER_SIDE
                    if column == 1
                    else None
                ),
                right=(
                    GROUP_BORDER_SIDE
                    if column == last_column
                    else None
                ),
                top=(
                    GROUP_BORDER_SIDE
                    if row == first_row
                    else None
                ),
                bottom=(
                    GROUP_BORDER_SIDE
                    if row == last_row
                    else None
                ),
            )


def write_duplicate_groups(
    source_worksheet,
    destination_worksheet,
    duplicate_groups: Dict[int, List[int]],
    points: List[Pt],
    maximum_source_column: int,
    duplicate_group_column: int,
) -> int:
    """
    Write duplicate groups into consecutive rows.

    Groups are ordered by their first appearance in the source worksheet.
    Rows within each group also retain their original source order.
    """
    destination_row = 2

    ordered_groups = sorted(
        duplicate_groups.items(),
        key=lambda group_item: min(
            points[point_index].excel_row
            for point_index in group_item[1]
        ),
    )

    for group_id, point_indexes in ordered_groups:
        group_first_row = destination_row

        ordered_point_indexes = sorted(
            point_indexes,
            key=lambda point_index: points[
                point_index
            ].excel_row,
        )

        for point_index in ordered_point_indexes:
            source_row = points[point_index].excel_row

            copy_row_values(
                source_worksheet=source_worksheet,
                destination_worksheet=destination_worksheet,
                source_row=source_row,
                destination_row=destination_row,
                maximum_column=maximum_source_column,
            )

            destination_worksheet.cell(
                row=destination_row,
                column=duplicate_group_column,
                value=group_id,
            )

            destination_row += 1

        group_last_row = destination_row - 1
        group_color = GROUP_COLORS[
            (group_id - 1) % len(GROUP_COLORS)
        ]

        apply_group_fill(
            worksheet=destination_worksheet,
            first_row=group_first_row,
            last_row=group_last_row,
            last_column=duplicate_group_column,
            color=group_color,
        )

        outline_group(
            worksheet=destination_worksheet,
            first_row=group_first_row,
            last_row=group_last_row,
            last_column=duplicate_group_column,
        )

    return destination_row - 2


def create_report_sheet(
    output_workbook: Workbook,
    report_rows: List[Dict[str, object]],
    tolerance_um: float,
) -> None:
    """Create the summary report at the beginning of the output workbook."""
    report_worksheet = output_workbook.create_sheet(
        "Duplicate_Report",
        0,
    )

    report_headers = [
        "Sheet",
        "PointsChecked",
        "DuplicateRows",
        "DuplicateGroups",
        "Note",
    ]

    for column, header in enumerate(report_headers, start=1):
        header_cell = report_worksheet.cell(
            row=1,
            column=column,
            value=header,
        )
        header_cell.fill = HEADER_FILL
        header_cell.font = BOLD_FONT

    report_row_number = 2

    for report_data in report_rows:
        report_worksheet.cell(
            row=report_row_number,
            column=1,
            value=report_data["Sheet"],
        )
        report_worksheet.cell(
            row=report_row_number,
            column=2,
            value=report_data["PointsChecked"],
        )
        report_worksheet.cell(
            row=report_row_number,
            column=3,
            value=report_data["DuplicateRows"],
        )
        report_worksheet.cell(
            row=report_row_number,
            column=4,
            value=report_data["DuplicateGroups"],
        )
        report_worksheet.cell(
            row=report_row_number,
            column=5,
            value=report_data["Note"],
        )

        report_row_number += 1

    tolerance_label_row = report_row_number + 1

    report_worksheet.cell(
        row=tolerance_label_row,
        column=1,
        value="Tolerance (µm)",
    )
    report_worksheet.cell(
        row=tolerance_label_row,
        column=2,
        value=tolerance_um,
    )

    report_worksheet.cell(
        row=tolerance_label_row + 1,
        column=1,
        value="Require different field?",
    )
    report_worksheet.cell(
        row=tolerance_label_row + 1,
        column=2,
        value=str(REQUIRE_DIFFERENT_FIELD),
    )


# ===========================
# Main
# ===========================
def main() -> None:
    input_path = pick_file()
    print(f"Selected: {input_path}", flush=True)

    print("Loading workbook...", flush=True)
    input_workbook = load_workbook(input_path)

    print(
        f"Loaded. Sheets: {len(input_workbook.sheetnames)}",
        flush=True,
    )

    selected_sheets = select_sheets_gui(
        input_workbook.sheetnames
    )

    if not selected_sheets:
        raise SystemExit("No sheets selected.")

    print(
        f"Selected sheets ({len(selected_sheets)}): "
        f"{selected_sheets}",
        flush=True,
    )

    tolerance_um = ask_for_tolerance()

    print(
        f"Tolerance = {tolerance_um} µm | "
        f"Require different field = {REQUIRE_DIFFERENT_FIELD}",
        flush=True,
    )

    output_workbook = Workbook()
    output_workbook.remove(output_workbook.active)

    report_rows: List[Dict[str, object]] = []

    for sheet_name in selected_sheets:
        source_worksheet = input_workbook[sheet_name]

        print(
            f"\nProcessing: {sheet_name}",
            flush=True,
        )

        (
            note,
            duplicate_groups,
            points,
        ) = build_duplicate_groups(
            worksheet=source_worksheet,
            tolerance_um=tolerance_um,
        )

        maximum_source_column = source_worksheet.max_column

        if note and note != "Not enough points":
            print(f"  NOTE: {note}", flush=True)

        destination_worksheet = output_workbook.create_sheet(
            sheet_name[:31]
        )

        duplicate_group_column = copy_headers(
            source_worksheet=source_worksheet,
            destination_worksheet=destination_worksheet,
            maximum_column=maximum_source_column,
        )

        if not duplicate_groups:
            report_rows.append(
                {
                    "Sheet": sheet_name,
                    "PointsChecked": len(points),
                    "DuplicateRows": 0,
                    "DuplicateGroups": 0,
                    "Note": note or "",
                }
            )
            continue

        duplicate_row_count = write_duplicate_groups(
            source_worksheet=source_worksheet,
            destination_worksheet=destination_worksheet,
            duplicate_groups=duplicate_groups,
            points=points,
            maximum_source_column=maximum_source_column,
            duplicate_group_column=duplicate_group_column,
        )

        duplicate_group_count = len(duplicate_groups)

        report_rows.append(
            {
                "Sheet": sheet_name,
                "PointsChecked": len(points),
                "DuplicateRows": duplicate_row_count,
                "DuplicateGroups": duplicate_group_count,
                "Note": note or "",
            }
        )

        print(
            f"  Duplicate groups: {duplicate_group_count} | "
            f"Rows kept: {duplicate_row_count}",
            flush=True,
        )

    create_report_sheet(
        output_workbook=output_workbook,
        report_rows=report_rows,
        tolerance_um=tolerance_um,
    )

    output_path = input_path.with_name(
        input_path.stem + "_DUPS_ONLY.xlsx"
    )

    print(f"\nSaving: {output_path}", flush=True)
    output_workbook.save(output_path)

    print("Done.", flush=True)


if __name__ == "__main__":
    main()