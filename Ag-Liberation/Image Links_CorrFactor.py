from __future__ import annotations

import ctypes
import os
import re
from collections import defaultdict
from ctypes import wintypes
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import tkinter as tk
from tkinter import filedialog, messagebox

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

IMG_EXTS = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp"}
TARGET_IMAGE_HEADER = "Image"
TARGET_MICROX_HEADER = "Micro X"
LINK_HEADER = "Open Image"
NEW_HEADERS = ["Area Orig", "Corr %", "New Brea", "New Len", "Asso"]
MIN_COL_WIDTH = 8
MAX_COL_WIDTH = 28


class LiveLogger:
    def __init__(self, root: tk.Tk):
        self.win = tk.Toplevel(root)
        self.win.title("Image Links + CorrFactor - Live Progress")
        self.win.geometry("980x520")

        self.text = tk.Text(self.win, wrap="word")
        self.text.pack(fill="both", expand=True)

        self.write("=== Live Progress ===")

    def write(self, msg: str):
        self.text.insert("end", msg + "\n")
        self.text.see("end")
        self.win.update_idletasks()
        self.win.update()


def canonical_key(s: str) -> str:
    s = str(s or "").strip().lower()
    s = s.replace("&", "_").replace("-", "_")
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"[^a-z0-9_]", "_", s)
    s = re.sub(r"_+", "_", s)
    return s.strip("_")


def canonical_variants(s: str) -> List[str]:
    base = canonical_key(s)

    if not base:
        return []

    return list(
        {
            base,
            base.replace("_", ""),
            base.replace("_and_", "_"),
            base.replace("and", ""),
        }
    )


def normalize_text(value) -> str:
    return re.sub(
        r"\s+",
        "",
        str(value or "").strip().lower(),
    )


def normalize_micro_x(value) -> str:
    """
    Normalize numeric and text Micro X values without losing precision.

    Examples:
        16921.52       -> "16921.52"
        "16921.520"    -> "16921.52"
        "16921.521"    -> "16921.521"
    """
    text = str(value or "").strip()

    if not text:
        return ""

    try:
        number = Decimal(text)
    except InvalidOperation:
        return normalize_text(text)

    if not number.is_finite():
        return normalize_text(text)

    return format(number.normalize(), "f")


def clean_key(value) -> str:
    s = str(value or "").strip().lower()
    s = re.sub(r"\.[a-z0-9]{2,5}$", "", s)
    s = s.replace("&", "and")
    s = re.sub(r"[^a-z0-9]+", "", s)
    return s


def choose_workbook(root: tk.Tk) -> Path:
    selected_path = filedialog.askopenfilename(
        parent=root,
        title="Select workbook",
        filetypes=[("Excel", "*.xlsx *.xlsm")],
    )

    if not selected_path:
        raise RuntimeError("No workbook selected")

    return Path(selected_path)


def resolve_folder_for_sheet(
    sheet_name: str,
    folder_map: Dict[str, Path],
) -> Optional[Path]:
    sheet_key = canonical_key(sheet_name)

    if sheet_key in folder_map:
        return folder_map[sheet_key]

    for variant in canonical_variants(sheet_name):
        if variant in folder_map:
            return folder_map[variant]

    for folder_key, folder_path in folder_map.items():
        if sheet_key and (
            sheet_key in folder_key
            or folder_key in sheet_key
        ):
            return folder_path

    return None


def build_category_folder_map(
    root_dir: Path,
) -> Dict[str, Path]:
    folder_map = {}

    for folder_path in root_dir.iterdir():
        cropped_folder = folder_path / "cropped"

        if (
            folder_path.is_dir()
            and cropped_folder.is_dir()
        ):
            folder_map[
                canonical_key(folder_path.name)
            ] = folder_path

            for variant in canonical_variants(
                folder_path.name
            ):
                folder_map.setdefault(
                    variant,
                    folder_path,
                )

    return folder_map


def build_cropped_index(cropped_dir: Path):
    """
    Index supported images using the complete Micro X token before
    the first underscore.

    Example:
        16921.520_27410.070_F13330.09.jpg

    Indexed Micro X:
        16921.520
    """
    exact = {}
    by_clean = defaultdict(list)
    by_x_exact = defaultdict(list)

    for image_path in cropped_dir.iterdir():
        if (
            image_path.is_file()
            and image_path.suffix.lower() in IMG_EXTS
        ):
            exact[
                normalize_text(image_path.stem)
            ] = image_path

            by_clean[
                clean_key(image_path.stem)
            ].append(image_path)

            micro_x_token = image_path.stem.split(
                "_",
                1,
            )[0].strip()

            if micro_x_token:
                normalized_micro_x = normalize_micro_x(
                    micro_x_token
                )

                by_x_exact[
                    normalized_micro_x
                ].append(image_path)

    return exact, by_clean, by_x_exact


def find_best_image(
    micro_x_value,
    exact,
    by_clean,
    by_x_exact,
    allow_shared_field_match: bool = False,
):
    """
    Find an image by matching the complete Excel Micro X value with
    the complete filename token before the first underscore.

    If exactly one image matches, that image is returned.

    If multiple images match, one deterministic image is returned only
    when the worksheet row belongs to a repeated nonblank Field subgroup
    for that same Micro X.
    """
    if micro_x_value is None:
        return None

    raw_micro_x = str(micro_x_value).strip()

    if not raw_micro_x:
        return None

    # Remove only a recognized image extension if one was included in
    # the Excel value. Do not treat a decimal part such as ".520" as
    # a file extension.
    suffix = Path(raw_micro_x).suffix.lower()

    if suffix in IMG_EXTS:
        raw_micro_x = raw_micro_x[:-len(suffix)]

    normalized_micro_x = normalize_micro_x(
        raw_micro_x
    )

    matching_images = by_x_exact.get(
        normalized_micro_x,
        [],
    )

    if len(matching_images) == 1:
        return matching_images[0]

    if (
        len(matching_images) > 1
        and allow_shared_field_match
    ):
        # Choose deterministically so repeated runs use the same file.
        return sorted(
            matching_images,
            key=lambda image_path: image_path.name.lower(),
        )[0]

    return None


def get_header_positions(
    ws,
) -> Dict[str, List[int]]:
    positions: Dict[str, List[int]] = defaultdict(list)

    for cell in ws[1]:
        if cell.value is None:
            continue

        header_name = str(cell.value).strip()
        positions[header_name].append(cell.column)

    return dict(positions)


def choose_pairs_gui(
    root: tk.Tk,
    pairs: List[Tuple[str, Path]],
    logger: LiveLogger,
) -> List[Tuple[str, Path]]:
    logger.write(
        f"Sheets with matched folders: {len(pairs)}"
    )

    if not pairs:
        return []

    window = tk.Toplevel(root)
    window.title("Choose categories to process")
    window.geometry("920x620")

    tk.Label(
        window,
        text="Select categories to process (checkboxes).",
        anchor="w",
    ).pack(
        fill="x",
        padx=10,
        pady=8,
    )

    outer_frame = tk.Frame(window)
    outer_frame.pack(
        fill="both",
        expand=True,
        padx=10,
        pady=10,
    )

    canvas = tk.Canvas(outer_frame)
    inner_frame = tk.Frame(canvas)

    scrollbar = tk.Scrollbar(
        outer_frame,
        orient="vertical",
        command=canvas.yview,
    )

    canvas.configure(
        yscrollcommand=scrollbar.set,
    )

    scrollbar.pack(
        side="right",
        fill="y",
    )

    canvas.pack(
        side="left",
        fill="both",
        expand=True,
    )

    canvas.create_window(
        (0, 0),
        window=inner_frame,
        anchor="nw",
    )

    inner_frame.bind(
        "<Configure>",
        lambda event: canvas.configure(
            scrollregion=canvas.bbox("all")
        ),
    )

    tk.Label(
        inner_frame,
        text="Use",
        font=("TkDefaultFont", 9, "bold"),
    ).grid(
        row=0,
        column=0,
        sticky="w",
        padx=(2, 10),
    )

    tk.Label(
        inner_frame,
        text="Sheet",
        font=("TkDefaultFont", 9, "bold"),
    ).grid(
        row=0,
        column=1,
        sticky="w",
        padx=(0, 10),
    )

    tk.Label(
        inner_frame,
        text="Folder",
        font=("TkDefaultFont", 9, "bold"),
    ).grid(
        row=0,
        column=2,
        sticky="w",
    )

    selection_variables = []

    for row_number, (
        sheet_name,
        folder_path,
    ) in enumerate(pairs, start=1):
        variable = tk.BooleanVar(
            master=window,
            value=True,
        )

        selection_variables.append(variable)

        tk.Checkbutton(
            inner_frame,
            variable=variable,
        ).grid(
            row=row_number,
            column=0,
            sticky="w",
        )

        tk.Label(
            inner_frame,
            text=sheet_name,
        ).grid(
            row=row_number,
            column=1,
            sticky="w",
            padx=(0, 10),
        )

        tk.Label(
            inner_frame,
            text=str(folder_path),
            justify="left",
            wraplength=520,
        ).grid(
            row=row_number,
            column=2,
            sticky="w",
        )

    result = {"pairs": pairs}

    def select_all():
        for variable in selection_variables:
            variable.set(True)

    def select_none():
        for variable in selection_variables:
            variable.set(False)

    def confirm_selection():
        result["pairs"] = [
            pair
            for pair, variable in zip(
                pairs,
                selection_variables,
            )
            if variable.get()
        ]

        window.destroy()

    def cancel_selection():
        result["pairs"] = []
        window.destroy()

    button_frame = tk.Frame(window)
    button_frame.pack(
        fill="x",
        padx=10,
        pady=(0, 10),
    )

    tk.Button(
        button_frame,
        text="Select All",
        command=select_all,
    ).pack(side="left")

    tk.Button(
        button_frame,
        text="Select None",
        command=select_none,
    ).pack(
        side="left",
        padx=5,
    )

    tk.Button(
        button_frame,
        text="OK",
        command=confirm_selection,
    ).pack(side="right")

    tk.Button(
        button_frame,
        text="Cancel",
        command=cancel_selection,
    ).pack(
        side="right",
        padx=5,
    )

    window.lift()
    window.attributes("-topmost", True)

    window.after(
        250,
        lambda: window.attributes(
            "-topmost",
            False,
        ),
    )

    window.transient(root)
    window.grab_set()
    window.wait_window()

    return result["pairs"]


def choose_visible_columns_listbox(
    root: tk.Tk,
    headers: List[str],
) -> Optional[set]:
    """
    Allow the user to select columns to show.

    Unselected columns will be hidden.
    All columns are selected by default.
    Cancel means show all.
    """
    if not headers:
        return None

    window = tk.Toplevel(root)
    window.title("Choose columns to SHOW")
    window.geometry("760x620")
    window.lift()
    window.attributes("-topmost", True)

    window.after(
        300,
        lambda: window.attributes(
            "-topmost",
            False,
        ),
    )

    tk.Label(
        window,
        text=(
            "Select columns to SHOW "
            "(others will be hidden).\n"
            "Default is ALL selected."
        ),
        anchor="w",
        justify="left",
    ).pack(
        fill="x",
        padx=10,
        pady=(10, 5),
    )

    frame = tk.Frame(window)
    frame.pack(
        fill="both",
        expand=True,
        padx=10,
        pady=10,
    )

    scrollbar = tk.Scrollbar(frame)
    scrollbar.pack(
        side="right",
        fill="y",
    )

    listbox = tk.Listbox(
        frame,
        selectmode="extended",
        yscrollcommand=scrollbar.set,
    )

    for header in headers:
        listbox.insert("end", header)

    listbox.selection_set(0, "end")

    listbox.pack(
        side="left",
        fill="both",
        expand=True,
    )

    scrollbar.config(
        command=listbox.yview,
    )

    result = {"value": None}

    def select_all():
        listbox.selection_set(0, "end")

    def select_none():
        listbox.selection_clear(0, "end")

    def confirm_selection():
        selected_indices = listbox.curselection()

        result["value"] = {
            headers[index]
            for index in selected_indices
        }

        window.destroy()

    def cancel_selection():
        result["value"] = None
        window.destroy()

    button_frame = tk.Frame(window)
    button_frame.pack(
        fill="x",
        padx=10,
        pady=(0, 10),
    )

    tk.Button(
        button_frame,
        text="Select All",
        command=select_all,
    ).pack(side="left")

    tk.Button(
        button_frame,
        text="Select None",
        command=select_none,
    ).pack(
        side="left",
        padx=5,
    )

    tk.Button(
        button_frame,
        text="OK",
        command=confirm_selection,
    ).pack(side="right")

    tk.Button(
        button_frame,
        text="Cancel (show all)",
        command=cancel_selection,
    ).pack(
        side="right",
        padx=5,
    )

    window.transient(root)
    window.grab_set()
    window.wait_window()

    return result["value"]


def shrink_columns_keep_visible(ws):
    for column_index in range(
        1,
        ws.max_column + 1,
    ):
        column_letter = get_column_letter(
            column_index
        )

        dimension = ws.column_dimensions[
            column_letter
        ]

        if bool(
            getattr(dimension, "hidden", False)
        ):
            continue

        header = ws.cell(
            row=1,
            column=column_index,
        ).value

        header_text = (
            str(header).strip()
            if header is not None
            else ""
        )

        if header_text == TARGET_IMAGE_HEADER:
            continue

        ws.cell(
            row=1,
            column=column_index,
        ).alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )

        first_word = (
            re.split(r"\s+", header_text)[0]
            if header_text
            else ""
        )

        dimension.width = max(
            MIN_COL_WIDTH,
            min(
                MAX_COL_WIDTH,
                len(first_word) + 1,
            ),
        )


def freeze_top_header_row(ws):
    # Freeze only the first row, with no frozen columns.
    ws.freeze_panes = "A2"


def ensure_new_columns_and_formulas(
    ws,
    area_orig_col: int,
    corr_col: int,
    new_brea_col: int,
    new_len_col: int,
    asso_col: int,
    breadth_col: int,
    length_col: int,
):
    for row_number in range(
        2,
        ws.max_row + 1,
    ):
        # Corr % remains user editable.
        corr_reference = (
            f"{get_column_letter(corr_col)}"
            f"{row_number}"
        )

        area_reference = (
            f"{get_column_letter(area_orig_col)}"
            f"{row_number}"
        )

        breadth_reference = (
            f"{get_column_letter(breadth_col)}"
            f"{row_number}"
        )

        length_reference = (
            f"{get_column_letter(length_col)}"
            f"{row_number}"
        )

        ws.cell(
            row=row_number,
            column=new_brea_col,
        ).value = (
            f'=IF(OR({corr_reference}="",'
            f'{corr_reference}=0),'
            f'{breadth_reference},'
            f'SQRT({area_reference}*'
            f'{corr_reference}/100))'
        )

        ws.cell(
            row=row_number,
            column=new_len_col,
        ).value = (
            f'=IF(OR({corr_reference}="",'
            f'{corr_reference}=0),'
            f'{length_reference},'
            f'SQRT({area_reference}*'
            f'{corr_reference}/100))'
        )

        # Asso remains blank and editable.


def align_new_columns_center(
    ws,
    new_columns: List[int],
):
    for column_index in new_columns:
        for row_number in range(
            1,
            ws.max_row + 1,
        ):
            ws.cell(
                row=row_number,
                column=column_index,
            ).alignment = Alignment(
                horizontal="center",
                vertical="center",
            )


def process_sheet(
    ws,
    cropped_dir: Path,
    link_base_dir: Path,
):
    headers = get_header_positions(ws)

    if (
        TARGET_IMAGE_HEADER not in headers
        or TARGET_MICROX_HEADER not in headers
    ):
        return (
            0,
            0,
            "Missing Image/Micro X headers",
        )

    # Preserve hidden-like states by original column index.
    # This supports duplicate headers and zero-width hidden columns.
    hidden_indices = set()
    hidden_widths = {}

    for column_index in range(
        1,
        ws.max_column + 1,
    ):
        column_letter = get_column_letter(
            column_index
        )

        dimension = ws.column_dimensions[
            column_letter
        ]

        width = dimension.width

        is_hidden_like = (
            bool(dimension.hidden)
            or (
                width is not None
                and float(width) == 0.0
            )
        )

        if is_hidden_like:
            hidden_indices.add(column_index)
            hidden_widths[column_index] = width

    image_column = headers[
        TARGET_IMAGE_HEADER
    ][0]

    inserted_count = 0

    if LINK_HEADER in headers:
        link_column = headers[LINK_HEADER][0]
    else:
        insert_at = image_column + 1
        inserted_count = 1 + len(NEW_HEADERS)

        ws.insert_cols(
            insert_at,
            amount=inserted_count,
        )

        ws.cell(
            row=1,
            column=insert_at,
        ).value = LINK_HEADER

        for offset, header_name in enumerate(
            NEW_HEADERS,
            start=1,
        ):
            ws.cell(
                row=1,
                column=insert_at + offset,
            ).value = header_name

        link_column = insert_at

    # Refresh every column position after possible insertion.
    headers = get_header_positions(ws)

    # Micro X moves when the new columns are inserted.
    micro_x_column = headers[
        TARGET_MICROX_HEADER
    ][0]

    area_columns = headers.get("Area", [])

    if len(area_columns) < 2:
        return (
            0,
            0,
            "Could not find second Area column",
        )

    area_second_column = area_columns[1]

    breadth_column = headers.get(
        "Breadth (um)",
        [None],
    )[0]

    length_column = headers.get(
        "Length",
        [None],
    )[0]

    if (
        breadth_column is None
        or length_column is None
    ):
        return (
            0,
            0,
            "Missing Breadth (um) or Length column",
        )

    area_orig_column = headers["Area Orig"][0]
    corr_column = headers["Corr %"][0]
    new_brea_column = headers["New Brea"][0]
    new_len_column = headers["New Len"][0]
    asso_column = headers["Asso"][0]

    field_column = headers.get(
        "Field",
        [None],
    )[0]

    for header_name in [
        LINK_HEADER,
    ] + NEW_HEADERS:
        column_number = headers[
            header_name
        ][0]

        ws.cell(
            row=1,
            column=column_number,
        ).font = Font(bold=True)

    (
        exact_index,
        clean_index,
        micro_x_index,
    ) = build_cropped_index(cropped_dir)

    linked = 0
    missing = 0

    # Count each nonblank Field subgroup separately within each Micro X.
    #
    # Example:
    #   Micro X 16921.52, Fields [7, 7, 8]
    #
    # Field 7 count = 2, so those two rows may share an image.
    # Field 8 count = 1, so that row must remain Missing.
    micro_x_field_counts = defaultdict(
        lambda: defaultdict(int)
    )

    if field_column is not None:
        for row_number in range(
            2,
            ws.max_row + 1,
        ):
            micro_x_value = ws.cell(
                row=row_number,
                column=micro_x_column,
            ).value

            field_value = ws.cell(
                row=row_number,
                column=field_column,
            ).value

            micro_x_key = normalize_micro_x(
                micro_x_value
            )

            field_key = (
                normalize_text(field_value)
                if field_value is not None
                else ""
            )

            if micro_x_key and field_key:
                micro_x_field_counts[
                    micro_x_key
                ][field_key] += 1

    for row_number in range(
        2,
        ws.max_row + 1,
    ):
        micro_x_value = ws.cell(
            row=row_number,
            column=micro_x_column,
        ).value

        field_value = (
            ws.cell(
                row=row_number,
                column=field_column,
            ).value
            if field_column is not None
            else None
        )

        micro_x_key = normalize_micro_x(
            micro_x_value
        )

        field_key = (
            normalize_text(field_value)
            if field_value is not None
            else ""
        )

        # This row may use an ambiguous Micro X image only when its own
        # nonblank Field occurs at least twice for the same Micro X.
        allow_shared_field_match = bool(
            field_key
            and micro_x_field_counts.get(
                micro_x_key,
                {},
            ).get(
                field_key,
                0,
            ) >= 2
        )

        image_path = find_best_image(
            micro_x_value,
            exact_index,
            clean_index,
            micro_x_index,
            allow_shared_field_match,
        )

        link_cell = ws.cell(
            row=row_number,
            column=link_column,
        )

        if image_path is None:
            link_cell.value = "Missing"
            link_cell.hyperlink = None
            missing += 1
        else:
            # Use an absolute path. On Windows, convert a mapped-drive
            # path to UNC when Windows can resolve it.
            absolute_path = str(
                image_path.resolve()
            )

            try:
                if os.name == "nt":
                    universal_name_info_level = (
                        0x00000001
                    )
                    error_more_data = 234

                    get_universal_name = (
                        ctypes.windll.mpr
                        .WNetGetUniversalNameW
                    )

                    get_universal_name.argtypes = [
                        wintypes.LPCWSTR,
                        wintypes.DWORD,
                        wintypes.LPVOID,
                        ctypes.POINTER(
                            wintypes.DWORD
                        ),
                    ]

                    size = wintypes.DWORD(0)

                    first_result = get_universal_name(
                        absolute_path,
                        universal_name_info_level,
                        None,
                        ctypes.byref(size),
                    )

                    if (
                        first_result == error_more_data
                        and size.value > 0
                    ):
                        buffer = (
                            ctypes.create_string_buffer(
                                size.value
                            )
                        )

                        second_result = (
                            get_universal_name(
                                absolute_path,
                                universal_name_info_level,
                                ctypes.byref(buffer),
                                ctypes.byref(size),
                            )
                        )

                        if second_result == 0:
                            unc_pointer = ctypes.cast(
                                buffer,
                                ctypes.POINTER(
                                    wintypes.LPWSTR
                                ),
                            )

                            unc_value = (
                                unc_pointer
                                .contents
                                .value
                            )

                            if unc_value:
                                absolute_path = unc_value

            except Exception:
                # Keep the absolute path if UNC conversion is unavailable.
                pass

            link_path = absolute_path.replace(
                "/",
                "\\",
            )

            formula_path = link_path.replace(
                '"',
                '""',
            )

            link_cell.value = (
                f'=HYPERLINK('
                f'"{formula_path}", '
                f'"Open Image")'
            )

            link_cell.hyperlink = None

            link_cell.font = Font(
                color="0563C1",
                underline="single",
            )

            link_cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
            )

            linked += 1

        ws.cell(
            row=row_number,
            column=area_orig_column,
        ).value = ws.cell(
            row=row_number,
            column=area_second_column,
        ).value

    ensure_new_columns_and_formulas(
        ws,
        area_orig_column,
        corr_column,
        new_brea_column,
        new_len_column,
        asso_column,
        breadth_column,
        length_column,
    )

    align_new_columns_center(
        ws,
        [
            area_orig_column,
            corr_column,
            new_brea_column,
            new_len_column,
            asso_column,
        ],
    )

    freeze_top_header_row(ws)
    shrink_columns_keep_visible(ws)

    # Reapply original hidden states after insertion while compensating
    # for the new-column index shift.
    for old_index in hidden_indices:
        if (
            inserted_count
            and old_index >= link_column
        ):
            new_index = (
                old_index + inserted_count
            )
        else:
            new_index = old_index

        new_dimension = ws.column_dimensions[
            get_column_letter(new_index)
        ]

        new_dimension.hidden = True

        if hidden_widths.get(old_index) is not None:
            new_dimension.width = hidden_widths[
                old_index
            ]

    return linked, missing, "OK"


def apply_column_visibility_by_choice(
    ws,
    visible_set: Optional[set],
):
    if visible_set is None:
        return

    for column_index in range(
        1,
        ws.max_column + 1,
    ):
        header = ws.cell(
            row=1,
            column=column_index,
        ).value

        header_text = (
            str(header).strip()
            if header is not None
            else ""
        )

        ws.column_dimensions[
            get_column_letter(column_index)
        ].hidden = (
            header_text not in visible_set
        )


def get_visibility_target_sheet_names(
    sheet_names: List[str],
) -> set:
    # Only sheets number 4 through the third-last are affected.
    if len(sheet_names) < 6:
        return set()

    return set(sheet_names[3:-2])


def main():
    root = tk.Tk()
    root.withdraw()

    try:
        logger = LiveLogger(root)

        workbook_path = choose_workbook(root)

        logger.write(
            f"Workbook selected: {workbook_path}"
        )

        root_directory = workbook_path.parent

        logger.write(
            f"Auto-using workbook directory "
            f"as category root: {root_directory}"
        )

        workbook = load_workbook(
            workbook_path
        )

        folder_map = build_category_folder_map(
            root_directory
        )

        logger.write(
            f"Detected category folders with "
            f"cropped: {len(folder_map)}"
        )

        target_visibility_sheet_names = (
            get_visibility_target_sheet_names(
                workbook.sheetnames
            )
        )

        all_headers = []

        for worksheet in workbook.worksheets:
            if (
                worksheet.title
                not in target_visibility_sheet_names
            ):
                continue

            for cell in worksheet[1]:
                if cell.value is None:
                    continue

                header_name = str(
                    cell.value
                ).strip()

                if (
                    header_name
                    and header_name not in all_headers
                ):
                    all_headers.append(
                        header_name
                    )

        for extra_header in [
            LINK_HEADER,
        ] + NEW_HEADERS:
            if extra_header not in all_headers:
                all_headers.append(
                    extra_header
                )

        if target_visibility_sheet_names:
            logger.write(
                "Opening visible-column selector GUI "
                "(applies sheet #4 to 3rd-last only)..."
            )

            root.deiconify()

            visible_set = (
                choose_visible_columns_listbox(
                    root,
                    all_headers,
                )
            )

            root.withdraw()

            logger.write(
                "Column visibility selection captured."
            )
        else:
            visible_set = None

            logger.write(
                "Workbook has fewer than 6 sheets; "
                "skipping visibility GUI."
            )

        pairs = []

        for sheet_name in workbook.sheetnames:
            category_folder = (
                resolve_folder_for_sheet(
                    sheet_name,
                    folder_map,
                )
            )

            if category_folder is not None:
                pairs.append(
                    (
                        sheet_name,
                        category_folder,
                    )
                )

        logger.write(
            "Opening category selector GUI..."
        )

        root.deiconify()

        selected_pairs = choose_pairs_gui(
            root,
            pairs,
            logger,
        )

        root.withdraw()

        if not selected_pairs:
            logger.write(
                "No categories selected. Exiting."
            )

            messagebox.showwarning(
                "No selection",
                "No categories selected.",
                parent=root,
            )
            return

        total_linked = 0
        total_missing = 0
        processed = 0
        skipped = []

        for (
            sheet_name,
            category_folder,
        ) in selected_pairs:
            logger.write(
                f"Processing: {sheet_name}"
            )

            linked, missing, status = (
                process_sheet(
                    workbook[sheet_name],
                    category_folder / "cropped",
                    workbook_path.parent,
                )
            )

            if status != "OK":
                skipped.append(
                    f"{sheet_name}: {status}"
                )

                logger.write(
                    f"  Skipped: {status}"
                )
                continue

            processed += 1
            total_linked += linked
            total_missing += missing

            if (
                sheet_name
                in target_visibility_sheet_names
            ):
                apply_column_visibility_by_choice(
                    workbook[sheet_name],
                    visible_set,
                )

            logger.write(
                f"  Done links={linked}, "
                f"missing={missing}"
            )

        output_path = workbook_path.with_name(
            f"{workbook_path.stem}"
            f"_with_quick_links"
            f"{workbook_path.suffix}"
        )

        logger.write(
            f"Saving: {output_path}"
        )

        workbook.save(output_path)

        logger.write("Save complete.")

        formula_info = (
            "New Brea/New Len formula:\n"
            "IF(Corr% blank or 0, keep original "
            "Breadth/Length, else "
            "SQRT(Area Orig * Corr% / 100))"
        )

        summary = (
            f"Saved: {output_path}\n"
            f"Processed sheets: {processed}\n"
            f"Links added: {total_linked}\n"
            f"Missing links: {total_missing}\n\n"
            f"{formula_info}"
        )

        if skipped:
            summary += (
                "\n\nSkipped:\n"
                + "\n".join(skipped[:20])
            )

        messagebox.showinfo(
            "Done",
            summary,
            parent=root,
        )

    finally:
        root.destroy()


if __name__ == "__main__":
    main()