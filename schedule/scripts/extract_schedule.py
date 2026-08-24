from pathlib import Path
import json
import re
import sys

from openpyxl import load_workbook


DEFAULT_WORKBOOK = "Studio_schedule2026.xlsx"
DEFAULT_OUTPUT = "schedule/data/schedule.json"


def get_workbook_path() -> Path:
    if len(sys.argv) > 1:
        return Path(sys.argv[1])

    return Path(DEFAULT_WORKBOOK)


def normalize_defined_name_formula(value: str) -> str:
    value = value.strip()

    if value.startswith("="):
        value = value[1:]

    return value


def is_direct_range_reference(value: str) -> bool:
    return "!" in value


def parse_direct_range(value: str):
    match = re.match(
        r"^'?(.+?)'?!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$",
        value,
    )

    if not match:
        raise ValueError(
            f"Could not parse range reference: {value}"
        )

    sheet_name = match.group(1)

    cell_range = (
        f"{match.group(2)}{match.group(3)}:"
        f"{match.group(4)}{match.group(5)}"
    )

    return sheet_name, cell_range


def resolve_defined_name(workbook, name: str, visited=None):
    if visited is None:
        visited = []

    if name in visited:
        chain = " -> ".join(visited + [name])

        raise ValueError(
            f"Circular defined-name reference detected: {chain}"
        )

    visited = visited + [name]

    defined_name = workbook.defined_names.get(name)

    if defined_name is None:
        raise KeyError(
            f"Defined name not found: {name}"
        )

    reference = normalize_defined_name_formula(
        defined_name.attr_text
    )

    if is_direct_range_reference(reference):
        sheet_name, cell_range = parse_direct_range(reference)

        if sheet_name not in workbook.sheetnames:
            raise KeyError(
                f"Defined name {name} refers to missing worksheet: "
                f"{sheet_name}"
            )

        return {
            "defined_name": name,
            "sheet": sheet_name,
            "range": cell_range,
        }

    resolved = resolve_defined_name(
        workbook,
        reference,
        visited,
    )

    return {
        "defined_name": name,
        "points_to": reference,
        "sheet": resolved["sheet"],
        "range": resolved["range"],
    }


def web_name_sort_key(name: str):
    if name == "Web_ThisWeek":
        return 0

    if name == "Web_NextWeek":
        return 1

    match = re.match(r"Web_WeekPlus(\d+)$", name)

    if match:
        return int(match.group(1))

    return 999


def find_merge_for_cell(worksheet, cell):
    """
    Return merge information for a cell.

    Only the upper-left cell of a merged range is emitted.
    The other covered cells are omitted from JSON.
    """
    for merged_range in worksheet.merged_cells.ranges:
        if cell.coordinate not in merged_range:
            continue

        if (
            cell.row == merged_range.min_row
            and cell.column == merged_range.min_col
        ):
            return {
                "rowspan":
                    merged_range.max_row
                    - merged_range.min_row
                    + 1,
                "colspan":
                    merged_range.max_col
                    - merged_range.min_col
                    + 1,
            }

        return {
            "covered": True,
        }

    return None


def get_row_tone(first_cell, is_header):
    """
    Translate the workbook's alternating day fills into
    simple semantic tones.

    In this workbook:
      theme 9 = green
      theme 7 = blue
    """
    if is_header:
        return "neutral"

    fill = first_cell.fill

    if fill.fill_type == "solid":
        color = fill.fgColor

        if color.type == "theme":
            if color.theme == 9:
                return "green"

            if color.theme == 7:
                return "blue"

    return "neutral"


def is_unavailable_cell(cell):
    """
    Detect Excel's hatched unavailable cells.
    """
    return cell.fill.fill_type == "gray125"


def classify_cell(
    cell,
    text,
    relative_column,
    is_header,
):
    """
    Assign a semantic role used later by CSS/JavaScript.

    Relative columns:
      0 = day/date
      1 = time
      2-6 = apparatus/client slots
      7 = instructor
    """
    clean_text = text.strip()
    upper_text = clean_text.upper()

    if is_header:
        return "column-header"

    if is_unavailable_cell(cell):
        return "unavailable"

    if upper_text == "OPEN":
        return "open"

    if upper_text.startswith("NO CLASSES -"):
        return "no-classes-block"

    if upper_text.startswith("CLASS CANCELED"):
        return "class-canceled"

    if upper_text.startswith("NO CLASS"):
        return "no-class"

    if upper_text.startswith("MAT:"):
        return "mat"

    if relative_column == 0:
        return "day"

    if relative_column == 1:
        return "time"

    if relative_column == 7:
        return "instructor"

    if 2 <= relative_column <= 6:
        return "client"

    return "cell"


def extract_geometry(worksheet, cells):
    """
    Preserve the fixed Excel geometry once per week,
    rather than repeating it on every cell.
    """
    first_row = cells[0]

    column_widths = []

    for cell in first_row:
        column_letter = cell.column_letter

        width = worksheet.column_dimensions[
            column_letter
        ].width

        column_widths.append(width)

    row_heights = []

    for excel_row in cells:
        row_number = excel_row[0].row

        height = worksheet.row_dimensions[
            row_number
        ].height

        row_heights.append(height)

    return {
        "columnWidths": column_widths,
        "rowHeights": row_heights,
    }


def extract_week(workbook, web_name):
    resolved = resolve_defined_name(
        workbook,
        web_name,
    )

    worksheet = workbook[resolved["sheet"]]
    cells = worksheet[resolved["range"]]

    rows = []

    for row_index, excel_row in enumerate(cells):
        is_header = row_index == 0

        row_tone = get_row_tone(
            excel_row[0],
            is_header,
        )

        row_data = {
            "tone": row_tone,
            "cells": [],
        }

        for relative_column, cell in enumerate(excel_row):
            merge_info = find_merge_for_cell(
                worksheet,
                cell,
            )

            if (
                merge_info
                and merge_info.get("covered")
            ):
                continue

            value = cell.value

            if value is None:
                text = ""
            else:
                text = str(value)

            cell_data = {
                "text": text,
                "style": classify_cell(
                    cell,
                    text,
                    relative_column,
                    is_header,
                ),
            }

            if cell.font.bold:
                cell_data["bold"] = True

            if merge_info:
                if merge_info["rowspan"] > 1:
                    cell_data["rowspan"] = (
                        merge_info["rowspan"]
                    )

                if merge_info["colspan"] > 1:
                    cell_data["colspan"] = (
                        merge_info["colspan"]
                    )

            row_data["cells"].append(
                cell_data
            )

        rows.append(row_data)

    geometry = extract_geometry(
        worksheet,
        cells,
    )

    return {
        "key": web_name,
        "sourceName": resolved.get(
            "points_to",
            web_name,
        ),
        "sourceSheet": resolved["sheet"],
        "sourceRange": resolved["range"],
        "rowCount": len(cells),
        "columnCount": len(cells[0]),
        "columnWidths": geometry["columnWidths"],
        "rowHeights": geometry["rowHeights"],
        "rows": rows,
    }


def validate_week(week):
    """
    Basic sanity checks before allowing the week
    into the published JSON.
    """
    if week["columnCount"] != 8:
        raise ValueError(
            f"{week['key']} contains "
            f"{week['columnCount']} columns; expected 8."
        )

    if week["rowCount"] < 2:
        raise ValueError(
            f"{week['key']} contains too few rows."
        )

    if not week["rows"]:
        raise ValueError(
            f"{week['key']} contains no schedule data."
        )

    header_cells = week["rows"][0]["cells"]

    if len(header_cells) < 8:
        raise ValueError(
            f"{week['key']} header is incomplete."
        )


def main():
    workbook_path = get_workbook_path()
    output_path = Path(DEFAULT_OUTPUT)

    print(
        f"Opening workbook: {workbook_path}"
    )

    if not workbook_path.exists():
        raise FileNotFoundError(
            f"Workbook does not exist: "
            f"{workbook_path}"
        )

    workbook = load_workbook(
        workbook_path,
        data_only=True,
        read_only=False,
    )

    web_names = sorted(
        [
            name
            for name in workbook.defined_names.keys()
            if name.startswith("Web_")
        ],
        key=web_name_sort_key,
    )

    if not web_names:
        raise RuntimeError(
            "No workbook-level defined names beginning "
            "with 'Web_' were found."
        )

    print(
        f"Found {len(web_names)} "
        f"web schedule aliases."
    )

    weeks = []

    for web_name in web_names:
        resolved = resolve_defined_name(
            workbook,
            web_name,
        )

        intermediate = resolved.get("points_to")

        if intermediate:
            print(
                f"{web_name}"
                f" -> {intermediate}"
                f" -> {resolved['sheet']}"
                f"!{resolved['range']}"
            )
        else:
            print(
                f"{web_name}"
                f" -> {resolved['sheet']}"
                f"!{resolved['range']}"
            )

        week = extract_week(
            workbook,
            web_name,
        )

        validate_week(week)

        weeks.append(week)

    schedule_data = {
        "schemaVersion": 2,
        "weeks": weeks,
    }

    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    with output_path.open(
        "w",
        encoding="utf-8",
    ) as output_file:
        json.dump(
            schedule_data,
            output_file,
            indent=2,
            ensure_ascii=False,
        )

        output_file.write("\n")

    print()
    print(
        f"Created: {output_path}"
    )

    print(
        f"Extracted {len(weeks)} "
        f"web schedule weeks."
    )

    print(
        "SUCCESS: Lean public schedule JSON "
        "was generated."
    )


if __name__ == "__main__":
    main()