from pathlib import Path
import json
import re
import sys

from openpyxl import load_workbook
from openpyxl.styles.colors import COLOR_INDEXED


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


def color_to_hex(color):
    """
    Convert common Excel colors to #RRGGBB.

    Theme colors are intentionally identified separately for now.
    We don't want to guess at theme transformations during this
    first extraction pass.
    """
    if color is None:
        return None

    if color.type == "rgb" and color.rgb:
        return f"#{color.rgb[-6:]}"

    if color.type == "indexed":
        index = color.indexed

        if (
            index is not None
            and isinstance(index, int)
            and index < len(COLOR_INDEXED)
        ):
            value = COLOR_INDEXED[index]

            if value:
                return f"#{value[-6:]}"

    if color.type == "theme":
        return {
            "theme": color.theme,
            "tint": color.tint,
        }

    return None


def border_style(side):
    if side is None or side.style is None:
        return None

    return {
        "style": side.style,
        "color": color_to_hex(side.color),
    }


def extract_cell_style(cell):
    fill = cell.fill

    fill_info = {
        "type": fill.fill_type,
        "foreground": color_to_hex(fill.fgColor),
        "background": color_to_hex(fill.bgColor),
    }

    return {
        "bold": bool(cell.font.bold),
        "italic": bool(cell.font.italic),
        "fontSize": cell.font.sz,
        "fill": fill_info,
        "alignment": {
            "horizontal": cell.alignment.horizontal,
            "vertical": cell.alignment.vertical,
            "wrapText": bool(cell.alignment.wrap_text),
        },
        "borders": {
            "left": border_style(cell.border.left),
            "right": border_style(cell.border.right),
            "top": border_style(cell.border.top),
            "bottom": border_style(cell.border.bottom),
        },
    }


def find_merge_for_cell(worksheet, cell):
    """
    Return merge information only for the upper-left cell
    of a merged range.

    Covered cells do not need to be represented separately
    in the JSON.
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


def extract_week(workbook, web_name):
    resolved = resolve_defined_name(
        workbook,
        web_name,
    )

    worksheet = workbook[resolved["sheet"]]

    cells = worksheet[resolved["range"]]

    rows = []

    for excel_row in cells:
        row_data = []

        for cell in excel_row:
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
                value = ""
            else:
                value = str(value)

            cell_data = {
                "text": value,
                "style": extract_cell_style(cell),
            }

            if merge_info:
                if merge_info["rowspan"] > 1:
                    cell_data["rowspan"] = (
                        merge_info["rowspan"]
                    )

                if merge_info["colspan"] > 1:
                    cell_data["colspan"] = (
                        merge_info["colspan"]
                    )

            row_data.append(cell_data)

        rows.append(row_data)

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
        "rows": rows,
    }


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

        weeks.append(
            extract_week(
                workbook,
                web_name,
            )
        )

    schedule_data = {
        "schemaVersion": 1,
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
        "SUCCESS: Public schedule JSON "
        "was generated."
    )


if __name__ == "__main__":
    main()