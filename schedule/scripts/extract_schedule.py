from pathlib import Path
import re
import sys

from openpyxl import load_workbook


DEFAULT_WORKBOOK = "Studio_schedule2026.xlsx"


def get_workbook_path() -> Path:
    """
    Use a workbook path supplied on the command line if present.
    Otherwise use the filename downloaded by the GitHub workflow.
    """
    if len(sys.argv) > 1:
        return Path(sys.argv[1])

    return Path(DEFAULT_WORKBOOK)


def normalize_defined_name_formula(value: str) -> str:
    """
    Excel defined names may be returned with or without a leading '='.
    Normalize them for easier processing.
    """
    value = value.strip()

    if value.startswith("="):
        value = value[1:]

    return value


def is_direct_range_reference(value: str) -> bool:
    """
    Determine whether a defined name points directly to a worksheet range.

    Examples:
        Aug!$A$50:$H$64
        'Some Sheet'!$A$1:$H$20
    """
    return "!" in value


def parse_direct_range(value: str):
    """
    Convert an Excel reference such as:

        Aug!$A$50:$H$64

    into:

        ("Aug", "A50:H64")
    """
    match = re.match(
        r"^'?(.+?)'?!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$",
        value,
    )

    if not match:
        raise ValueError(f"Could not parse range reference: {value}")

    sheet_name = match.group(1)

    cell_range = (
        f"{match.group(2)}{match.group(3)}:"
        f"{match.group(4)}{match.group(5)}"
    )

    return sheet_name, cell_range


def resolve_defined_name(workbook, name: str, visited=None):
    """
    Resolve a workbook defined name recursively.

    This handles the structure already used in the schedule workbook:

        Web_ThisWeek -> AugWk4 -> Aug!A50:H64
    """
    if visited is None:
        visited = set()

    if name in visited:
        chain = " -> ".join(list(visited) + [name])
        raise ValueError(f"Circular defined-name reference detected: {chain}")

    visited.add(name)

    defined_name = workbook.defined_names.get(name)

    if defined_name is None:
        raise KeyError(f"Defined name not found: {name}")

    reference = normalize_defined_name_formula(defined_name.attr_text)

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

    # Otherwise this defined name points to another defined name.
    resolved = resolve_defined_name(
        workbook,
        reference,
        visited.copy(),
    )

    return {
        "defined_name": name,
        "points_to": reference,
        "sheet": resolved["sheet"],
        "range": resolved["range"],
    }


def web_name_sort_key(name: str):
    """
    Put the web aliases in their intended display order.

    Web_ThisWeek
    Web_NextWeek
    Web_WeekPlus2
    Web_WeekPlus3
    Web_WeekPlus4
    ...
    """
    if name == "Web_ThisWeek":
        return 0

    if name == "Web_NextWeek":
        return 1

    match = re.match(r"Web_WeekPlus(\d+)$", name)

    if match:
        return int(match.group(1))

    return 999


def main():
    workbook_path = get_workbook_path()

    print(f"Opening workbook: {workbook_path}")

    if not workbook_path.exists():
        raise FileNotFoundError(
            f"Workbook does not exist: {workbook_path}"
        )

    workbook = load_workbook(
        workbook_path,
        data_only=False,
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
            "No workbook-level defined names beginning with 'Web_' were found."
        )

    print()
    print(f"Found {len(web_names)} web schedule aliases:")
    print()

    for name in web_names:
        resolved = resolve_defined_name(workbook, name)

        intermediate = resolved.get("points_to")

        if intermediate:
            print(
                f"{name}"
                f" -> {intermediate}"
                f" -> {resolved['sheet']}!{resolved['range']}"
            )
        else:
            print(
                f"{name}"
                f" -> {resolved['sheet']}!{resolved['range']}"
            )

    print()
    print(
        "SUCCESS: All Web_ schedule aliases resolved "
        "to valid workbook ranges."
    )


if __name__ == "__main__":
    main()