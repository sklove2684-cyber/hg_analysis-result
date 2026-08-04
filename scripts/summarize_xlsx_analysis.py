from __future__ import annotations

import argparse
import json
from pathlib import Path
import re


def column_number(ref: str) -> int:
    letters = re.match(r"[A-Z]+", ref).group()
    value = 0
    for char in letters:
        value = value * 26 + ord(char) - 64
    return value


def cell_text(cell: dict) -> str:
    if cell["formula"] is not None:
        return f"={cell['formula']}" if cell["formula"] else "=<shared formula>"
    return repr(cell["value"])


def print_ranges(
    data: dict,
    sheet_name: str,
    ranges: list[tuple[int, int]],
    min_column: int | None = None,
    max_column: int | None = None,
    include_empty: bool = False,
) -> None:
    sheet = next(value for value in data["sheets"] if value["name"] == sheet_name)
    cells = {cell["ref"]: cell for cell in sheet["cells"]}
    print(f"### {sheet_name}")
    for low, high in ranges:
        print(f"RANGE {low}:{high}")
        for row in range(low, high + 1):
            values = []
            for ref, cell in cells.items():
                match = re.fullmatch(r"[A-Z]+(\d+)", ref)
                if int(match.group(1)) != row:
                    continue
                column = column_number(ref)
                if min_column is not None and column < min_column:
                    continue
                if max_column is not None and column > max_column:
                    continue
                if not include_empty and cell["value"] in (None, "") and cell["formula"] is None:
                    continue
                suffix = f" [style {cell['style_id']}]" if include_empty else ""
                values.append((column, f"{ref}={cell_text(cell)}{suffix}"))
            if values:
                print(row, " | ".join(text for _, text in sorted(values)))


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("analysis", type=Path)
    parser.add_argument("--sheet", required=True)
    parser.add_argument("--ranges", required=True, help="Example: 1-20,34-60")
    parser.add_argument("--columns", help="Example: F-AA")
    parser.add_argument("--include-empty", action="store_true")
    args = parser.parse_args()
    data = json.loads(args.analysis.read_text(encoding="utf-8"))
    ranges = []
    for part in args.ranges.split(","):
        low, high = part.split("-", 1)
        ranges.append((int(low), int(high)))
    min_column = max_column = None
    if args.columns:
        start, end = args.columns.split("-", 1)
        min_column = column_number(start)
        max_column = column_number(end)
    print_ranges(data, args.sheet, ranges, min_column, max_column, args.include_empty)


if __name__ == "__main__":
    main()
