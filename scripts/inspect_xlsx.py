from __future__ import annotations

import argparse
from collections import Counter, defaultdict
import json
from pathlib import Path, PurePosixPath
import re
from xml.etree import ElementTree as ET
import zipfile


MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
DRAW = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"
DRAW_TEXT = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"m": MAIN, "r": REL, "pr": PKG_REL, "xdr": DRAW, "c": CHART, "a": DRAW_TEXT}


CELL_REF_RE = re.compile(
    r"(?:(?:'([^']+)'|([A-Za-z0-9가-힣_]+))!)?"
    r"\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?"
)
FUNCTION_RE = re.compile(r"\b([A-Z][A-Z0-9_.]*)\s*\(", re.IGNORECASE)


def col_number(value: str) -> int:
    number = 0
    for char in value:
        number = number * 26 + ord(char) - 64
    return number


def col_name(number: int) -> str:
    output = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        output = chr(65 + remainder) + output
    return output


def split_ref(value: str) -> tuple[int, int]:
    match = re.fullmatch(r"([A-Z]+)(\d+)", value)
    if not match:
        raise ValueError(value)
    return col_number(match.group(1)), int(match.group(2))


def normalized_part(base: str, target: str) -> str:
    base_dir = PurePosixPath(base).parent
    parts: list[str] = []
    source = PurePosixPath(target.lstrip("/")) if target.startswith("/") else base_dir / target
    for part in source.parts:
        if part == "..":
            if parts:
                parts.pop()
        elif part != ".":
            parts.append(part)
    return "/".join(parts)


def relationships(archive: zipfile.ZipFile, path: str) -> dict[str, str]:
    part = PurePosixPath(path)
    rel_path = str(part.parent / "_rels" / f"{part.name}.rels")
    if rel_path not in archive.namelist():
        return {}
    root = ET.fromstring(archive.read(rel_path))
    return {
        node.attrib["Id"]: normalized_part(path, node.attrib["Target"])
        for node in root
    }


def read_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    return ["".join(node.text or "" for node in si.findall(".//m:t", NS)) for si in root]


def read_styles(archive: zipfile.ZipFile) -> dict:
    root = ET.fromstring(archive.read("xl/styles.xml"))
    custom_formats = {
        int(node.attrib["numFmtId"]): node.attrib["formatCode"]
        for node in root.findall("m:numFmts/m:numFmt", NS)
    }
    xfs = []
    for index, node in enumerate(root.findall("m:cellXfs/m:xf", NS)):
        alignment = node.find("m:alignment", NS)
        xfs.append(
            {
                "style_id": index,
                "numFmtId": int(node.attrib.get("numFmtId", 0)),
                "numFmt": custom_formats.get(int(node.attrib.get("numFmtId", 0))),
                "fontId": int(node.attrib.get("fontId", 0)),
                "fillId": int(node.attrib.get("fillId", 0)),
                "borderId": int(node.attrib.get("borderId", 0)),
                "alignment": alignment.attrib if alignment is not None else {},
            }
        )
    return {
        "fonts": len(root.findall("m:fonts/m:font", NS)),
        "fills": len(root.findall("m:fills/m:fill", NS)),
        "borders": len(root.findall("m:borders/m:border", NS)),
        "cell_styles": len(xfs),
        "named_cell_styles": len(root.findall("m:cellStyles/m:cellStyle", NS)),
        "custom_number_formats": custom_formats,
        "cell_xfs": xfs,
    }


def cell_value(node: ET.Element, shared_strings: list[str]):
    cell_type = node.attrib.get("t")
    value_node = node.find("m:v", NS)
    if cell_type == "inlineStr":
        return "".join(x.text or "" for x in node.findall(".//m:t", NS))
    if value_node is None or value_node.text is None:
        return None
    value = value_node.text
    if cell_type == "s":
        return shared_strings[int(value)]
    if cell_type in {"str", "e"}:
        return value
    if cell_type == "b":
        return value == "1"
    try:
        return int(value) if re.fullmatch(r"-?\d+", value) else float(value)
    except ValueError:
        return value


def read_chart(archive: zipfile.ZipFile, path: str) -> dict:
    root = ET.fromstring(archive.read(path))
    title = " ".join(
        text.strip() for text in (node.text or "" for node in root.findall(".//a:t", NS)) if text.strip()
    )
    formulas = [node.text for node in root.findall(".//c:f", NS) if node.text]
    chart_types = [
        node.tag.rsplit("}", 1)[-1]
        for node in root.findall(".//c:plotArea/*", NS)
        if node.tag.endswith("Chart")
    ]
    return {"path": path, "title": title or None, "chart_types": chart_types, "series_formulas": formulas}


def read_drawing(archive: zipfile.ZipFile, path: str) -> dict:
    root = ET.fromstring(archive.read(path))
    rels = relationships(archive, path)
    anchors = []
    for anchor in list(root):
        from_node = anchor.find("xdr:from", NS)
        to_node = anchor.find("xdr:to", NS)
        chart = anchor.find(".//c:chart", NS)
        image = anchor.find(".//a:blip", NS)
        anchors.append(
            {
                "type": anchor.tag.rsplit("}", 1)[-1],
                "from": {
                    "col": int(from_node.findtext("xdr:col", "0", NS)) if from_node is not None else None,
                    "row": int(from_node.findtext("xdr:row", "0", NS)) if from_node is not None else None,
                },
                "to": {
                    "col": int(to_node.findtext("xdr:col", "0", NS)) if to_node is not None else None,
                    "row": int(to_node.findtext("xdr:row", "0", NS)) if to_node is not None else None,
                },
                "chart_path": rels.get(chart.attrib.get(f"{{{REL}}}id")) if chart is not None else None,
                "image_path": rels.get(image.attrib.get(f"{{{REL}}}embed")) if image is not None else None,
            }
        )
    return {"path": path, "anchors": anchors}


def iter_formula_references(formula: str, formula_sheet: str):
    for match in CELL_REF_RE.finditer(formula):
        sheet = match.group(1) or match.group(2) or formula_sheet
        start_col, start_row = col_number(match.group(3)), int(match.group(4))
        end_col = col_number(match.group(5)) if match.group(5) else start_col
        end_row = int(match.group(6)) if match.group(6) else start_row
        if (end_col - start_col + 1) * (end_row - start_row + 1) > 2000:
            continue
        for row in range(min(start_row, end_row), max(start_row, end_row) + 1):
            for column in range(min(start_col, end_col), max(start_col, end_col) + 1):
                yield sheet, f"{col_name(column)}{row}"


def inspect_workbook(path: Path) -> dict:
    with zipfile.ZipFile(path) as archive:
        shared_strings = read_shared_strings(archive)
        styles = read_styles(archive)
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        workbook_rels = relationships(archive, "xl/workbook.xml")
        sheets: list[dict] = []
        cell_index: dict[tuple[str, str], dict] = {}
        inbound: Counter[tuple[str, str]] = Counter()
        all_formulas: list[tuple[str, str, str]] = []
        for sheet_node in workbook.find("m:sheets", NS):
            name = sheet_node.attrib["name"]
            sheet_path = workbook_rels[sheet_node.attrib[f"{{{REL}}}id"]]
            root = ET.fromstring(archive.read(sheet_path))
            sheet_rels = relationships(archive, sheet_path)
            cells = []
            formula_functions: Counter[str] = Counter()
            formula_errors = Counter()
            style_usage = Counter()
            shared_formula_masters = []
            for cell in root.findall(".//m:c", NS):
                ref = cell.attrib.get("r")
                formula_node = cell.find("m:f", NS)
                formula = (formula_node.text or "") if formula_node is not None else None
                value = cell_value(cell, shared_strings)
                style_id = int(cell.attrib.get("s", 0))
                record = {
                    "ref": ref,
                    "value": value,
                    "formula": formula,
                    "formula_type": formula_node.attrib.get("t") if formula_node is not None else None,
                    "style_id": style_id,
                    "cell_type": cell.attrib.get("t"),
                }
                cells.append(record)
                cell_index[(name, ref)] = record
                style_usage[style_id] += 1
                if formula is not None:
                    if formula_node.attrib.get("t") == "shared" and formula_node.text:
                        shared_formula_masters.append(
                            {
                                "ref": ref,
                                "shared_index": formula_node.attrib.get("si"),
                                "applies_to": formula_node.attrib.get("ref"),
                                "formula": formula,
                            }
                        )
                    if formula:
                        all_formulas.append((name, ref, formula))
                    formula_functions.update(value.upper() for value in FUNCTION_RE.findall(formula))
                    for error in ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A"):
                        if error in formula:
                            formula_errors[error] += 1
            dimension = root.find("m:dimension", NS)
            sheet_view = root.find("m:sheetViews/m:sheetView", NS)
            auto_filter = root.find("m:autoFilter", NS)
            page_setup = root.find("m:pageSetup", NS)
            page_margins = root.find("m:pageMargins", NS)
            print_options = root.find("m:printOptions", NS)
            drawing_node = root.find("m:drawing", NS)
            drawing_path = (
                sheet_rels.get(drawing_node.attrib[f"{{{REL}}}id"])
                if drawing_node is not None else None
            )
            drawing = read_drawing(archive, drawing_path) if drawing_path else None
            charts = []
            if drawing:
                for anchor in drawing["anchors"]:
                    chart_path = anchor["chart_path"]
                    if chart_path:
                        chart = read_chart(archive, chart_path)
                        chart["anchor"] = {"from": anchor["from"], "to": anchor["to"]}
                        charts.append(chart)
            sheets.append(
                {
                    "name": name,
                    "state": sheet_node.attrib.get("state", "visible"),
                    "xml_path": sheet_path,
                    "dimension": dimension.attrib.get("ref") if dimension is not None else None,
                    "cell_count": len(cells),
                    "formula_count": sum(cell["formula"] is not None for cell in cells),
                    "hardcoded_numeric_count": sum(
                        cell["formula"] is None and isinstance(cell["value"], (int, float))
                        for cell in cells
                    ),
                    "formula_functions": formula_functions.most_common(),
                    "shared_formula_masters": shared_formula_masters,
                    "formula_errors": dict(formula_errors),
                    "style_usage": style_usage.most_common(),
                    "merged_ranges": [
                        node.attrib["ref"] for node in root.findall("m:mergeCells/m:mergeCell", NS)
                    ],
                    "hidden_rows": [
                        int(node.attrib["r"]) for node in root.findall("m:sheetData/m:row", NS)
                        if node.attrib.get("hidden") == "1"
                    ],
                    "hidden_columns": [
                        {"min": int(node.attrib["min"]), "max": int(node.attrib["max"])}
                        for node in root.findall("m:cols/m:col", NS)
                        if node.attrib.get("hidden") == "1"
                    ],
                    "custom_row_heights": [
                        {"row": int(node.attrib["r"]), "height": float(node.attrib["ht"])}
                        for node in root.findall("m:sheetData/m:row", NS)
                        if "ht" in node.attrib
                    ],
                    "column_definitions": [node.attrib for node in root.findall("m:cols/m:col", NS)],
                    "conditional_formatting": [
                        node.attrib.get("sqref")
                        for node in root.findall("m:conditionalFormatting", NS)
                    ],
                    "data_validations": [
                        node.attrib for node in root.findall("m:dataValidations/m:dataValidation", NS)
                    ],
                    "auto_filter": auto_filter.attrib if auto_filter is not None else None,
                    "sheet_view": sheet_view.attrib if sheet_view is not None else {},
                    "page_setup": page_setup.attrib if page_setup is not None else {},
                    "page_margins": page_margins.attrib if page_margins is not None else {},
                    "print_options": print_options.attrib if print_options is not None else {},
                    "drawing": drawing,
                    "charts": charts,
                    "cells": cells,
                }
            )
        for formula_sheet, _, formula in all_formulas:
            for target in iter_formula_references(formula, formula_sheet):
                inbound[target] += 1
        candidates = []
        referenced_blanks = []
        for (sheet_name, ref), count in inbound.most_common():
            cell = cell_index.get((sheet_name, ref))
            if not cell or cell["formula"] is not None:
                continue
            column, row = split_ref(ref)
            context = []
            for c, r in ((column - 2, row), (column - 1, row), (column + 1, row),
                         (column + 2, row), (column, row - 1), (column, row + 1)):
                if c < 1 or r < 1:
                    continue
                neighbor_ref = f"{col_name(c)}{r}"
                neighbor = cell_index.get((sheet_name, neighbor_ref))
                if neighbor and neighbor["value"] not in (None, ""):
                    context.append({"ref": neighbor_ref, "value": neighbor["value"], "formula": neighbor["formula"]})
            if cell["value"] in (None, ""):
                referenced_blanks.append(
                    {
                        "sheet": sheet_name, "ref": ref, "style_id": cell["style_id"],
                        "inbound_formula_references": count, "context": context,
                    }
                )
                continue
            if not isinstance(cell["value"], (int, float)):
                continue
            candidates.append(
                {
                    "sheet": sheet_name, "ref": ref, "value": cell["value"],
                    "style_id": cell["style_id"], "inbound_formula_references": count,
                    "context": context,
                }
            )
        defined_names_node = workbook.find("m:definedNames", NS)
        defined_names = []
        if defined_names_node is not None:
            defined_names = [
                {"name": node.attrib.get("name"), "localSheetId": node.attrib.get("localSheetId"), "ref": node.text}
                for node in defined_names_node
            ]
        calc = workbook.find("m:calcPr", NS)
        return {
            "file": str(path),
            "file_size": path.stat().st_size,
            "zip_entry_count": len(archive.namelist()),
            "sheet_count": len(sheets),
            "defined_names": defined_names,
            "calculation_properties": calc.attrib if calc is not None else {},
            "styles": styles,
            "sheets": sheets,
            "input_candidates": candidates,
            "referenced_blank_candidates": referenced_blanks,
            "external_links": sorted(name for name in archive.namelist() if name.startswith("xl/externalLinks/")),
            "media": sorted(name for name in archive.namelist() if name.startswith("xl/media/")),
            "vba": "xl/vbaProject.bin" in archive.namelist(),
        }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("xlsx", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    result = inspect_workbook(args.xlsx)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(
        json.dumps(
            {
                "sheets": result["sheet_count"],
                "formulas": sum(sheet["formula_count"] for sheet in result["sheets"]),
                "charts": sum(len(sheet["charts"]) for sheet in result["sheets"]),
                "input_candidates": len(result["input_candidates"]),
            },
            ensure_ascii=False,
        )
    )


if __name__ == "__main__":
    main()
