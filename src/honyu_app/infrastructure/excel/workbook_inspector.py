from __future__ import annotations

from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET
import zipfile

from honyu_app.domain.errors import ValidationError
from honyu_app.services.excel_template_service import (
    ExcelTemplateSnapshot,
    TemplateCell,
)


MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"m": MAIN, "r": REL, "pr": PKG_REL}


def _normalized_part(base: str, target: str) -> str:
    base_dir = PurePosixPath(base).parent
    source = PurePosixPath(target.lstrip("/")) if target.startswith("/") else base_dir / target
    parts: list[str] = []
    for part in source.parts:
        if part == "..":
            if parts:
                parts.pop()
        elif part != ".":
            parts.append(part)
    return "/".join(parts)


def _relationships(archive: zipfile.ZipFile, path: str) -> dict[str, str]:
    part = PurePosixPath(path)
    rel_path = str(part.parent / "_rels" / f"{part.name}.rels")
    if rel_path not in archive.namelist():
        return {}
    root = ET.fromstring(archive.read(rel_path))
    return {
        node.attrib["Id"]: _normalized_part(path, node.attrib["Target"])
        for node in root
    }


def _shared_strings(archive: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    return ["".join(node.text or "" for node in item.findall(".//m:t", NS)) for item in root]


def _cell_value(node: ET.Element, shared: list[str]) -> tuple[object | None, str]:
    cell_type = node.attrib.get("t")
    value_node = node.find("m:v", NS)
    if cell_type == "inlineStr":
        return "".join(value.text or "" for value in node.findall(".//m:t", NS)), "string"
    if value_node is None or value_node.text is None:
        return None, "blank"
    raw = value_node.text
    if cell_type == "s":
        return shared[int(raw)], "string"
    if cell_type in {"str", "e"}:
        return raw, "error" if cell_type == "e" else "string"
    if cell_type == "b":
        return raw == "1", "boolean"
    try:
        return int(raw), "numeric"
    except ValueError:
        try:
            return float(raw), "numeric"
        except ValueError:
            return raw, "string"


class XlsxTemplateInspector:
    """Reads only workbook and cell metadata needed by the Excel preview."""

    def inspect(self, path: Path) -> ExcelTemplateSnapshot:
        path = Path(path)
        if not path.is_file():
            raise ValidationError(f"Excel 템플릿을 찾을 수 없습니다: {path}")
        try:
            with zipfile.ZipFile(path) as archive:
                shared = _shared_strings(archive)
                workbook_path = "xl/workbook.xml"
                workbook = ET.fromstring(archive.read(workbook_path))
                workbook_rels = _relationships(archive, workbook_path)
                sheet_names: list[str] = []
                cells: dict[tuple[str, str], TemplateCell] = {}
                for sheet_node in workbook.findall("m:sheets/m:sheet", NS):
                    sheet_name = sheet_node.attrib["name"]
                    sheet_names.append(sheet_name)
                    relation_id = sheet_node.attrib[f"{{{REL}}}id"]
                    sheet_path = workbook_rels[relation_id]
                    root = ET.fromstring(archive.read(sheet_path))
                    for node in root.findall(".//m:c", NS):
                        address = node.attrib["r"]
                        formula_node = node.find("m:f", NS)
                        formula = None
                        if formula_node is not None:
                            formula = formula_node.text or "<shared-formula>"
                        value, value_type = _cell_value(node, shared)
                        cells[(sheet_name, address)] = TemplateCell(
                            sheet=sheet_name,
                            address=address,
                            exists=True,
                            value=value,
                            value_type="formula" if formula_node is not None else value_type,
                            formula=formula,
                            style_id=int(node.attrib.get("s", 0)),
                        )
                return ExcelTemplateSnapshot(path, tuple(sheet_names), cells)
        except (OSError, KeyError, ET.ParseError, zipfile.BadZipFile) as exc:
            raise ValidationError(f"Excel 템플릿 구조를 읽을 수 없습니다: {path}") from exc
