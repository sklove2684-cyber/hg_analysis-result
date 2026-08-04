from __future__ import annotations

from copy import deepcopy
from pathlib import Path
import re
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

from honyu_app.domain.errors import ExcelExportError
from honyu_app.domain.models import ExcelCellWrite


MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"m": MAIN, "r": REL, "p": PKG_REL}
ET.register_namespace("", MAIN)
ET.register_namespace("r", REL)


def _column_number(address: str) -> int:
    match = re.fullmatch(r"([A-Z]+)([1-9]\d*)", address.upper())
    if not match:
        raise ExcelExportError(f"잘못된 Excel 셀 주소입니다: {address}")
    value = 0
    for char in match.group(1):
        value = value * 26 + ord(char) - 64
    return value


def _row_number(address: str) -> int:
    return int(re.search(r"\d+$", address).group())


class XlsxXmlCellWriter:
    """Writes numeric inputs without asking a spreadsheet library to rebuild XLSX."""

    def write_copy(
        self,
        template_path: Path,
        output_path: Path,
        writes: list[ExcelCellWrite],
    ) -> None:
        template_path = Path(template_path)
        output_path = Path(output_path)
        if template_path.resolve() == output_path.resolve():
            raise ExcelExportError("원본 Excel 파일에는 직접 쓸 수 없습니다.")
        if not template_path.is_file():
            raise ExcelExportError(f"원본 Excel 파일이 없습니다: {template_path}")
        if output_path.exists():
            raise ExcelExportError(f"출력 파일이 이미 있습니다: {output_path}")

        grouped: dict[str, list[ExcelCellWrite]] = {}
        seen: set[tuple[str, str]] = set()
        for item in writes:
            address = item.address.upper()
            key = (item.sheet, address)
            if key in seen:
                raise ExcelExportError(f"중복 입력 셀입니다: {item.sheet}!{address}")
            seen.add(key)
            grouped.setdefault(item.sheet, []).append(
                ExcelCellWrite(item.sheet, address, int(item.value))
            )

        try:
            with ZipFile(template_path, "r") as source:
                parts = {info.filename: source.read(info.filename) for info in source.infolist()}
                infos = {info.filename: info for info in source.infolist()}
            sheet_parts = self._sheet_parts(parts)
            for sheet, sheet_writes in grouped.items():
                part = sheet_parts.get(sheet)
                if part is None:
                    raise ExcelExportError(f"Excel 시트가 없습니다: {sheet}")
                parts[part] = self._write_sheet(parts[part], sheet, sheet_writes)
            parts["xl/workbook.xml"] = self._set_calculation_mode(parts["xl/workbook.xml"])

            with ZipFile(output_path, "x", compression=ZIP_DEFLATED) as target:
                for name, data in parts.items():
                    info = deepcopy(infos[name])
                    target.writestr(info, data)
        except ExcelExportError:
            if output_path.exists():
                output_path.unlink()
            raise
        except Exception as exc:
            if output_path.exists():
                output_path.unlink()
            raise ExcelExportError(f"Excel 숫자 입력에 실패했습니다: {exc}") from exc

    @staticmethod
    def _sheet_parts(parts: dict[str, bytes]) -> dict[str, str]:
        workbook = ET.fromstring(parts["xl/workbook.xml"])
        rels = ET.fromstring(parts["xl/_rels/workbook.xml.rels"])
        targets = {
            rel.attrib["Id"]: rel.attrib["Target"]
            for rel in rels.findall("p:Relationship", NS)
        }
        result: dict[str, str] = {}
        for sheet in workbook.findall("m:sheets/m:sheet", NS):
            target = targets[sheet.attrib[f"{{{REL}}}id"]].replace("\\", "/")
            if target.startswith("/"):
                part = target.lstrip("/")
            elif target.startswith("xl/"):
                part = target
            else:
                part = f"xl/{target}"
            result[sheet.attrib["name"]] = part
        return result

    def _write_sheet(
        self, xml: bytes, sheet_name: str, writes: list[ExcelCellWrite]
    ) -> bytes:
        root = ET.fromstring(xml)
        sheet_data = root.find("m:sheetData", NS)
        if sheet_data is None:
            raise ExcelExportError(f"sheetData가 없습니다: {sheet_name}")
        rows = {int(row.attrib["r"]): row for row in sheet_data.findall("m:row", NS)}

        for item in writes:
            row_no = _row_number(item.address)
            row = rows.get(row_no)
            if row is None:
                raise ExcelExportError(f"입력 대상 행이 없습니다: {sheet_name}!{item.address}")
            cells = {cell.attrib["r"].upper(): cell for cell in row.findall("m:c", NS)}
            cell = cells.get(item.address)
            if cell is None:
                cell = ET.Element(f"{{{MAIN}}}c", {"r": item.address})
                style = self._style_for_new_cell(row, item.address)
                if style is not None:
                    cell.set("s", style)
                position = sum(
                    _column_number(existing.attrib["r"]) < _column_number(item.address)
                    for existing in row.findall("m:c", NS)
                )
                row.insert(position, cell)
            if cell.find("m:f", NS) is not None:
                raise ExcelExportError(f"수식 셀에는 입력할 수 없습니다: {sheet_name}!{item.address}")
            cell.attrib.pop("t", None)
            for child in list(cell):
                if child.tag in {f"{{{MAIN}}}v", f"{{{MAIN}}}is"}:
                    cell.remove(child)
            value = ET.SubElement(cell, f"{{{MAIN}}}v")
            value.text = str(item.value)
        return ET.tostring(root, encoding="utf-8", xml_declaration=True)

    @staticmethod
    def _style_for_new_cell(row: ET.Element, address: str) -> str | None:
        cells = {cell.attrib["r"].upper(): cell for cell in row.findall("m:c", NS)}
        column = re.match(r"[A-Z]+", address).group()
        paired = {"U": "V", "V": "U", "Z": "AA", "AA": "Z"}.get(column)
        candidates = []
        if paired:
            candidates.append(f"{paired}{_row_number(address)}")
        candidates.extend((f"F{_row_number(address)}", f"R{_row_number(address)}"))
        for candidate in candidates:
            cell = cells.get(candidate)
            if cell is not None and "s" in cell.attrib:
                return cell.attrib["s"]
        nearest = min(
            cells.values(),
            key=lambda cell: abs(_column_number(cell.attrib["r"]) - _column_number(address)),
            default=None,
        )
        return nearest.attrib.get("s") if nearest is not None else None

    @staticmethod
    def _set_calculation_mode(xml: bytes) -> bytes:
        root = ET.fromstring(xml)
        calc = root.find("m:calcPr", NS)
        if calc is None:
            calc = ET.SubElement(root, f"{{{MAIN}}}calcPr")
        calc.set("calcMode", "auto")
        calc.set("fullCalcOnLoad", "1")
        calc.set("forceFullCalc", "1")
        return ET.tostring(root, encoding="utf-8", xml_declaration=True)
