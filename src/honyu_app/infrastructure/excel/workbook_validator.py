from __future__ import annotations

from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from honyu_app.domain.models import ExcelCellWrite, WorkbookValidationResult
from honyu_app.infrastructure.excel.xml_cell_writer import MAIN, NS, XlsxXmlCellWriter


def _canonical(element: ET.Element | None) -> object:
    if element is None:
        return None
    return (
        element.tag,
        tuple(sorted(element.attrib.items())),
        (element.text or "").strip(),
        tuple(_canonical(child) for child in element),
    )


def _cell_signature(cell: ET.Element, *, ignore_formula_cache: bool = False) -> object:
    formula = cell.find("m:f", NS)
    value = cell.find("m:v", NS)
    inline = cell.find("m:is", NS)
    return (
        cell.attrib.get("s"),
        cell.attrib.get("t"),
        _canonical(formula),
        None if ignore_formula_cache and formula is not None else (value.text if value is not None else None),
        _canonical(inline),
    )


class XlsxWorkbookValidator:
    """Checks that an export changed only approved input values and formula caches."""

    SHEET_STRUCTURES = (
        "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols",
        "sheetProtection", "protectedRanges", "scenarios", "autoFilter",
        "sortState", "dataConsolidate", "customSheetViews", "mergeCells",
        "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks",
        "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks",
        "colBreaks", "customProperties", "cellWatches", "ignoredErrors",
        "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "picture",
        "oleObjects", "controls", "webPublishItems", "tableParts",
    )
    PRESERVED_PREFIXES = (
        "xl/charts/", "xl/drawings/", "xl/media/", "xl/externalLinks/",
        "xl/ctrlProps/", "xl/embeddings/", "xl/printerSettings/",
    )

    def validate(
        self,
        original_path: Path,
        result_path: Path,
        writes: list[ExcelCellWrite],
        *,
        after_excel_recalculation: bool,
    ) -> WorkbookValidationResult:
        errors: list[str] = []
        try:
            original = self._parts(Path(original_path))
            result = self._parts(Path(result_path))
            original_sheets = XlsxXmlCellWriter._sheet_parts(original)
            result_sheets = XlsxXmlCellWriter._sheet_parts(result)
            if list(original_sheets) != list(result_sheets):
                errors.append("시트 이름 또는 순서가 변경되었습니다.")
            self._compare_workbook(original, result, errors)
            targets = {(item.sheet, item.address.upper()): item.value for item in writes}
            for sheet in original_sheets.keys() & result_sheets.keys():
                self._compare_sheet(
                    sheet,
                    original[original_sheets[sheet]],
                    result[result_sheets[sheet]],
                    targets,
                    after_excel_recalculation,
                    errors,
                )
            self._compare_preserved_parts(original, result, errors)
            self._check_targets(result, result_sheets, targets, errors)
        except Exception as exc:
            errors.append(f"통합문서 검증 중 오류가 발생했습니다: {exc}")
        return WorkbookValidationResult(not errors, tuple(errors))

    @staticmethod
    def _parts(path: Path) -> dict[str, bytes]:
        with ZipFile(path, "r") as archive:
            return {name: archive.read(name) for name in archive.namelist()}

    @staticmethod
    def _compare_workbook(
        original: dict[str, bytes], result: dict[str, bytes], errors: list[str]
    ) -> None:
        left = ET.fromstring(original["xl/workbook.xml"])
        right = ET.fromstring(result["xl/workbook.xml"])
        for path, label in (
            ("m:sheets", "시트 정의"),
            ("m:definedNames", "정의된 이름/인쇄 영역"),
            ("m:workbookProtection", "통합문서 보호"),
        ):
            if _canonical(left.find(path, NS)) != _canonical(right.find(path, NS)):
                errors.append(f"{label}이 변경되었습니다.")
        if "xl/styles.xml" not in result:
            errors.append("스타일 정의가 사라졌습니다.")
        elif _canonical(ET.fromstring(original["xl/styles.xml"])) != _canonical(
            ET.fromstring(result["xl/styles.xml"])
        ):
            errors.append("스타일 정의가 변경되었습니다.")

    def _compare_sheet(
        self,
        sheet: str,
        original_xml: bytes,
        result_xml: bytes,
        targets: dict[tuple[str, str], int],
        after_recalc: bool,
        errors: list[str],
    ) -> None:
        left = ET.fromstring(original_xml)
        right = ET.fromstring(result_xml)
        for tag in self.SHEET_STRUCTURES:
            left_items = left.findall(f"m:{tag}", NS)
            right_items = right.findall(f"m:{tag}", NS)
            if tuple(_canonical(item) for item in left_items) != tuple(
                _canonical(item) for item in right_items
            ):
                errors.append(f"{sheet}: {tag} 구조가 변경되었습니다.")

        left_rows = {
            row.attrib["r"]: row for row in left.findall("m:sheetData/m:row", NS)
        }
        right_rows = {
            row.attrib["r"]: row for row in right.findall("m:sheetData/m:row", NS)
        }
        if set(left_rows) != set(right_rows):
            errors.append(f"{sheet}: 행 구성이 변경되었습니다.")
        for row_no in left_rows.keys() & right_rows.keys():
            if left_rows[row_no].attrib != right_rows[row_no].attrib:
                errors.append(f"{sheet}!{row_no}행의 높이/숨김/서식이 변경되었습니다.")
            left_cells = {cell.attrib["r"]: cell for cell in left_rows[row_no].findall("m:c", NS)}
            right_cells = {cell.attrib["r"]: cell for cell in right_rows[row_no].findall("m:c", NS)}
            approved = {address for target_sheet, address in targets if target_sheet == sheet}
            unexpected_missing = set(left_cells) - set(right_cells)
            unexpected_new = set(right_cells) - set(left_cells) - approved
            if unexpected_missing or unexpected_new:
                errors.append(
                    f"{sheet}!{row_no}행의 셀 구성이 변경되었습니다: "
                    f"삭제 {sorted(unexpected_missing)}, 추가 {sorted(unexpected_new)}"
                )
            for address in left_cells.keys() & right_cells.keys():
                if address in approved:
                    if left_cells[address].attrib.get("s") != right_cells[address].attrib.get("s"):
                        errors.append(f"{sheet}!{address} 입력 셀의 서식이 변경되었습니다.")
                    continue
                left_signature = _cell_signature(
                    left_cells[address], ignore_formula_cache=after_recalc
                )
                right_signature = _cell_signature(
                    right_cells[address], ignore_formula_cache=after_recalc
                )
                if left_signature != right_signature:
                    errors.append(f"{sheet}!{address} 값·수식·서식이 변경되었습니다.")

    def _compare_preserved_parts(
        self, original: dict[str, bytes], result: dict[str, bytes], errors: list[str]
    ) -> None:
        left_names = {
            name for name in original if name.startswith(self.PRESERVED_PREFIXES)
        }
        right_names = {
            name for name in result if name.startswith(self.PRESERVED_PREFIXES)
        }
        if left_names != right_names:
            errors.append("차트·그림·외부 연결 부품 구성이 변경되었습니다.")
            return
        for name in sorted(left_names):
            if name.endswith(".xml") or name.endswith(".rels"):
                if _canonical(ET.fromstring(original[name])) != _canonical(
                    ET.fromstring(result[name])
                ):
                    errors.append(f"보존 대상 Excel 부품이 변경되었습니다: {name}")
            elif original[name] != result[name]:
                errors.append(f"보존 대상 바이너리 부품이 변경되었습니다: {name}")

    @staticmethod
    def _check_targets(
        result: dict[str, bytes],
        sheet_parts: dict[str, str],
        targets: dict[tuple[str, str], int],
        errors: list[str],
    ) -> None:
        parsed: dict[str, ET.Element] = {}
        for (sheet, address), expected in targets.items():
            if sheet not in sheet_parts:
                errors.append(f"입력 대상 시트가 없습니다: {sheet}")
                continue
            root = parsed.setdefault(sheet, ET.fromstring(result[sheet_parts[sheet]]))
            cell = root.find(f".//m:c[@r='{address}']", NS)
            if cell is None:
                errors.append(f"입력 대상 셀이 없습니다: {sheet}!{address}")
                continue
            if cell.find("m:f", NS) is not None:
                errors.append(f"입력 대상이 수식 셀이 되었습니다: {sheet}!{address}")
            value = cell.find("m:v", NS)
            if cell.attrib.get("t") is not None or value is None or value.text != str(expected):
                errors.append(f"입력값이 일치하지 않습니다: {sheet}!{address}")
