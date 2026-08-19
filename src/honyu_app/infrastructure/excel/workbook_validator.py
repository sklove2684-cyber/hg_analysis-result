from __future__ import annotations

from collections import Counter
from pathlib import Path
import re
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


def _chart_structure(xml: bytes) -> object:
    root = ET.fromstring(xml)
    cache_names = {"numCache", "strCache", "multiLvlStrCache"}
    for parent in root.iter():
        for child in list(parent):
            if child.tag.rsplit("}", 1)[-1] in cache_names:
                parent.remove(child)
    return _canonical(root)


def _cell_signature(
    cell: ET.Element,
    styles: tuple[object, ...],
    *,
    ignore_formula_cache: bool = False,
) -> object:
    formula = cell.find("m:f", NS)
    value = cell.find("m:v", NS)
    inline = cell.find("m:is", NS)
    style_id = int(cell.attrib.get("s", "0"))
    style = styles[style_id] if style_id < len(styles) else ("missing-style", style_id)
    return (
        style,
        None if ignore_formula_cache and formula is not None else cell.attrib.get("t"),
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
            original_styles = self._style_signatures(original)
            result_styles = self._style_signatures(result)
            if list(original_sheets) != list(result_sheets):
                errors.append("시트 이름 또는 순서가 변경되었습니다.")
            self._compare_workbook(original, result, errors)
            self._compare_styles(original, result, errors)
            targets = {(item.sheet, item.address.upper()): item.value for item in writes}
            for sheet in original_sheets.keys() & result_sheets.keys():
                self._compare_sheet(
                    sheet,
                    original[original_sheets[sheet]],
                    result[result_sheets[sheet]],
                    targets,
                    original_styles,
                    result_styles,
                    after_excel_recalculation,
                    errors,
                )
            self._compare_preserved_parts(original, result, errors)
            self._check_targets(result, result_sheets, targets, errors)
            self._check_office_namespaces(result, result_sheets, errors)
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

    @staticmethod
    def _style_signatures(parts: dict[str, bytes]) -> tuple[object, ...]:
        root = ET.fromstring(parts["xl/styles.xml"])
        cell_xfs = root.find("m:cellXfs", NS)
        return tuple(_canonical(child) for child in cell_xfs) if cell_xfs is not None else ()

    @staticmethod
    def _compare_styles(
        original: dict[str, bytes], result: dict[str, bytes], errors: list[str]
    ) -> None:
        if "xl/styles.xml" not in result:
            errors.append("스타일 정의가 사라졌습니다.")
            return
        left = ET.fromstring(original["xl/styles.xml"])
        right = ET.fromstring(result["xl/styles.xml"])
        for tag in (
            "numFmts", "fonts", "fills", "borders", "cellStyleXfs",
            "cellXfs", "cellStyles", "dxfs", "tableStyles",
        ):
            left_parent = left.find(f"m:{tag}", NS)
            right_parent = right.find(f"m:{tag}", NS)
            left_items = Counter(
                _canonical(child) for child in (left_parent if left_parent is not None else ())
            )
            right_items = Counter(
                _canonical(child) for child in (right_parent if right_parent is not None else ())
            )
            if left_items != right_items:
                errors.append(f"스타일 구성 요소가 변경되었습니다: {tag}")

    def _compare_sheet(
        self,
        sheet: str,
        original_xml: bytes,
        result_xml: bytes,
        targets: dict[tuple[str, str], int],
        original_styles: tuple[object, ...],
        result_styles: tuple[object, ...],
        after_recalc: bool,
        errors: list[str],
    ) -> None:
        left = ET.fromstring(original_xml)
        right = ET.fromstring(result_xml)
        for tag in self.SHEET_STRUCTURES:
            left_items = left.findall(f"m:{tag}", NS)
            right_items = right.findall(f"m:{tag}", NS)
            if tag == "mergeCells":
                left_merges = Counter(
                    item.attrib.get("ref")
                    for parent in left_items
                    for item in parent.findall("m:mergeCell", NS)
                )
                right_merges = Counter(
                    item.attrib.get("ref")
                    for parent in right_items
                    for item in parent.findall("m:mergeCell", NS)
                )
                if left_merges != right_merges:
                    errors.append(f"{sheet}: 병합 셀 구성이 변경되었습니다.")
                continue
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
                    left_style = _cell_signature(
                        left_cells[address], original_styles, ignore_formula_cache=True
                    )[0]
                    right_style = _cell_signature(
                        right_cells[address], result_styles, ignore_formula_cache=True
                    )[0]
                    if left_style != right_style:
                        errors.append(f"{sheet}!{address} 입력 셀의 서식이 변경되었습니다.")
                    continue
                left_signature = _cell_signature(
                    left_cells[address],
                    original_styles,
                    ignore_formula_cache=after_recalc,
                )
                right_signature = _cell_signature(
                    right_cells[address],
                    result_styles,
                    ignore_formula_cache=after_recalc,
                )
                if left_signature != right_signature:
                    errors.append(f"{sheet}!{address} 값·수식·서식이 변경되었습니다.")
                if after_recalc:
                    left_error = self._formula_error(left_cells[address])
                    right_error = self._formula_error(right_cells[address])
                    if right_error is not None and right_error != left_error:
                        errors.append(
                            f"{sheet}!{address} 수식에 새 오류가 발생했습니다: {right_error}"
                        )

    @staticmethod
    def _formula_error(cell: ET.Element) -> str | None:
        if cell.find("m:f", NS) is None or cell.attrib.get("t") != "e":
            return None
        value = cell.find("m:v", NS)
        return value.text if value is not None else "UNKNOWN"

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
            if name.startswith("xl/printerSettings/"):
                continue
            if name.endswith(".xml") or name.endswith(".rels"):
                left = (
                    _chart_structure(original[name])
                    if name.startswith("xl/charts/") and name.endswith(".xml")
                    else _canonical(ET.fromstring(original[name]))
                )
                right = (
                    _chart_structure(result[name])
                    if name.startswith("xl/charts/") and name.endswith(".xml")
                    else _canonical(ET.fromstring(result[name]))
                )
                if left != right:
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

    @staticmethod
    def _check_office_namespaces(
        parts: dict[str, bytes],
        sheet_parts: dict[str, str],
        errors: list[str],
    ) -> None:
        names = ["xl/workbook.xml", *sheet_parts.values()]
        for name in names:
            xml = parts[name]
            root_match = re.search(rb"<(?!\?)[^>]+>", xml)
            if root_match is None:
                errors.append(f"Excel XML 루트 요소가 손상되었습니다: {name}")
                continue
            opening = root_match.group(0)
            declared = {
                prefix.decode("ascii")
                for prefix in re.findall(rb'xmlns:([A-Za-z_][\w.-]*)="[^"]+"', opening)
            }
            references: set[str] = set()
            for value in re.findall(rb'(?:Ignorable|Requires)="([^"]+)"', xml):
                references.update(value.decode("ascii").split())
            missing = references - declared
            if missing:
                errors.append(
                    f"Microsoft Excel 호환 네임스페이스가 누락되었습니다: "
                    f"{name} ({', '.join(sorted(missing))})"
                )
