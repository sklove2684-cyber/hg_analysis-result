from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from decimal import Decimal
import logging
import re

from honyu_app.domain.enums import SampleType, ValidationSeverity
from honyu_app.domain.models import ExcelPreviewIssue, ExcelPreviewResult, ExcelPreviewRow
from honyu_app.infrastructure.pdf.heavy_metal_layout import ELEMENT_SYMBOLS


def position(address):
    match = re.fullmatch(r"([A-Z]+)([0-9]+)", address)
    column = 0
    for char in match[1]:
        column = column * 26 + ord(char) - 64
    return column, int(match[2])


def column_name(column):
    name = ""
    while column:
        column, remainder = divmod(column - 1, 26)
        name = chr(65 + remainder) + name
    return name


def normalized(value):
    return re.sub(r"\s+", "", str(value if value is not None else ""))


@dataclass(frozen=True)
class RecoveryBlock:
    sheet: str
    element: str
    label: str
    base_row: int
    blank_column: int
    recovery_column: int
    group_no: int | None

    def target(self, level, replicate):
        column = self.blank_column if level == "blank" else self.recovery_column
        offset = {"blank": 0, "low": 0, "mid": 3, "high": 6}[level]
        return f"{column_name(column)}{self.base_row + offset + replicate - 1}"


def find_recovery_blocks(snapshot):
    """Find only input headers and complete low/mid/high triplets on any sheet."""
    blocks = []
    for sheet in snapshot.sheet_names:
        cells = {position(address): cell for (cell_sheet, address), cell
                 in snapshot.cells.items() if cell_sheet == sheet}
        blanks = [(col, row) for (col, row), cell in cells.items()
                  if normalized(cell.value).casefold() == "blank"]
        detections = [(col, row) for (col, row), cell in cells.items()
                      if normalized(cell.value) in {"검출량", "검출농도"}]
        headers = [(bc, dc, br) for bc, br in blanks for dc, dr in detections
                   if br == dr and bc < dc]
        for (label_col, base_row), cell in cells.items():
            label = normalized(cell.value)
            if label in {"NaOH", "KOH"}:
                element = {"NaOH": "Na", "KOH": "K"}[label]
            else:
                match = re.fullmatch(r"([A-Z][a-z]?)(?:-([0-9]+))?", label)
                if not match or match[1] not in ELEMENT_SYMBOLS:
                    continue
                element = match[1]
            candidates = []
            for bc, dc, header_row in headers:
                if not (label_col < bc and header_row < base_row):
                    continue
                for level_col in range(label_col + 1, bc):
                    if all(normalized(cells.get((level_col, base_row + offset), None).value
                                      if cells.get((level_col, base_row + offset)) else None) == level
                           for offset, level in ((0, "저"), (3, "중"), (6, "고"))):
                        candidates.append((bc - label_col, -header_row, bc, dc, header_row))
            if not candidates:
                continue
            # Pick the closest columns, then the nearest preceding header row.
            _, _, bc, dc, header_row = min(candidates)
            group_no = None
            # Region IDs precede the element label or sit beside its header.
            # The suffix of Fe-10 etc. is an element label, never a row/group offset.
            for row in (base_row, header_row):
                for col in range(label_col - 1, 0, -1):
                    candidate = cells.get((col, row))
                    value = normalized(candidate.value) if candidate else ""
                    if re.fullmatch(r"[1-9][0-9]*", value):
                        group_no = int(value)
                        break
                if group_no is not None:
                    break
            blocks.append(RecoveryBlock(sheet, element, f"{column_name(label_col)}{base_row}",
                                        base_row, bc, dc, group_no))
    return blocks


def select_block(snapshot, blocks, element, group_no):
    candidates = [block for block in blocks if block.element == element]
    if group_no is not None and any(b.group_no is not None for b in candidates):
        candidates = [b for b in candidates if b.group_no == group_no]
    if len(candidates) > 1:
        active = [b for b in candidates if b.sheet == snapshot.active_sheet]
        if len(active) == 1:
            candidates = active
    return candidates


def preview_recovery(batch, template_path, snapshot):
    result = ExcelPreviewResult(template_path, "N/A")
    blocks = find_recovery_blocks(snapshot)
    values = batch.heavy_metal_recovery_values
    if not values:
        result.issues.append(ExcelPreviewIssue(
            ValidationSeverity.ERROR, "HEAVY_METAL_WRITE_COUNT_MISMATCH",
            "PDF에서 중금속 회수율 값을 확인할 수 없습니다.",
        ))
        return result
    groups = list(dict.fromkeys((v.element, v.group_no) for v in values))
    selected = {}
    for element, group_no in groups:
        matches = select_block(snapshot, blocks, element, group_no)
        if len(matches) != 1:
            code = ("HEAVY_METAL_ELEMENT_NOT_IN_TEMPLATE" if not matches
                    else "HEAVY_METAL_RECOVERY_BLOCK_AMBIGUOUS")
            result.issues.append(ExcelPreviewIssue(
                ValidationSeverity.ERROR, code,
                f"PDF 원소 {element} (G#={group_no})의 회수율 블록을 유일하게 찾을 수 없습니다.",
            ))
        else:
            selected[(element, group_no)] = matches[0]
    expected = {(element, group, level, replicate)
                for element, group in groups for level in ("blank", "low", "mid", "high")
                for replicate in (1, 2, 3)}
    counts = Counter((v.element, v.group_no, v.level, v.replicate_no) for v in values)
    if set(counts) != expected or any(count != 1 for count in counts.values()):
        result.issues.append(ExcelPreviewIssue(
            ValidationSeverity.ERROR, "HEAVY_METAL_WRITE_COUNT_MISMATCH",
            "중금속 회수율은 원소/구역별 BLANK 3개와 저/중/고 각 3개여야 합니다.",
        ))
    if result.issues:
        return result
    excel_only = sorted({b.element for b in blocks} - {v.element for v in values})
    if excel_only:
        result.issues.append(ExcelPreviewIssue(
            ValidationSeverity.WARNING, "HEAVY_METAL_EXCEL_ONLY_ELEMENTS",
            "PDF에 없어 입력하지 않는 Excel 원소: " + ", ".join(excel_only),
        ))
    used = set()
    for value in values:
        block = selected[(value.element, value.group_no)]
        address = block.target(value.level, value.replicate_no)
        cell = snapshot.cell(block.sheet, address)
        if cell.has_formula:
            result.issues.append(ExcelPreviewIssue(
                ValidationSeverity.ERROR, "TARGET_CELL_HAS_FORMULA",
                f"수식 셀에는 입력할 수 없습니다: {block.sheet}!{address}",
                value.sample_name, block.sheet, address,
            ))
            continue
        col, row = position(address)
        merged = False
        for sheet, cell_range in snapshot.merged_ranges:
            if sheet != block.sheet:
                continue
            first, last = cell_range.split(":") if ":" in cell_range else (cell_range, cell_range)
            left, top = position(first)
            right, bottom = position(last)
            if (left, top) != (right, bottom) and left <= col <= right and top <= row <= bottom:
                merged = True
                break
        if merged or (block.sheet, address) in used:
            result.issues.append(ExcelPreviewIssue(
                ValidationSeverity.ERROR, "HEAVY_METAL_TARGET_NOT_UNIQUE",
                f"병합 또는 중복 입력 셀입니다: {block.sheet}!{address}",
                value.sample_name, block.sheet, address,
            ))
            continue
        used.add((block.sheet, address))
        result.rows.append(ExcelPreviewRow(
            sample_name=value.sample_name,
            sample_type=SampleType.RECOVERY_BLANK if value.level == "blank" else SampleType.RECOVERY,
            material=value.element, peak_no=value.replicate_no,
            retention_time=Decimal("0"), area_raw=value.value, applied_area=value.value,
            target_sheet=block.sheet, target_cell=address,
            existing_value_type=cell.value_type, existing_has_formula=cell.has_formula,
            message=f"G#={value.group_no}" + (" L" if value.below_limit else ""),
        ))
    for key, block in selected.items():
        logging.getLogger(__name__).info("Recovery block %s: %s", key, block)
    return result
