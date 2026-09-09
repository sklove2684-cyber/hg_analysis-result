from __future__ import annotations

from collections import Counter
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from collections.abc import Callable
import hashlib
import re

import pdfplumber

from honyu_app.domain.enums import ReviewStatus
from honyu_app.domain.errors import ExtractionCancelledError, ValidationError
from honyu_app.domain.models import AnalysisBatch, HeavyMetalRecoveryValue, SourceFile


ELEMENTS = ("Fe", "Mn", "Al", "Cr", "Sn", "Ti", "Cu", "Zr", "Zn", "Ni")
RECOVERY_NAMES = {"저": "low", "중": "mid", "고": "high"}


class HeavyMetalResultsParser:
    name = "heavy-metal-pdfplumber"
    version = "1.0.0"
    layout_id = "list-of-results-quant-average-10-v1"

    def parse(self, pdf_path: Path, *, analysis_type: str, analysis_no_start: int,
              analysis_no_end: int, progress_callback: Callable[[int, int], None] | None = None,
              cancel_check: Callable[[], bool] | None = None) -> AnalysisBatch:
        path = Path(pdf_path)
        if not path.is_file():
            raise ValidationError(f"PDF 파일을 찾을 수 없습니다: {path}")
        raw = path.read_bytes()
        values: list[HeavyMetalRecoveryValue] = []
        blank_no = 0
        with pdfplumber.open(path) as pdf:
            page_count = len(pdf.pages)
            for page_no, page in enumerate(pdf.pages, 1):
                if cancel_check and cancel_check():
                    raise ExtractionCancelledError("PDF 추출이 취소되었습니다.")
                for table in page.extract_tables():
                    if not table:
                        continue
                    header = next(
                        (row for row in table if row and any("Quant" in (c or "") for c in row)),
                        None,
                    )
                    if header is None:
                        continue
                    columns = []
                    for element in ELEMENTS:
                        found = [i for i, cell in enumerate(header) if element in (cell or "") and "Quant" in (cell or "")]
                        if len(found) != 1:
                            raise ValidationError("PDF_LAYOUT_MISMATCH: Quant Average 10개 원소 헤더를 확인할 수 없습니다.")
                        columns.append(found[0])
                    for source_row, row in enumerate(table, 1):
                        if len(row) <= max(columns):
                            continue
                        sample_name = (row[3] or "").strip() if len(row) > 3 else ""
                        if sample_name == "회수율-B":
                            blank_no += 1
                            level, replicate = "blank", blank_no
                        else:
                            match = re.fullmatch(r"([저중고])([123])", sample_name)
                            if not match:
                                continue
                            level, replicate = RECOVERY_NAMES[match.group(1)], int(match.group(2))
                        if replicate not in {1, 2, 3}:
                            raise ValidationError("PDF_LAYOUT_MISMATCH: 회수율 반복 번호가 1~3 범위를 벗어났습니다.")
                        for element, col in zip(ELEMENTS, columns):
                            value, below_limit = self._quant_value(row[col], page_no, source_row)
                            values.append(HeavyMetalRecoveryValue(
                                element, sample_name, level, replicate, value,
                                below_limit, page_no, source_row,
                            ))
                if progress_callback:
                    progress_callback(page_no, page_count)
        keys = [(v.element, v.level, v.replicate_no) for v in values]
        duplicates = [key for key, count in Counter(keys).items() if count != 1]
        if len(values) != 120 or duplicates:
            raise ValidationError(
                f"PDF_LAYOUT_MISMATCH: 회수율 값은 정확히 120개여야 합니다 "
                f"(추출 {len(values)}개, 중복 {len(duplicates)}개)."
            )
        now = datetime.now().astimezone()
        return AnalysisBatch(
            batch_code=f"{analysis_type}-{analysis_no_start}-{analysis_no_end}-{now:%Y%m%d%H%M%S}",
            source_file=SourceFile(path.name, path.resolve(), hashlib.sha256(raw).hexdigest(),
                                   len(raw), page_count),
            analysis_type=analysis_type, analysis_no_start=analysis_no_start,
            analysis_no_end=analysis_no_end, parser_name=self.name,
            parser_version=self.version, parser_layout_id=self.layout_id,
            extracted_at=now, heavy_metal_recovery_values=values,
            warning_count=0, review_status=ReviewStatus.PENDING,
        )

    @staticmethod
    def _quant_value(raw: str | None, page_no: int, source_row: int) -> tuple[Decimal, bool]:
        text = (raw or "").strip()
        below_limit = bool(re.search(r"L\s*$", text, re.I))
        numeric = re.sub(r"[mMgG/L\s]", "", text)
        if numeric.count(".") > 1 and numeric.startswith(("-0.0.", "+0.0.", "0.0.")):
            numeric = numeric.replace("0.0.", "0.", 1)
        try:
            return Decimal(numeric), below_limit
        except InvalidOperation as exc:
            raise ValidationError(
                f"PDF_LAYOUT_MISMATCH: {page_no}페이지 {source_row}행 Quant Average가 숫자가 아닙니다: {text}"
            ) from exc
