from __future__ import annotations

from collections import defaultdict
from pathlib import Path
import re
from uuid import UUID

from honyu_app.domain.enums import (
    ConcentrationLevel,
    ExcelPreviewStatus,
    ExcludeReason,
    SampleType,
    StdMethod,
    ValidationSeverity,
)
from honyu_app.domain.errors import ValidationError
from honyu_app.domain.models import (
    AnalysisBatch,
    ExcelPreviewIssue,
    ExcelPreviewResult,
    ExcelPreviewRow,
    Peak,
    Sample,
)
from honyu_app.services.database_service import DatabaseService
from honyu_app.services.excel_template_service import (
    ExcelTemplateService,
    ExcelTemplateSnapshot,
)


MATERIAL_COLUMNS = {
    "n-hexane": "F",
    "acetone": "G",
    "E.A": "H",
    "MIBK": "I",
    "Toluene": "J",
    "B.A": "K",
    "E.B": "L",
    "p-xylene": "M",
    "m-xylene": "N",
    "o-xylene": "O",
    "styrene": "P",
    "c-hexanone": "Q",
}
RECOVERY_COLUMNS = {
    material: chr(ord("B") + index)
    for index, material in enumerate(MATERIAL_COLUMNS)
}
STD_REPLICATES = {
    StdMethod.A: (1, 2, 3, 4, 5),
    StdMethod.B: (1, 2, 3, 4, 6),
}
RECOVERY_ROW_START = {
    ConcentrationLevel.LOW: 37,
    ConcentrationLevel.MID: 40,
    ConcentrationLevel.HIGH: 43,
}
REQUIRED_SHEETS = ("검량선", "area", "최종결과", "회수율", "STD제조")


class PreviewExcelExportService:
    def __init__(
        self,
        database: DatabaseService,
        template_service: ExcelTemplateService,
    ) -> None:
        self._database = database
        self._template_service = template_service

    def preview(
        self, batch_id: UUID, template_path: Path, std_method: StdMethod | str
    ) -> ExcelPreviewResult:
        batch = self._database.get_batch_detail(batch_id)
        return self.preview_batch(batch, template_path, std_method)

    def preview_batch(
        self,
        batch: AnalysisBatch,
        template_path: Path,
        std_method: StdMethod | str,
    ) -> ExcelPreviewResult:
        try:
            method = StdMethod(std_method)
        except ValueError as exc:
            raise ValidationError("STD 방식은 A 또는 B여야 합니다.") from exc
        snapshot = self._template_service.inspect(Path(template_path))
        result = ExcelPreviewResult(Path(template_path), method.value)
        self._validate_template(snapshot, result)
        worker_rows = self._worker_row_index(snapshot)

        for sample in batch.samples:
            self._map_sample(sample, method, snapshot, worker_rows, result)
        self._detect_target_collisions(result)
        return result

    def _validate_template(
        self, snapshot: ExcelTemplateSnapshot, result: ExcelPreviewResult
    ) -> None:
        missing = [name for name in REQUIRED_SHEETS if name not in snapshot.sheet_names]
        if missing:
            result.issues.append(
                ExcelPreviewIssue(
                    ValidationSeverity.ERROR,
                    "TEMPLATE_SHEET_MISSING",
                    f"필수 시트가 없습니다: {', '.join(missing)}",
                )
            )

    @staticmethod
    def _worker_key(value: object | None) -> str | None:
        if value is None or isinstance(value, bool):
            return None
        if isinstance(value, float) and value.is_integer():
            value = int(value)
        match = re.search(r"(\d+)\s*$", str(value).strip())
        return match.group(1) if match else None

    def _worker_row_index(
        self, snapshot: ExcelTemplateSnapshot
    ) -> dict[str, list[int]]:
        index: dict[str, list[int]] = defaultdict(list)
        for row in range(37, 184):
            key = self._worker_key(snapshot.cell("area", f"A{row}").value)
            if key is not None:
                index[key].append(row)
        return dict(index)

    def _map_sample(
        self,
        sample: Sample,
        method: StdMethod,
        snapshot: ExcelTemplateSnapshot,
        worker_rows: dict[str, list[int]],
        result: ExcelPreviewResult,
    ) -> None:
        row = self._sample_target_row(sample, method, worker_rows, result)
        if row is None:
            reason = self._sample_exclusion_reason(sample, method)
            for peak in sample.peaks:
                if reason is not None:
                    result.rows.append(
                        self._row_for_peak(
                            sample,
                            peak,
                            status=ExcelPreviewStatus.EXCLUDED,
                            exclude_reason=reason,
                        )
                    )
                else:
                    result.rows.append(
                        self._row_for_peak(
                            sample,
                            peak,
                            status=ExcelPreviewStatus.ERROR,
                            message="Sample의 Excel 목표 행을 결정할 수 없습니다.",
                        )
                    )
            return

        eligible = [
            peak
            for peak in sample.peaks
            if peak.include_for_excel and peak.material_standard not in {None, "CS2"}
        ]
        for peak in sample.peaks:
            if peak not in eligible:
                result.rows.append(
                    self._row_for_peak(
                        sample,
                        peak,
                        status=ExcelPreviewStatus.EXCLUDED,
                        exclude_reason=(
                            peak.exclude_reason.value
                            if peak.exclude_reason
                            else "NOT_EXCEL_ELIGIBLE"
                        ),
                    )
                )

        dibk = [peak for peak in eligible if peak.material_standard == "DIBK"]
        single = [peak for peak in eligible if peak.material_standard != "DIBK"]
        for peak in single:
            column = self._material_column(sample, peak.material_standard)
            if column is None:
                self._append_error_row(
                    result,
                    sample,
                    peak,
                    "UNSUPPORTED_MATERIAL",
                    f"Excel 핵심 물질 열이 없습니다: {peak.material_standard}",
                )
                continue
            self._append_mapped_row(result, snapshot, sample, peak, column, row)

        ranked = sorted(
            dibk,
            key=lambda peak: (
                -self._applied_area(peak),
                peak.peak_no,
                peak.retention_time,
            ),
        )
        slots = self._dibk_slots(sample)
        for rank, peak in enumerate(ranked, start=1):
            if rank > 2:
                result.rows.append(
                    self._row_for_peak(
                        sample,
                        peak,
                        dibk_area_rank=rank,
                        status=ExcelPreviewStatus.EXCLUDED,
                        exclude_reason=ExcludeReason.DIBK_AREA_NOT_TOP2.value,
                        message="DIBK 적용 Area 상위 2개 밖의 피크",
                    )
                )
                continue
            self._append_mapped_row(
                result,
                snapshot,
                sample,
                peak,
                slots[rank - 1],
                row,
                dibk_area_rank=rank,
            )

    def _sample_target_row(
        self,
        sample: Sample,
        method: StdMethod,
        worker_rows: dict[str, list[int]],
        result: ExcelPreviewResult,
    ) -> int | None:
        if sample.sample_type is SampleType.STD:
            selected = STD_REPLICATES[method]
            if sample.replicate_no not in selected:
                return None
            return 15 + selected.index(sample.replicate_no)
        if sample.sample_type is SampleType.RECOVERY:
            if sample.concentration_level not in RECOVERY_ROW_START:
                self._sample_error(
                    result, sample, "RECOVERY_LEVEL_MISSING", "회수율 농도 구분이 없습니다."
                )
                return None
            if sample.replicate_no not in {1, 2, 3}:
                self._sample_error(
                    result, sample, "RECOVERY_REPLICATE_INVALID", "회수율 반복번호는 1~3이어야 합니다."
                )
                return None
            return RECOVERY_ROW_START[sample.concentration_level] + sample.replicate_no - 1
        if sample.sample_type is SampleType.NUMERIC:
            key = sample.worker_match_key or self._worker_key(sample.sample_name_normalized)
            matches = worker_rows.get(str(key), []) if key else []
            if len(matches) != 1:
                self._sample_error(
                    result,
                    sample,
                    "WORKER_ROW_NOT_UNIQUE",
                    f"작업자 키 {key or '(없음)'}와 일치하는 area 분석번호 행: {matches}",
                )
                return None
            return matches[0]
        return None

    @staticmethod
    def _sample_exclusion_reason(sample: Sample, method: StdMethod) -> str | None:
        if sample.sample_type is SampleType.STD:
            selected = STD_REPLICATES[method]
            if sample.replicate_no not in selected:
                return f"STD_METHOD_{method.value}_NOT_SELECTED"
        if sample.sample_type in {
            SampleType.BLANK,
            SampleType.RECOVERY_BLANK,
            SampleType.UNKNOWN,
        }:
            return f"SAMPLE_TYPE_{sample.sample_type.value}"
        return None

    @staticmethod
    def _material_column(sample: Sample, material: str | None) -> str | None:
        if sample.sample_type is SampleType.RECOVERY:
            return RECOVERY_COLUMNS.get(material or "")
        return MATERIAL_COLUMNS.get(material or "")

    @staticmethod
    def _dibk_slots(sample: Sample) -> tuple[str, str]:
        return ("U", "V") if sample.sample_type is SampleType.RECOVERY else ("Z", "AA")

    @staticmethod
    def _target_sheet(sample: Sample) -> str:
        return "회수율" if sample.sample_type is SampleType.RECOVERY else "area"

    def _applied_area(self, peak: Peak) -> int:
        corrections = self._database.list_peak_corrections(peak.peak_id)
        return corrections[-1].area_after if corrections else peak.area_raw

    def _row_for_peak(
        self,
        sample: Sample,
        peak: Peak,
        *,
        dibk_area_rank: int | None = None,
        target_sheet: str | None = None,
        target_cell: str | None = None,
        existing_value_type: str | None = None,
        existing_has_formula: bool = False,
        status: ExcelPreviewStatus = ExcelPreviewStatus.MAPPED,
        exclude_reason: str | None = None,
        message: str | None = None,
    ) -> ExcelPreviewRow:
        return ExcelPreviewRow(
            sample_name=sample.sample_name_raw,
            sample_type=sample.sample_type,
            material=peak.material_standard,
            peak_no=peak.peak_no,
            retention_time=peak.retention_time,
            area_raw=peak.area_raw,
            applied_area=self._applied_area(peak),
            dibk_area_rank=dibk_area_rank,
            target_sheet=target_sheet,
            target_cell=target_cell,
            existing_value_type=existing_value_type,
            existing_has_formula=existing_has_formula,
            status=status,
            exclude_reason=exclude_reason,
            message=message,
        )

    def _append_mapped_row(
        self,
        result: ExcelPreviewResult,
        snapshot: ExcelTemplateSnapshot,
        sample: Sample,
        peak: Peak,
        column: str,
        row: int,
        *,
        dibk_area_rank: int | None = None,
    ) -> None:
        sheet = self._target_sheet(sample)
        address = f"{column}{row}"
        cell = snapshot.cell(sheet, address)
        preview = self._row_for_peak(
            sample,
            peak,
            dibk_area_rank=dibk_area_rank,
            target_sheet=sheet,
            target_cell=address,
            existing_value_type=cell.value_type,
            existing_has_formula=cell.has_formula,
        )
        if cell.has_formula:
            preview.status = ExcelPreviewStatus.ERROR
            preview.message = "입력 대상 셀이 수식 셀입니다."
            result.issues.append(
                ExcelPreviewIssue(
                    ValidationSeverity.ERROR,
                    "TARGET_IS_FORMULA",
                    preview.message,
                    sample.sample_name_raw,
                    sheet,
                    address,
                )
            )
        result.rows.append(preview)

    def _append_error_row(
        self,
        result: ExcelPreviewResult,
        sample: Sample,
        peak: Peak,
        code: str,
        message: str,
    ) -> None:
        result.rows.append(
            self._row_for_peak(
                sample, peak, status=ExcelPreviewStatus.ERROR, message=message
            )
        )
        result.issues.append(
            ExcelPreviewIssue(
                ValidationSeverity.ERROR, code, message, sample.sample_name_raw
            )
        )

    @staticmethod
    def _sample_error(
        result: ExcelPreviewResult, sample: Sample, code: str, message: str
    ) -> None:
        result.issues.append(
            ExcelPreviewIssue(
                ValidationSeverity.ERROR, code, message, sample.sample_name_raw
            )
        )

    @staticmethod
    def _detect_target_collisions(result: ExcelPreviewResult) -> None:
        mapped: dict[tuple[str, str], list[ExcelPreviewRow]] = defaultdict(list)
        for row in result.rows:
            if (
                row.status is ExcelPreviewStatus.MAPPED
                and row.target_sheet
                and row.target_cell
            ):
                mapped[(row.target_sheet, row.target_cell)].append(row)
        for (sheet, cell), rows in mapped.items():
            if len(rows) < 2:
                continue
            message = f"동일한 입력 셀에 {len(rows)}개 Peak가 매핑되었습니다."
            for row in rows:
                row.status = ExcelPreviewStatus.ERROR
                row.message = message
            result.issues.append(
                ExcelPreviewIssue(
                    ValidationSeverity.ERROR,
                    "TARGET_COLLISION",
                    message,
                    target_sheet=sheet,
                    target_cell=cell,
                )
            )
