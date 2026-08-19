from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal
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


LEGACY_MATERIAL_COLUMNS = {
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
LEGACY_RECOVERY_COLUMNS = {
    material: chr(ord("B") + index)
    for index, material in enumerate(LEGACY_MATERIAL_COLUMNS)
}
STD_REPLICATES = {
    StdMethod.A: (1, 2, 3, 4, 5),
    StdMethod.B: (1, 2, 3, 4, 6),
}
LEGACY_RECOVERY_ROW_START = {
    ConcentrationLevel.LOW: 37,
    ConcentrationLevel.MID: 40,
    ConcentrationLevel.HIGH: 43,
}


@dataclass(frozen=True)
class TemplateProfile:
    name: str
    required_sheets: tuple[str, ...]
    area_sheet: str
    std_columns: dict[str, str]
    numeric_columns: dict[str, str]
    recovery_columns: dict[str, str]
    std_row_start: int
    recovery_row_start: dict[ConcentrationLevel, int]
    worker_row_start: int
    worker_row_end: int
    dibk_std_slots: tuple[str, ...] = ()
    dibk_recovery_slots: tuple[str, ...] = ()


LEGACY_PROFILE = TemplateProfile(
    name="혼유",
    required_sheets=("검량선", "area", "최종결과", "회수율", "STD제조"),
    area_sheet="area",
    std_columns=LEGACY_MATERIAL_COLUMNS,
    numeric_columns=LEGACY_MATERIAL_COLUMNS,
    recovery_columns=LEGACY_RECOVERY_COLUMNS,
    std_row_start=15,
    recovery_row_start=LEGACY_RECOVERY_ROW_START,
    worker_row_start=37,
    worker_row_end=183,
    dibk_std_slots=("Z", "AA"),
    dibk_recovery_slots=("U", "V"),
)

ONE_COLUMN_PROFILE = TemplateProfile(
    name="1컬럼혼유",
    required_sheets=("검량선", "area입력", "회수율", "STD제조"),
    area_sheet="area입력",
    std_columns={
        "methyl acetate": "G",
        "c-hexane": "J",
        "n-heptane": "M",
        "isobutyl acetate": "P",
    },
    numeric_columns={
        "methyl acetate": "F",
        "c-hexane": "I",
        "n-heptane": "L",
        "isobutyl acetate": "O",
    },
    recovery_columns={
        "methyl acetate": "B",
        "c-hexane": "C",
        "n-heptane": "D",
        "isobutyl acetate": "E",
    },
    std_row_start=5,
    recovery_row_start={
        ConcentrationLevel.LOW: 30,
        ConcentrationLevel.MID: 33,
        ConcentrationLevel.HIGH: 36,
    },
    worker_row_start=21,
    worker_row_end=287,
)

ALCOHOL_PROFILE = TemplateProfile(
    name="알콜",
    required_sheets=("검량선", "area입력", "회수율", "STD제조"),
    area_sheet="area입력",
    std_columns={
        "IBA": "G",
        "n-BTOH": "J",
    },
    numeric_columns={
        "IBA": "F",
        "n-BTOH": "I",
    },
    recovery_columns={
        "IBA": "B",
        "n-BTOH": "C",
    },
    std_row_start=5,
    recovery_row_start={
        ConcentrationLevel.LOW: 30,
        ConcentrationLevel.MID: 33,
        ConcentrationLevel.HIGH: 36,
    },
    worker_row_start=21,
    worker_row_end=287,
)

ONE_COLUMN_TARGET_RETENTION_TIMES = {
    "methyl acetate": Decimal("2.551"),
    "c-hexane": Decimal("3.730"),
    "n-heptane": Decimal("4.195"),
    "isobutyl acetate": Decimal("4.910"),
}

ALCOHOL_TARGET_RETENTION_TIMES = {
    "IBA": Decimal("3.391"),
    "n-BTOH": Decimal("3.858"),
}


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
        profile = self._template_profile(snapshot, result)
        if profile is None:
            return result
        worker_rows = self._worker_row_index(snapshot, profile)

        for sample in batch.samples:
            self._map_sample(sample, method, snapshot, profile, worker_rows, result)
        self._detect_target_collisions(result)
        return result

    @staticmethod
    def _template_profile(
        snapshot: ExcelTemplateSnapshot, result: ExcelPreviewResult
    ) -> TemplateProfile | None:
        if ONE_COLUMN_PROFILE.area_sheet in snapshot.sheet_names:
            iba_header = str(
                snapshot.cell(ONE_COLUMN_PROFILE.area_sheet, "F3").value or ""
            ).strip().casefold()
            btoh_header = str(
                snapshot.cell(ONE_COLUMN_PROFILE.area_sheet, "I3").value or ""
            ).strip().casefold()
            if iba_header == "iba" and btoh_header in {"1-btoh", "n-btoh"}:
                profile = ALCOHOL_PROFILE
            else:
                profile = ONE_COLUMN_PROFILE
        elif LEGACY_PROFILE.area_sheet in snapshot.sheet_names:
            profile = LEGACY_PROFILE
        else:
            result.issues.append(
                ExcelPreviewIssue(
                    ValidationSeverity.ERROR,
                    "TEMPLATE_PROFILE_UNSUPPORTED",
                    "지원하는 Excel 양식이 아닙니다. area 또는 area입력 시트가 필요합니다.",
                )
            )
            return None
        missing = [
            name for name in profile.required_sheets if name not in snapshot.sheet_names
        ]
        if missing:
            result.issues.append(
                ExcelPreviewIssue(
                    ValidationSeverity.ERROR,
                    "TEMPLATE_SHEET_MISSING",
                    f"필수 시트가 없습니다: {', '.join(missing)}",
                )
            )
        return profile

    @staticmethod
    def _worker_key(value: object | None) -> str | None:
        if value is None or isinstance(value, bool):
            return None
        if isinstance(value, float) and value.is_integer():
            value = int(value)
        match = re.search(r"(\d+)\s*$", str(value).strip())
        return match.group(1) if match else None

    @staticmethod
    def _sample_worker_key(sample: Sample) -> str | None:
        if sample.worker_match_key:
            return sample.worker_match_key
        name = re.sub(r"\s+", "", sample.sample_name_normalized).replace(":", "")
        match = re.fullmatch(r"(\d+)(?:\D.*)?", name)
        return match.group(1) if match else None

    def _worker_row_index(
        self, snapshot: ExcelTemplateSnapshot, profile: TemplateProfile
    ) -> dict[str, list[int]]:
        index: dict[str, list[int]] = defaultdict(list)
        for row in range(profile.worker_row_start, profile.worker_row_end + 1):
            analysis_cell = snapshot.cell(profile.area_sheet, f"A{row}")
            # Some templates contain a second, calculated table that mirrors the
            # input table.  Its analysis numbers are formulas and must not be
            # treated as writable worker rows.
            if analysis_cell.has_formula:
                continue
            key = self._worker_key(analysis_cell.value)
            if key is not None:
                index[key].append(row)
        return dict(index)

    def _map_sample(
        self,
        sample: Sample,
        method: StdMethod,
        snapshot: ExcelTemplateSnapshot,
        profile: TemplateProfile,
        worker_rows: dict[str, list[int]],
        result: ExcelPreviewResult,
    ) -> None:
        row = self._sample_target_row(sample, method, profile, worker_rows, result)
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
        single_groups: dict[str, list[Peak]] = defaultdict(list)
        for peak in eligible:
            if peak.material_standard != "DIBK":
                single_groups[peak.material_standard or ""].append(peak)
        for material, peaks in single_groups.items():
            if profile == ONE_COLUMN_PROFILE:
                target_rt = ONE_COLUMN_TARGET_RETENTION_TIMES.get(material)
            elif profile == ALCOHOL_PROFILE:
                target_rt = ALCOHOL_TARGET_RETENTION_TIMES.get(material)
            else:
                target_rt = None
            if target_rt is None:
                ranked_single = sorted(
                    peaks,
                    key=lambda peak: (
                        -self._applied_area(peak),
                        peak.peak_no,
                        peak.retention_time,
                    ),
                )
                residual_reason = ExcludeReason.MATERIAL_AREA_NOT_TOP1.value
                residual_message = "동일 물질 중 적용 Area가 가장 큰 피크가 아님"
            else:
                ranked_single = sorted(
                    peaks,
                    key=lambda peak: (
                        abs(peak.retention_time - target_rt),
                        peak.peak_no,
                        -self._applied_area(peak),
                    ),
                )
                residual_reason = ExcludeReason.MATERIAL_RT_NOT_CLOSEST.value
                residual_message = (
                    f"동일 물질 중 기준 RT {target_rt}에 가장 가까운 피크가 아님"
                )
            peak = ranked_single[0]
            column = self._material_column(profile, sample, material)
            if column is None:
                self._append_error_row(
                    result,
                    sample,
                    peak,
                    "UNSUPPORTED_MATERIAL",
                    f"Excel 핵심 물질 열이 없습니다: {peak.material_standard}",
                )
                continue
            self._append_mapped_row(
                result, snapshot, profile, sample, peak, column, row
            )
            for residual in ranked_single[1:]:
                result.rows.append(
                    self._row_for_peak(
                        sample,
                        residual,
                        status=ExcelPreviewStatus.EXCLUDED,
                        exclude_reason=residual_reason,
                        message=residual_message,
                    )
                )

        ranked = sorted(
            dibk,
            key=lambda peak: (
                -self._applied_area(peak),
                peak.peak_no,
                peak.retention_time,
            ),
        )
        slots = self._dibk_slots(profile, sample)
        if dibk and len(slots) < 2:
            for peak in dibk:
                self._append_error_row(
                    result,
                    sample,
                    peak,
                    "UNSUPPORTED_MATERIAL",
                    f"{profile.name} Excel 양식에는 DIBK 입력 열이 없습니다.",
                )
            return
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
                profile,
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
        profile: TemplateProfile,
        worker_rows: dict[str, list[int]],
        result: ExcelPreviewResult,
    ) -> int | None:
        if sample.sample_type is SampleType.STD:
            selected = STD_REPLICATES[method]
            if sample.replicate_no not in selected:
                return None
            return profile.std_row_start + selected.index(sample.replicate_no)
        if sample.sample_type is SampleType.RECOVERY:
            if sample.concentration_level not in profile.recovery_row_start:
                self._sample_error(
                    result, sample, "RECOVERY_LEVEL_MISSING", "회수율 농도 구분이 없습니다."
                )
                return None
            if sample.replicate_no not in {1, 2, 3}:
                self._sample_error(
                    result, sample, "RECOVERY_REPLICATE_INVALID", "회수율 반복번호는 1~3이어야 합니다."
                )
                return None
            return (
                profile.recovery_row_start[sample.concentration_level]
                + sample.replicate_no
                - 1
            )
        key = self._sample_worker_key(sample)
        if sample.sample_type is SampleType.NUMERIC or (
            sample.sample_type is SampleType.UNKNOWN and key is not None
        ):
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

    @classmethod
    def _sample_exclusion_reason(cls, sample: Sample, method: StdMethod) -> str | None:
        if sample.sample_type is SampleType.STD:
            selected = STD_REPLICATES[method]
            if sample.replicate_no not in selected:
                return f"STD_METHOD_{method.value}_NOT_SELECTED"
        if sample.sample_type is SampleType.UNKNOWN and cls._sample_worker_key(sample):
            return None
        if sample.sample_type in {
            SampleType.BLANK,
            SampleType.RECOVERY_BLANK,
            SampleType.UNKNOWN,
        }:
            return f"SAMPLE_TYPE_{sample.sample_type.value}"
        return None

    @classmethod
    def _material_column(
        cls, profile: TemplateProfile, sample: Sample, material: str | None
    ) -> str | None:
        if sample.sample_type is SampleType.RECOVERY:
            return profile.recovery_columns.get(material or "")
        if sample.sample_type is SampleType.NUMERIC or (
            sample.sample_type is SampleType.UNKNOWN and cls._sample_worker_key(sample)
        ):
            return profile.numeric_columns.get(material or "")
        return profile.std_columns.get(material or "")

    @staticmethod
    def _dibk_slots(
        profile: TemplateProfile, sample: Sample
    ) -> tuple[str, ...]:
        return (
            profile.dibk_recovery_slots
            if sample.sample_type is SampleType.RECOVERY
            else profile.dibk_std_slots
        )

    @staticmethod
    def _target_sheet(profile: TemplateProfile, sample: Sample) -> str:
        return "회수율" if sample.sample_type is SampleType.RECOVERY else profile.area_sheet

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
        profile: TemplateProfile,
        sample: Sample,
        peak: Peak,
        column: str,
        row: int,
        *,
        dibk_area_rank: int | None = None,
    ) -> None:
        sheet = self._target_sheet(profile, sample)
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
