from __future__ import annotations

import csv
from collections import Counter
from pathlib import Path
from uuid import UUID

from honyu_app.domain.commands import AddPeakCorrectionCommand, SaveAnalysisBatchCommand
from honyu_app.domain.enums import ExcludeReason, ReviewStatus, SampleType
from honyu_app.domain.errors import ValidationError
from honyu_app.domain.models import AnalysisBatch, Peak, Sample
from honyu_app.domain.models import PeakCorrection
from honyu_app.domain.results import SaveAnalysisBatchResult
from honyu_app.domain.queries import BatchSearchQuery
from honyu_app.domain.results import BatchSummary
from honyu_app.infrastructure.pdf.material_normalizer import MaterialNormalizer
from honyu_app.services.database_service import DatabaseService


class ReviewExtractionService:
    def __init__(
        self, database: DatabaseService, normalizer: MaterialNormalizer | None = None
    ) -> None:
        self._database = database
        self._normalizer = normalizer or MaterialNormalizer()

    @property
    def standard_names(self) -> tuple[str, ...]:
        return self._normalizer.standard_names

    @staticmethod
    def _find_peak(batch: AnalysisBatch, peak_id: UUID) -> tuple[Sample, Peak]:
        for sample in batch.samples:
            for peak in sample.peaks:
                if peak.peak_id == peak_id:
                    return sample, peak
        raise ValidationError(f"Peak를 찾을 수 없습니다: {peak_id}")

    def set_material_mapping(
        self, batch: AnalysisBatch, peak_id: UUID, standard_name: str
    ) -> None:
        if standard_name not in self._normalizer.standard_names:
            raise ValidationError(f"등록되지 않은 표준 물질명입니다: {standard_name}")
        sample, peak = self._find_peak(batch, peak_id)
        peak.material_standard = standard_name
        if sample.sample_type is SampleType.BLANK:
            peak.include_for_excel = False
            peak.exclude_reason = ExcludeReason.BLANK_SAMPLE
        elif sample.sample_type is SampleType.RECOVERY_BLANK:
            peak.include_for_excel = False
            peak.exclude_reason = ExcludeReason.RECOVERY_BLANK
        elif standard_name == "CS2":
            peak.include_for_excel = False
            peak.exclude_reason = ExcludeReason.INTERNAL_STANDARD_CS2
        else:
            peak.include_for_excel = True
            peak.exclude_reason = None

    def set_peak_included(
        self, batch: AnalysisBatch, peak_id: UUID, included: bool
    ) -> None:
        sample, peak = self._find_peak(batch, peak_id)
        if included:
            if sample.is_blank:
                raise ValidationError("Blank Sample의 Peak는 Excel 입력에 포함할 수 없습니다.")
            if peak.material_standard is None:
                raise ValidationError("표준 물질명이 없는 Peak는 포함할 수 없습니다.")
            if peak.material_standard == "CS2":
                raise ValidationError("CS2는 Excel 입력에 포함할 수 없습니다.")
            peak.include_for_excel = True
            peak.exclude_reason = None
        else:
            peak.include_for_excel = False
            peak.exclude_reason = ExcludeReason.USER_EXCLUDED

    @staticmethod
    def complete_review(batch: AnalysisBatch) -> None:
        unknown = [
            peak
            for sample in batch.samples
            for peak in sample.peaks
            if peak.exclude_reason is ExcludeReason.UNKNOWN_MATERIAL
        ]
        if unknown:
            pages = sorted({peak.source_page for peak in unknown})
            material_counts = Counter(
                (peak.material_raw or "(물질명 없음)").strip() for peak in unknown
            )
            materials = ", ".join(
                f"{name} {count}개"
                for name, count in sorted(material_counts.items())
            )
            raise ValidationError(
                f"미등록 물질 Peak {len(unknown)}개를 먼저 처리해야 합니다. "
                f"물질별: {materials}. 페이지: {pages}"
            )
        batch.review_status = ReviewStatus.REVIEWED

    def save_batch(self, batch: AnalysisBatch) -> SaveAnalysisBatchResult:
        if batch.review_status is not ReviewStatus.REVIEWED:
            raise ValidationError("사용자 검토 완료 후에만 DB에 저장할 수 있습니다.")
        duplicate = self._database.check_duplicate(batch.source_file.file_hash)
        if duplicate.is_duplicate:
            raise ValidationError(
                f"동일 PDF가 이미 저장되어 있습니다: {duplicate.existing_batch_code}"
            )
        result = self._database.save_analysis_batch(SaveAnalysisBatchCommand(batch))
        batch.review_status = ReviewStatus.SAVED
        return result

    def add_area_correction(
        self,
        peak_id: UUID,
        area_after: int,
        reason: str,
        device_id: str,
    ) -> PeakCorrection:
        history = self._database.list_peak_corrections(peak_id)
        expected_revision = history[-1].revision_no if history else 0
        return self._database.add_peak_correction(
            AddPeakCorrectionCommand(
                peak_id, area_after, reason, device_id, expected_revision
            )
        ).correction

    def list_area_corrections(self, peak_id: UUID) -> list[PeakCorrection]:
        return self._database.list_peak_corrections(peak_id)

    def list_saved_batches(self) -> list[BatchSummary]:
        return self._database.search_batches(BatchSearchQuery())

    def load_saved_batch(self, batch_id: UUID) -> AnalysisBatch:
        batch = self._database.get_batch_detail(batch_id)
        batch.review_status = ReviewStatus.SAVED
        return batch

    @staticmethod
    def export_csv(batch: AnalysisBatch, output_file: Path) -> None:
        output_file.parent.mkdir(parents=True, exist_ok=True)
        with output_file.open("w", encoding="utf-8-sig", newline="") as stream:
            writer = csv.writer(stream)
            writer.writerow(
                [
                    "page_no", "sample_name_raw", "sample_name_normalized", "sample_type",
                    "replicate_no", "concentration_level", "peak_no", "retention_time",
                    "area_raw", "height", "material_raw", "material_standard",
                    "peak_group_no", "include_for_excel", "exclude_reason",
                ]
            )
            for sample in batch.samples:
                for peak in sample.peaks:
                    writer.writerow(
                        [
                            sample.page_no, sample.sample_name_raw,
                            sample.sample_name_normalized, sample.sample_type.value,
                            sample.replicate_no,
                            sample.concentration_level.value if sample.concentration_level else "",
                            peak.peak_no, str(peak.retention_time), peak.area_raw, peak.height,
                            peak.material_raw or "", peak.material_standard or "",
                            peak.peak_group_no or "", peak.include_for_excel,
                            peak.exclude_reason.value if peak.exclude_reason else "",
                        ]
                    )
