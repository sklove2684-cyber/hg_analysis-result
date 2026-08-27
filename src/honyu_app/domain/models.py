from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from uuid import UUID, uuid4

from .enums import (
    ConcentrationLevel,
    ExcelPreviewStatus,
    ExcludeReason,
    ReviewStatus,
    SampleType,
    ValidationSeverity,
)


@dataclass(frozen=True, slots=True)
class SourceFile:
    original_name: str
    full_path: Path
    file_hash: str
    file_size: int
    page_count: int


@dataclass(slots=True)
class Peak:
    peak_no: int
    retention_time: Decimal
    area_raw: int
    height: int | None = None
    material_raw: str | None = None
    material_standard: str | None = None
    peak_group_no: int | None = None
    include_for_excel: bool = True
    exclude_reason: ExcludeReason | None = None
    source_page: int = 0
    peak_id: UUID = field(default_factory=uuid4)

    def __post_init__(self) -> None:
        if self.peak_no < 1:
            raise ValueError("peak_no must be at least 1")
        if self.retention_time < 0:
            raise ValueError("retention_time cannot be negative")
        if self.area_raw < 0:
            raise ValueError("area_raw cannot be negative")


@dataclass(slots=True)
class Sample:
    page_no: int
    sample_name_raw: str
    sample_name_normalized: str
    sample_type: SampleType
    data_filename: str | None = None
    method_filename: str | None = None
    batch_filename: str | None = None
    acquired_at: datetime | None = None
    concentration_level: ConcentrationLevel | None = None
    replicate_no: int | None = None
    worker_match_key: str | None = None
    is_blank: bool = False
    total_area: int | None = None
    peaks: list[Peak] = field(default_factory=list)
    sample_id: UUID = field(default_factory=uuid4)


@dataclass(slots=True)
class AnalysisBatch:
    batch_code: str
    source_file: SourceFile
    analysis_type: str
    analysis_no_start: int
    analysis_no_end: int
    parser_name: str
    parser_version: str
    parser_layout_id: str
    extracted_at: datetime
    samples: list[Sample] = field(default_factory=list)
    warning_count: int = 0
    review_status: ReviewStatus = ReviewStatus.PENDING
    workplace: str | None = None
    year: int | None = None
    period: str | None = None
    device_id: str | None = None
    analyst: str | None = None
    batch_id: UUID = field(default_factory=uuid4)
    # Set only after the user explicitly approves replacing a saved batch for
    # the same source PDF.  This is workflow metadata and is not persisted.
    replacement_for_batch_id: UUID | None = None


@dataclass(frozen=True, slots=True)
class PeakCorrection:
    correction_id: UUID
    peak_id: UUID
    area_before: int
    area_after: int
    reason: str
    corrected_at: datetime
    device_id: str
    revision_no: int


@dataclass(slots=True)
class ExcelPreviewRow:
    sample_name: str
    sample_type: SampleType
    material: str | None
    peak_no: int
    retention_time: Decimal
    area_raw: int
    applied_area: int
    dibk_area_rank: int | None = None
    target_sheet: str | None = None
    target_cell: str | None = None
    existing_value_type: str | None = None
    existing_has_formula: bool = False
    status: ExcelPreviewStatus = ExcelPreviewStatus.MAPPED
    exclude_reason: str | None = None
    message: str | None = None


@dataclass(frozen=True, slots=True)
class ExcelPreviewIssue:
    severity: ValidationSeverity
    code: str
    message: str
    sample_name: str | None = None
    target_sheet: str | None = None
    target_cell: str | None = None


@dataclass(slots=True)
class ExcelPreviewResult:
    template_path: Path
    std_method: str
    rows: list[ExcelPreviewRow] = field(default_factory=list)
    issues: list[ExcelPreviewIssue] = field(default_factory=list)

    @property
    def can_generate(self) -> bool:
        return not any(
            issue.severity is ValidationSeverity.ERROR for issue in self.issues
        )

    @property
    def mapped_count(self) -> int:
        return sum(row.status is ExcelPreviewStatus.MAPPED for row in self.rows)

    @property
    def excluded_count(self) -> int:
        return sum(row.status is ExcelPreviewStatus.EXCLUDED for row in self.rows)

    @property
    def error_count(self) -> int:
        return sum(
            issue.severity is ValidationSeverity.ERROR for issue in self.issues
        )


@dataclass(frozen=True, slots=True)
class ExcelCellWrite:
    sheet: str
    address: str
    value: int


@dataclass(frozen=True, slots=True)
class WorkbookValidationResult:
    valid: bool
    errors: tuple[str, ...] = ()


@dataclass(frozen=True, slots=True)
class ExcelExportResult:
    output_path: Path
    mapped_cell_count: int
    validation_passed: bool
    recalculated: bool
    export_job_id: UUID | None = None
