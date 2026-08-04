from dataclasses import dataclass
from uuid import UUID

from .models import AnalysisBatch, PeakCorrection


@dataclass(frozen=True, slots=True)
class ConnectionStatus:
    connected: bool
    mode: str
    message: str


@dataclass(frozen=True, slots=True)
class DuplicateCheckResult:
    is_duplicate: bool
    existing_batch_id: UUID | None = None
    existing_batch_code: str | None = None


@dataclass(frozen=True, slots=True)
class SaveAnalysisBatchResult:
    batch_id: UUID
    batch_code: str
    saved: bool


@dataclass(frozen=True, slots=True)
class BatchSummary:
    batch_id: UUID
    batch_code: str
    pdf_filename: str
    analysis_type: str
    review_status: str
    workplace: str | None = None
    year: int | None = None
    period: str | None = None
    analysis_no_start: int | None = None
    analysis_no_end: int | None = None
    parser_version: str | None = None
    correction_count: int = 0
    export_count: int = 0


AnalysisBatchDetail = AnalysisBatch


@dataclass(frozen=True, slots=True)
class PeakCorrectionResult:
    correction: PeakCorrection


@dataclass(frozen=True, slots=True)
class ExportJobResult:
    export_job_id: UUID
    saved: bool


@dataclass(frozen=True, slots=True)
class SharedFolderConnectionStatus:
    connected: bool
    active_base_path: str | None
    used_fallback: bool
    attempted_paths: tuple[str, ...]
    message: str


@dataclass(frozen=True, slots=True)
class FolderValidationResult:
    valid: bool
    path: str
    message: str
