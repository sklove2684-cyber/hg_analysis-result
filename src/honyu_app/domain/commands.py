from dataclasses import dataclass
from uuid import UUID

from .models import AnalysisBatch


@dataclass(frozen=True, slots=True)
class SaveAnalysisBatchCommand:
    batch: AnalysisBatch


@dataclass(frozen=True, slots=True)
class AddPeakCorrectionCommand:
    peak_id: UUID
    area_after: int
    reason: str
    device_id: str
    expected_revision_no: int


@dataclass(frozen=True, slots=True)
class SaveExportJobCommand:
    batch_id: UUID
    template_path: str
    output_path: str
    std_method: str
    device_id: str

