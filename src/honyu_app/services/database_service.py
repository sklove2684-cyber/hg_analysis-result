from typing import Protocol
from uuid import UUID

from honyu_app.domain.commands import (
    AddPeakCorrectionCommand,
    SaveAnalysisBatchCommand,
    SaveExportJobCommand,
)
from honyu_app.domain.models import AnalysisBatch, PeakCorrection
from honyu_app.domain.queries import BatchSearchQuery
from honyu_app.domain.results import (
    BatchSummary,
    ConnectionStatus,
    DuplicateCheckResult,
    ExportJobResult,
    PeakCorrectionResult,
    SaveAnalysisBatchResult,
)


class DatabaseService(Protocol):
    def check_connection(self) -> ConnectionStatus: ...

    def check_duplicate(self, file_hash: str) -> DuplicateCheckResult: ...

    def save_analysis_batch(
        self, command: SaveAnalysisBatchCommand
    ) -> SaveAnalysisBatchResult: ...

    def search_batches(self, query: BatchSearchQuery) -> list[BatchSummary]: ...

    def get_batch_detail(self, batch_id: UUID) -> AnalysisBatch: ...

    def add_peak_correction(
        self, command: AddPeakCorrectionCommand
    ) -> PeakCorrectionResult: ...

    def list_peak_corrections(self, peak_id: UUID) -> list[PeakCorrection]: ...

    def save_export_job(self, command: SaveExportJobCommand) -> ExportJobResult: ...

