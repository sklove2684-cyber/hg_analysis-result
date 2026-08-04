from pathlib import Path
from typing import Protocol

from honyu_app.domain.models import AnalysisBatch


class PdfParser(Protocol):
    name: str
    version: str
    layout_id: str

    def parse(
        self,
        pdf_path: Path,
        *,
        analysis_type: str,
        analysis_no_start: int,
        analysis_no_end: int,
    ) -> AnalysisBatch: ...

