from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class BatchSearchQuery:
    workplace: str | None = None
    year: int | None = None
    period: str | None = None
    analysis_type: str | None = None
    analysis_no_start: int | None = None
    analysis_no_end: int | None = None
    pdf_filename: str | None = None
    analyst: str | None = None
    sample_name: str | None = None
    material_name: str | None = None

