from __future__ import annotations

from dataclasses import dataclass
import re
from collections.abc import Mapping

from honyu_app.domain.enums import ExcludeReason, SampleType


_ANALYSIS_SAMPLE = re.compile(r"^(?P<number>\d+)(?:-(?P<suffix>.+))?$")
_DATE_QC_SAMPLE = re.compile(r"^\d{4}b[\w-]+$", re.IGNORECASE)


@dataclass(frozen=True, slots=True)
class SampleNumberDecision:
    analysis_number: str | None
    exclude_reason: str | None = None


def classify_sample_number(
    sample_name: str,
    sample_type: SampleType,
    *,
    is_blank: bool = False,
    explicit_sequence_overrides: Mapping[str, int] | None = None,
) -> SampleNumberDecision:
    """Return an explicit analysis number or a non-analysis exclusion reason.

    Sequence overrides are deliberately opt-in. They are reserved for a reviewed
    PDF where names such as 1..5 are known to represent a different continuous
    analysis-number sequence; no sequence is inferred by this function.
    """
    normalized = " ".join(sample_name.strip().split())
    compact = re.sub(r"\s+", "", normalized)
    upper = compact.upper()

    if explicit_sequence_overrides and normalized in explicit_sequence_overrides:
        return SampleNumberDecision(str(explicit_sequence_overrides[normalized]))
    if (
        is_blank
        or sample_type in {SampleType.BLANK, SampleType.RECOVERY_BLANK}
        or upper == "BLANK"
    ):
        return SampleNumberDecision(None, ExcludeReason.BLANK_SAMPLE.value)
    if upper.startswith("B"):
        return SampleNumberDecision(None, ExcludeReason.QC_SAMPLE.value)
    if _DATE_QC_SAMPLE.fullmatch(compact) or "QC" in upper or "CONTROL" in upper:
        return SampleNumberDecision(None, ExcludeReason.QC_SAMPLE.value)

    match = _ANALYSIS_SAMPLE.fullmatch(normalized)
    if match:
        return SampleNumberDecision(match.group("number"))
    return SampleNumberDecision(None, ExcludeReason.NON_ANALYSIS_SAMPLE.value)


def extract_excel_analysis_number(value: object | None) -> str | None:
    """Extract only the final analysis-number component from an Excel cell."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    match = re.search(r"(\d+)\s*$", str(value).strip())
    return match.group(1) if match else None
