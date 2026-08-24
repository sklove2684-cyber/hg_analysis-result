from __future__ import annotations

from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
import hashlib
import re
from collections.abc import Callable

import pdfplumber

from honyu_app.domain.enums import (
    ConcentrationLevel,
    ExcludeReason,
    ReviewStatus,
    SampleType,
)
from honyu_app.domain.errors import ExtractionCancelledError, ValidationError
from honyu_app.domain.models import AnalysisBatch, Peak, Sample, SourceFile
from honyu_app.config.analysis_types import (
    runtime_material_inference_allowed_for,
    supported_canonical_names_for,
)
from honyu_app.infrastructure.pdf.material_normalizer import MaterialNormalizer


class LabSolutionsParser:
    name = "labsolutions-pdfplumber"
    version = "1.0.0"
    layout_id = "labsolutions-analysis-report-8col-v1"

    def __init__(self, normalizer: MaterialNormalizer | None = None) -> None:
        self._normalizer = normalizer or MaterialNormalizer()

    def parse(
        self,
        pdf_path: Path,
        *,
        analysis_type: str,
        analysis_no_start: int,
        analysis_no_end: int,
        progress_callback: Callable[[int, int], None] | None = None,
        cancel_check: Callable[[], bool] | None = None,
    ) -> AnalysisBatch:
        pdf_path = Path(pdf_path)
        if not pdf_path.is_file():
            raise ValidationError(f"PDF 파일을 찾을 수 없습니다: {pdf_path}")
        if analysis_no_start > analysis_no_end:
            raise ValidationError("분석번호 시작값은 종료값보다 클 수 없습니다.")
        file_bytes = pdf_path.read_bytes()
        file_hash = hashlib.sha256(file_bytes).hexdigest()
        samples: list[Sample] = []
        # An empty supported-material list means "nothing is eligible yet", not
        # "allow every canonical known by the global normalizer".
        allowed_materials = set(supported_canonical_names_for(analysis_type)) | {"CS2"}
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
            for page_no, page in enumerate(pdf.pages, 1):
                if cancel_check and cancel_check():
                    raise ExtractionCancelledError("PDF 추출이 취소되었습니다.")
                text = page.extract_text() or ""
                table = self._find_peak_table(page.extract_tables(), page_no)
                if self._field(text, "Sample Name"):
                    metadata = self._extract_metadata(text, page_no)
                    sample = self._build_sample(
                        page_no, metadata, table, allowed_materials
                    )
                    samples.append(sample)
                    added_peaks = sample.peaks
                else:
                    if not samples:
                        raise ValidationError(
                            f"PDF {page_no}페이지: Sample 정보가 없는 연속 Peak Table 앞에 "
                            "원본 Sample 페이지가 없습니다."
                        )
                    added_peaks = self._append_continuation_page(
                        samples[-1], page_no, table, allowed_materials
                    )
                if progress_callback:
                    progress_callback(page_no, total_pages)
            page_count = total_pages
        inferred_aliases = self._infer_session_material_aliases(analysis_type, samples)
        if inferred_aliases:
            allowed_materials = set(inferred_aliases.values()) | {"CS2"}
            self._apply_session_material_aliases(
                samples, inferred_aliases, allowed_materials
            )
        warning_count = sum(
            peak.exclude_reason is ExcludeReason.UNKNOWN_MATERIAL
            for sample in samples
            for peak in sample.peaks
        )
        source = SourceFile(
            original_name=pdf_path.name,
            full_path=pdf_path.resolve(),
            file_hash=file_hash,
            file_size=len(file_bytes),
            page_count=page_count,
        )
        extracted_at = datetime.now().astimezone()
        return AnalysisBatch(
            batch_code=(
                f"{analysis_type}-{analysis_no_start}-{analysis_no_end}-"
                f"{extracted_at:%Y%m%d%H%M%S}"
            ),
            source_file=source,
            analysis_type=analysis_type,
            analysis_no_start=analysis_no_start,
            analysis_no_end=analysis_no_end,
            parser_name=self.name,
            parser_version=self.version,
            parser_layout_id=self.layout_id,
            extracted_at=extracted_at,
            samples=samples,
            warning_count=warning_count,
            review_status=ReviewStatus.PENDING,
        )

    @staticmethod
    def _material_key(value: str) -> str:
        return re.sub(r"\s+", "", value).casefold()

    def _infer_session_material_aliases(
        self, analysis_type: str, samples: list[Sample]
    ) -> dict[str, str]:
        """Infer only an exact, repeated single-material match for this parse run.

        This intentionally does not persist aliases or guess abbreviations. Complex
        pending analyses therefore remain unresolved until their registry is defined.
        """
        if not runtime_material_inference_allowed_for(analysis_type):
            return {}
        analysis_key = self._material_key(analysis_type)
        std_samples_by_raw: dict[str, set[int]] = {}
        recovery_samples_by_raw: dict[str, set[int]] = {}
        for sample in samples:
            if sample.sample_type not in {SampleType.STD, SampleType.RECOVERY}:
                continue
            target = (
                std_samples_by_raw
                if sample.sample_type is SampleType.STD
                else recovery_samples_by_raw
            )
            for peak in sample.peaks:
                if not peak.material_raw or peak.material_standard is not None:
                    continue
                raw_key = self._material_key(peak.material_raw)
                target.setdefault(raw_key, set()).add(sample.page_no)

        exact = [
            raw_key
            for raw_key, pages in std_samples_by_raw.items()
            if raw_key == analysis_key and len(pages) >= 2
        ]
        if len(exact) != 1:
            return {}
        raw_key = exact[0]
        # Recovery confirmation strengthens the evidence when present, but is not
        # mandatory because some valid single-material reports contain only STD.
        _ = recovery_samples_by_raw.get(raw_key, set())
        return {raw_key: analysis_type}

    def _apply_session_material_aliases(
        self,
        samples: list[Sample],
        aliases: dict[str, str],
        allowed_materials: set[str],
    ) -> None:
        for sample in samples:
            for peak in sample.peaks:
                if peak.material_raw:
                    inferred = aliases.get(self._material_key(peak.material_raw))
                    if inferred is not None:
                        peak.material_standard = inferred
                reason = self._exclude_reason(
                    sample.sample_name_raw,
                    sample.sample_type,
                    peak.material_raw,
                    peak.material_standard,
                    allowed_materials,
                )
                peak.exclude_reason = reason
                peak.include_for_excel = reason is None

    @staticmethod
    def extract_analysis_range(filename: str) -> tuple[int, int] | None:
        matches = re.findall(r"(?<!\d)(\d+)\s*-\s*(\d+)(?!\d)", filename)
        if not matches:
            return None
        start, end = matches[-1]
        return int(start), int(end)

    @staticmethod
    def _field(text: str, label: str, required: bool = False) -> str | None:
        match = re.search(rf"^{re.escape(label)}\s*:\s*(.*?)\s*$", text, re.MULTILINE)
        value = match.group(1).strip() if match else None
        if required and not value:
            raise ValidationError(f"필수 PDF 필드를 찾을 수 없습니다: {label}")
        return value

    def _extract_metadata(self, text: str, page_no: int) -> dict[str, object]:
        try:
            sample_name = self._field(text, "Sample Name", required=True)
            assert sample_name is not None
            return {
                "sample_name": sample_name,
                "data_filename": self._field(text, "Data Filename"),
                "method_filename": self._field(text, "Method Filename"),
                "batch_filename": self._field(text, "Batch Filename"),
                "acquired_at": self._parse_acquired_at(text),
            }
        except ValidationError as exc:
            raise ValidationError(f"PDF {page_no}페이지: {exc}") from exc

    @staticmethod
    def _parse_acquired_at(text: str) -> datetime | None:
        match = re.search(
            r"Date Acquired\s*:\s*(\d{4}-\d{2}-\d{2})\s*"
            r"(오전|오후)\s*(\d{1,2}):(\d{2}):(\d{2})",
            text,
        )
        if not match:
            return None
        day, meridiem, hour_text, minute, second = match.groups()
        hour = int(hour_text)
        if meridiem == "오전" and hour == 12:
            hour = 0
        elif meridiem == "오후" and hour < 12:
            hour += 12
        base = datetime.strptime(day, "%Y-%m-%d")
        return base.replace(hour=hour, minute=int(minute), second=int(second))

    @staticmethod
    def _find_peak_table(tables: list[list[list[str | None]]], page_no: int) -> list[list[str | None]]:
        expected = ["Peak#", "Ret. Time", "Area", "Height", "Conc.", "Unit", "Mark", "Name"]
        for table in tables:
            for index, row in enumerate(table):
                normalized = [(cell or "").strip() for cell in row]
                if normalized == expected:
                    rows: list[list[str | None]] = []
                    for candidate in table[index + 1 :]:
                        first = (candidate[0] or "").strip() if candidate else ""
                        if first == "Total":
                            break
                        if first.isdigit():
                            rows.append(candidate)
                    if rows:
                        return rows
        raise ValidationError(f"PDF {page_no}페이지: 8열 Peak Table을 찾을 수 없습니다.")

    def _build_sample(
        self,
        page_no: int,
        metadata: dict[str, object],
        rows: list[list[str | None]],
        allowed_materials: set[str] | None,
    ) -> Sample:
        raw_name = str(metadata["sample_name"])
        sample_type, level, replicate, worker_key, is_blank = self._classify_sample(raw_name)
        peaks = self._build_peaks(
            page_no, raw_name, sample_type, rows, allowed_materials=allowed_materials
        )
        return Sample(
            page_no=page_no,
            sample_name_raw=raw_name,
            sample_name_normalized=self._normalize_sample_name(raw_name),
            sample_type=sample_type,
            data_filename=metadata["data_filename"],
            method_filename=metadata["method_filename"],
            batch_filename=metadata["batch_filename"],
            acquired_at=metadata["acquired_at"],
            concentration_level=level,
            replicate_no=replicate,
            worker_match_key=worker_key,
            is_blank=is_blank,
            peaks=peaks,
        )

    def _append_continuation_page(
        self,
        sample: Sample,
        page_no: int,
        rows: list[list[str | None]],
        allowed_materials: set[str] | None,
    ) -> list[Peak]:
        existing_peak_numbers = {peak.peak_no for peak in sample.peaks}
        added = self._build_peaks(
            page_no,
            sample.sample_name_raw,
            sample.sample_type,
            rows,
            allowed_materials=allowed_materials,
            dibk_group_start=sum(
                peak.material_standard == "DIBK" for peak in sample.peaks
            ),
        )
        duplicate_numbers = existing_peak_numbers.intersection(
            peak.peak_no for peak in added
        )
        if duplicate_numbers:
            raise ValidationError(
                f"PDF {page_no}페이지: 연속 Peak Table의 Peak 번호가 앞 페이지와 "
                f"중복됩니다: {sorted(duplicate_numbers)}"
            )
        sample.peaks.extend(added)
        return added

    def _build_peaks(
        self,
        page_no: int,
        raw_name: str,
        sample_type: SampleType,
        rows: list[list[str | None]],
        *,
        allowed_materials: set[str] | None = None,
        dibk_group_start: int = 0,
    ) -> list[Peak]:
        peaks: list[Peak] = []
        dibk_group = dibk_group_start
        for row in rows:
            if len(row) != 8:
                raise ValidationError(
                    f"PDF {page_no}페이지: Peak Table 열 수가 8개가 아닙니다."
                )
            raw_material = (row[7] or "").strip() or None
            standard = self._normalizer.normalize(raw_material)
            reason = self._exclude_reason(
                raw_name,
                sample_type,
                raw_material,
                standard,
                allowed_materials,
            )
            if standard == "DIBK":
                dibk_group += 1
                group_no = dibk_group
            else:
                group_no = None
            try:
                peak = Peak(
                    peak_no=int((row[0] or "").strip()),
                    retention_time=Decimal((row[1] or "").strip()),
                    area_raw=int((row[2] or "").replace(",", "").strip()),
                    height=int((row[3] or "").replace(",", "").strip()),
                    material_raw=raw_material,
                    material_standard=standard,
                    peak_group_no=group_no,
                    include_for_excel=reason is None,
                    exclude_reason=reason,
                    source_page=page_no,
                )
            except (ValueError, InvalidOperation) as exc:
                raise ValidationError(
                    f"PDF {page_no}페이지 Peak {row[0]} 숫자 변환 오류"
                ) from exc
            peaks.append(peak)
        return peaks

    @staticmethod
    def _normalize_sample_name(value: str) -> str:
        return re.sub(r"\s+", "", value).replace(":", "")

    def _classify_sample(
        self, raw_name: str
    ) -> tuple[SampleType, ConcentrationLevel | None, int | None, str | None, bool]:
        name = self._normalize_sample_name(raw_name)
        if name.casefold() == "blank":
            return SampleType.BLANK, None, None, None, True
        std = re.fullmatch(r"STD([1-6])", name, re.IGNORECASE)
        if std:
            return SampleType.STD, None, int(std.group(1)), None, False
        if name.endswith("B"):
            return SampleType.RECOVERY_BLANK, None, None, None, True
        recovery = re.fullmatch(r"(저|중|고)(\d+)", name)
        if recovery:
            levels = {
                "저": ConcentrationLevel.LOW,
                "중": ConcentrationLevel.MID,
                "고": ConcentrationLevel.HIGH,
            }
            return SampleType.RECOVERY, levels[recovery.group(1)], int(recovery.group(2)), None, False
        numeric = re.fullmatch(r"(\d+)(?:\D.*)?", name)
        if numeric:
            return SampleType.NUMERIC, None, None, numeric.group(1), False
        return SampleType.UNKNOWN, None, None, None, False

    @staticmethod
    def _exclude_reason(
        raw_sample_name: str,
        sample_type: SampleType,
        raw_material: str | None,
        standard: str | None,
        allowed_materials: set[str] | None = None,
    ) -> ExcludeReason | None:
        if sample_type is SampleType.BLANK:
            return ExcludeReason.BLANK_SAMPLE
        if sample_type is SampleType.RECOVERY_BLANK:
            return ExcludeReason.RECOVERY_BLANK
        if LabSolutionsParser._normalize_sample_name(raw_sample_name).endswith("B"):
            return ExcludeReason.SAMPLE_NAME_ENDS_WITH_B
        if not raw_material:
            return ExcludeReason.UNNAMED_PEAK
        if standard is None:
            return ExcludeReason.UNKNOWN_MATERIAL
        if standard == "CS2":
            return ExcludeReason.INTERNAL_STANDARD_CS2
        if allowed_materials is not None and standard not in allowed_materials:
            return ExcludeReason.MATERIAL_NOT_SUPPORTED_FOR_ANALYSIS
        return None
