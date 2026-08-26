from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path
import re
from uuid import UUID

from honyu_app.config.analysis_types import excel_profile_key_for, has_excel_profile
from honyu_app.application.sample_number_matching import (
    SampleNumberDecision,
    classify_sample_number,
    extract_excel_analysis_number,
)
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
    key: str
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
    std_replicates_override: tuple[int, ...] = ()


LEGACY_PROFILE = TemplateProfile(
    key="mixture",
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
    key="mixture_one_column",
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
    key="alcohol_2",
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

MEK_PROFILE = TemplateProfile(
    key="mek",
    name="MEK",
    required_sheets=("검량선", "LOD(area입력)", "회수율", "std"),
    area_sheet="LOD(area입력)",
    std_columns={"methyl ethyl ketone": "F"},
    numeric_columns={"methyl ethyl ketone": "F"},
    recovery_columns={"methyl ethyl ketone": "B"},
    std_row_start=6,
    recovery_row_start={
        ConcentrationLevel.LOW: 28,
        ConcentrationLevel.MID: 31,
        ConcentrationLevel.HIGH: 34,
    },
    worker_row_start=21,
    worker_row_end=66,
)

ACETIC_ACID_PROFILE = TemplateProfile(
    key="acetic_acid",
    name="초산",
    required_sheets=("검량선", "LOD(area입력)", "회수율", "std"),
    area_sheet="LOD(area입력)",
    std_columns={"초산": "F"},
    numeric_columns={"초산": "F"},
    recovery_columns={"초산": "B"},
    std_row_start=4,
    recovery_row_start={
        ConcentrationLevel.LOW: 28,
        ConcentrationLevel.MID: 31,
        ConcentrationLevel.HIGH: 34,
    },
    worker_row_start=19,
    worker_row_end=52,
)

ETHYLENE_GLYCOL_PROFILE = TemplateProfile(
    key="ethylene_glycol",
    name="에틸렌글리콜",
    required_sheets=("검량선", "LOD(area입력)", "회수율", "std"),
    area_sheet="LOD(area입력)",
    std_columns={"Ethylene glycol": "F"},
    numeric_columns={"Ethylene glycol": "F"},
    recovery_columns={"Ethylene glycol": "B"},
    std_row_start=4,
    recovery_row_start={
        ConcentrationLevel.LOW: 28,
        ConcentrationLevel.MID: 31,
        ConcentrationLevel.HIGH: 34,
    },
    worker_row_start=19,
    worker_row_end=40,
)

DIETHYL_ETHER_PROFILE = TemplateProfile(
    key="diethyl_ether",
    name="디에틸에테르",
    required_sheets=("검량선", "LOD(area입력)", "회수율", "std"),
    area_sheet="LOD(area입력)",
    std_columns={"Diethyl ether": "F"},
    numeric_columns={"Diethyl ether": "F"},
    recovery_columns={"Diethyl ether": "B"},
    std_row_start=4,
    recovery_row_start={
        ConcentrationLevel.LOW: 28,
        ConcentrationLevel.MID: 31,
        ConcentrationLevel.HIGH: 34,
    },
    worker_row_start=19,
    worker_row_end=122,
    # The source PDF has six named STD injections, while the workbook has five
    # non-zero calibration levels. STD2 and STD3 represent the same level; the
    # approved template levels correspond to source STD1, STD2, STD4, STD5, STD6.
    std_replicates_override=(1, 2, 4, 5, 6),
)

ACN_PROFILE = TemplateProfile(
    key="acn",
    name="ACN",
    required_sheets=("검량선", "결과입력(area입력)", "회수율", "STD제조"),
    area_sheet="결과입력(area입력)",
    std_columns={"Acetonitrile": "F"},
    numeric_columns={"Acetonitrile": "F"},
    recovery_columns={"Acetonitrile": "D"},
    std_row_start=14,
    recovery_row_start={
        ConcentrationLevel.LOW: 29,
        ConcentrationLevel.MID: 32,
        ConcentrationLevel.HIGH: 35,
    },
    worker_row_start=26,
    worker_row_end=79,
)

BC_PROFILE = TemplateProfile(
    key="bc",
    name="B.C",
    required_sheets=("검량선", "결과입력(area입력)", "회수율", "STD제조"),
    area_sheet="결과입력(area입력)",
    std_columns={"2-Butoxyethanol": "F"},
    numeric_columns={"2-Butoxyethanol": "F"},
    recovery_columns={"2-Butoxyethanol": "B"},
    std_row_start=14,
    recovery_row_start={
        ConcentrationLevel.LOW: 29,
        ConcentrationLevel.MID: 32,
        ConcentrationLevel.HIGH: 35,
    },
    worker_row_start=26,
    worker_row_end=76,
)

G2_PROFILE = TemplateProfile(
    key="mixture_g2",
    name="(혼유-G2) THF,CFM,벤젠,클로로벤젠",
    required_sheets=("검량선", "area입력", "회수율", "STD제조"),
    area_sheet="area입력",
    std_columns={
        "THF": "G",
        "CFM": "J",
        "벤젠": "M",
        "클로로벤젠": "P",
    },
    numeric_columns={
        "THF": "F",
        "CFM": "I",
        "벤젠": "L",
        "클로로벤젠": "O",
    },
    recovery_columns={
        "THF": "B",
        "CFM": "C",
        "벤젠": "D",
        "클로로벤젠": "E",
    },
    std_row_start=5,
    recovery_row_start={
        ConcentrationLevel.LOW: 30,
        ConcentrationLevel.MID: 33,
        ConcentrationLevel.HIGH: 36,
    },
    worker_row_start=21,
    worker_row_end=41,
)

ISOAMYL_N_PROPYL_ACETATE_PROFILE = TemplateProfile(
    key="isoamyl_n_propyl_acetate",
    name="이소아밀,n-프로필 아세테이트",
    required_sheets=("검량선", "LOD(area입력)", "회수율", "std"),
    area_sheet="LOD(area입력)",
    std_columns={
        "이소아밀 아세테이트": "G",
        "n-프로필 아세테이트": "J",
    },
    numeric_columns={
        "이소아밀 아세테이트": "F",
        "n-프로필 아세테이트": "I",
    },
    recovery_columns={
        "이소아밀 아세테이트": "B",
        "n-프로필 아세테이트": "C",
    },
    std_row_start=5,
    recovery_row_start={
        ConcentrationLevel.LOW: 28,
        ConcentrationLevel.MID: 31,
        ConcentrationLevel.HIGH: 34,
    },
    worker_row_start=20,
    worker_row_end=100,
)

CELLOSOLVE_PROFILE = TemplateProfile(
    key="cellosolve",
    name="셀로솔브",
    required_sheets=("검량선", "area입력", "회수율", "STD제조"),
    area_sheet="area입력",
    std_columns={
        "2-Butoxyethanol": "G",
        "2-Butoxyethyl acetate": "J",
        "2-Ethoxyethanol": "M",
        "2-Ethoxyethyl acetate": "P",
    },
    numeric_columns={
        "2-Butoxyethanol": "F",
        "2-Butoxyethyl acetate": "I",
        "2-Ethoxyethanol": "L",
        "2-Ethoxyethyl acetate": "O",
    },
    recovery_columns={
        "2-Butoxyethanol": "B",
        "2-Butoxyethyl acetate": "C",
        "2-Ethoxyethanol": "D",
        "2-Ethoxyethyl acetate": "E",
    },
    std_row_start=5,
    recovery_row_start={
        ConcentrationLevel.LOW: 30,
        ConcentrationLevel.MID: 33,
        ConcentrationLevel.HIGH: 36,
    },
    worker_row_start=21,
    worker_row_end=57,
)

G3_PROFILE = TemplateProfile(
    key="mixture_g3",
    name="(혼유-G3) 1,2-디클로로에틸렌,퍼클로로에틸렌,프로판,에탄",
    required_sheets=("검량선", "area입력", "회수율", "STD제조"),
    area_sheet="area입력",
    std_columns={
        "1,2-Dichloroethylene": "G",
        "Trichloroethylene": "J",
        "Tetrachloroethylene": "M",
        "1,2-Dichloropropane": "P",
        "1,2-Dichloroethane": "S",
    },
    numeric_columns={
        "1,2-Dichloroethylene": "F",
        "Trichloroethylene": "I",
        "Tetrachloroethylene": "L",
        "1,2-Dichloropropane": "O",
        "1,2-Dichloroethane": "R",
    },
    recovery_columns={
        "1,2-Dichloroethylene": "B",
        "Trichloroethylene": "C",
        "Tetrachloroethylene": "D",
        "1,2-Dichloropropane": "E",
        "1,2-Dichloroethane": "F",
    },
    std_row_start=5,
    recovery_row_start={
        ConcentrationLevel.LOW: 30,
        ConcentrationLevel.MID: 33,
        ConcentrationLevel.HIGH: 36,
    },
    worker_row_start=21,
    worker_row_end=41,
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

MEK_TARGET_RETENTION_TIMES = {
    "methyl ethyl ketone": Decimal("3.915"),
}

ACETIC_ACID_TARGET_RETENTION_TIMES = {
    "초산": Decimal("9.125"),
}
ETHYLENE_GLYCOL_TARGET_RETENTION_TIMES = {
    "Ethylene glycol": Decimal("3.389"),
}
DIETHYL_ETHER_TARGET_RETENTION_TIMES = {
    "Diethyl ether": Decimal("1.305"),
}
ACN_TARGET_RETENTION_TIMES = {
    "Acetonitrile": Decimal("2.031"),
}
BC_TARGET_RETENTION_TIMES = {
    "2-Butoxyethanol": Decimal("6.128"),
}

G2_TARGET_RETENTION_TIMES = {
    "THF": Decimal("3.700"),
    "벤젠": Decimal("4.831"),
    "CFM": Decimal("6.387"),
    "클로로벤젠": Decimal("11.880"),
}

ISOAMYL_N_PROPYL_ACETATE_TARGET_RETENTION_TIMES = {
    "n-프로필 아세테이트": Decimal("4.901"),
    "이소아밀 아세테이트": Decimal("7.362"),
}

CELLOSOLVE_TARGET_RETENTION_TIMES = {
    "2-Ethoxyethanol": Decimal("6.453"),
    "2-Ethoxyethyl acetate": Decimal("7.145"),
    "2-Butoxyethanol": Decimal("8.117"),
    "2-Butoxyethyl acetate": Decimal("8.737"),
}

G3_TARGET_RETENTION_TIMES = {
    "1,2-Dichloroethylene": Decimal("3.698"),
    "Trichloroethylene": Decimal("5.806"),
    "Tetrachloroethylene": Decimal("6.437"),
    "1,2-Dichloropropane": Decimal("7.005"),
    "1,2-Dichloroethane": Decimal("7.675"),
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
        if not has_excel_profile(batch.analysis_type):
            result.issues.append(
                ExcelPreviewIssue(
                    ValidationSeverity.ERROR,
                    "EXCEL_PROFILE_NOT_REGISTERED",
                    f"아직 등록되지 않은 Excel 양식/프로필입니다: {batch.analysis_type}",
                )
            )
            return result
        profile = self._template_profile(snapshot, result)
        if profile is None:
            return result
        expected_profile_key = excel_profile_key_for(batch.analysis_type)
        if expected_profile_key != profile.key:
            result.issues.append(
                ExcelPreviewIssue(
                    ValidationSeverity.ERROR,
                    "TEMPLATE_PROFILE_MISMATCH",
                    f"분석종류 '{batch.analysis_type}'에 맞는 Excel 양식이 아닙니다. "
                    f"선택된 파일은 '{profile.name}' 양식으로 확인됐습니다.",
                )
            )
            return result
        worker_rows = self._worker_row_index(snapshot, profile)
        target_analysis_numbers = set(worker_rows)
        if batch.analysis_no_start <= batch.analysis_no_end:
            target_analysis_numbers.update(
                str(number)
                for number in range(batch.analysis_no_start, batch.analysis_no_end + 1)
            )
        excluded_std_samples = self._select_legacy_std_set(
            batch.samples, method, profile, result
        )

        for sample in batch.samples:
            if sample.sample_id in excluded_std_samples:
                for peak in sample.peaks:
                    result.rows.append(
                        self._row_for_peak(
                            sample,
                            peak,
                            status=ExcelPreviewStatus.EXCLUDED,
                            exclude_reason=ExcludeReason.DUPLICATE_STD_SET.value,
                            message="최초 완전 STD 세트 이후의 추가 STD 재확인 세트",
                        )
                    )
                continue
            self._map_sample(
                sample,
                method,
                snapshot,
                profile,
                worker_rows,
                target_analysis_numbers,
                result,
            )
        self._detect_target_collisions(result)
        return result

    @classmethod
    def _select_legacy_std_set(
        cls,
        samples: list[Sample],
        method: StdMethod,
        profile: TemplateProfile,
        result: ExcelPreviewResult,
    ) -> set[UUID]:
        """Select one complete STD run when a profile can contain recheck sets.

        혼유 and 디에틸에테르 source PDFs can contain a complete calibration run
        followed by partial recheck standards. Other profiles retain their
        established behaviour. A single isolated STD is left alone so focused
        review/test batches remain valid.
        """
        if profile not in (LEGACY_PROFILE, DIETHYL_ETHER_PROFILE):
            return set()

        std_samples = [sample for sample in samples if sample.sample_type is SampleType.STD]
        replicate_counts: dict[int, int] = defaultdict(int)
        for sample in std_samples:
            if sample.replicate_no is not None:
                replicate_counts[sample.replicate_no] += 1
        needs_set_selection = len(std_samples) > 1 or any(
            count > 1 for count in replicate_counts.values()
        )
        if not needs_set_selection:
            return set()

        required = cls._std_replicates(profile, method)
        runs: list[list[Sample]] = []
        current: list[Sample] = []
        for sample in samples:
            if sample.sample_type is SampleType.STD:
                if (
                    current
                    and sample.replicate_no is not None
                    and current[-1].replicate_no is not None
                    and sample.replicate_no <= current[-1].replicate_no
                ):
                    runs.append(current)
                    current = []
                current.append(sample)
            elif current:
                runs.append(current)
                current = []
        if current:
            runs.append(current)

        complete_sequences = {(1, 2, 3, 4, 5), (1, 2, 3, 4, 5, 6)}
        complete_runs = [
            run
            for run in runs
            if tuple(sample.replicate_no for sample in run) in complete_sequences
            and set(required).issubset(
                {sample.replicate_no for sample in run}
            )
        ]
        if len(complete_runs) == 1:
            selected_ids = {sample.sample_id for sample in complete_runs[0]}
            return {
                sample.sample_id
                for sample in std_samples
                if sample.sample_id not in selected_ids
            }

        code = "STD_SET_AMBIGUOUS" if len(complete_runs) > 1 else "STD_SET_INCOMPLETE"
        message = (
            "완전한 STD 세트가 2개 이상이므로 기본 검량선 세트를 자동 선택할 수 없습니다."
            if len(complete_runs) > 1
            else f"연속된 완전 STD 세트({', '.join(f'STD{n}' for n in required)})가 없습니다."
        )
        result.issues.append(
            ExcelPreviewIssue(ValidationSeverity.ERROR, code, message)
        )
        return set()

    @staticmethod
    def _template_profile(
        snapshot: ExcelTemplateSnapshot, result: ExcelPreviewResult
    ) -> TemplateProfile | None:
        if ACN_PROFILE.area_sheet in snapshot.sheet_names:
            single_header = (
                str(snapshot.cell(ACN_PROFILE.area_sheet, "F2").value or "")
                .strip()
                .casefold()
            )
            if single_header == "acetonitrile":
                profile = ACN_PROFILE
            elif single_header == "2-부톡시에탄올":
                profile = BC_PROFILE
            else:
                result.issues.append(
                    ExcelPreviewIssue(
                        ValidationSeverity.ERROR,
                        "TEMPLATE_PROFILE_UNSUPPORTED",
                        "결과입력(area입력) 시트의 F2 물질 헤더가 지원 대상이 아닙니다.",
                    )
                )
                return None
        elif MEK_PROFILE.area_sheet in snapshot.sheet_names:
            acetic_header = (
                str(snapshot.cell(MEK_PROFILE.area_sheet, "E2").value or "")
                .strip()
                .casefold()
            )
            lod_headers = tuple(
                str(snapshot.cell(MEK_PROFILE.area_sheet, address).value or "")
                .strip()
                .casefold()
                for address in ("F3", "I3")
            )
            if acetic_header == "acetic acid":
                profile = ACETIC_ACID_PROFILE
            elif acetic_header == "ethylene glycol":
                profile = ETHYLENE_GLYCOL_PROFILE
            elif acetic_header == "diethyl ether":
                profile = DIETHYL_ETHER_PROFILE
            elif lod_headers == ("isoamyl acetate", "n-propyl acetae"):
                profile = ISOAMYL_N_PROPYL_ACETATE_PROFILE
            else:
                profile = MEK_PROFILE
        elif ONE_COLUMN_PROFILE.area_sheet in snapshot.sheet_names:
            headers = tuple(
                str(snapshot.cell(ONE_COLUMN_PROFILE.area_sheet, address).value or "")
                .strip()
                .casefold()
                for address in ("F3", "I3", "L3", "O3")
            )
            compact_headers = tuple(re.sub(r"\s+", "", header) for header in headers)
            g3_headers = tuple(
                re.sub(
                    r"\s+",
                    "",
                    str(snapshot.cell(ONE_COLUMN_PROFILE.area_sheet, address).value or "")
                    .strip()
                    .casefold(),
                )
                for address in ("F3", "I3", "L3", "O3", "R3")
            )
            if g3_headers == (
                "1,2-dichloroethylene",
                "trichloroethylene",
                "tetrachloroethylene",
                "1,2-dichloropropane",
                "1,2-dichloroethane",
            ):
                profile = G3_PROFILE
            elif compact_headers == (
                "2-butoxyethanol(egbe)",
                "2-butoxyethylacetate(egbea)",
                "2-ethoxyethanol(egee)",
                "2-ethoxyethylacetate(egeea)",
            ):
                profile = CELLOSOLVE_PROFILE
            elif headers == ("thf", "cfm", "benzene", "chlorobenzene"):
                profile = G2_PROFILE
            else:
                iba_header, btoh_header = headers[:2]
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
                    "지원하는 Excel 양식이 아닙니다. area, area입력 또는 "
                    "LOD(area입력) 시트가 필요합니다.",
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
        return extract_excel_analysis_number(value)

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
        target_analysis_numbers: set[str],
        result: ExcelPreviewResult,
    ) -> None:
        number_decision = classify_sample_number(
            sample.sample_name_normalized,
            sample.sample_type,
            is_blank=sample.is_blank,
        )
        # A sample containing only unnamed/excluded peaks has nothing that can
        # be written.  Do not turn a missing template row into a validation
        # error before applying those established peak-exclusion rules.
        if not any(
            peak.include_for_excel and peak.material_standard not in {None, "CS2"}
            for peak in sample.peaks
        ):
            for peak in sample.peaks:
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
            return
        if (
            sample.sample_type not in {SampleType.STD, SampleType.RECOVERY}
            and number_decision.analysis_number is not None
            and str(int(number_decision.analysis_number))
            not in target_analysis_numbers
        ):
            for peak in sample.peaks:
                result.rows.append(
                    self._row_for_peak(
                        sample,
                        peak,
                        status=ExcelPreviewStatus.EXCLUDED,
                        exclude_reason=ExcludeReason.NON_TARGET_SAMPLE.value,
                        message=(
                            f"분석번호 {number_decision.analysis_number}는 현재 배치 범위와 "
                            "Excel 분석번호 목록에 포함되지 않습니다."
                        ),
                    )
                )
            return
        row = self._sample_target_row(
            sample, method, profile, worker_rows, result, number_decision
        )
        if row is None:
            reason = self._sample_exclusion_reason(sample, method, profile)
            if (
                number_decision.exclude_reason is not None
                and sample.sample_type not in {SampleType.STD, SampleType.RECOVERY}
            ):
                reason = number_decision.exclude_reason
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
            elif profile == MEK_PROFILE:
                target_rt = MEK_TARGET_RETENTION_TIMES.get(material)
            elif profile == ACETIC_ACID_PROFILE:
                target_rt = ACETIC_ACID_TARGET_RETENTION_TIMES.get(material)
            elif profile == ETHYLENE_GLYCOL_PROFILE:
                target_rt = ETHYLENE_GLYCOL_TARGET_RETENTION_TIMES.get(material)
            elif profile == DIETHYL_ETHER_PROFILE:
                target_rt = DIETHYL_ETHER_TARGET_RETENTION_TIMES.get(material)
            elif profile == ACN_PROFILE:
                target_rt = ACN_TARGET_RETENTION_TIMES.get(material)
            elif profile == BC_PROFILE:
                target_rt = BC_TARGET_RETENTION_TIMES.get(material)
            elif profile == G2_PROFILE:
                target_rt = G2_TARGET_RETENTION_TIMES.get(material)
            elif profile == ISOAMYL_N_PROPYL_ACETATE_PROFILE:
                target_rt = ISOAMYL_N_PROPYL_ACETATE_TARGET_RETENTION_TIMES.get(material)
            elif profile == CELLOSOLVE_PROFILE:
                target_rt = CELLOSOLVE_TARGET_RETENTION_TIMES.get(material)
            elif profile == G3_PROFILE:
                target_rt = G3_TARGET_RETENTION_TIMES.get(material)
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
        number_decision: SampleNumberDecision,
    ) -> int | None:
        if sample.sample_type is SampleType.STD:
            selected = self._std_replicates(profile, method)
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
        key = number_decision.analysis_number
        if key is not None:
            matches = worker_rows.get(str(key), []) if key else []
            if not matches:
                self._sample_error(
                    result,
                    sample,
                    "WORKER_ROW_NOT_FOUND",
                    f"분석번호 {key}와 일치하는 {profile.area_sheet} 분석번호 행이 없습니다.",
                )
                return None
            if len(matches) > 1:
                self._sample_error(
                    result,
                    sample,
                    "WORKER_ROW_NOT_UNIQUE",
                    f"분석번호 {key}와 일치하는 {profile.area_sheet} 분석번호 행: {matches}",
                )
                return None
            return matches[0]
        return None

    @staticmethod
    def _std_replicates(
        profile: TemplateProfile, method: StdMethod
    ) -> tuple[int, ...]:
        if profile.std_replicates_override:
            return profile.std_replicates_override
        return (
            (1, 2, 3, 4, 5)
            if profile in (MEK_PROFILE, ETHYLENE_GLYCOL_PROFILE, BC_PROFILE)
            else STD_REPLICATES[method]
        )

    @classmethod
    def _sample_exclusion_reason(
        cls, sample: Sample, method: StdMethod, profile: TemplateProfile
    ) -> str | None:
        if sample.sample_type is SampleType.STD:
            selected = cls._std_replicates(profile, method)
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
