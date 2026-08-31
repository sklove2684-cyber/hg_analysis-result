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
    worker_column: str = "A"
    dibk_std_slots: tuple[str, ...] = ()
    dibk_recovery_slots: tuple[str, ...] = ()
    use_runtime_std_rt: bool = False
    worker_key_mode: str = "last"
    special_kind: str | None = None


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
    use_runtime_std_rt=True,
)

IPA_PROFILE = TemplateProfile(
    key="ipa",
    name="IPA",
    required_sheets=("검량선", "LOD(area입력)", "회수율", "std"),
    area_sheet="LOD(area입력)",
    std_columns={"Isopropyl alcohol": "J"},
    numeric_columns={"Isopropyl alcohol": "J"},
    recovery_columns={"Isopropyl alcohol": "B"},
    std_row_start=5,
    recovery_row_start={
        ConcentrationLevel.LOW: 28,
        ConcentrationLevel.MID: 31,
        ConcentrationLevel.HIGH: 34,
    },
    worker_row_start=20,
    worker_row_end=128,
    worker_column="E",
    use_runtime_std_rt=True,
)


IPA_AREA_PROFILE = TemplateProfile(
    key="ipa",
    name="IPA",
    required_sheets=("검량선", "area", "회수율", "std"),
    area_sheet="area",
    std_columns={"Isopropyl alcohol": "J"},
    numeric_columns={"Isopropyl alcohol": "J"},
    recovery_columns={"Isopropyl alcohol": "B"},
    std_row_start=5,
    recovery_row_start={
        ConcentrationLevel.LOW: 28,
        ConcentrationLevel.MID: 31,
        ConcentrationLevel.HIGH: 34,
    },
    worker_row_start=20,
    worker_row_end=46,
    worker_column="E",
    use_runtime_std_rt=True,
)


METHANOL_PROFILE = TemplateProfile(
    key="methanol",
    name="메탄올A",
    required_sheets=("검량선", "LOD(area입력)", "회수율", "std"),
    area_sheet="LOD(area입력)",
    std_columns={"Methanol": "F"},
    numeric_columns={"Methanol": "F"},
    recovery_columns={"Methanol": "B"},
    std_row_start=4,
    recovery_row_start={
        ConcentrationLevel.LOW: 28,
        ConcentrationLevel.MID: 31,
        ConcentrationLevel.HIGH: 34,
    },
    worker_row_start=19,
    worker_row_end=80,
    use_runtime_std_rt=True,
)


PHENOL_PROFILE = TemplateProfile(
    key="phenol",
    name="페놀",
    required_sheets=("검량선", "LOD(area입력)", "회수율", "std"),
    area_sheet="LOD(area입력)",
    std_columns={"Phenol": "F"},
    numeric_columns={"Phenol": "F"},
    recovery_columns={"Phenol": "B"},
    std_row_start=4,
    recovery_row_start={
        ConcentrationLevel.LOW: 28,
        ConcentrationLevel.MID: 31,
        ConcentrationLevel.HIGH: 34,
    },
    worker_row_start=19,
    worker_row_end=122,
    use_runtime_std_rt=True,
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

DMF_DMA_PROFILE = TemplateProfile(
    key="dmf_dma",
    name="DMF,DMA",
    required_sheets=("검량선", "LOD(area입력)", "회수율", "std"),
    area_sheet="LOD(area입력)",
    std_columns={"DMF": "G", "DMA": "J"},
    numeric_columns={"DMF": "F", "DMA": "I"},
    recovery_columns={"DMF": "B", "DMA": "C"},
    std_row_start=5,
    recovery_row_start={
        ConcentrationLevel.LOW: 28,
        ConcentrationLevel.MID: 31,
        ConcentrationLevel.HIGH: 34,
    },
    worker_row_start=20,
    worker_row_end=123,
    use_runtime_std_rt=True,
)

ALCOHOL_4_PROFILE = TemplateProfile(
    key="alcohol_4",
    name="알콜4",
    required_sheets=("검량선", "area입력", "회수율", "STD제조"),
    area_sheet="area입력",
    std_columns={"IBA": "G", "n-BTOH": "J", "IAA": "M", "2-BTOH": "P"},
    numeric_columns={"IBA": "F", "n-BTOH": "I", "IAA": "L", "2-BTOH": "O"},
    recovery_columns={"IBA": "B", "n-BTOH": "C", "IAA": "D", "2-BTOH": "E"},
    std_row_start=5,
    recovery_row_start={
        ConcentrationLevel.LOW: 30,
        ConcentrationLevel.MID: 33,
        ConcentrationLevel.HIGH: 36,
    },
    worker_row_start=21,
    worker_row_end=287,
    use_runtime_std_rt=True,
)

PROPYLENE_OXIDE_PROFILE = TemplateProfile(
    key="propylene_oxide",
    name="1,2-에폭시프로판(산화프로필렌)",
    required_sheets=("검량선", "결과입력(area입력)", "회수율", "STD제조"),
    area_sheet="결과입력(area입력)",
    std_columns={"Propylene oxide": "F"},
    numeric_columns={"Propylene oxide": "F"},
    recovery_columns={"Propylene oxide": "B"},
    std_row_start=14,
    recovery_row_start={
        ConcentrationLevel.LOW: 29,
        ConcentrationLevel.MID: 32,
        ConcentrationLevel.HIGH: 35,
    },
    worker_row_start=26,
    worker_row_end=79,
    use_runtime_std_rt=True,
)

DICHLOROMETHANE_PROFILE = TemplateProfile(
    key="dichloromethane",
    name="디클로로메탄(MC)",
    required_sheets=("검량선", "LOD(area입력)", "회수율", "std"),
    area_sheet="LOD(area입력)",
    std_columns={"Dichloromethane": "F"},
    numeric_columns={"Dichloromethane": "F"},
    recovery_columns={"Dichloromethane": "B"},
    std_row_start=4,
    recovery_row_start={
        ConcentrationLevel.LOW: 28,
        ConcentrationLevel.MID: 31,
        ConcentrationLevel.HIGH: 34,
    },
    worker_row_start=19,
    worker_row_end=122,
    use_runtime_std_rt=True,
    worker_key_mode="compound",
)

METHYL_N_AMYL_KETONE_PROFILE = TemplateProfile(
    key="methyl_n_amyl_ketone",
    name="메틸 n아밀케톤",
    required_sheets=("검량선", "LOD(area입력)", "회수율", "std"),
    area_sheet="LOD(area입력)",
    std_columns={"메틸 n-아밀케톤": "E"},
    numeric_columns={"메틸 n-아밀케톤": "E"},
    recovery_columns={"메틸 n-아밀케톤": "B"},
    std_row_start=4,
    recovery_row_start={
        ConcentrationLevel.LOW: 28,
        ConcentrationLevel.MID: 31,
        ConcentrationLevel.HIGH: 34,
    },
    worker_row_start=19,
    worker_row_end=122,
    use_runtime_std_rt=True,
)

VINYL_ACETATE_PROFILE = TemplateProfile(
    key="vinyl_acetate",
    name="비닐아세테이트",
    required_sheets=("검량선", "LOD(area입력)", "회수율", "std"),
    area_sheet="LOD(area입력)",
    std_columns={"Vinyl acetate": "F"},
    numeric_columns={"Vinyl acetate": "F"},
    recovery_columns={"Vinyl acetate": "B"},
    std_row_start=4,
    recovery_row_start={
        ConcentrationLevel.LOW: 28,
        ConcentrationLevel.MID: 31,
        ConcentrationLevel.HIGH: 34,
    },
    worker_row_start=20,
    worker_row_end=122,
    use_runtime_std_rt=True,
)

ISOPROPYL_ACETATE_PROFILE = TemplateProfile(
    key="isopropyl_acetate",
    name="이소프로필 아세테이트",
    required_sheets=("검량선", "LOD(area입력)", "회수율", "std"),
    area_sheet="LOD(area입력)",
    std_columns={"Isopropyl acetate": "E"},
    numeric_columns={"Isopropyl acetate": "E"},
    recovery_columns={"Isopropyl acetate": "B"},
    std_row_start=4,
    recovery_row_start={
        ConcentrationLevel.LOW: 28,
        ConcentrationLevel.MID: 31,
        ConcentrationLevel.HIGH: 34,
    },
    worker_row_start=20,
    worker_row_end=122,
    use_runtime_std_rt=True,
)

PYRIDINE_PROFILE = TemplateProfile(
    key="pyridine",
    name="피리딘",
    required_sheets=("검량선", "LOD(area입력)", "회수율", "std"),
    area_sheet="LOD(area입력)",
    std_columns={"Pyridine": "F"},
    numeric_columns={"Pyridine": "F"},
    recovery_columns={"Pyridine": "B"},
    std_row_start=4,
    recovery_row_start={
        ConcentrationLevel.LOW: 28,
        ConcentrationLevel.MID: 31,
        ConcentrationLevel.HIGH: 34,
    },
    worker_row_start=19,
    worker_row_end=122,
    use_runtime_std_rt=True,
)

STODDARD_SOLVENT_PROFILE = TemplateProfile(
    key="stoddard_solvent",
    name="스토다드솔벤트",
    required_sheets=("검량선", "LOD(area입력)", "회수율", "std"),
    area_sheet="LOD(area입력)",
    std_columns={"Stoddard total": "F", "Stoddard CS2": "G", "Stoddard residual": "H"},
    numeric_columns={"Stoddard solvent": "E"},
    recovery_columns={"Stoddard total": "D", "Stoddard CS2": "E", "Stoddard residual": "F"},
    std_row_start=4,
    recovery_row_start={
        ConcentrationLevel.LOW: 28,
        ConcentrationLevel.MID: 31,
        ConcentrationLevel.HIGH: 34,
    },
    worker_row_start=19,
    worker_row_end=122,
    special_kind="stoddard",
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
        expected_profile_key = excel_profile_key_for(batch.analysis_type)
        profile = self._template_profile(snapshot, result, expected_profile_key)
        if profile is None:
            return result
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
        runtime_target_retention_times = self._runtime_target_retention_times(
            batch.samples, excluded_std_samples, method, profile
        )
        if profile.use_runtime_std_rt:
            missing_runtime_materials = [
                material
                for material in profile.std_columns
                if material not in runtime_target_retention_times
                and any(
                    peak.include_for_excel and peak.material_standard == material
                    for sample in batch.samples
                    for peak in sample.peaks
                )
            ]
            for material in missing_runtime_materials:
                result.issues.append(
                    ExcelPreviewIssue(
                        ValidationSeverity.ERROR,
                        "STD_TARGET_RT_NOT_FOUND",
                        f"선택된 STD 세트에서 {material} 기준 RT를 확인할 수 없습니다.",
                    )
                )

        stoddard_cs2_rt = (
            self._stoddard_cs2_reference_rt(batch.samples)
            if profile.special_kind == "stoddard"
            else None
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
                runtime_target_retention_times,
                result,
                stoddard_cs2_rt=stoddard_cs2_rt,
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

        A source PDF can contain a complete calibration run followed by partial
        recheck standards.  Set selection is only activated when a replicate
        number repeats, so ordinary STD1~5/6 runs retain their normal behaviour.
        """
        std_samples = [sample for sample in samples if sample.sample_type is SampleType.STD]
        replicate_counts: dict[int, int] = defaultdict(int)
        for sample in std_samples:
            if sample.replicate_no is not None:
                replicate_counts[sample.replicate_no] += 1
        needs_set_selection = any(count > 1 for count in replicate_counts.values())
        if not needs_set_selection:
            required = cls._std_replicates(profile, method)
            available = {
                sample.replicate_no
                for sample in std_samples
                if sample.replicate_no is not None
            }
            if (
                len(std_samples) >= 2
                and len(std_samples) == len(samples)
                and not set(required).issubset(available)
            ):
                result.issues.append(
                    ExcelPreviewIssue(
                        ValidationSeverity.ERROR,
                        "STD_SET_INCOMPLETE",
                        f"완전한 STD 세트({', '.join(f'STD{n}' for n in required)})가 없습니다.",
                    )
                )
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

    @classmethod
    def _runtime_target_retention_times(
        cls,
        samples: list[Sample],
        excluded_sample_ids: set[UUID],
        method: StdMethod,
        profile: TemplateProfile,
    ) -> dict[str, Decimal]:
        """Use observed STD RTs directly; do not average or infer by Area."""
        if not profile.use_runtime_std_rt:
            return {}
        result: dict[str, Decimal] = {}
        for replicate_no in cls._std_replicates(profile, method):
            for sample in samples:
                if (
                    sample.sample_type is not SampleType.STD
                    or sample.sample_id in excluded_sample_ids
                    or sample.replicate_no != replicate_no
                ):
                    continue
                for material in profile.std_columns:
                    if material in result:
                        continue
                    peaks = [
                        peak
                        for peak in sample.peaks
                        if peak.include_for_excel
                        and peak.material_standard == material
                    ]
                    if len(peaks) == 1:
                        result[material] = peaks[0].retention_time
        return result

    @staticmethod
    def _template_profile(
        snapshot: ExcelTemplateSnapshot,
        result: ExcelPreviewResult,
        expected_profile_key: str | None = None,
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
            elif single_header == "1,2-epoxypropane":
                profile = PROPYLENE_OXIDE_PROFILE
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
            primary_header = (
                str(snapshot.cell(MEK_PROFILE.area_sheet, "E2").value or "")
                .strip()
                .casefold()
            )
            alternate_header = (
                str(snapshot.cell(MEK_PROFILE.area_sheet, "D2").value or "")
                .strip()
                .casefold()
            )
            lod_headers = tuple(
                str(snapshot.cell(MEK_PROFILE.area_sheet, address).value or "")
                .strip()
                .casefold()
                for address in ("F3", "I3")
            )
            ipa_headers = tuple(
                str(snapshot.cell(MEK_PROFILE.area_sheet, address).value or "")
                .strip()
                .casefold()
                for address in ("I3", "J3")
            )
            if primary_header == "acetic acid":
                profile = ACETIC_ACID_PROFILE
            elif primary_header == "ethylene glycol":
                profile = ETHYLENE_GLYCOL_PROFILE
            elif primary_header == "methanol":
                profile = METHANOL_PROFILE
            elif primary_header == "phenol":
                profile = PHENOL_PROFILE
            elif primary_header == "diethyl ether":
                profile = DIETHYL_ETHER_PROFILE
            elif primary_header == "methylene chloride":
                profile = DICHLOROMETHANE_PROFILE
            elif primary_header == "vinyl acetate":
                profile = VINYL_ACETATE_PROFILE
            elif primary_header == "pyridine":
                profile = PYRIDINE_PROFILE
            elif ipa_headers == ("ipa", "area"):
                profile = IPA_PROFILE
            elif alternate_header == "메틸 n-아밀케톤":
                profile = METHYL_N_AMYL_KETONE_PROFILE
            elif alternate_header == "isopropyl acetate":
                profile = ISOPROPYL_ACETATE_PROFILE
            elif alternate_header == "stoddard solvent":
                profile = STODDARD_SOLVENT_PROFILE
            elif lod_headers == ("dmf", "dma"):
                profile = DMF_DMA_PROFILE
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
            elif headers == ("iba", "1-btoh", "iaa", "2-btoh"):
                # Alcohol-2 and Alcohol-4 source workbooks share this physical
                # four-column layout.  The selected analysis registry entry is
                # the only reliable discriminator; Alcohol-2 intentionally
                # writes only IBA/n-BTOH and preserves the other two columns.
                profile = (
                    ALCOHOL_PROFILE
                    if expected_profile_key == ALCOHOL_PROFILE.key
                    else ALCOHOL_4_PROFILE
                )
            else:
                iba_header, btoh_header = headers[:2]
                if iba_header == "iba" and btoh_header in {"1-btoh", "n-btoh"}:
                    profile = ALCOHOL_PROFILE
                else:
                    profile = ONE_COLUMN_PROFILE
        elif LEGACY_PROFILE.area_sheet in snapshot.sheet_names:
            area_headers = tuple(
                str(snapshot.cell(LEGACY_PROFILE.area_sheet, address).value or "")
                .strip()
                .casefold()
                for address in ("I3", "J3")
            )
            # The original IPA workbook uses the same physical ``area`` sheet
            # name as the mixture workbook.  Its material/header pair is the
            # reliable discriminator and must take precedence over the generic
            # legacy-mixture fallback.
            profile = (
                IPA_AREA_PROFILE
                if area_headers == ("ipa", "area")
                else LEGACY_PROFILE
            )
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

    @classmethod
    def _profile_worker_key(
        cls, value: object | None, profile: TemplateProfile
    ) -> str | None:
        if profile.worker_key_mode == "compound" and value is not None:
            match = re.search(r"(\d+)\s*-\s*(\d+)\s*$", str(value).strip())
            if match:
                return f"{match.group(1)}-{match.group(2)}"
        return cls._worker_key(value)

    @staticmethod
    def _sample_number_decision(
        sample: Sample, profile: TemplateProfile
    ) -> SampleNumberDecision:
        if (
            profile.worker_key_mode == "compound"
            and sample.sample_type not in {SampleType.BLANK, SampleType.RECOVERY_BLANK}
        ):
            match = re.fullmatch(
                r"(\d+)\((\d+)\)", re.sub(r"\s+", "", sample.sample_name_raw)
            )
            if match:
                return SampleNumberDecision(f"{match.group(1)}-{match.group(2)}")
        return classify_sample_number(
            sample.sample_name_normalized,
            sample.sample_type,
            is_blank=sample.is_blank,
        )

    def _stoddard_cs2_reference_rt(
        self, samples: list[Sample]
    ) -> Decimal | None:
        for sample_type in (SampleType.BLANK, SampleType.STD):
            for sample in samples:
                if sample.sample_type is not sample_type:
                    continue
                unnamed = [peak for peak in sample.peaks if peak.material_raw is None]
                if unnamed:
                    return max(
                        unnamed,
                        key=lambda peak: (self._applied_area(peak), -peak.peak_no),
                    ).retention_time
        return None

    def _map_stoddard_sample(
        self,
        sample: Sample,
        snapshot: ExcelTemplateSnapshot,
        profile: TemplateProfile,
        row: int,
        cs2_reference_rt: Decimal | None,
        result: ExcelPreviewResult,
    ) -> None:
        if cs2_reference_rt is None:
            self._sample_error(
                result,
                sample,
                "STODDARD_CS2_RT_NOT_FOUND",
                "BLANK 또는 STD에서 스토다드솔벤트 CS2 기준 RT를 확인할 수 없습니다.",
            )
            return
        candidates = [peak for peak in sample.peaks if peak.material_raw is None]
        if not candidates:
            self._sample_error(
                result,
                sample,
                "STODDARD_CS2_PEAK_NOT_FOUND",
                "CS2 기준 RT에 대응하는 이름 없는 Peak가 없습니다.",
            )
            return
        cs2_peak = min(
            candidates,
            key=lambda peak: (
                abs(peak.retention_time - cs2_reference_rt),
                peak.peak_no,
            ),
        )
        total = (
            sample.total_area
            if sample.total_area is not None
            else sum(self._applied_area(peak) for peak in sample.peaks)
        )
        residual = sum(
            self._applied_area(peak)
            for peak in sample.peaks
            if peak is not cs2_peak and peak.material_standard != "Stoddard solvent"
        )
        solvent = total - self._applied_area(cs2_peak) - residual

        for peak in sample.peaks:
            result.rows.append(
                self._row_for_peak(
                    sample,
                    peak,
                    status=ExcelPreviewStatus.EXCLUDED,
                    exclude_reason="STODDARD_COMPONENT_SOURCE",
                    message="Total/CS2/제외 계산의 원본 Peak",
                )
            )

        values = (
            (("Stoddard solvent", solvent),)
            if sample.sample_type in {SampleType.NUMERIC, SampleType.UNKNOWN}
            else (
                ("Stoddard total", total),
                ("Stoddard CS2", self._applied_area(cs2_peak)),
                ("Stoddard residual", residual),
            )
        )
        for offset, (material, area) in enumerate(values, start=1):
            synthetic = Peak(
                peak_no=max((peak.peak_no for peak in sample.peaks), default=0) + offset,
                retention_time=cs2_peak.retention_time,
                area_raw=area,
                material_raw=material,
                material_standard=material,
                include_for_excel=True,
                source_page=sample.page_no,
            )
            column = self._material_column(profile, sample, material)
            if column is None:
                self._append_error_row(
                    result,
                    sample,
                    synthetic,
                    "UNSUPPORTED_MATERIAL",
                    f"스토다드솔벤트 Excel 입력 열이 없습니다: {material}",
                )
                continue
            if (
                sample.sample_type in {SampleType.NUMERIC, SampleType.UNKNOWN}
                and material == "Stoddard solvent"
                and area == 0
            ):
                sheet = self._target_sheet(profile, sample)
                address = f"{column}{row}"
                cell = snapshot.cell(sheet, address)
                result.rows.append(
                    self._row_for_peak(
                        sample,
                        synthetic,
                        target_sheet=sheet,
                        target_cell=address,
                        existing_value_type=cell.value_type,
                        existing_has_formula=cell.has_formula,
                        status=ExcelPreviewStatus.EXCLUDED,
                        exclude_reason=ExcludeReason.STODDARD_ND_PRESERVED.value,
                        message="계산 Area가 0이므로 원본 Excel의 N.D 셀을 수정하지 않음",
                    )
                )
                continue
            self._append_mapped_row(
                result, snapshot, profile, sample, synthetic, column, row
            )

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
            analysis_cell = snapshot.cell(
                profile.area_sheet, f"{profile.worker_column}{row}"
            )
            # Some templates contain a second, calculated table that mirrors the
            # input table.  Its analysis numbers are formulas and must not be
            # treated as writable worker rows.
            if analysis_cell.has_formula:
                continue
            key = self._profile_worker_key(analysis_cell.value, profile)
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
        runtime_target_retention_times: dict[str, Decimal],
        result: ExcelPreviewResult,
        *,
        stoddard_cs2_rt: Decimal | None = None,
    ) -> None:
        number_decision = self._sample_number_decision(sample, profile)
        # A sample containing only unnamed/excluded peaks has nothing that can
        # be written.  Do not turn a missing template row into a validation
        # error before applying those established peak-exclusion rules.
        if profile.special_kind != "stoddard" and not any(
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
            and number_decision.analysis_number not in target_analysis_numbers
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

        if profile.special_kind == "stoddard":
            self._map_stoddard_sample(
                sample,
                snapshot,
                profile,
                row,
                stoddard_cs2_rt,
                result,
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
            target_rt = runtime_target_retention_times.get(material)
            if target_rt is not None:
                pass
            elif profile == ONE_COLUMN_PROFILE:
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
                target_rt = None
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
            if target_rt is None and profile.use_runtime_std_rt:
                for candidate in peaks:
                    result.rows.append(
                        self._row_for_peak(
                            sample,
                            candidate,
                            status=ExcelPreviewStatus.ERROR,
                            message=f"STD에서 {material} 기준 RT를 확인할 수 없습니다.",
                        )
                    )
                continue
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
