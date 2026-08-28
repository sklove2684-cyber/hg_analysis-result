from __future__ import annotations

from dataclasses import dataclass
import re


@dataclass(frozen=True)
class MaterialDefinition:
    canonical_name: str
    aliases: tuple[str, ...]


@dataclass(frozen=True)
class AnalysisTypeDefinition:
    key: str
    display_name: str
    supported_materials: tuple[MaterialDefinition, ...] = ()
    excel_profile_key: str | None = None
    materials_pending: bool = False
    allow_runtime_single_material_inference: bool = False


def _material(canonical_name: str, *aliases: str) -> MaterialDefinition:
    return MaterialDefinition(canonical_name, aliases)


MIXTURE_MATERIALS = (
    _material("n-hexane", "헥산", "n-hexane"),
    _material("acetone", "아세톤", "acetone"),
    _material("E.A", "e.a"),
    _material("MIBK", "mibk"),
    _material("Toluene", "tol", "toluene"),
    _material("B.A", "b.a"),
    _material("E.B", "e.b"),
    _material("p-xylene", "p", "p-xylene"),
    _material("m-xylene", "m", "m-xylene"),
    _material("o-xylene", "o", "o-xylene"),
    _material("styrene", "스티렌", "styrene"),
    _material("c-hexanone", "시클로헥사논", "c-hexanone"),
    _material("DIBK", "dibk"),
)
ONE_COLUMN_MATERIALS = (
    _material("methyl acetate", "메틸아세테이트", "초산메틸", "methyl acetate", "methyl acatate"),
    _material("c-hexane", "시클로헥산", "사이클로헥산", "cyclohexane", "c-hexane"),
    _material("n-heptane", "헵탄", "n-헵탄", "heptane", "n-heptane"),
    _material("isobutyl acetate", "이소부틸아세테이트", "초산이소부틸", "isobutyl acetate", "isobutyl acetae"),
)
ALCOHOL_2_MATERIALS = (
    _material("IBA", "iba"),
    _material("n-BTOH", "n-btoh"),
)
MEK_MATERIALS = (
    _material("methyl ethyl ketone", "mek", "methyl ethyl ketone", "메틸에틸케톤", "2-butanone"),
)
IPA_MATERIALS = (
    _material("Isopropyl alcohol", "ipa", "isopropyl alcohol"),
)
METHANOL_MATERIALS = (
    _material("Methanol", "메탄올", "메탄올A", "methanol", "MeOH"),
)
ACN_MATERIALS = (
    _material("Acetonitrile", "acn", "acetonitrile"),
)
PHENOL_MATERIALS = (
    _material("Phenol", "페놀", "phenol"),
)
PROPYLENE_OXIDE_MATERIALS = (
    _material(
        "Propylene oxide",
        "1,2에폭시프로판",
        "1,2-에폭시프로판",
        "산화프로필렌",
        "propylene oxide",
    ),
)
DIETHYL_ETHER_MATERIALS = (
    _material("Diethyl ether", "디에틸에테르", "diethyl ether"),
)
DICHLOROMETHANE_MATERIALS = (
    _material(
        "Dichloromethane",
        "디클로로메탄",
        "M.C",
        "MC",
        "dichloromethane",
        "Methylene chloride",
    ),
)
CHLOROMETHANE_MATERIALS = (
    _material("Chloromethane", "메틸클로라이드", "chloromethane"),
)
VINYL_ACETATE_MATERIALS = (
    _material("Vinyl acetate", "비닐아세테이트", "vinyl acetate"),
)
ISOPROPYL_ACETATE_MATERIALS = (
    _material(
        "Isopropyl acetate",
        "초산이소프로필",
        "이소프로필 아세테이트",
        "isopropyl acetate",
    ),
)
ACETIC_ACID_MATERIALS = (
    _material("초산", "초산", "acetic acid"),
)
PYRIDINE_MATERIALS = (
    _material("Pyridine", "피리딘", "pyridine"),
)
ETHYLENE_GLYCOL_MATERIALS = (
    _material("Ethylene glycol", "에틸렌글리콜", "Ethylene glycol", "E.G"),
)
G2_MATERIALS = (
    _material("THF", "thf"),
    _material("CFM", "cfm"),
    _material("벤젠", "벤젠"),
    _material("클로로벤젠", "클로로벤젠"),
)
ISOAMYL_N_PROPYL_ACETATE_MATERIALS = (
    _material("n-프로필 아세테이트", "초산프로필"),
    _material("이소아밀 아세테이트", "초산이소아밀"),
)
CELLOSOLVE_MATERIALS = (
    _material("2-Butoxyethanol", "부톡시에탄올(BC)", "2-부톡시에탄올", "BC"),
    _material("2-Butoxyethyl acetate", "부톡시에틸아세테이트"),
    _material("2-Ethoxyethanol", "에톡시에탄올"),
    _material("2-Ethoxyethyl acetate", "에톡시에틸아세테이트"),
)
BC_MATERIALS = (CELLOSOLVE_MATERIALS[0],)
TETRACHLOROETHYLENE_MATERIAL = _material(
    "Tetrachloroethylene", "pce", "테트라클로로에틸렌", "tetrachloroethylene"
)
G3_MATERIALS = (
    _material("1,2-Dichloroethylene", "1,2디클로로에틸렌"),
    _material("Trichloroethylene", "tce"),
    TETRACHLOROETHYLENE_MATERIAL,
    _material("1,2-Dichloropropane", "1,2디클로로프로판"),
    _material("1,2-Dichloroethane", "1,2디클로로에탄"),
)
DMF_DMA_MATERIALS = (
    _material("DMF", "DMF"),
    _material("DMA", "DMA"),
)
METHYL_N_AMYL_KETONE_MATERIALS = (
    _material(
        "메틸 n-아밀케톤",
        "메틸 n아밀케톤",
        "메틸 n-아밀케톤",
    ),
)
ALCOHOL_4_MATERIALS = (
    _material("IBA", "IBA"),
    _material("n-BTOH", "n-부탄올", "1-BTOH", "n-BTOH"),
    _material("IAA", "IAA"),
    _material("2-BTOH", "2-부탄올", "2-BTOH"),
)
STODDARD_SOLVENT_MATERIALS = (
    _material("Stoddard solvent", "스토다드솔벤트", "Stoddard solvent"),
)
COMMON_MATERIALS = (
    _material("CS2", "cs2"),
    _material("Formic acid", "formic acid", "개미산"),
    _material("Carbon tetrachloride", "사염화탄소", "carbon tetrachloride"),
)


ANALYSIS_TYPES: tuple[AnalysisTypeDefinition, ...] = (
    AnalysisTypeDefinition("mixture", "혼유", MIXTURE_MATERIALS, "mixture"),
    AnalysisTypeDefinition("mixture_one_column", "1컬럼혼유", ONE_COLUMN_MATERIALS, "mixture_one_column"),
    AnalysisTypeDefinition("mek", "MEK", MEK_MATERIALS, "mek"),
    AnalysisTypeDefinition("alcohol_2_iba_1_btoh", "(알콜2) IBA,1-BTOH", ALCOHOL_2_MATERIALS, "alcohol_2"),
    AnalysisTypeDefinition("ipa", "IPA", IPA_MATERIALS, "ipa"),
    AnalysisTypeDefinition("methanol", "메탄올A", METHANOL_MATERIALS, "methanol"),
    AnalysisTypeDefinition("acn", "ACN", ACN_MATERIALS, "acn"),
    AnalysisTypeDefinition("bc", "B.C", BC_MATERIALS, "bc"),
    AnalysisTypeDefinition("dmf_dma", "DMF,DMA", DMF_DMA_MATERIALS, "dmf_dma"),
    AnalysisTypeDefinition(
        "isoamyl_n_propyl_acetate",
        "이소아밀,n-프로필 아세테이트",
        ISOAMYL_N_PROPYL_ACETATE_MATERIALS,
        "isoamyl_n_propyl_acetate",
    ),
    AnalysisTypeDefinition("phenol", "페놀", PHENOL_MATERIALS),
    AnalysisTypeDefinition(
        "propylene_oxide",
        "1,2-에폭시프로판(산화프로필렌)",
        PROPYLENE_OXIDE_MATERIALS,
        "propylene_oxide",
    ),
    AnalysisTypeDefinition(
        "diethyl_ether",
        "디에틸에테르",
        DIETHYL_ETHER_MATERIALS,
        "diethyl_ether",
    ),
    AnalysisTypeDefinition("thf_diethylene_oxide", "THF(Diethylene oxide)", materials_pending=True),
    AnalysisTypeDefinition(
        "pce", "PCE(테트라클로로에틸렌)", (TETRACHLOROETHYLENE_MATERIAL,)
    ),
    AnalysisTypeDefinition(
        "dichloromethane", "디클로로메탄(MC)", DICHLOROMETHANE_MATERIALS,
        "dichloromethane",
    ),
    AnalysisTypeDefinition(
        "methyl_n_amyl_ketone",
        "메틸 n아밀케톤",
        METHYL_N_AMYL_KETONE_MATERIALS,
        "methyl_n_amyl_ketone",
    ),
    AnalysisTypeDefinition(
        "chloromethane",
        "메틸클로라이드(Chloromethane)",
        CHLOROMETHANE_MATERIALS,
    ),
    AnalysisTypeDefinition(
        "vinyl_acetate", "비닐아세테이트", VINYL_ACETATE_MATERIALS,
        "vinyl_acetate",
    ),
    AnalysisTypeDefinition("cellosolve", "셀로솔브", CELLOSOLVE_MATERIALS, "cellosolve"),
    AnalysisTypeDefinition("alcohol_4", "알콜4", ALCOHOL_4_MATERIALS, "alcohol_4"),
    AnalysisTypeDefinition(
        "stoddard_solvent", "스토다드솔벤트",
        STODDARD_SOLVENT_MATERIALS, "stoddard_solvent",
    ),
    AnalysisTypeDefinition(
        "isopropyl_acetate", "이소프로필 아세테이트", ISOPROPYL_ACETATE_MATERIALS,
        "isopropyl_acetate",
    ),
    AnalysisTypeDefinition(
        "acetic_acid", "초산", ACETIC_ACID_MATERIALS, "acetic_acid"
    ),
    AnalysisTypeDefinition("pyridine", "피리딘", PYRIDINE_MATERIALS, "pyridine"),
    AnalysisTypeDefinition(
        "ethylene_glycol",
        "에틸렌글리콜",
        ETHYLENE_GLYCOL_MATERIALS,
        "ethylene_glycol",
    ),
    AnalysisTypeDefinition("mixture_g2", "(혼유-G2) THF,CFM,벤젠,클로로벤젠", G2_MATERIALS, "mixture_g2"),
    AnalysisTypeDefinition(
        "mixture_g3",
        "(혼유-G3) 1,2-디클로로에틸렌,퍼클로로에틸렌,프로판,에탄",
        G3_MATERIALS,
        "mixture_g3",
    ),
)

ANALYSIS_TYPE_NAMES = tuple(item.display_name for item in ANALYSIS_TYPES)
_BY_NAME = {item.display_name: item for item in ANALYSIS_TYPES}
_BY_KEY = {item.key: item for item in ANALYSIS_TYPES}

# G2 automation deliberately excludes carbon tetrachloride even when an older
# workbook mentions it.  Future G2 Excel profiles must use this allow-list.
G2_SUPPORTED_MATERIALS = tuple(item.canonical_name for item in G2_MATERIALS)


def validate_analysis_type_registry(
    definitions: tuple[AnalysisTypeDefinition, ...] = ANALYSIS_TYPES,
) -> None:
    for definition in definitions:
        if not definition.supported_materials and not definition.materials_pending:
            raise ValueError(
                f"분석종류 '{definition.display_name}'에 지원 물질이 없습니다. "
                "supported_materials를 등록하거나 materials_pending=True를 명시하세요."
            )
        if definition.supported_materials and definition.materials_pending:
            raise ValueError(
                f"분석종류 '{definition.display_name}'는 지원 물질과 pending을 동시에 설정할 수 없습니다."
            )
        if (
            definition.allow_runtime_single_material_inference
            and not definition.materials_pending
        ):
            raise ValueError(
                f"분석종류 '{definition.display_name}'는 pending 상태가 아닌데 "
                "런타임 물질 추론이 활성화돼 있습니다."
            )


def normalize_material_alias_key(value: str) -> str:
    """Normalize harmless notation differences while preserving chemical locants."""
    compact = " ".join(value.strip().split()).casefold()
    return compact.replace(" ", "").replace(".", "").replace("-", "")


def material_aliases(
    definitions: tuple[AnalysisTypeDefinition, ...] = ANALYSIS_TYPES,
    common_materials: tuple[MaterialDefinition, ...] = COMMON_MATERIALS,
) -> dict[str, str]:
    aliases: dict[str, str] = {}
    for material in (
        *common_materials,
        *(item for definition in definitions for item in definition.supported_materials),
    ):
        for alias in material.aliases:
            key = normalize_material_alias_key(alias)
            existing = aliases.get(key)
            if existing is not None and existing != material.canonical_name:
                raise ValueError(
                    f"물질 alias '{alias}'가 '{existing}'와 '{material.canonical_name}'에 중복 등록됐습니다."
                )
            aliases[key] = material.canonical_name
    return aliases


validate_analysis_type_registry()


def analysis_type_key(display_name: str) -> str | None:
    definition = _BY_NAME.get(display_name)
    return definition.key if definition else None


def analysis_type_display_name(key: str) -> str | None:
    definition = _BY_KEY.get(key)
    return definition.display_name if definition else None


def supported_materials_for(analysis_type: str) -> tuple[MaterialDefinition, ...]:
    definition = _BY_NAME.get(analysis_type)
    return definition.supported_materials if definition else ()


def supported_canonical_names_for(analysis_type: str) -> frozenset[str]:
    return frozenset(
        material.canonical_name for material in supported_materials_for(analysis_type)
    )


def material_supported_for_analysis(
    analysis_type: str, canonical_name: str | None
) -> bool:
    """Return whether a known canonical is eligible for this analysis.

    Global canonical recognition is intentionally independent from analysis-level
    eligibility.  The sole exception is the explicitly enabled exact-name runtime
    inference used by selected pending single-material analyses.
    """
    if canonical_name is None:
        return False
    if canonical_name in supported_canonical_names_for(analysis_type):
        return True
    if not runtime_material_inference_allowed_for(analysis_type):
        return False
    return re.sub(r"\s+", "", canonical_name).casefold() == re.sub(
        r"\s+", "", analysis_type
    ).casefold()


def materials_pending_for(analysis_type: str) -> bool:
    definition = _BY_NAME.get(analysis_type)
    return bool(definition and definition.materials_pending)


def runtime_material_inference_allowed_for(analysis_type: str) -> bool:
    definition = _BY_NAME.get(analysis_type)
    return bool(definition and definition.allow_runtime_single_material_inference)


def has_excel_profile(analysis_type: str) -> bool:
    definition = _BY_NAME.get(analysis_type)
    return bool(definition and definition.excel_profile_key)


def excel_profile_key_for(analysis_type: str) -> str | None:
    definition = _BY_NAME.get(analysis_type)
    return definition.excel_profile_key if definition else None


def infer_analysis_type(
    filename: str,
    method_filenames: tuple[str, ...] = (),
    materials: tuple[str, ...] = (),
) -> str | None:
    evidence = " ".join((filename, *method_filenames, *materials)).casefold()
    filename_evidence = filename.casefold()
    # Specific names precede broad families. Short or ambiguous names are not guessed.
    rules: tuple[tuple[str, tuple[str, ...]], ...] = (
        ("(혼유-G2) THF,CFM,벤젠,클로로벤젠", ("혼유-g2",)),
        (
            "(혼유-G3) 1,2-디클로로에틸렌,퍼클로로에틸렌,프로판,에탄",
            ("혼유-g3", "g3 혼유", "g3-혼유"),
        ),
        ("(알콜2) IBA,1-BTOH", ("알콜(2)", "알콜2")),
        ("알콜4", ("알콜(4)", "알콜4")),
        ("DMF,DMA", ("dmf,dma", "dmf dma", "(dmf)")),
        ("이소아밀,n-프로필 아세테이트", ("이소아밀,n-프로필 아세테이트",)),
        ("1,2-에폭시프로판(산화프로필렌)", ("산화프로필렌", "1,2-에폭시프로판")),
        ("THF(Diethylene oxide)", ("diethylene oxide",)),
        ("PCE(테트라클로로에틸렌)", ("테트라클로로에틸렌",)),
        ("디클로로메탄(MC)", ("디클로로메탄",)),
        ("메틸클로라이드(Chloromethane)", ("chloromethane", "메틸클로라이드")),
        ("메틸 n아밀케톤", ("메틸 n아밀케톤", "메틸 n-아밀케톤")),
        ("비닐아세테이트", ("비닐아세테이트",)),
        ("스토다드솔벤트", ("스토다드솔벤트",)),
        ("이소프로필 아세테이트", ("이소프로필 아세테이트", "이소프로필아세테이트")),
        ("디에틸에테르", ("디에틸에테르",)),
        ("셀로솔브", ("셀로솔브",)),
        ("초산", ("초산", "acetic acid")),
        ("메탄올A", ("메탄올a", "메탄올", "methanol")),
        ("ACN", ("acn", "acetonitrile")),
        ("페놀", ("페놀", "phenol")),
        ("피리딘", ("피리딘", "pyridine")),
        ("에틸렌글리콜", ("에틸렌글리콜", "ethylene glycol")),
        ("MEK", ("mek", "methyl ethyl ketone", "메틸에틸케톤")),
        ("1컬럼혼유", ("1컬럼", "methyl acetate", "c-hexane", "n-heptane", "isobutyl acetate")),
        ("혼유", ("혼유",)),
    )
    if "iba" in evidence and any(token in evidence for token in ("1-btoh", "n-btoh")):
        return "(알콜2) IBA,1-BTOH"
    if re.search(r"(?<![0-9a-z])ipa(?![0-9a-z])", filename_evidence):
        return "IPA"
    for display_name, tokens in rules:
        if any(token in evidence for token in tokens):
            return display_name
    if re.search(r"(?<![0-9a-z])ipa(?![0-9a-z])", evidence):
        return "IPA"
    return None
