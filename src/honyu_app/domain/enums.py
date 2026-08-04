from enum import StrEnum


class SampleType(StrEnum):
    BLANK = "BLANK"
    STD = "STD"
    RECOVERY_BLANK = "RECOVERY_BLANK"
    RECOVERY = "RECOVERY"
    NUMERIC = "NUMERIC"
    UNKNOWN = "UNKNOWN"


class ConcentrationLevel(StrEnum):
    LOW = "LOW"
    MID = "MID"
    HIGH = "HIGH"


class ReviewStatus(StrEnum):
    PENDING = "PENDING"
    REVIEWED = "REVIEWED"
    SAVED = "SAVED"


class ExcludeReason(StrEnum):
    BLANK_SAMPLE = "BLANK_SAMPLE"
    RECOVERY_BLANK = "RECOVERY_BLANK"
    SAMPLE_NAME_ENDS_WITH_B = "SAMPLE_NAME_ENDS_WITH_B"
    INTERNAL_STANDARD_CS2 = "INTERNAL_STANDARD_CS2"
    UNNAMED_PEAK = "UNNAMED_PEAK"
    UNKNOWN_MATERIAL = "UNKNOWN_MATERIAL"
    USER_EXCLUDED = "USER_EXCLUDED"
    DIBK_AREA_NOT_TOP2 = "DIBK_AREA_NOT_TOP2"


class StdMethod(StrEnum):
    A = "A"
    B = "B"


class ExcelPreviewStatus(StrEnum):
    MAPPED = "MAPPED"
    EXCLUDED = "EXCLUDED"
    ERROR = "ERROR"


class ValidationSeverity(StrEnum):
    WARNING = "WARNING"
    ERROR = "ERROR"


class DatabaseMode(StrEnum):
    MOCK = "mock"
    SUPABASE = "supabase"


class HalfYear(StrEnum):
    FIRST = "상반기"
    SECOND = "하반기"

    @property
    def folder_suffix(self) -> str:
        return "상" if self is HalfYear.FIRST else "하"
