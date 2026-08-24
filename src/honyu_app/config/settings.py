from dataclasses import dataclass
import os

from honyu_app.domain.enums import DatabaseMode
from honyu_app.config.paths import default_local_export_dir


@dataclass(frozen=True, slots=True)
class Settings:
    database_mode: DatabaseMode
    log_level: str
    unc_base_path: str
    z_fallback_path: str
    local_export_path: str


def load_settings() -> Settings:
    return Settings(
        database_mode=DatabaseMode(os.getenv("HONYU_DATABASE_MODE", "mock")),
        log_level=os.getenv("HONYU_LOG_LEVEL", "INFO"),
        unc_base_path=os.getenv(
            "HONYU_UNC_BASE_PATH",
            r"\\172.30.1.100\data\분석결과(작업장별)\작업장",
        ),
        z_fallback_path=os.getenv(
            "HONYU_Z_FALLBACK_PATH",
            r"Z:\분석결과(작업장별)\작업장",
        ),
        local_export_path=os.getenv(
            "HONYU_LOCAL_EXPORT_PATH", str(default_local_export_dir())
        ),
    )
