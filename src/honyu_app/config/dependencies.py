from dataclasses import dataclass

from honyu_app.config.paths import mock_db_dir
from honyu_app.config.settings import Settings
from honyu_app.infrastructure.database.mock_database_service import MockDatabaseService
from honyu_app.services.database_service import DatabaseService


@dataclass(frozen=True, slots=True)
class ApplicationServices:
    database: DatabaseService


def build_services(settings: Settings) -> ApplicationServices:
    if settings.database_mode.value == "mock":
        return ApplicationServices(
            database=MockDatabaseService(mock_db_dir() / "honyu_mock.db")
        )
    raise RuntimeError("Supabase 연결은 PHASE 10에서 활성화됩니다.")
