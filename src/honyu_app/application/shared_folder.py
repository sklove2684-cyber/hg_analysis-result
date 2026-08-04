from dataclasses import dataclass
from pathlib import Path

from honyu_app.domain.enums import HalfYear
from honyu_app.domain.results import FolderValidationResult, SharedFolderConnectionStatus
from honyu_app.infrastructure.storage.local_settings import (
    LocalSettingsStore,
    RecentFolderSelection,
)
from honyu_app.services.shared_folder_service import SharedFolderService


@dataclass(frozen=True, slots=True)
class SharedFolderViewState:
    connection: SharedFolderConnectionStatus
    workplaces: tuple[str, ...]
    recent: RecentFolderSelection


class SharedFolderController:
    def __init__(
        self, service: SharedFolderService, settings_store: LocalSettingsStore
    ) -> None:
        self._service = service
        self._settings_store = settings_store

    def refresh(self) -> SharedFolderViewState:
        status = self._service.check_connection()
        workplaces = tuple(self._service.list_workplaces()) if status.connected else ()
        return SharedFolderViewState(
            connection=status,
            workplaces=workplaces,
            recent=self._settings_store.load_recent_selection(),
        )

    def validate_period(
        self, workplace: str, year: int, half: HalfYear
    ) -> FolderValidationResult:
        return self._service.validate_period_path(workplace, year, half)

    def select_final_folder(
        self, folder: Path, *, workplace: str, year: int, half: HalfYear
    ) -> FolderValidationResult:
        result = self._service.validate_final_folder(
            folder, workplace=workplace, year=year, half=half
        )
        if result.valid:
            self._settings_store.save_recent_selection(
                RecentFolderSelection(workplace, year, half.value, result.path)
            )
        return result
