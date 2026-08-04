from pathlib import Path
from typing import Protocol

from honyu_app.domain.enums import HalfYear
from honyu_app.domain.results import FolderValidationResult, SharedFolderConnectionStatus


class SharedFolderService(Protocol):
    def check_connection(self) -> SharedFolderConnectionStatus: ...

    def list_workplaces(self) -> list[str]: ...

    def build_period_path(self, workplace: str, year: int, half: HalfYear) -> Path: ...

    def validate_period_path(
        self, workplace: str, year: int, half: HalfYear
    ) -> FolderValidationResult: ...

    def validate_final_folder(
        self, folder: Path, *, workplace: str, year: int, half: HalfYear
    ) -> FolderValidationResult: ...

