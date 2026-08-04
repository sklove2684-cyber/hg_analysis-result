from __future__ import annotations

from dataclasses import asdict, dataclass
import json
from pathlib import Path


@dataclass(frozen=True, slots=True)
class RecentFolderSelection:
    workplace: str | None = None
    year: int | None = None
    half: str | None = None
    final_folder: str | None = None


class LocalSettingsStore:
    def __init__(self, settings_file: Path) -> None:
        self._settings_file = settings_file

    def load_recent_selection(self) -> RecentFolderSelection:
        if not self._settings_file.is_file():
            return RecentFolderSelection()
        try:
            payload = json.loads(self._settings_file.read_text(encoding="utf-8"))
            value = payload.get("recent_folder_selection", {})
            return RecentFolderSelection(
                workplace=value.get("workplace"),
                year=value.get("year"),
                half=value.get("half"),
                final_folder=value.get("final_folder"),
            )
        except (OSError, ValueError, TypeError, AttributeError):
            return RecentFolderSelection()

    def save_recent_selection(self, selection: RecentFolderSelection) -> None:
        self._settings_file.parent.mkdir(parents=True, exist_ok=True)
        payload = {"recent_folder_selection": asdict(selection)}
        temporary = self._settings_file.with_suffix(".tmp")
        temporary.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        temporary.replace(self._settings_file)
