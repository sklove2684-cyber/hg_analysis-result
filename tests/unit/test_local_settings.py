from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from honyu_app.infrastructure.storage.local_settings import (
    LocalSettingsStore,
    RecentFolderSelection,
)


class LocalSettingsStoreTests(unittest.TestCase):
    def test_round_trip_recent_selection(self) -> None:
        with TemporaryDirectory() as temp:
            store = LocalSettingsStore(Path(temp) / "settings.json")
            expected = RecentFolderSelection("작업장", 2026, "상반기", r"Z:\결과")
            store.save_recent_selection(expected)
            self.assertEqual(store.load_recent_selection(), expected)

    def test_missing_file_returns_empty_selection(self) -> None:
        with TemporaryDirectory() as temp:
            store = LocalSettingsStore(Path(temp) / "settings.json")
            self.assertEqual(store.load_recent_selection(), RecentFolderSelection())
