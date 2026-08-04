from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from honyu_app.domain.enums import HalfYear
from honyu_app.infrastructure.filesystem.shared_folder_service import (
    WindowsSharedFolderService,
)


class SharedFolderServiceTests(unittest.TestCase):
    def test_unc_is_preferred_and_workplaces_are_sorted(self) -> None:
        with TemporaryDirectory() as temp:
            base = Path(temp) / "unc"
            fallback = Path(temp) / "fallback"
            base.mkdir()
            fallback.mkdir()
            for name in ("작업장10", "작업장2", "Alpha"):
                (base / name).mkdir()
            service = WindowsSharedFolderService(str(base), str(fallback))
            status = service.check_connection()
            self.assertTrue(status.connected)
            self.assertFalse(status.used_fallback)
            self.assertEqual(service.list_workplaces(), ["Alpha", "작업장2", "작업장10"])

    def test_z_path_is_used_only_as_fallback(self) -> None:
        with TemporaryDirectory() as temp:
            missing = Path(temp) / "missing"
            fallback = Path(temp) / "fallback"
            fallback.mkdir()
            service = WindowsSharedFolderService(str(missing), str(fallback))
            status = service.check_connection()
            self.assertTrue(status.connected)
            self.assertTrue(status.used_fallback)
            self.assertEqual(status.active_base_path, str(fallback))

    def test_period_path_uses_two_digit_year_and_half_suffix(self) -> None:
        with TemporaryDirectory() as temp:
            base = Path(temp) / "base"
            base.mkdir()
            service = WindowsSharedFolderService(str(base), str(Path(temp) / "missing"))
            service.check_connection()
            self.assertEqual(
                service.build_period_path("작업장", 2026, HalfYear.FIRST),
                base / "작업장" / "26상",
            )
            self.assertEqual(
                service.build_period_path("작업장", 2026, HalfYear.SECOND),
                base / "작업장" / "26하",
            )

    def test_missing_period_is_not_created(self) -> None:
        with TemporaryDirectory() as temp:
            base = Path(temp) / "base"
            base.mkdir()
            service = WindowsSharedFolderService(str(base), str(Path(temp) / "missing"))
            service.check_connection()
            result = service.validate_period_path("작업장", 2026, HalfYear.FIRST)
            self.assertFalse(result.valid)
            self.assertFalse((base / "작업장" / "26상").exists())

    def test_final_folder_must_be_below_period_folder(self) -> None:
        with TemporaryDirectory() as temp:
            base = Path(temp) / "base"
            period = base / "작업장" / "26상"
            final = period / "결과"
            outside = base / "다른폴더"
            final.mkdir(parents=True)
            outside.mkdir()
            service = WindowsSharedFolderService(str(base), str(Path(temp) / "missing"))
            service.check_connection()
            self.assertTrue(
                service.validate_final_folder(
                    final, workplace="작업장", year=2026, half=HalfYear.FIRST
                ).valid
            )
            self.assertFalse(
                service.validate_final_folder(
                    outside, workplace="작업장", year=2026, half=HalfYear.FIRST
                ).valid
            )
