from __future__ import annotations

import importlib.util
import os
from pathlib import Path
import unittest


PYSIDE_AVAILABLE = importlib.util.find_spec("PySide6") is not None

if PYSIDE_AVAILABLE:
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    from PySide6.QtCore import QEventLoop, QTimer
    from PySide6.QtWidgets import QApplication

    from honyu_app.application.environment_diagnostics import (
        EnvironmentDiagnosticReport,
    )
    from honyu_app.ui.pages.settings_page import SettingsPage


@unittest.skipUnless(PYSIDE_AVAILABLE, "PySide6가 설치된 GUI 환경이 아닙니다.")
class EnvironmentDiagnosticUiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    @staticmethod
    def _report(excel_path: str = "") -> EnvironmentDiagnosticReport:
        return EnvironmentDiagnosticReport(
            text=f"diagnostic\nExcel={excel_path}",
            local_sha="a" * 40,
            origin_sha="a" * 40,
            git_matches=True,
            worktree_clean=True,
            pdf_path="선택되지 않음",
            pdf_sha256="선택되지 않음",
            excel_path=excel_path or "선택되지 않음",
            excel_sha256="B" * 64 if excel_path else "선택되지 않음",
            excel_profile_key="ipa" if excel_path else "선택되지 않음",
            excel_profile_name="IPA - area형" if excel_path else "선택되지 않음",
        )

    def _wait_for_refresh(self, page: SettingsPage) -> None:
        thread = page._diagnostic_thread
        self.assertIsNotNone(thread)
        loop = QEventLoop()
        thread.finished.connect(loop.quit)
        QTimer.singleShot(3000, loop.quit)
        loop.exec()
        self.assertIsNone(page._diagnostic_thread)

    def test_selected_excel_change_refreshes_report_immediately(self) -> None:
        calls: list[tuple[str, str]] = []

        class FakeService:
            def collect(self, pdf_path="", excel_path=""):
                calls.append((pdf_path, excel_path))
                return EnvironmentDiagnosticUiTests._report(excel_path)

        page = SettingsPage(None, FakeService())
        page.show()
        QApplication.processEvents()
        first = str(Path("company") / "first.xlsx")
        second = str(Path("company") / "second.xlsx")

        page.set_selected_excel(first)
        self._wait_for_refresh(page)
        self.assertEqual(calls[-1], ("", first))
        self.assertIn(first, page.diagnostic_text.toPlainText())

        page.set_selected_excel(second)
        self._wait_for_refresh(page)
        self.assertEqual(calls[-1], ("", second))
        self.assertIn(second, page.diagnostic_text.toPlainText())
        page.deleteLater()

    def test_copy_button_copies_exact_diagnostic_text(self) -> None:
        class FakeService:
            def collect(self, pdf_path="", excel_path=""):
                return EnvironmentDiagnosticUiTests._report(excel_path)

        page = SettingsPage(None, FakeService())
        report = self._report(r"\\server\share\actual.xlsx")
        page._apply_diagnostic_report(report)
        page.copy_diagnostic_button.click()

        self.assertEqual(QApplication.clipboard().text(), report.text)
        self.assertIn("클립보드에 복사", page.diagnostic_status.text())
        page.deleteLater()


if __name__ == "__main__":
    unittest.main()
