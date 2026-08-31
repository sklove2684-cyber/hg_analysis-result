from __future__ import annotations

from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
import unittest
from unittest.mock import patch
from uuid import uuid4

from honyu_app.application.environment_diagnostics import (
    EnvironmentDiagnosticService,
    FAILURE,
    file_sha256,
)
from honyu_app.domain.results import BatchSummary
from honyu_app.services.excel_template_service import (
    ExcelTemplateSnapshot,
    TemplateCell,
)


class FakeDatabase:
    def __init__(self, path: Path) -> None:
        self.database_file = path

    def check_connection(self):
        return SimpleNamespace(mode="mock")

    def search_batches(self, _query):
        return [
            BatchSummary(
                batch_id=uuid4(),
                batch_code="IPA-120-167",
                pdf_filename="IPA 120-167.pdf",
                analysis_type="IPA",
                review_status="SAVED",
            )
        ]


class FakeInspector:
    def __init__(self, snapshots: dict[str, ExcelTemplateSnapshot]) -> None:
        self.snapshots = snapshots
        self.paths: list[Path] = []

    def inspect(self, path: Path) -> ExcelTemplateSnapshot:
        self.paths.append(path)
        return self.snapshots[path.name]


def workbook_snapshot(
    path: Path,
    sheets: tuple[str, ...],
    *cells: TemplateCell,
) -> ExcelTemplateSnapshot:
    return ExcelTemplateSnapshot(
        path,
        sheets,
        {(cell.sheet, cell.address): cell for cell in cells},
    )


class EnvironmentDiagnosticServiceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.database_path = self.root / "actual.db"
        self.database_path.write_bytes(b"database-content")
        self.pdf = self.root / "IPA 120-167.pdf"
        self.pdf.write_bytes(b"pdf-content")
        for name in ("lod.xlsx", "area.xlsx", "mixture.xlsx"):
            (self.root / name).write_bytes(name.encode())
        self.snapshots = {
            "lod.xlsx": workbook_snapshot(
                self.root / "lod.xlsx",
                ("검량선", "LOD(area입력)", "회수율", "std"),
                TemplateCell("LOD(area입력)", "I3", True, "IPA", "string"),
                TemplateCell("LOD(area입력)", "J3", True, "area", "string"),
            ),
            "area.xlsx": workbook_snapshot(
                self.root / "area.xlsx",
                ("검량선", "area", "회수율", "std", "Sheet1", "Sheet2"),
                TemplateCell("area", "I3", True, "IPA", "string"),
                TemplateCell("area", "J3", True, "area", "string"),
            ),
            "mixture.xlsx": workbook_snapshot(
                self.root / "mixture.xlsx",
                ("검량선", "area", "최종결과", "회수율", "STD제조"),
                TemplateCell("area", "I3", True, "MIBK", "string"),
                TemplateCell("area", "J3", True, "Toluene", "string"),
            ),
        }
        self.inspector = FakeInspector(self.snapshots)
        self.service = EnvironmentDiagnosticService(
            FakeDatabase(self.database_path),
            project_directory=self.root,
            inspector=self.inspector,
        )

    def _successful_git(self, *arguments: str) -> str:
        values = {
            ("branch", "--show-current"): "main",
            ("rev-parse", "HEAD"): "a" * 40,
            ("log", "-1", "--format=%s"): "diagnostic commit",
            ("rev-parse", "refs/remotes/origin/main"): "a" * 40,
            ("status", "--porcelain"): "",
        }
        return values[arguments]

    def test_git_db_pdf_and_hash_information_is_displayed(self) -> None:
        with patch.object(self.service, "_git", side_effect=self._successful_git):
            report = self.service.collect(str(self.pdf), "")

        self.assertTrue(report.git_matches)
        self.assertTrue(report.worktree_clean)
        self.assertIn("[정상] Local == origin/main", report.text)
        self.assertIn(str(self.database_path.resolve()), report.text)
        self.assertIn("DB 종류: mock", report.text)
        self.assertIn("IPA-120-167", report.text)
        self.assertIn("IPA 120-167.pdf", report.text)
        self.assertEqual(report.pdf_sha256, sha256(b"pdf-content").hexdigest().upper())

    def test_git_failure_is_isolated_and_report_still_completes(self) -> None:
        with patch.object(self.service, "_git", side_effect=OSError("git missing")):
            report = self.service.collect()

        self.assertIsNone(report.git_matches)
        self.assertEqual(report.local_sha, FAILURE)
        self.assertIn("Git commit: 확인 실패", report.text)
        self.assertIn("DB 경로:", report.text)

    def test_excel_profile_detection_covers_both_ipa_layouts_and_mixture(self) -> None:
        expected = {
            "lod.xlsx": ("ipa", "IPA - LOD(area입력)형", "LOD(area입력)!I3 = IPA"),
            "area.xlsx": ("ipa", "IPA - area형", "area!I3 = IPA"),
            "mixture.xlsx": ("mixture", "혼유", "area!I3 = MIBK"),
        }
        for filename, (key, name, evidence) in expected.items():
            with self.subTest(filename=filename), patch.object(
                self.service, "_git", side_effect=self._successful_git
            ):
                selected = self.root / filename
                report = self.service.collect(excel_path=str(selected))

                self.assertEqual(report.excel_path, str(selected))
                self.assertEqual(report.excel_profile_key, key)
                self.assertEqual(report.excel_profile_name, name)
                self.assertIn(evidence, report.text)
                self.assertEqual(
                    report.excel_sha256,
                    sha256(filename.encode()).hexdigest().upper(),
                )
                self.assertEqual(self.inspector.paths[-1], selected)

    def test_selected_excel_access_failure_does_not_raise(self) -> None:
        unavailable = r"\\server\company\(IPA) 120-167.xlsx"
        with patch.object(self.service, "_git", side_effect=self._successful_git):
            report = self.service.collect(excel_path=unavailable)

        self.assertEqual(report.excel_path, unavailable)
        self.assertEqual(report.excel_sha256, "접근 실패")
        self.assertIn(f"전체 경로: {unavailable}", report.text)
        self.assertIn("실제 시트 목록: 접근 실패", report.text)

    def test_file_sha256_streaming_matches_reference(self) -> None:
        path = self.root / "hash.bin"
        content = (b"0123456789" * 1000) + b"end"
        path.write_bytes(content)

        self.assertEqual(file_sha256(path, chunk_size=17), sha256(content).hexdigest().upper())


if __name__ == "__main__":
    unittest.main()
