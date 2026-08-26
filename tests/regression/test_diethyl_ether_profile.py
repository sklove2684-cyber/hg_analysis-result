from __future__ import annotations

import os
from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from zipfile import ZipFile

from honyu_app.application.create_excel_export import CreateExcelExportService
from honyu_app.application.preview_excel_export import PreviewExcelExportService
from honyu_app.application.review_extraction import ReviewExtractionService
from honyu_app.domain.enums import ExcelPreviewStatus, SampleType
from honyu_app.infrastructure.database.mock_database_service import MockDatabaseService
from honyu_app.infrastructure.excel.workbook_inspector import XlsxTemplateInspector
from honyu_app.infrastructure.excel.workbook_validator import XlsxWorkbookValidator
from honyu_app.infrastructure.excel.xml_cell_writer import XlsxXmlCellWriter
from honyu_app.infrastructure.pdf.labsolutions_parser import LabSolutionsParser


def _candidate_directories() -> tuple[Path, ...]:
    configured = os.environ.get("HONYU_DIETHYL_TEST_DIR")
    return tuple(
        directory
        for directory in (
            Path(configured) if configured else None,
            Path(__file__).parents[3] / "TEST",
            Path.home() / "Desktop" / "분석프로그램",
        )
        if directory is not None
    )


def _find_file(name: str) -> Path:
    for directory in _candidate_directories():
        candidate = directory / name
        if candidate.is_file():
            return candidate
    return Path()


PDF = _find_file("디에틸에테르 152,153@완료.pdf")
TEMPLATE = _find_file("(디에틸에테르) 152,153.xlsx")
MIXTURE_TEMPLATE = _find_file("(혼유) 601-690.xlsx")


@unittest.skipUnless(PDF.is_file() and TEMPLATE.is_file(), "디에틸에테르 실제 PDF/XLSX가 없습니다.")
class DiethylEtherActualFileTests(unittest.TestCase):
    def _saved_batch(self, database: MockDatabaseService):
        batch = LabSolutionsParser().parse(
            PDF,
            analysis_type="디에틸에테르",
            analysis_no_start=152,
            analysis_no_end=153,
        )
        review = ReviewExtractionService(database)
        review.complete_review(batch)
        return review.save_batch(batch)

    def _preview(self, std_method: str = "A"):
        temporary = TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        database = MockDatabaseService(Path(temporary.name) / "diethyl.db")
        saved = self._saved_batch(database)
        result = PreviewExcelExportService(database, XlsxTemplateInspector()).preview(
            saved.batch_id, TEMPLATE, std_method
        )
        return database, saved, result

    def test_actual_pdf_uses_common_a_b_std_selection_without_area_judgment(self) -> None:
        expected_by_method = {
            "A": {
                "F4": 26539,
                "F5": 62642,
                "F6": 63347,
                "F7": 280075,
                "F8": 534342,
            },
            "B": {
                "F4": 26539,
                "F5": 62642,
                "F6": 63347,
                "F7": 280075,
                "F8": 705737,
            },
        }
        for method, expected_std in expected_by_method.items():
            with self.subTest(method=method):
                _database, _saved, result = self._preview(method)
                mapped = {
                    row.target_cell: row.applied_area
                    for row in result.rows
                    if row.status is ExcelPreviewStatus.MAPPED
                    and row.target_sheet == "LOD(area입력)"
                    and row.target_cell in expected_std
                }

                self.assertTrue(result.can_generate, result.issues)
                self.assertEqual(result.error_count, 0)
                self.assertEqual(result.mapped_count, 14)
                self.assertEqual(result.excluded_count, 30)
                self.assertEqual(mapped, expected_std)
                saved_batch = _database.get_batch_detail(_saved.batch_id)
                std1 = next(
                    sample for sample in saved_batch.samples
                    if sample.sample_type is SampleType.STD
                    and sample.replicate_no == 1
                )
                std_target = next(
                    peak for peak in std1.peaks
                    if peak.material_standard == "Diethyl ether"
                )
                self.assertEqual(std_target.retention_time, Decimal("1.305"))
                duplicate_recheck = [
                    row
                    for row in result.rows
                    if row.sample_name == "STD2"
                    and row.applied_area == 61218
                    and row.status is ExcelPreviewStatus.EXCLUDED
                ]
                self.assertEqual(len(duplicate_recheck), 1)
                self.assertEqual(
                    duplicate_recheck[0].exclude_reason, "DUPLICATE_STD_SET"
                )

    def test_actual_export_preserves_formulas_styles_merges_and_charts(self) -> None:
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            database = MockDatabaseService(root / "diethyl.db")
            saved = self._saved_batch(database)
            output = root / "diethyl-result.xlsx"

            result = CreateExcelExportService(
                database,
                PreviewExcelExportService(database, XlsxTemplateInspector()),
                XlsxXmlCellWriter(),
                XlsxWorkbookValidator(),
                object(),
                recalculate_with_excel=False,
            ).create(saved.batch_id, TEMPLATE, output, "A", "REGRESSION")

            self.assertTrue(result.validation_passed)
            self.assertFalse(result.recalculated)
            self.assertEqual(result.mapped_cell_count, 14)
            before = XlsxTemplateInspector().inspect(TEMPLATE)
            after = XlsxTemplateInspector().inspect(output)
            expected = {
                "F4": 26539,
                "F5": 62642,
                "F6": 63347,
                "F7": 280075,
                "F8": 534342,
            }
            for address, value in expected.items():
                cell = after.cell("LOD(area입력)", address)
                self.assertEqual(cell.value, value)
                self.assertEqual(cell.value_type, "numeric")
            self.assertEqual(
                {key: cell.formula for key, cell in before.cells.items() if cell.has_formula},
                {key: cell.formula for key, cell in after.cells.items() if cell.has_formula},
            )
            self.assertEqual(
                {key: cell.style_id for key, cell in before.cells.items()},
                {key: cell.style_id for key, cell in after.cells.items()},
            )
            with ZipFile(TEMPLATE) as original, ZipFile(output) as generated:
                preserved = tuple(
                    name
                    for name in original.namelist()
                    if name.startswith(("xl/charts/", "xl/drawings/", "xl/media/"))
                )
                self.assertTrue(any(name.startswith("xl/charts/") for name in preserved))
                for name in preserved:
                    self.assertEqual(original.read(name), generated.read(name), name)

    @unittest.skipUnless(MIXTURE_TEMPLATE.is_file(), "비교용 혼유 Excel 양식이 없습니다.")
    def test_actual_mixture_workbook_is_rejected_as_profile_mismatch(self) -> None:
        with TemporaryDirectory() as temporary:
            database = MockDatabaseService(Path(temporary) / "diethyl.db")
            saved = self._saved_batch(database)
            result = PreviewExcelExportService(
                database, XlsxTemplateInspector()
            ).preview(saved.batch_id, MIXTURE_TEMPLATE, "A")

        self.assertFalse(result.can_generate)
        self.assertEqual(result.issues[0].code, "TEMPLATE_PROFILE_MISMATCH")
        self.assertIn("디에틸에테르", result.issues[0].message)
        self.assertIn("혼유", result.issues[0].message)


if __name__ == "__main__":
    unittest.main()
