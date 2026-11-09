from __future__ import annotations

import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch
from zipfile import ZipFile

from honyu_app.application.create_excel_export import CreateExcelExportService
from honyu_app.application.preview_excel_export import PreviewExcelExportService
from honyu_app.application.review_extraction import ReviewExtractionService
from honyu_app.domain.enums import ExcelPreviewStatus
from honyu_app.infrastructure.database.mock_database_service import MockDatabaseService
from honyu_app.infrastructure.excel.workbook_inspector import XlsxTemplateInspector
from honyu_app.infrastructure.excel.workbook_validator import XlsxWorkbookValidator
from honyu_app.infrastructure.excel.xml_cell_writer import XlsxXmlCellWriter
from honyu_app.infrastructure.pdf.labsolutions_parser import LabSolutionsParser


ACTUAL_DIR = Path(os.environ.get(
    "HONYU_ISOPROPYL_ACETATE_530_TEST_DIR",
    r"\\172.30.1.100\data\분석결과(사업장별)★\자동화프로그램@@\11월\11.09\이소프로필아세테이트 530-531",
))
PDF = ACTUAL_DIR / "이소프로필아세테이트 530-531.pdf"
TEMPLATE = ACTUAL_DIR / "(이소프로필 아세테이트) 530-531.xlsx"


def actual_files_available() -> bool:
    try:
        return PDF.is_file() and TEMPLATE.is_file()
    except OSError:
        return False


@unittest.skipUnless(actual_files_available(), "이소프로필 아세테이트 530-531 실제 파일이 필요합니다.")
class IsopropylAcetate530To531ActualRegressionTests(unittest.TestCase):
    def test_pdf_database_preview_and_xlsx_generation(self) -> None:
        batch = LabSolutionsParser().parse(
            PDF,
            analysis_type="이소프로필 아세테이트",
            analysis_no_start=530,
            analysis_no_end=531,
        )
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            database = MockDatabaseService(root / "actual.db")
            review = ReviewExtractionService(database)
            review.complete_review(batch)
            saved = review.save_batch(batch)
            inspector = XlsxTemplateInspector()
            preview_service = PreviewExcelExportService(database, inspector)
            preview = preview_service.preview(saved.batch_id, TEMPLATE, "A")

            self.assertTrue(preview.can_generate, preview.issues)
            self.assertEqual(16, preview.mapped_count)
            self.assertEqual(164, preview.excluded_count)
            self.assertEqual(0, preview.error_count)
            mapped = {
                (row.target_sheet, row.target_cell): row.applied_area
                for row in preview.rows if row.status is ExcelPreviewStatus.MAPPED
            }
            self.assertEqual({
                ("LOD(area입력)", "F4"): 11818,
                ("LOD(area입력)", "F5"): 24726,
                ("LOD(area입력)", "F6"): 62034,
                ("LOD(area입력)", "F7"): 126109,
                ("LOD(area입력)", "F8"): 245476,
                ("회수율", "B28"): 32641,
                ("회수율", "B29"): 32573,
                ("회수율", "B30"): 31980,
                ("회수율", "B31"): 92569,
                ("회수율", "B32"): 93419,
                ("회수율", "B33"): 92130,
                ("회수율", "B34"): 223606,
                ("회수율", "B35"): 223348,
                ("회수율", "B36"): 221798,
                ("LOD(area입력)", "F19"): 10000,
                ("LOD(area입력)", "F20"): 4605,
            }, mapped)
            self.assertFalse(any(
                sheet == "LOD(area입력)" and cell.startswith("E")
                for sheet, cell in mapped
            ))

            output = root / "isopropyl-acetate-result.xlsx"
            with patch.dict(os.environ, {"LOCALAPPDATA": str(root / "localappdata")}):
                created = CreateExcelExportService(
                    database,
                    preview_service,
                    XlsxXmlCellWriter(),
                    XlsxWorkbookValidator(),
                    object(),
                    recalculate_with_excel=False,
                ).create(
                    saved.batch_id,
                    TEMPLATE,
                    output,
                    "A",
                    "ISOPROPYL-ACETATE-REGRESSION",
                )
            self.assertTrue(created.validation_passed)
            self.assertEqual(16, created.mapped_cell_count)
            with ZipFile(output) as archive:
                self.assertIsNone(archive.testzip())
            result = inspector.inspect(output)
            for (sheet, address), expected in mapped.items():
                self.assertEqual(expected, result.cell(sheet, address).value)


if __name__ == "__main__":
    unittest.main()
