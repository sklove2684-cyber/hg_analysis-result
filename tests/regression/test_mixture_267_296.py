from copy import deepcopy
from decimal import Decimal
import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from honyu_app.application.preview_excel_export import (
    LEGACY_PROFILE,
    PreviewExcelExportService,
)
from honyu_app.domain.commands import SaveAnalysisBatchCommand
from honyu_app.domain.enums import ExcelPreviewStatus, StdMethod
from honyu_app.domain.models import ExcelCellWrite, ExcelPreviewResult
from honyu_app.infrastructure.database.mock_database_service import MockDatabaseService
from honyu_app.infrastructure.excel.workbook_inspector import XlsxTemplateInspector
from honyu_app.infrastructure.excel.workbook_validator import XlsxWorkbookValidator
from honyu_app.infrastructure.excel.xml_cell_writer import XlsxXmlCellWriter
from honyu_app.infrastructure.pdf.labsolutions_parser import LabSolutionsParser


def _actual_file(name: str) -> Path:
    configured = os.environ.get("MIXTURE_267_296_TEST_DIR")
    directories = (
        Path(configured) if configured else None,
        Path(
            r"\\172.30.1.100\data\분석결과(사업장별)★"
            r"\자동화프로그램@@\09.21"
        ),
        Path(__file__).resolve().parents[3] / "TEST" / "혼유 267-296",
    )
    return next(
        (
            directory / name
            for directory in directories
            if directory is not None and (directory / name).is_file()
        ),
        Path(),
    )


ACTUAL_PDF = _actual_file("혼유 267-296@완료.pdf")
ACTUAL_XLSX = _actual_file("(혼유) 267-296.xlsx")


@unittest.skipUnless(
    ACTUAL_PDF.is_file() and ACTUAL_XLSX.is_file(),
    "혼유 267-296 실제 PDF/Excel이 없습니다.",
)
class Mixture267296ActualRegressionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.batch = LabSolutionsParser().parse(
            ACTUAL_PDF,
            analysis_type="혼유",
            analysis_no_start=267,
            analysis_no_end=296,
        )

    def test_pdf_db_preview_and_final_xlsx(self) -> None:
        with TemporaryDirectory() as temporary:
            database = MockDatabaseService(Path(temporary) / "actual.db")
            saved = database.save_analysis_batch(
                SaveAnalysisBatchCommand(deepcopy(self.batch))
            )
            service = PreviewExcelExportService(database, XlsxTemplateInspector())
            previews = {}

            for method in (StdMethod.A, StdMethod.B):
                selection = ExcelPreviewResult(ACTUAL_XLSX, method)
                excluded = service._select_legacy_std_set(
                    self.batch.samples, method, LEGACY_PROFILE, selection
                )
                ambiguous: set[str] = set()
                targets = service._runtime_target_retention_times(
                    self.batch.samples,
                    excluded,
                    method,
                    LEGACY_PROFILE,
                    ambiguous_materials=ambiguous,
                )
                self.assertEqual(targets["o-xylene"], Decimal("10.685"))
                self.assertEqual(ambiguous, set())

                preview = service.preview(saved.batch_id, ACTUAL_XLSX, method)
                previews[method] = preview
                self.assertTrue(preview.can_generate, preview.issues)
                self.assertEqual(preview.error_count, 0)
                self.assertFalse(
                    any(
                        issue.code == "STD_TARGET_RT_NOT_FOUND"
                        and issue.material == "o-xylene"
                        for issue in preview.issues
                    )
                )

            preview = previews[StdMethod.A]
            writes = [
                ExcelCellWrite(row.target_sheet, row.target_cell, row.applied_area)
                for row in preview.rows
                if row.status is ExcelPreviewStatus.MAPPED
            ]
            output = Path(temporary) / "mixture-267-296-result.xlsx"
            XlsxXmlCellWriter().write_copy(ACTUAL_XLSX, output, writes)
            validation = XlsxWorkbookValidator().validate(
                ACTUAL_XLSX,
                output,
                writes,
                after_excel_recalculation=False,
            )
            self.assertTrue(validation.valid, validation.errors)


if __name__ == "__main__":
    unittest.main()
