from copy import deepcopy
import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from honyu_app.application.create_excel_export import CreateExcelExportService
from honyu_app.application.preview_excel_export import (
    DMF_DMA_PROFILE,
    MEK_PROFILE,
    PreviewExcelExportService,
)
from honyu_app.domain.commands import SaveAnalysisBatchCommand
from honyu_app.domain.enums import StdMethod
from honyu_app.domain.models import ExcelPreviewResult
from honyu_app.infrastructure.database.mock_database_service import MockDatabaseService
from honyu_app.infrastructure.excel.excel_recalculator import ExcelComRecalculator
from honyu_app.infrastructure.excel.workbook_inspector import XlsxTemplateInspector
from honyu_app.infrastructure.excel.workbook_validator import XlsxWorkbookValidator
from honyu_app.infrastructure.excel.xml_cell_writer import XlsxXmlCellWriter
from honyu_app.infrastructure.pdf.labsolutions_parser import LabSolutionsParser


def _actual_file(name: str) -> Path:
    configured = os.environ.get("DMF_DMA_335_336_TEST_DIR")
    directories = (
        Path(configured) if configured else None,
        Path(
            r"\\172.30.1.100\data\분석결과(사업장별)★"
            r"\자동화프로그램@@\09.23"
        ),
        Path(__file__).resolve().parents[3] / "TEST" / "DMF,DMA 335-336",
    )
    return next(
        (
            directory / name
            for directory in directories
            if directory is not None and (directory / name).is_file()
        ),
        Path(),
    )


ACTUAL_PDF = _actual_file("DMF 335-336@완료.pdf")
ACTUAL_XLSX = _actual_file("(DMF,DMA) 335-336.xlsx")
ACTUAL_MEK_XLSX = (
    Path(__file__).resolve().parents[3] / "TEST" / "(MEK).xlsx"
)


@unittest.skipUnless(
    ACTUAL_PDF.is_file() and ACTUAL_XLSX.is_file(),
    "DMF,DMA 335-336 실제 PDF/Excel이 없습니다.",
)
class DmfDma335336ActualRegressionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.batch = LabSolutionsParser().parse(
            ACTUAL_PDF,
            analysis_type="DMF,DMA",
            analysis_no_start=335,
            analysis_no_end=336,
        )

    def test_actual_headers_detect_dmf_dma_profile(self) -> None:
        snapshot = XlsxTemplateInspector().inspect(ACTUAL_XLSX)
        self.assertEqual(
            snapshot.cell("LOD(area입력)", "F3").value,
            "Dimethylformamide",
        )
        self.assertEqual(
            snapshot.cell("LOD(area입력)", "I3").value,
            "N,N-Dimethyl acetamide",
        )
        result = ExcelPreviewResult(ACTUAL_XLSX, StdMethod.A)
        profile = PreviewExcelExportService._template_profile(snapshot, result)
        self.assertIs(profile, DMF_DMA_PROFILE)
        self.assertEqual(result.issues, [])

    def test_pdf_db_preview_and_final_xlsx(self) -> None:
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            database = MockDatabaseService(root / "actual.db")
            saved = database.save_analysis_batch(
                SaveAnalysisBatchCommand(deepcopy(self.batch))
            )
            preview_service = PreviewExcelExportService(
                database, XlsxTemplateInspector()
            )
            preview = preview_service.preview(
                saved.batch_id, ACTUAL_XLSX, StdMethod.A
            )
            self.assertEqual(preview.mapped_count, 27)
            self.assertEqual(preview.excluded_count, 38)
            self.assertEqual(preview.error_count, 0)
            self.assertFalse(
                any(
                    issue.code == "TEMPLATE_PROFILE_MISMATCH"
                    for issue in preview.issues
                )
            )

            output = root / "DMF-DMA-335-336-result.xlsx"
            created = CreateExcelExportService(
                database,
                preview_service,
                XlsxXmlCellWriter(),
                XlsxWorkbookValidator(),
                ExcelComRecalculator(),
            ).create(
                saved.batch_id,
                ACTUAL_XLSX,
                output,
                StdMethod.A,
                "actual-regression",
            )
            self.assertTrue(created.validation_passed)
            self.assertEqual(created.mapped_cell_count, 27)
            self.assertTrue(output.is_file())


@unittest.skipUnless(ACTUAL_MEK_XLSX.is_file(), "실제 MEK Excel이 없습니다.")
class MekActualTemplateRegressionTests(unittest.TestCase):
    def test_actual_mek_template_keeps_mek_profile(self) -> None:
        snapshot = XlsxTemplateInspector().inspect(ACTUAL_MEK_XLSX)
        self.assertEqual(
            snapshot.cell("LOD(area입력)", "E4").value,
            "Methyl ethyl ketone",
        )
        self.assertEqual(snapshot.cell("LOD(area입력)", "F4").value, "area")
        result = ExcelPreviewResult(ACTUAL_MEK_XLSX, StdMethod.A)
        profile = PreviewExcelExportService._template_profile(snapshot, result)
        self.assertIs(profile, MEK_PROFILE)
        self.assertEqual(result.issues, [])


if __name__ == "__main__":
    unittest.main()
