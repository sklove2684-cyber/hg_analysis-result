from __future__ import annotations

from copy import deepcopy
import importlib.util
import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest


PYSIDE_AVAILABLE = importlib.util.find_spec("PySide6") is not None
SAMPLE_XLSX = Path(__file__).parents[3] / "TEST" / "(혼유) 틀.xlsx"

if PYSIDE_AVAILABLE:
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    from PySide6.QtWidgets import QApplication

    from honyu_app.application.preview_excel_export import PreviewExcelExportService
    from honyu_app.application.review_extraction import ReviewExtractionService
    from honyu_app.domain.commands import SaveAnalysisBatchCommand
    from honyu_app.domain.enums import ReviewStatus
    from honyu_app.infrastructure.database.mock_database_service import MockDatabaseService
    from honyu_app.infrastructure.excel.workbook_inspector import XlsxTemplateInspector
    from honyu_app.ui.pages.database_page import DatabasePage
    from honyu_app.ui.pages.excel_export_page import ExcelExportPage
    from honyu_app.ui.pages.extraction_review_page import ExtractionReviewPage
    from tests.unit.test_mock_database_service import make_batch


@unittest.skipUnless(PYSIDE_AVAILABLE, "PySide6가 설치된 GUI 환경이 아닙니다.")
class ReviewToExcelUiWorkflowTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self) -> None:
        self.temp = TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.database = MockDatabaseService(Path(self.temp.name) / "workflow.db")
        self.service = ReviewExtractionService(self.database)

    def test_new_batch_buttons_advance_review_save_excel(self) -> None:
        batch = make_batch(file_hash="b" * 64, batch_code="NEW-BATCH")
        batch.review_status = ReviewStatus.PENDING
        page = ExtractionReviewPage(self.service)
        page.load_batch(batch)

        self.assertTrue(page.complete_button.isEnabled())
        self.assertFalse(page.save_button.isEnabled())
        self.assertFalse(page.excel_button.isEnabled())

        page.complete_button.click()
        self.assertEqual(batch.review_status, ReviewStatus.REVIEWED)
        self.assertFalse(page.complete_button.isEnabled())
        self.assertTrue(page.save_button.isEnabled())

        page.save_button.click()
        self.assertEqual(batch.review_status, ReviewStatus.SAVED)
        self.assertFalse(page.save_button.isEnabled())
        self.assertTrue(page.excel_button.isEnabled())
        page.deleteLater()

    def test_saved_batch_can_be_reopened_and_sent_to_excel(self) -> None:
        batch = make_batch(file_hash="c" * 64, batch_code="SAVED-BATCH")
        saved = self.database.save_analysis_batch(SaveAnalysisBatchCommand(batch))
        page = ExtractionReviewPage(self.service)
        self.assertEqual(page.saved_batches.count(), 1)
        page.saved_batches.setCurrentIndex(0)
        page.open_saved_batch()

        self.assertEqual(page._batch.batch_id, saved.batch_id)
        self.assertEqual(page._batch.review_status, ReviewStatus.SAVED)
        self.assertFalse(page.complete_button.isEnabled())
        self.assertFalse(page.save_button.isEnabled())
        self.assertTrue(page.excel_button.isEnabled())
        self.assertEqual(page.complete_button.text(), "검토 완료됨")
        self.assertEqual(page.save_button.text(), "DB 저장됨")
        requested = []
        page.excel_requested.connect(requested.append)
        page.excel_button.click()
        self.assertEqual(requested[0].batch_id, saved.batch_id)
        page.deleteLater()

    def test_database_page_routes_saved_data_to_review_and_excel(self) -> None:
        batch = make_batch(file_hash="d" * 64, batch_code="DB-BATCH")
        saved = self.database.save_analysis_batch(SaveAnalysisBatchCommand(batch))
        page = DatabasePage(self.database)
        review_requests = []
        excel_requests = []
        page.review_requested.connect(review_requests.append)
        page.excel_requested.connect(excel_requests.append)

        self.assertEqual(page.table.rowCount(), 1)
        page.table.selectRow(0)
        page.open_review()
        page.open_excel()
        self.assertEqual(review_requests[0].batch_id, saved.batch_id)
        self.assertEqual(excel_requests[0].batch_id, saved.batch_id)
        page.deleteLater()

    @unittest.skipUnless(SAMPLE_XLSX.is_file(), "샘플 Excel 양식이 없습니다.")
    def test_excel_page_receives_saved_batch_and_enables_generation_after_preview(self) -> None:
        batch = make_batch(file_hash="e" * 64, batch_code="EXCEL-BATCH")
        saved = self.database.save_analysis_batch(SaveAnalysisBatchCommand(batch))
        loaded = self.database.get_batch_detail(saved.batch_id)
        preview = PreviewExcelExportService(self.database, XlsxTemplateInspector())
        page = ExcelExportPage(self.database, preview, None)
        page.load_batch(loaded)
        output = Path(self.temp.name) / "result.xlsx"
        page.template_path.setText(str(SAMPLE_XLSX))
        page.output_path.setText(str(output))
        page.preview()

        self.assertEqual(page.batch_combo.currentData(), saved.batch_id)
        self.assertIsNotNone(page._result)
        self.assertTrue(page._result.can_generate)
        self.assertEqual(page._result.mapped_count, 1)
        self.assertTrue(page.create_button.isEnabled())
        page.deleteLater()

    @unittest.skipUnless(SAMPLE_XLSX.is_file(), "샘플 Excel 양식이 없습니다.")
    def test_existing_result_gets_a_new_name_and_keeps_create_enabled(self) -> None:
        batch = make_batch(file_hash="f" * 64, batch_code="EXISTING-OUTPUT")
        saved = self.database.save_analysis_batch(SaveAnalysisBatchCommand(batch))
        preview = PreviewExcelExportService(self.database, XlsxTemplateInspector())
        page = ExcelExportPage(self.database, preview, None)
        page.load_batch(self.database.get_batch_detail(saved.batch_id))
        existing = Path(self.temp.name) / "result.xlsx"
        existing.touch()
        page.template_path.setText(str(SAMPLE_XLSX))
        page.output_path.setText(str(existing))

        page.preview()

        renamed = Path(page.output_path.text())
        self.assertNotEqual(renamed, existing)
        self.assertFalse(renamed.exists())
        self.assertTrue(page.create_button.isEnabled())
        page.deleteLater()


if __name__ == "__main__":
    unittest.main()
