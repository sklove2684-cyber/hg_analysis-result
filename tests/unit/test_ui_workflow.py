from __future__ import annotations

from copy import deepcopy
import importlib.util
import os
from pathlib import Path
from tempfile import TemporaryDirectory
import time
from types import SimpleNamespace
import unittest
from unittest.mock import patch
from uuid import uuid4


PYSIDE_AVAILABLE = importlib.util.find_spec("PySide6") is not None
SAMPLE_XLSX = Path(__file__).parents[3] / "TEST" / "(혼유) 틀.xlsx"

if PYSIDE_AVAILABLE:
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    from PySide6.QtCore import QEventLoop, QThread, QTimer
    from PySide6.QtWidgets import QApplication, QFileDialog

    from honyu_app.application.preview_excel_export import PreviewExcelExportService
    from honyu_app.application.review_extraction import ReviewExtractionService
    from honyu_app.domain.commands import SaveAnalysisBatchCommand
    from honyu_app.domain.enums import ReviewStatus
    from honyu_app.domain.models import SourceFile
    from honyu_app.infrastructure.database.mock_database_service import MockDatabaseService
    from honyu_app.infrastructure.excel.workbook_inspector import XlsxTemplateInspector
    from honyu_app.infrastructure.pdf.labsolutions_parser import LabSolutionsParser
    from honyu_app.ui.main_window import MainWindow
    from honyu_app.ui.pages.database_page import DatabasePage
    from honyu_app.ui.pages.excel_export_page import ExcelCreationWorker, ExcelExportPage
    from honyu_app.ui.pages.extraction_review_page import ExtractionReviewPage
    from honyu_app.ui.pages.pdf_registration_page import PdfRegistrationPage
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

    def test_excel_creation_worker_keeps_gui_event_loop_responsive(self) -> None:
        class SlowCreateService:
            def create(self, *_args):
                time.sleep(0.2)
                return "created"

        thread = QThread()
        worker = ExcelCreationWorker(
            SlowCreateService(),
            uuid4(),
            Path(self.temp.name) / "template.xlsx",
            Path(self.temp.name) / "result.xlsx",
            "A",
            "test-device",
        )
        worker.moveToThread(thread)
        loop = QEventLoop()
        events = []
        worker.completed.connect(lambda result: events.append(("completed", result)))
        worker.finished.connect(thread.quit)
        worker.finished.connect(loop.quit)
        thread.started.connect(worker.run)
        QTimer.singleShot(25, lambda: events.append(("timer", None)))

        thread.start()
        loop.exec()
        thread.wait(1000)

        self.assertEqual(events[0][0], "timer")
        self.assertEqual(events[1], ("completed", "created"))
        worker.deleteLater()
        thread.deleteLater()

    def test_excel_page_local_storage_mode_does_not_disable_generation(self) -> None:
        local = Path(self.temp.name) / "local_exports"
        company = Path(self.temp.name) / "company_exports"
        local.mkdir()
        company.mkdir()

        class LocalController:
            directory = local
            mode = "local"

            def export_directory(self):
                return self.directory, SimpleNamespace(storage_mode=self.mode)

        controller = LocalController()
        page = ExcelExportPage(self.database, None, None, controller)
        page.refresh_storage_mode()
        page._result = SimpleNamespace(can_generate=True)
        page.output_path.setText(str(local / "result.xlsx"))
        page._update_create_button()

        self.assertIn("로컬 저장 모드", page.storage_mode.text())
        self.assertIn(str(local), page.storage_mode.text())
        self.assertTrue(page.create_button.isEnabled())

        controller.directory = company
        controller.mode = "company"
        page.refresh_storage_mode()
        self.assertIn("회사 공유폴더 저장", page.storage_mode.text())
        self.assertEqual(Path(page.output_path.text()), company / "result.xlsx")
        self.assertTrue(page.create_button.isEnabled())
        page.deleteLater()

    def test_output_filename_changes_from_column_to_alcohol_template(self) -> None:
        exports = Path(self.temp.name) / "exports"
        exports.mkdir()
        column = Path(self.temp.name) / "(1컬럼) 120, 123-130.xlsx"
        alcohol = Path(self.temp.name) / "(알콜2) 74-119 빈양식.xlsx"
        column.touch()
        alcohol.touch()
        page = ExcelExportPage(self.database, None, None)
        page._default_export_directory = exports

        page.template_path.setText(str(column))
        self.assertEqual(
            Path(page.output_path.text()).name,
            "(1컬럼) 120, 123-130_결과.xlsx",
        )
        page.template_path.setText(str(alcohol))
        self.assertEqual(
            Path(page.output_path.text()).name,
            "(알콜2) 74-119_결과.xlsx",
        )
        page.batch_combo.addItem("1컬럼 배치", uuid4())
        page.batch_combo.addItem("알콜 74-119 배치", uuid4())
        page.output_path.setText(str(exports / "(1컬럼) 120, 123-130_결과.xlsx"))
        page.batch_combo.setCurrentIndex(1)
        self.assertEqual(
            Path(page.output_path.text()).name,
            "(알콜2) 74-119_결과.xlsx",
        )
        page.deleteLater()

    def test_output_filename_changes_from_alcohol_to_honyu_and_avoids_collision(self) -> None:
        exports = Path(self.temp.name) / "exports"
        exports.mkdir()
        alcohol = Path(self.temp.name) / "(알콜2) 74-119 빈양식.xlsx"
        honyu = Path(self.temp.name) / "(혼유) 120-130.xlsx"
        alcohol.touch()
        honyu.touch()
        existing = exports / "(혼유) 120-130_결과.xlsx"
        existing.touch()
        page = ExcelExportPage(self.database, None, None)
        page._default_export_directory = exports

        page.template_path.setText(str(alcohol))
        self.assertEqual(Path(page.output_path.text()).name, "(알콜2) 74-119_결과.xlsx")
        page.template_path.setText(str(honyu))
        generated = Path(page.output_path.text()).name
        self.assertRegex(
            generated,
            r"^\(혼유\) 120-130_결과_\d{8}_\d{6}(?:_\d+)?\.xlsx$",
        )
        page.deleteLater()

    def test_analysis_type_batch_code_and_excel_name_follow_each_new_work(self) -> None:
        registration = PdfRegistrationPage(None, LabSolutionsParser(), self.database)
        excel = ExcelExportPage(self.database, None, None)
        exports = Path(self.temp.name) / "exports"
        exports.mkdir()
        excel._default_export_directory = exports
        cases = (
            ("MEK 74-119.pdf", "알콜", "MEK", 74, 119, "(MEK) 74-119.xlsx"),
            (
                "1컬럼혼유 120, 123-130 병합완료.pdf", "MEK", "1컬럼혼유",
                120, 130, "(1컬럼) 120, 123-130.xlsx",
            ),
            (
                "알콜(2) 74-119.pdf", "1컬럼혼유", "(알콜2) IBA,1-BTOH",
                74, 119, "(알콜2) 74-119 빈양식.xlsx",
            ),
        )
        emitted = []
        registration.extraction_ready.connect(emitted.append)
        for index, (pdf_name, stale_type, expected_type, start, end, template_name) in enumerate(cases):
            batch = make_batch(file_hash=str(index + 1) * 64, batch_code="OLD-BATCH")
            batch.source_file = SourceFile(
                original_name=pdf_name,
                full_path=Path(self.temp.name) / pdf_name,
                file_hash=batch.source_file.file_hash,
                file_size=batch.source_file.file_size,
                page_count=batch.source_file.page_count,
            )
            batch.analysis_type = stale_type
            batch.batch_code = f"{stale_type}-OLD"
            batch.analysis_no_start = start
            batch.analysis_no_end = end
            registration._on_extraction_completed(batch)

            self.assertEqual(batch.analysis_type, expected_type)
            self.assertTrue(batch.batch_code.startswith(f"{expected_type}-{start}-{end}-"))
            self.assertIs(emitted[-1], batch)
            template = Path(self.temp.name) / template_name
            template.touch()
            excel.template_path.setText(str(template))
            expected_stem = template.stem.removesuffix(" 빈양식") + "_결과"
            self.assertTrue(Path(excel.output_path.text()).stem.startswith(expected_stem))
        registration.deleteLater()
        excel.deleteLater()

    def test_analysis_type_combo_uses_registry_and_manual_choice_wins(self) -> None:
        from honyu_app.config.analysis_types import ANALYSIS_TYPE_NAMES

        registration = PdfRegistrationPage(None, LabSolutionsParser(), self.database)
        self.assertEqual(
            tuple(registration.analysis_type.itemText(index) for index in range(registration.analysis_type.count())),
            ANALYSIS_TYPE_NAMES,
        )
        registration.analysis_type.setCurrentText("IPA")
        registration._mark_analysis_type_user_selected()
        source = make_batch(file_hash="9" * 64, batch_code="IPA-74-119")
        source.analysis_type = "IPA"
        source.source_file = SourceFile(
            original_name="MEK 74-119.pdf",
            full_path=Path(self.temp.name) / "MEK 74-119.pdf",
            file_hash=source.source_file.file_hash,
            file_size=source.source_file.file_size,
            page_count=source.source_file.page_count,
        )

        registration._on_extraction_completed(source)

        self.assertEqual(source.analysis_type, "IPA")
        self.assertEqual(registration.analysis_type.currentText(), "IPA")
        registration.deleteLater()

    def test_selecting_new_pdf_clears_review_and_excel_work_state(self) -> None:
        registration = PdfRegistrationPage(None, LabSolutionsParser(), self.database)
        review = ExtractionReviewPage(self.service)
        excel = ExcelExportPage(self.database, None, None)
        registration.new_work_started.connect(review.reset_for_new_work)
        registration.new_work_started.connect(excel.reset_for_new_work)
        review.load_batch(make_batch(batch_code="ALCOHOL-OLD"))
        excel.batch_combo.addItem("이전 알콜 배치", uuid4())
        excel.template_path.setText(str(Path(self.temp.name) / "old-alcohol.xlsx"))
        excel.output_path.setText(str(Path(self.temp.name) / "old-result.xlsx"))
        excel._result = SimpleNamespace(can_generate=True)
        new_pdf = Path(self.temp.name) / "MEK 74-119.pdf"
        new_pdf.touch()

        with patch.object(
            QFileDialog, "getOpenFileName", return_value=(str(new_pdf), "PDF 파일 (*.pdf)")
        ):
            registration.choose_pdf()

        self.assertEqual(registration.analysis_type.currentText(), "MEK")
        self.assertEqual((registration.start_no.value(), registration.end_no.value()), (74, 119))
        self.assertIsNone(review._batch)
        self.assertEqual(review.saved_batches.currentIndex(), 0)
        self.assertIsNone(review.saved_batches.currentData())
        self.assertEqual(review.saved_batches.currentText(), "저장된 배치를 선택하세요")
        self.assertEqual(excel.batch_combo.currentIndex(), -1)
        self.assertEqual(excel.template_path.text(), "")
        self.assertEqual(excel.output_path.text(), "")
        self.assertIsNone(excel._result)
        registration.deleteLater()
        review.deleteLater()
        excel.deleteLater()

    def test_changing_analysis_type_for_loaded_pdf_clears_downstream_state(self) -> None:
        registration = PdfRegistrationPage(None, LabSolutionsParser(), self.database)
        review = ExtractionReviewPage(self.service)
        excel = ExcelExportPage(self.database, None, None)
        registration.new_work_started.connect(review.reset_for_new_work)
        registration.new_work_started.connect(excel.reset_for_new_work)
        current_pdf = Path(self.temp.name) / "current.pdf"
        current_pdf.touch()
        registration.pdf_path.setText(str(current_pdf))
        review.load_batch(make_batch(batch_code="STALE-BATCH"))
        excel.batch_combo.addItem("이전 배치", uuid4())
        excel.template_path.setText(str(Path(self.temp.name) / "old.xlsx"))
        excel.output_path.setText(str(Path(self.temp.name) / "old-result.xlsx"))
        excel._result = SimpleNamespace(can_generate=True)

        registration.analysis_type.setCurrentText("DMF,DMA")
        registration._mark_analysis_type_user_selected()

        self.assertTrue(registration._analysis_type_user_selected)
        self.assertIsNone(review._batch)
        self.assertEqual(review.saved_batches.currentIndex(), 0)
        self.assertEqual(excel.batch_combo.currentIndex(), -1)
        self.assertEqual(excel.template_path.text(), "")
        self.assertEqual(excel.output_path.text(), "")
        self.assertIsNone(excel._result)
        registration.deleteLater()
        review.deleteLater()
        excel.deleteLater()

    def test_main_window_does_not_wait_for_slow_shared_folder_check(self) -> None:
        local = Path(self.temp.name) / "exports"
        local.mkdir()
        state = SimpleNamespace(
            connection=SimpleNamespace(
                connected=False,
                storage_mode="local",
                active_base_path=str(local),
                message="로컬 저장 모드",
                attempted_paths=("missing-unc", "missing-z"),
            ),
            workplaces=(),
            recent=SimpleNamespace(workplace=None, year=None, half=None, final_folder=None),
        )

        class SlowController:
            def refresh(self):
                time.sleep(0.2)
                return state

            def export_directory_from_state(self, _state):
                return local, state.connection

        started = time.perf_counter()
        window = MainWindow(SlowController(), LabSolutionsParser(), self.database)
        construction_seconds = time.perf_counter() - started
        # The mocked synchronous network wait is 0.20s.  Keep the assertion
        # below that boundary while allowing normal Windows/CI scheduler jitter.
        self.assertLess(construction_seconds, 0.20)
        self.assertIn("확인 대기", window._excel_page.storage_mode.text())

        events = []
        loop = QEventLoop()
        window.start_background_initialization()
        window._storage_thread.finished.connect(loop.quit)
        QTimer.singleShot(25, lambda: events.append("gui-responsive"))
        loop.exec()

        self.assertEqual(events, ["gui-responsive"])
        self.assertIn("로컬 저장 모드", window._excel_page.storage_mode.text())
        window.deleteLater()

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
        self.assertEqual(page.saved_batches.count(), 2)
        self.assertIsNone(page.saved_batches.currentData())
        page.saved_batches.setCurrentIndex(1)
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
