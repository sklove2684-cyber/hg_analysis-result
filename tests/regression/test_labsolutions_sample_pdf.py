from pathlib import Path
from tempfile import TemporaryDirectory
from copy import deepcopy
import unittest

from honyu_app.application.review_extraction import ReviewExtractionService
from honyu_app.application.preview_excel_export import PreviewExcelExportService
from honyu_app.domain.enums import ExcelPreviewStatus, ExcludeReason, SampleType, StdMethod
from honyu_app.domain.errors import ExtractionCancelledError
from honyu_app.domain.queries import BatchSearchQuery
from honyu_app.infrastructure.database.mock_database_service import MockDatabaseService
from honyu_app.infrastructure.excel.workbook_inspector import XlsxTemplateInspector
from honyu_app.infrastructure.excel.workbook_validator import XlsxWorkbookValidator
from honyu_app.infrastructure.excel.xml_cell_writer import XlsxXmlCellWriter
from honyu_app.domain.models import ExcelCellWrite
from honyu_app.infrastructure.pdf.labsolutions_parser import LabSolutionsParser


SAMPLE_PDF = (
    Path(__file__).resolve().parents[3]
    / "TEST"
    / "혼유 39-73 병합완료.pdf"
)
SAMPLE_XLSX = Path(__file__).resolve().parents[3] / "TEST" / "(혼유) 틀.xlsx"


@unittest.skipUnless(SAMPLE_PDF.is_file(), f"샘플 PDF 없음: {SAMPLE_PDF}")
class LabSolutionsSamplePdfRegressionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.batch = LabSolutionsParser().parse(
            SAMPLE_PDF,
            analysis_type="혼유",
            analysis_no_start=39,
            analysis_no_end=73,
        )

    def test_page_sample_order(self) -> None:
        self.assertEqual(
            [sample.sample_name_normalized for sample in self.batch.samples],
            [
                "BLANK", "STD1", "STD2", "STD3", "STD4", "STD5", "STD6",
                "회수율-B", "저1", "중1", "고1", "회수율-B", "저2", "중2",
                "고2", "회수율-B", "저3", "중3", "고3",
            ],
        )

    def test_sample_classification(self) -> None:
        self.assertEqual(self.batch.samples[0].sample_type, SampleType.BLANK)
        self.assertTrue(all(s.sample_type is SampleType.STD for s in self.batch.samples[1:7]))
        self.assertEqual(self.batch.samples[8].sample_type, SampleType.RECOVERY)
        self.assertEqual(self.batch.samples[8].replicate_no, 1)
        self.assertEqual(self.batch.samples[16].replicate_no, 3)

    def test_peak_table_columns_stay_aligned_with_unnamed_peak(self) -> None:
        std2 = self.batch.samples[2]
        residual = std2.peaks[1]
        self.assertEqual(residual.peak_no, 2)
        self.assertEqual(str(residual.retention_time), "2.519")
        self.assertEqual(residual.area_raw, 1616)
        self.assertIsNone(residual.material_raw)
        self.assertEqual(residual.exclude_reason, ExcludeReason.UNNAMED_PEAK)
        cs2 = std2.peaks[2]
        self.assertEqual(cs2.peak_no, 3)
        self.assertEqual(cs2.material_standard, "CS2")

    def test_known_material_and_integer_area(self) -> None:
        std1_first = self.batch.samples[1].peaks[0]
        self.assertEqual(std1_first.material_raw, "헥산")
        self.assertEqual(std1_first.material_standard, "n-hexane")
        self.assertEqual(std1_first.area_raw, 24350)
        self.assertIsInstance(std1_first.area_raw, int)

    def test_unnamed_residual_peaks_do_not_raise_review_warning(self) -> None:
        self.assertEqual(self.batch.warning_count, 0)

    def test_dibk_peaks_remain_separate_and_grouped(self) -> None:
        std4_dibk = [
            peak for peak in self.batch.samples[4].peaks
            if peak.material_standard == "DIBK"
        ]
        self.assertEqual([peak.peak_group_no for peak in std4_dibk], [1, 2, 3])
        self.assertEqual(len({peak.peak_id for peak in std4_dibk}), 3)

    def test_filename_range_and_hash(self) -> None:
        self.assertEqual(
            LabSolutionsParser.extract_analysis_range(SAMPLE_PDF.name), (39, 73)
        )
        self.assertEqual(len(self.batch.source_file.file_hash), 64)

    def test_extraction_can_be_cancelled_between_pages(self) -> None:
        with self.assertRaises(ExtractionCancelledError):
            LabSolutionsParser().parse(
                SAMPLE_PDF,
                analysis_type="혼유",
                analysis_no_start=39,
                analysis_no_end=73,
                cancel_check=lambda: True,
            )

    def test_approved_sample_can_be_reviewed_and_saved_without_peak_count_rule(self) -> None:
        batch = deepcopy(self.batch)
        with TemporaryDirectory() as temp:
            database = MockDatabaseService(Path(temp) / "mock.db")
            review = ReviewExtractionService(database)
            review.complete_review(batch)
            result = review.save_batch(batch)
            loaded = database.get_batch_detail(result.batch_id)
            self.assertEqual(len(loaded.samples), 19)
            self.assertEqual(
                [peak.peak_no for peak in loaded.samples[2].peaks],
                list(range(1, len(loaded.samples[2].peaks) + 1)),
            )
            self.assertEqual(len(database.search_batches(BatchSearchQuery())), 1)

    @unittest.skipUnless(SAMPLE_XLSX.is_file(), f"샘플 Excel 없음: {SAMPLE_XLSX}")
    def test_sample_batch_preview_selects_only_dibk_area_top_two(self) -> None:
        batch = deepcopy(self.batch)
        with TemporaryDirectory() as temp:
            database = MockDatabaseService(Path(temp) / "mock.db")
            review = ReviewExtractionService(database)
            review.complete_review(batch)
            saved = review.save_batch(batch)
            preview = PreviewExcelExportService(
                database, XlsxTemplateInspector()
            ).preview(saved.batch_id, SAMPLE_XLSX, StdMethod.A)

        overflow = [
            row for row in preview.rows
            if row.exclude_reason == ExcludeReason.DIBK_AREA_NOT_TOP2.value
        ]
        selected_by_sample: dict[str, list] = {}
        for row in preview.rows:
            if row.material == "DIBK" and row.status is ExcelPreviewStatus.MAPPED:
                selected_by_sample.setdefault(row.sample_name, []).append(row)
        self.assertTrue(preview.can_generate)
        self.assertGreater(len(overflow), 0)
        self.assertTrue(all(len(rows) <= 2 for rows in selected_by_sample.values()))

    @unittest.skipUnless(SAMPLE_XLSX.is_file(), f"샘플 Excel 없음: {SAMPLE_XLSX}")
    def test_all_real_sample_mappings_write_without_structure_damage(self) -> None:
        batch = deepcopy(self.batch)
        with TemporaryDirectory() as temp:
            database = MockDatabaseService(Path(temp) / "mock.db")
            review = ReviewExtractionService(database)
            review.complete_review(batch)
            saved = review.save_batch(batch)
            preview = PreviewExcelExportService(
                database, XlsxTemplateInspector()
            ).preview(saved.batch_id, SAMPLE_XLSX, StdMethod.A)
            writes = [
                ExcelCellWrite(row.target_sheet, row.target_cell, row.applied_area)
                for row in preview.rows
                if row.status is ExcelPreviewStatus.MAPPED
            ]
            output = Path(temp) / "real-sample-pre-com.xlsx"
            XlsxXmlCellWriter().write_copy(SAMPLE_XLSX, output, writes)
            validation = XlsxWorkbookValidator().validate(
                SAMPLE_XLSX,
                output,
                writes,
                after_excel_recalculation=False,
            )

        self.assertTrue(preview.can_generate)
        self.assertEqual(len(writes), 196)
        self.assertTrue(validation.valid, validation.errors)
