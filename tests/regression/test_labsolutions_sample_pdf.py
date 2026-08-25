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
ONE_COLUMN_PDF = (
    Path(__file__).resolve().parents[3]
    / "TEST"
    / "1컬럼혼유 39-73병합완료.pdf"
)
ONE_COLUMN_XLSX = (
    Path(__file__).resolve().parents[3]
    / "TEST"
    / "(1컬럼혼유-틀).xlsx"
)
ALCOHOL_CONTINUATION_PDF = (
    Path(__file__).resolve().parents[3]
    / "TEST"
    / "알콜(2) 74-119.pdf"
)
ALCOHOL_XLSX = (
    Path(__file__).resolve().parents[3]
    / "TEST"
    / "(알콜2) 74-119.xlsx"
)
HONYU_120_167_PDF = next(
    (
        path
        for path in (
            Path(__file__).resolve().parents[3] / "TEST" / "혼유 120-167.pdf",
            Path(__file__).resolve().parents[3] / "TEST" / "혼유 120-167 병합.pdf",
        )
        if path.is_file()
    ),
    Path(__file__).resolve().parents[3] / "TEST" / "혼유 120-167.pdf",
)


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


@unittest.skipUnless(
    HONYU_120_167_PDF.is_file(),
    f"혼유 120-167 PDF 없음: {HONYU_120_167_PDF}",
)
class Honyu120167ContinuationRegressionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.batch = LabSolutionsParser().parse(
            HONYU_120_167_PDF,
            analysis_type="혼유",
            analysis_no_start=120,
            analysis_no_end=167,
        )

    def test_samples_two_and_three_merge_their_continuation_pages(self) -> None:
        sample_two = [
            sample for sample in self.batch.samples
            if sample.sample_name_normalized == "2"
        ]
        sample_three = [
            sample for sample in self.batch.samples
            if sample.sample_name_normalized == "3"
        ]

        self.assertEqual(len(sample_two), 1)
        self.assertEqual(len(sample_three), 1)
        self.assertEqual({peak.source_page for peak in sample_two[0].peaks}, {39, 40})
        self.assertEqual({peak.source_page for peak in sample_three[0].peaks}, {41, 42})
        self.assertEqual(
            [sample_two[0].peaks[0].peak_no, sample_two[0].peaks[-1].peak_no],
            [1, 82],
        )
        self.assertEqual(
            [sample_three[0].peaks[0].peak_no, sample_three[0].peaks[-1].peak_no],
            [1, 65],
        )

    def test_sample_126_is_single_and_126b_remains_a_separate_blank(self) -> None:
        sample_126 = [
            sample for sample in self.batch.samples
            if sample.sample_name_normalized == "126"
        ]
        sample_126b = [
            sample for sample in self.batch.samples
            if sample.sample_name_normalized == "126B"
        ]

        self.assertEqual(len(sample_126), 1)
        self.assertEqual(len(sample_126b), 1)
        self.assertEqual(sample_126[0].page_no, 36)
        self.assertEqual(sample_126b[0].page_no, 37)
        self.assertEqual(sample_126b[0].sample_type, SampleType.RECOVERY_BLANK)

    def test_continuation_pages_are_not_saved_as_separate_samples(self) -> None:
        batch = deepcopy(self.batch)
        with TemporaryDirectory() as temp:
            database = MockDatabaseService(Path(temp) / "continuation.db")
            ReviewExtractionService.complete_review(batch)
            saved = ReviewExtractionService(database).save_batch(batch)
            loaded = database.get_batch_detail(saved.batch_id)

        self.assertEqual(len(loaded.samples), len(self.batch.samples))
        self.assertEqual(
            len([s for s in loaded.samples if s.sample_name_normalized == "2"]),
            1,
        )
        self.assertEqual(
            len([s for s in loaded.samples if s.sample_name_normalized == "3"]),
            1,
        )
        self.assertNotIn(40, {sample.page_no for sample in loaded.samples})
        self.assertNotIn(42, {sample.page_no for sample in loaded.samples})


@unittest.skipUnless(
    ONE_COLUMN_PDF.is_file() and ONE_COLUMN_XLSX.is_file(),
    "1컬럼 PDF 또는 Excel 양식이 없습니다.",
)
class OneColumnSampleRegressionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.batch = LabSolutionsParser().parse(
            ONE_COLUMN_PDF,
            analysis_type="1컬럼혼유",
            analysis_no_start=39,
            analysis_no_end=73,
        )

    def test_materials_are_registered_without_review_warning(self) -> None:
        self.assertEqual(len(self.batch.samples), 19)
        self.assertEqual(sum(len(sample.peaks) for sample in self.batch.samples), 82)
        self.assertEqual(self.batch.warning_count, 0)
        self.assertEqual(
            {
                peak.material_standard
                for sample in self.batch.samples
                for peak in sample.peaks
                if peak.material_standard not in {None, "CS2"}
            },
            {"methyl acetate", "c-hexane", "n-heptane", "isobutyl acetate"},
        )

    def test_real_batch_maps_to_one_column_template(self) -> None:
        batch = deepcopy(self.batch)
        with TemporaryDirectory() as temp:
            database = MockDatabaseService(Path(temp) / "one-column.db")
            review = ReviewExtractionService(database)
            review.complete_review(batch)
            saved = review.save_batch(batch)
            preview = PreviewExcelExportService(
                database, XlsxTemplateInspector()
            ).preview(saved.batch_id, ONE_COLUMN_XLSX, StdMethod.A)

        mapped = [
            row for row in preview.rows
            if row.status is ExcelPreviewStatus.MAPPED
        ]
        residual = [
            row for row in preview.rows
            if row.exclude_reason == ExcludeReason.MATERIAL_RT_NOT_CLOSEST.value
        ]
        self.assertTrue(preview.can_generate)
        self.assertEqual(len(mapped), 56)
        self.assertEqual(
            (mapped[0].target_sheet, mapped[0].target_cell, mapped[0].applied_area),
            ("area입력", "G5", 8127),
        )
        self.assertEqual(
            [(row.sample_name, row.material, row.applied_area) for row in residual],
            [("STD5", "c-hexane", 1046)],
        )


@unittest.skipUnless(
    ALCOHOL_CONTINUATION_PDF.is_file(),
    f"연속 페이지 샘플 PDF 없음: {ALCOHOL_CONTINUATION_PDF}",
)
class LabSolutionsContinuationPageRegressionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.batch = LabSolutionsParser().parse(
            ALCOHOL_CONTINUATION_PDF,
            analysis_type="(알콜2) IBA,1-BTOH",
            analysis_no_start=74,
            analysis_no_end=119,
        )

    def test_continuation_pages_are_merged_into_the_previous_sample(self) -> None:
        self.assertEqual(self.batch.source_file.page_count, 63)
        self.assertEqual(len(self.batch.samples), 57)
        self.assertTrue(
            {29, 34, 38, 43, 49, 62}.isdisjoint(
                sample.page_no for sample in self.batch.samples
            )
        )

    def test_sample_90_keeps_all_peaks_from_pages_28_and_29(self) -> None:
        sample = next(
            item for item in self.batch.samples if item.worker_match_key == "90"
        )
        self.assertEqual(len(sample.peaks), 33)
        self.assertEqual([sample.peaks[0].peak_no, sample.peaks[-1].peak_no], [1, 33])
        self.assertEqual({peak.source_page for peak in sample.peaks}, {28, 29})

    def test_alcohol_materials_are_registered_without_review_warnings(self) -> None:
        self.assertEqual(self.batch.warning_count, 0)
        registered = [
            peak
            for sample in self.batch.samples
            for peak in sample.peaks
            if peak.material_standard in {"IBA", "n-BTOH"}
        ]
        self.assertEqual(len(registered), 138)
        self.assertTrue(all(peak.include_for_excel for peak in registered))

    @unittest.skipUnless(ALCOHOL_XLSX.is_file(), f"알콜 Excel 양식 없음: {ALCOHOL_XLSX}")
    def test_alcohol_batch_maps_to_its_template_by_closest_retention_time(self) -> None:
        batch = deepcopy(self.batch)
        with TemporaryDirectory() as temp:
            database = MockDatabaseService(Path(temp) / "alcohol.db")
            review = ReviewExtractionService(database)
            review.complete_review(batch)
            saved = review.save_batch(batch)
            preview = PreviewExcelExportService(
                database, XlsxTemplateInspector()
            ).preview(saved.batch_id, ALCOHOL_XLSX, StdMethod.A)

        mapped = [
            row for row in preview.rows
            if row.status is ExcelPreviewStatus.MAPPED
        ]
        sample_90 = [row for row in mapped if row.sample_name.startswith("90-")]
        self.assertTrue(preview.can_generate, preview.issues)
        self.assertEqual((preview.mapped_count, preview.excluded_count, preview.error_count), (70, 509, 0))
        self.assertFalse(
            any(issue.code == "UNSUPPORTED_MATERIAL" for issue in preview.issues)
        )
        self.assertEqual(
            [
                (row.material, str(row.retention_time), row.applied_area, row.target_cell)
                for row in sample_90
            ],
            [
                ("IBA", "3.352", 781127, "F37"),
                ("n-BTOH", "3.833", 443235, "I37"),
            ],
        )
