from copy import deepcopy
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from honyu_app.application.preview_excel_export import PreviewExcelExportService
from honyu_app.domain.commands import SaveAnalysisBatchCommand
from honyu_app.domain.enums import SampleType
from honyu_app.infrastructure.database.mock_database_service import MockDatabaseService
from honyu_app.infrastructure.excel.workbook_inspector import XlsxTemplateInspector
from honyu_app.infrastructure.pdf.labsolutions_parser import LabSolutionsParser


ROOT = Path(__file__).resolve().parents[3] / "TEST"
CASES = (
    ("혼유", "혼유 39-73 병합완료.pdf", "(혼유) 틀.xlsx", 39, 73),
    ("1컬럼혼유", "1컬럼혼유 39-73병합완료.pdf", "(1컬럼혼유-틀).xlsx", 39, 73),
)


class ActualRecoveryMappingRegressionTests(unittest.TestCase):
    def test_original_and_renumbered_pdf_extractions_write_identical_cells(self):
        for analysis, pdf, xlsx, start, end in CASES:
            with self.subTest(analysis=analysis):
                if not (ROOT / pdf).is_file() or not (ROOT / xlsx).is_file():
                    self.skipTest(f"Local fixture unavailable: {pdf}")
                source = LabSolutionsParser().parse(
                    ROOT / pdf, analysis_type=analysis,
                    analysis_no_start=start, analysis_no_end=end,
                )
                ids = list(dict.fromkeys(s.replicate_no for s in source.samples
                                         if s.sample_type is SampleType.RECOVERY))
                self.assertEqual([1, 2, 3], ids)
                with TemporaryDirectory() as directory:
                    database = MockDatabaseService(Path(directory) / "recovery.db")
                    saved = database.save_analysis_batch(SaveAnalysisBatchCommand(deepcopy(source)))
                    service = PreviewExcelExportService(database, XlsxTemplateInspector())
                    original = service.preview(saved.batch_id, ROOT / xlsx, "A")
                    self.assertTrue(original.can_generate, original.issues)
                    for new_ids in ((3, 4, 5), (7, 8, 9), (2, 4, 6)):
                        replacement = deepcopy(source)
                        mapping = dict(zip(ids, new_ids))
                        for sample in replacement.samples:
                            if sample.sample_type is SampleType.RECOVERY:
                                sample.replicate_no = mapping[sample.replicate_no]
                                sample.sample_name_raw = (sample.sample_name_raw[0]
                                                          + str(sample.replicate_no))
                                sample.sample_name_normalized = sample.sample_name_raw
                        database.replace_analysis_batch(saved.batch_id,
                                                        SaveAnalysisBatchCommand(replacement))
                        preview = service.preview(saved.batch_id, ROOT / xlsx, "A")
                        self.assertTrue(preview.can_generate, preview.issues)
                        self.assertEqual((original.mapped_count, original.excluded_count, 0),
                                         (preview.mapped_count, preview.excluded_count,
                                          preview.error_count))
                        self.assertEqual(
                            [(r.status, r.target_sheet, r.target_cell, r.applied_area)
                             for r in original.rows],
                            [(r.status, r.target_sheet, r.target_cell, r.applied_area)
                             for r in preview.rows],
                        )
