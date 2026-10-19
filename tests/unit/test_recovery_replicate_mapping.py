from pathlib import Path
import unittest

from honyu_app.application.preview_excel_export import PreviewExcelExportService
from honyu_app.domain.enums import ConcentrationLevel, SampleType, ExcelPreviewStatus
from honyu_app.domain.models import Sample, ExcelPreviewResult
from tests.unit.test_excel_preview import (
    FakeDatabase, FakeTemplateService, snapshot, batch, peak,
)


def recovery_samples(ids):
    return [Sample(
        page, f"{label}{source_id}", f"{label}{source_id}", SampleType.RECOVERY,
        concentration_level=level, replicate_no=source_id,
        peaks=[peak(1, page * 100, "acetone")],
    ) for page, (source_id, label, level) in enumerate(
        ((source_id, label, level) for source_id in ids for label, level in (
            ("저", ConcentrationLevel.LOW), ("중", ConcentrationLevel.MID),
            ("고", ConcentrationLevel.HIGH))), 1)]


class RecoveryReplicateMappingTests(unittest.TestCase):
    def preview(self, samples):
        # Bypass the complete-set helper used by older cell-specific fixtures.
        source = batch([])
        source.samples = samples
        return PreviewExcelExportService(
            FakeDatabase(), FakeTemplateService(snapshot())
        ).preview_batch(source, Path("template.xlsx"), "A")

    def test_source_ids_map_to_cells_in_first_appearance_order(self):
        for ids in ((1, 2, 3), (3, 4, 5), (7, 8, 9), (2, 4, 6),
                    (5, 6, 7), (10, 20, 30), (9, 2, 7)):
            with self.subTest(ids=ids):
                samples = recovery_samples(ids)
                result = self.preview(samples)
                self.assertTrue(result.can_generate, result.issues)
                mapped = [r for r in result.rows if r.status is ExcelPreviewStatus.MAPPED]
                self.assertEqual(9, len(mapped))
                for slot, source_id in enumerate(ids):
                    for label, start in (("저", 37), ("중", 40), ("고", 43)):
                        row = next(r for r in mapped if r.sample_name == f"{label}{source_id}")
                        self.assertEqual(f"C{start + slot}", row.target_cell)
                self.assertEqual([i for i in ids for _ in range(3)],
                                 [s.replicate_no for s in samples])

    def test_missing_or_duplicate_level_blocks_recovery_mapping(self):
        for missing in (ConcentrationLevel.MID, ConcentrationLevel.HIGH):
            samples = [s for s in recovery_samples((3, 4, 5))
                       if not (s.replicate_no == 4 and s.concentration_level is missing)]
            result = self.preview(samples)
            self.assertIn("RECOVERY_SET_INCOMPLETE", [i.code for i in result.issues])
            self.assertEqual(0, result.mapped_count)
        samples = recovery_samples((3, 4, 5))
        samples.append(recovery_samples((3,))[0])
        self.assertIn("RECOVERY_SET_INCOMPLETE", [i.code for i in self.preview(samples).issues])

    def test_two_or_four_sets_are_rejected(self):
        for ids in ((3, 4), (3, 4, 5, 6)):
            result = self.preview(recovery_samples(ids))
            self.assertIn("RECOVERY_SET_COUNT_MISMATCH", [i.code for i in result.issues])
            self.assertEqual(0, result.mapped_count)

    def test_nonpositive_replicate_is_rejected(self):
        result = self.preview(recovery_samples((0, 1, 2)))
        self.assertIn("RECOVERY_REPLICATE_INVALID", [i.code for i in result.issues])

    def test_blank_aliases_do_not_enter_recovery_sets(self):
        from honyu_app.infrastructure.pdf.labsolutions_parser import LabSolutionsParser
        samples = recovery_samples((3, 4, 5))
        for name in ("회수율-B", "회수율B", "회수율-BLANK", "회수율_blank"):
            kind, _, _, _, is_blank = LabSolutionsParser()._classify_sample(name)
            samples.append(Sample(1, name, name, kind, is_blank=is_blank))
        result = self.preview(samples)
        self.assertTrue(result.can_generate, result.issues)
        self.assertEqual(9, result.mapped_count)

    def test_db_replacement_preserves_source_ids_and_excel_cells(self):
        from tempfile import TemporaryDirectory
        from honyu_app.domain.commands import SaveAnalysisBatchCommand
        from honyu_app.infrastructure.database.mock_database_service import MockDatabaseService
        with TemporaryDirectory() as directory:
            database = MockDatabaseService(Path(directory) / "recovery.db")
            original = batch([])
            original.samples = recovery_samples((1, 2, 3))
            saved = database.save_analysis_batch(SaveAnalysisBatchCommand(original))
            service = PreviewExcelExportService(database, FakeTemplateService(snapshot()))
            before = service.preview(saved.batch_id, Path("template.xlsx"), "A")
            replacement = batch([])
            replacement.samples = recovery_samples((3, 4, 5))
            database.replace_analysis_batch(saved.batch_id, SaveAnalysisBatchCommand(replacement))
            loaded = database.get_batch_detail(saved.batch_id)
            self.assertEqual([3, 3, 3, 4, 4, 4, 5, 5, 5],
                             [s.replicate_no for s in loaded.samples])
            after = service.preview(saved.batch_id, Path("template.xlsx"), "A")
            self.assertTrue(after.can_generate, after.issues)
            self.assertEqual([(r.target_cell, r.applied_area) for r in before.rows],
                             [(r.target_cell, r.applied_area) for r in after.rows])
