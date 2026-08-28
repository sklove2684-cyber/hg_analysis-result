from __future__ import annotations

import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from honyu_app.application.review_extraction import ReviewExtractionService
from honyu_app.config.analysis_types import infer_analysis_type
from honyu_app.domain.enums import ExcludeReason
from honyu_app.domain.queries import BatchSearchQuery
from honyu_app.infrastructure.database.mock_database_service import MockDatabaseService
from honyu_app.infrastructure.pdf.labsolutions_parser import LabSolutionsParser


def _candidate_directories() -> tuple[Path, ...]:
    configured = os.environ.get("HONYU_PHENOL_TEST_DIR")
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


PDF = _find_file("(페놀) 256-305.pdf")


@unittest.skipUnless(PDF.is_file(), "페놀 실제 PDF가 없습니다.")
class PhenolAnalysisPriorityActualFileTests(unittest.TestCase):
    @staticmethod
    def _parse(analysis_type: str):
        return LabSolutionsParser().parse(
            PDF,
            analysis_type=analysis_type,
            analysis_no_start=256,
            analysis_no_end=305,
        )

    def test_filename_phenol_wins_over_detected_meoh(self) -> None:
        batch = self._parse("페놀")
        method_filenames = tuple(
            sample.method_filename for sample in batch.samples if sample.method_filename
        )
        materials = tuple(
            peak.material_standard or peak.material_raw or ""
            for sample in batch.samples
            for peak in sample.peaks
        )

        self.assertIn("Methanol", materials)
        self.assertIn("Phenol", materials)
        self.assertEqual(
            infer_analysis_type(PDF.name, method_filenames, materials), "페놀"
        )
        self.assertEqual(batch.warning_count, 0)
        self.assertEqual(
            sum(
                peak.exclude_reason is ExcludeReason.UNKNOWN_MATERIAL
                for sample in batch.samples
                for peak in sample.peaks
            ),
            0,
        )
        methanol = [
            peak
            for sample in batch.samples
            for peak in sample.peaks
            if peak.material_standard == "Methanol"
        ]
        self.assertEqual(len(methanol), 53)
        self.assertFalse(any(peak.include_for_excel for peak in methanol))
        self.assertEqual(
            sum(
                peak.exclude_reason
                is ExcludeReason.MATERIAL_NOT_SUPPORTED_FOR_ANALYSIS
                for peak in methanol
            ),
            35,
        )

    def test_wrong_methanol_batch_is_atomically_replaced_by_phenol_batch(self) -> None:
        with TemporaryDirectory() as temporary:
            database = MockDatabaseService(Path(temporary) / "phenol-replacement.db")
            review = ReviewExtractionService(database)

            wrong = self._parse("메탄올A")
            review.complete_review(wrong)
            saved_wrong = review.save_batch(wrong)

            corrected = self._parse("페놀")
            corrected.replacement_for_batch_id = saved_wrong.batch_id
            review.complete_review(corrected)
            saved_corrected = review.save_batch(corrected)

            self.assertEqual(saved_corrected.batch_id, saved_wrong.batch_id)
            rows = database.search_batches(BatchSearchQuery())
            self.assertEqual(len(rows), 1)
            self.assertEqual(rows[0].analysis_type, "페놀")
            loaded = database.get_batch_detail(saved_wrong.batch_id)
            self.assertEqual(loaded.analysis_type, "페놀")
            self.assertEqual(loaded.source_file.file_hash, wrong.source_file.file_hash)
            self.assertEqual(len(loaded.samples), len(corrected.samples))
            self.assertEqual(
                sum(len(sample.peaks) for sample in loaded.samples),
                sum(len(sample.peaks) for sample in corrected.samples),
            )
            phenol = [
                peak
                for sample in loaded.samples
                for peak in sample.peaks
                if peak.material_standard == "Phenol"
            ]
            self.assertEqual(len(phenol), 15)
            self.assertTrue(all(peak.include_for_excel for peak in phenol))


if __name__ == "__main__":
    unittest.main()
