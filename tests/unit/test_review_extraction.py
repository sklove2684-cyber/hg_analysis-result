from datetime import datetime, timezone
from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from honyu_app.application.review_extraction import ReviewExtractionService
from honyu_app.domain.enums import ExcludeReason, ReviewStatus, SampleType
from honyu_app.domain.errors import ValidationError
from honyu_app.domain.models import AnalysisBatch, Peak, Sample, SourceFile
from honyu_app.infrastructure.database.mock_database_service import MockDatabaseService


def review_batch() -> AnalysisBatch:
    peak = Peak(
        1, Decimal("2.5"), 100,
        material_raw="새물질", material_standard=None,
        include_for_excel=False, exclude_reason=ExcludeReason.UNKNOWN_MATERIAL,
        source_page=1,
    )
    return AnalysisBatch(
        batch_code="REVIEW-1",
        source_file=SourceFile("a.pdf", Path("a.pdf"), "b" * 64, 100, 1),
        analysis_type="혼유", analysis_no_start=1, analysis_no_end=1,
        parser_name="test", parser_version="1", parser_layout_id="test",
        extracted_at=datetime.now(timezone.utc),
        samples=[Sample(1, "STD1", "STD1", SampleType.STD, peaks=[peak])],
    )


class ReviewExtractionServiceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = TemporaryDirectory()
        self.database = MockDatabaseService(Path(self.temp.name) / "mock.db")
        self.service = ReviewExtractionService(self.database)

    def tearDown(self) -> None:
        self.temp.cleanup()

    def test_unknown_material_blocks_review(self) -> None:
        with self.assertRaisesRegex(
            ValidationError,
            r"미등록 물질 Peak 1개.*물질별: 새물질 1개.*페이지: \[1\]",
        ):
            self.service.complete_review(review_batch())

    def test_user_can_map_then_complete_and_save(self) -> None:
        batch = review_batch()
        peak = batch.samples[0].peaks[0]
        self.service.set_material_mapping(batch, peak.peak_id, "Toluene")
        self.service.complete_review(batch)
        self.assertEqual(batch.review_status, ReviewStatus.REVIEWED)
        result = self.service.save_batch(batch)
        self.assertTrue(result.saved)
        self.assertEqual(batch.review_status, ReviewStatus.SAVED)

    def test_user_exclusion_is_explicit(self) -> None:
        batch = review_batch()
        peak = batch.samples[0].peaks[0]
        self.service.set_material_mapping(batch, peak.peak_id, "Toluene")
        self.service.set_peak_included(batch, peak.peak_id, False)
        self.assertFalse(peak.include_for_excel)
        self.assertEqual(peak.exclude_reason, ExcludeReason.USER_EXCLUDED)

    def test_analysis_manual_mapping_stays_not_supported(self) -> None:
        batch = review_batch()
        batch.analysis_type = "B.C"
        peak = batch.samples[0].peaks[0]

        self.service.set_material_mapping(batch, peak.peak_id, "Methanol")

        self.assertEqual(peak.material_standard, "Methanol")
        self.assertFalse(peak.include_for_excel)
        self.assertEqual(
            peak.exclude_reason,
            ExcludeReason.MATERIAL_NOT_SUPPORTED_FOR_ANALYSIS,
        )
        with self.assertRaisesRegex(ValidationError, "현재 분석종류"):
            self.service.set_peak_included(batch, peak.peak_id, True)

    def test_csv_preserves_raw_and_standard_names(self) -> None:
        batch = review_batch()
        peak = batch.samples[0].peaks[0]
        self.service.set_material_mapping(batch, peak.peak_id, "Toluene")
        output = Path(self.temp.name) / "review.csv"
        self.service.export_csv(batch, output)
        text = output.read_text(encoding="utf-8-sig")
        self.assertIn("material_raw", text)
        self.assertIn("material_standard", text)
        self.assertIn("새물질,Toluene", text)

    def test_confirmed_replacement_reuses_existing_batch_without_duplicate(self) -> None:
        original = review_batch()
        original.samples[0].peaks[0].material_standard = "Toluene"
        original.samples[0].peaks[0].exclude_reason = None
        original.samples[0].peaks[0].include_for_excel = True
        original.review_status = ReviewStatus.REVIEWED
        saved = self.service.save_batch(original)

        replacement = review_batch()
        replacement.batch_code = "IPA-REPLACED"
        replacement.analysis_type = "IPA"
        replacement.samples[0].peaks[0].material_raw = "IPA"
        replacement.samples[0].peaks[0].material_standard = "Isopropyl alcohol"
        replacement.samples[0].peaks[0].exclude_reason = None
        replacement.samples[0].peaks[0].include_for_excel = True
        replacement.review_status = ReviewStatus.REVIEWED
        replacement.replacement_for_batch_id = saved.batch_id

        result = self.service.save_batch(replacement)

        self.assertEqual(result.batch_id, saved.batch_id)
        self.assertEqual(replacement.batch_id, saved.batch_id)
        self.assertIsNone(replacement.replacement_for_batch_id)
        loaded = self.database.get_batch_detail(saved.batch_id)
        self.assertEqual(loaded.analysis_type, "IPA")
        self.assertEqual(loaded.batch_code, "IPA-REPLACED")
        self.assertEqual(len(self.service.list_saved_batches()), 1)

    def test_duplicate_without_replacement_confirmation_remains_blocked(self) -> None:
        original = review_batch()
        original.samples[0].peaks[0].material_standard = "Toluene"
        original.samples[0].peaks[0].exclude_reason = None
        original.review_status = ReviewStatus.REVIEWED
        self.service.save_batch(original)
        duplicate = review_batch()
        duplicate.batch_code = "DUPLICATE"
        duplicate.review_status = ReviewStatus.REVIEWED

        with self.assertRaisesRegex(ValidationError, "동일 PDF"):
            self.service.save_batch(duplicate)
