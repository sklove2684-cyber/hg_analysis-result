from datetime import datetime, timezone
from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory
from contextlib import closing
import sqlite3
import unittest
from uuid import uuid4

from honyu_app.domain.commands import (
    AddPeakCorrectionCommand,
    SaveAnalysisBatchCommand,
    SaveExportJobCommand,
)
from honyu_app.domain.enums import ConcentrationLevel, ReviewStatus, SampleType
from honyu_app.domain.errors import DuplicateSourceFileError, RevisionConflictError
from honyu_app.domain.models import AnalysisBatch, Peak, Sample, SourceFile
from honyu_app.domain.queries import BatchSearchQuery
from honyu_app.infrastructure.database.mock_database_service import MockDatabaseService


def make_batch(*, file_hash: str = "a" * 64, batch_code: str = "BATCH-001") -> AnalysisBatch:
    source = SourceFile(
        original_name="혼유 39-73 병합완료.pdf",
        full_path=Path(r"C:\samples\혼유 39-73 병합완료.pdf"),
        file_hash=file_hash,
        file_size=2_452_519,
        page_count=19,
    )
    std_peak = Peak(
        peak_no=1,
        retention_time=Decimal("2.364"),
        area_raw=24_350,
        height=8_047,
        material_raw="헥산",
        material_standard="n-hexane",
        source_page=2,
    )
    sample = Sample(
        page_no=2,
        sample_name_raw="STD1",
        sample_name_normalized="STD1",
        sample_type=SampleType.STD,
        data_filename="혼유@001.gcd",
        concentration_level=ConcentrationLevel.LOW,
        replicate_no=1,
        peaks=[std_peak],
    )
    return AnalysisBatch(
        batch_code=batch_code,
        source_file=source,
        analysis_type="혼유",
        analysis_no_start=39,
        analysis_no_end=73,
        parser_name="labsolutions",
        parser_version="1.0.0",
        parser_layout_id="layout-1",
        extracted_at=datetime(2026, 8, 3, 12, 0, tzinfo=timezone.utc),
        samples=[sample],
        review_status=ReviewStatus.REVIEWED,
        workplace="구조제안",
        year=2026,
        period="상반기",
        device_id="PC-01",
        analyst="분석자",
    )


class MockDatabaseServiceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = TemporaryDirectory()
        self.db_file = Path(self.temp.name) / "mock.db"
        self.service = MockDatabaseService(self.db_file)

    def tearDown(self) -> None:
        self.temp.cleanup()

    def test_connection_and_schema(self) -> None:
        self.assertTrue(self.service.check_connection().connected)
        with closing(sqlite3.connect(self.db_file)) as connection:
            tables = {
                row[0]
                for row in connection.execute(
                    "SELECT name FROM sqlite_master WHERE type='table'"
                )
            }
        self.assertLessEqual(
            {
                "analysis_batches",
                "source_files",
                "samples",
                "peaks",
                "peak_corrections",
                "export_jobs",
            },
            tables,
        )

    def test_save_and_reload_full_batch(self) -> None:
        batch = make_batch()
        batch.samples[0].total_area = 1_109_240
        result = self.service.save_analysis_batch(SaveAnalysisBatchCommand(batch))
        loaded = self.service.get_batch_detail(result.batch_id)
        self.assertEqual(loaded.batch_code, batch.batch_code)
        self.assertEqual(loaded.source_file.file_hash, batch.source_file.file_hash)
        self.assertEqual(loaded.analyst, "분석자")
        self.assertEqual(loaded.samples[0].sample_type, SampleType.STD)
        self.assertEqual(loaded.samples[0].peaks[0].retention_time, Decimal("2.364"))
        self.assertEqual(loaded.samples[0].peaks[0].area_raw, 24_350)
        self.assertEqual(loaded.samples[0].total_area, 1_109_240)
        self.assertEqual(loaded.review_status, ReviewStatus.SAVED)

    def test_existing_reviewed_rows_are_migrated_to_saved(self) -> None:
        batch = make_batch()
        result = self.service.save_analysis_batch(SaveAnalysisBatchCommand(batch))
        with closing(sqlite3.connect(self.db_file)) as connection:
            connection.execute(
                "UPDATE analysis_batches SET review_status = 'REVIEWED' WHERE batch_id = ?",
                (str(result.batch_id),),
            )
            connection.commit()
        reopened = MockDatabaseService(self.db_file)
        self.assertEqual(
            reopened.get_batch_detail(result.batch_id).review_status,
            ReviewStatus.SAVED,
        )

    def test_duplicate_hash_is_blocked_even_with_different_filename(self) -> None:
        first = make_batch()
        self.service.save_analysis_batch(SaveAnalysisBatchCommand(first))
        second = make_batch(batch_code="BATCH-002")
        second.source_file = SourceFile(
            "다른이름.pdf",
            Path(r"C:\samples\다른이름.pdf"),
            first.source_file.file_hash,
            100,
            19,
        )
        with self.assertRaises(DuplicateSourceFileError):
            self.service.save_analysis_batch(SaveAnalysisBatchCommand(second))
        duplicate = self.service.check_duplicate(first.source_file.file_hash)
        self.assertTrue(duplicate.is_duplicate)
        self.assertEqual(duplicate.existing_batch_code, "BATCH-001")

    def test_transaction_rolls_back_all_rows_when_peak_unique_rule_fails(self) -> None:
        batch = make_batch()
        batch.samples[0].peaks.append(
            Peak(1, Decimal("9.999"), 999, source_page=2)
        )
        with self.assertRaises(sqlite3.IntegrityError):
            self.service.save_analysis_batch(SaveAnalysisBatchCommand(batch))
        self.assertFalse(self.service.check_duplicate(batch.source_file.file_hash).is_duplicate)
        self.assertEqual(self.service.search_batches(BatchSearchQuery()), [])

    def test_correction_preserves_area_raw_and_revisions(self) -> None:
        batch = make_batch()
        self.service.save_analysis_batch(SaveAnalysisBatchCommand(batch))
        peak_id = batch.samples[0].peaks[0].peak_id
        first = self.service.add_peak_correction(
            AddPeakCorrectionCommand(peak_id, 25_000, "적분선 확인", "PC-01", 0)
        ).correction
        second = self.service.add_peak_correction(
            AddPeakCorrectionCommand(peak_id, 25_100, "재확인", "PC-02", 1)
        ).correction
        self.assertEqual(first.area_before, 24_350)
        self.assertEqual(first.revision_no, 1)
        self.assertEqual(second.area_before, 25_000)
        self.assertEqual(second.revision_no, 2)
        loaded = self.service.get_batch_detail(batch.batch_id)
        self.assertEqual(loaded.samples[0].peaks[0].area_raw, 24_350)
        self.assertEqual(len(self.service.list_peak_corrections(peak_id)), 2)

    def test_stale_revision_is_rejected(self) -> None:
        batch = make_batch()
        self.service.save_analysis_batch(SaveAnalysisBatchCommand(batch))
        peak_id = batch.samples[0].peaks[0].peak_id
        self.service.add_peak_correction(
            AddPeakCorrectionCommand(peak_id, 25_000, "수정", "PC-01", 0)
        )
        with self.assertRaises(RevisionConflictError):
            self.service.add_peak_correction(
                AddPeakCorrectionCommand(peak_id, 26_000, "오래된 화면", "PC-02", 0)
            )

    def test_database_trigger_blocks_direct_area_raw_update(self) -> None:
        batch = make_batch()
        self.service.save_analysis_batch(SaveAnalysisBatchCommand(batch))
        with closing(sqlite3.connect(self.db_file)) as connection:
            with self.assertRaisesRegex(sqlite3.IntegrityError, "immutable"):
                connection.execute(
                    "UPDATE peaks SET area_raw = 1 WHERE peak_id = ?",
                    (str(batch.samples[0].peaks[0].peak_id),),
                )

    def test_search_by_batch_sample_and_material(self) -> None:
        batch = make_batch()
        self.service.save_analysis_batch(SaveAnalysisBatchCommand(batch))
        queries = (
            BatchSearchQuery(workplace="구조제안", year=2026),
            BatchSearchQuery(pdf_filename="39-73"),
            BatchSearchQuery(sample_name="STD1"),
            BatchSearchQuery(material_name="n-hexane"),
            BatchSearchQuery(analysis_no_start=50, analysis_no_end=60),
        )
        for query in queries:
            with self.subTest(query=query):
                self.assertEqual(len(self.service.search_batches(query)), 1)
        self.assertEqual(
            self.service.search_batches(BatchSearchQuery(material_name="Toluene")), []
        )

    def test_export_job_is_recorded_in_summary(self) -> None:
        batch = make_batch()
        self.service.save_analysis_batch(SaveAnalysisBatchCommand(batch))
        result = self.service.save_export_job(
            SaveExportJobCommand(
                batch.batch_id,
                r"C:\templates\틀.xlsx",
                r"C:\results\완료.xlsx",
                "A",
                "PC-01",
            )
        )
        self.assertTrue(result.saved)
        summary = self.service.search_batches(BatchSearchQuery())[0]
        self.assertEqual(summary.export_count, 1)

    def test_data_persists_after_service_restart(self) -> None:
        batch = make_batch()
        self.service.save_analysis_batch(SaveAnalysisBatchCommand(batch))
        restarted = MockDatabaseService(self.db_file)
        loaded = restarted.get_batch_detail(batch.batch_id)
        self.assertEqual(loaded.batch_code, "BATCH-001")
        self.assertEqual(len(loaded.samples), 1)
