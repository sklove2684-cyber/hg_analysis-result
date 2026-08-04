from __future__ import annotations

from datetime import datetime, timezone
from decimal import Decimal
from pathlib import Path
from contextlib import contextmanager
import sqlite3
from collections.abc import Iterator
from uuid import UUID, uuid4

from honyu_app.domain.commands import (
    AddPeakCorrectionCommand,
    SaveAnalysisBatchCommand,
    SaveExportJobCommand,
)
from honyu_app.domain.enums import ConcentrationLevel, ExcludeReason, ReviewStatus, SampleType
from honyu_app.domain.errors import (
    DatabaseUnavailableError,
    DuplicateSourceFileError,
    RecordNotFoundError,
    RevisionConflictError,
)
from honyu_app.domain.models import AnalysisBatch, Peak, PeakCorrection, Sample, SourceFile
from honyu_app.domain.queries import BatchSearchQuery
from honyu_app.domain.results import (
    BatchSummary,
    ConnectionStatus,
    DuplicateCheckResult,
    ExportJobResult,
    PeakCorrectionResult,
    SaveAnalysisBatchResult,
)


def _now() -> datetime:
    return datetime.now(timezone.utc)


def _optional_datetime(value: str | None) -> datetime | None:
    return datetime.fromisoformat(value) if value else None


def _optional_enum(enum_type, value):
    return enum_type(value) if value is not None else None


class MockDatabaseService:
    """SQLite implementation of the production DatabaseService contract."""

    def __init__(self, database_file: Path) -> None:
        self._database_file = Path(database_file)
        self._database_file.parent.mkdir(parents=True, exist_ok=True)
        self._initialize_schema()

    @property
    def database_file(self) -> Path:
        return self._database_file

    @contextmanager
    def _connect(self) -> Iterator[sqlite3.Connection]:
        connection = sqlite3.connect(self._database_file, timeout=10)
        try:
            connection.row_factory = sqlite3.Row
            connection.execute("PRAGMA foreign_keys = ON")
            connection.execute("PRAGMA busy_timeout = 10000")
            connection.execute("PRAGMA journal_mode = WAL")
            with connection:
                yield connection
        finally:
            connection.close()

    def _initialize_schema(self) -> None:
        schema = Path(__file__).with_name("mock_schema.sql").read_text(encoding="utf-8")
        try:
            with self._connect() as connection:
                connection.executescript(schema)
        except (OSError, sqlite3.Error) as exc:
            raise DatabaseUnavailableError(
                f"Mock DB를 초기화할 수 없습니다: {self._database_file}"
            ) from exc

    def check_connection(self) -> ConnectionStatus:
        try:
            with self._connect() as connection:
                connection.execute("SELECT 1").fetchone()
            return ConnectionStatus(True, "mock", f"Mock DB 연결됨: {self._database_file}")
        except sqlite3.Error as exc:
            return ConnectionStatus(False, "mock", f"Mock DB 연결 실패: {exc}")

    def check_duplicate(self, file_hash: str) -> DuplicateCheckResult:
        with self._connect() as connection:
            row = connection.execute(
                """
                SELECT b.batch_id, b.batch_code
                FROM source_files f
                JOIN analysis_batches b ON b.batch_id = f.batch_id
                WHERE f.file_hash = ?
                """,
                (file_hash,),
            ).fetchone()
        if row is None:
            return DuplicateCheckResult(False)
        return DuplicateCheckResult(True, UUID(row["batch_id"]), row["batch_code"])

    def save_analysis_batch(
        self, command: SaveAnalysisBatchCommand
    ) -> SaveAnalysisBatchResult:
        batch = command.batch
        duplicate = self.check_duplicate(batch.source_file.file_hash)
        if duplicate.is_duplicate:
            raise DuplicateSourceFileError(
                f"동일한 PDF가 이미 저장되어 있습니다: {duplicate.existing_batch_code}"
            )
        now = _now().isoformat()
        try:
            with self._connect() as connection:
                connection.execute(
                    """
                    INSERT INTO analysis_batches (
                        batch_id, batch_code, analysis_type, analysis_no_start,
                        analysis_no_end, parser_name, parser_version, parser_layout_id,
                        extracted_at, warning_count, review_status, workplace, year,
                        period, device_id, analyst, created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        str(batch.batch_id), batch.batch_code, batch.analysis_type,
                        batch.analysis_no_start, batch.analysis_no_end, batch.parser_name,
                        batch.parser_version, batch.parser_layout_id,
                        batch.extracted_at.isoformat(), batch.warning_count,
                        batch.review_status.value, batch.workplace, batch.year, batch.period,
                        batch.device_id, batch.analyst, now, now,
                    ),
                )
                connection.execute(
                    """
                    INSERT INTO source_files (
                        source_file_id, batch_id, original_name, full_path, file_hash,
                        file_size, page_count, created_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        str(uuid4()), str(batch.batch_id), batch.source_file.original_name,
                        str(batch.source_file.full_path), batch.source_file.file_hash,
                        batch.source_file.file_size, batch.source_file.page_count, now,
                    ),
                )
                for sample in batch.samples:
                    self._insert_sample(connection, batch.batch_id, sample, now)
        except sqlite3.IntegrityError as exc:
            if "file_hash" in str(exc) or "source_files.file_hash" in str(exc):
                raise DuplicateSourceFileError("동일한 PDF 해시가 이미 저장되어 있습니다.") from exc
            raise
        return SaveAnalysisBatchResult(batch.batch_id, batch.batch_code, True)

    @staticmethod
    def _insert_sample(
        connection: sqlite3.Connection, batch_id: UUID, sample: Sample, now: str
    ) -> None:
        connection.execute(
            """
            INSERT INTO samples (
                sample_id, batch_id, page_no, sample_name_raw, sample_name_normalized,
                data_filename, method_filename, batch_filename, acquired_at, sample_type,
                concentration_level, replicate_no, worker_match_key, is_blank, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                str(sample.sample_id), str(batch_id), sample.page_no, sample.sample_name_raw,
                sample.sample_name_normalized, sample.data_filename, sample.method_filename,
                sample.batch_filename,
                sample.acquired_at.isoformat() if sample.acquired_at else None,
                sample.sample_type.value,
                sample.concentration_level.value if sample.concentration_level else None,
                sample.replicate_no, sample.worker_match_key, int(sample.is_blank), now,
            ),
        )
        for peak in sample.peaks:
            connection.execute(
                """
                INSERT INTO peaks (
                    peak_id, sample_id, peak_no, retention_time, area_raw, height,
                    material_raw, material_standard, peak_group_no, include_for_excel,
                    exclude_reason, source_page, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    str(peak.peak_id), str(sample.sample_id), peak.peak_no,
                    str(peak.retention_time), peak.area_raw, peak.height, peak.material_raw,
                    peak.material_standard, peak.peak_group_no,
                    int(peak.include_for_excel),
                    peak.exclude_reason.value if peak.exclude_reason else None,
                    peak.source_page, now,
                ),
            )

    def search_batches(self, query: BatchSearchQuery) -> list[BatchSummary]:
        conditions: list[str] = []
        values: list[object] = []
        field_filters = (
            ("b.workplace", query.workplace),
            ("b.year", query.year),
            ("b.period", query.period),
            ("b.analysis_type", query.analysis_type),
            ("b.analyst", query.analyst),
        )
        for field, value in field_filters:
            if value is not None:
                conditions.append(f"{field} = ?")
                values.append(value)
        if query.analysis_no_start is not None:
            conditions.append("b.analysis_no_end >= ?")
            values.append(query.analysis_no_start)
        if query.analysis_no_end is not None:
            conditions.append("b.analysis_no_start <= ?")
            values.append(query.analysis_no_end)
        if query.pdf_filename:
            conditions.append("f.original_name LIKE ?")
            values.append(f"%{query.pdf_filename}%")
        if query.sample_name:
            conditions.append(
                "EXISTS (SELECT 1 FROM samples s WHERE s.batch_id = b.batch_id "
                "AND s.sample_name_normalized LIKE ?)"
            )
            values.append(f"%{query.sample_name}%")
        if query.material_name:
            conditions.append(
                "EXISTS (SELECT 1 FROM samples s JOIN peaks p ON p.sample_id = s.sample_id "
                "WHERE s.batch_id = b.batch_id AND p.material_standard LIKE ?)"
            )
            values.append(f"%{query.material_name}%")
        where = f"WHERE {' AND '.join(conditions)}" if conditions else ""
        sql = f"""
            SELECT b.*, f.original_name,
                   (SELECT COUNT(*) FROM peak_corrections c
                    JOIN peaks p ON p.peak_id = c.peak_id
                    JOIN samples s ON s.sample_id = p.sample_id
                    WHERE s.batch_id = b.batch_id) AS correction_count,
                   (SELECT COUNT(*) FROM export_jobs e
                    WHERE e.batch_id = b.batch_id) AS export_count
            FROM analysis_batches b
            JOIN source_files f ON f.batch_id = b.batch_id
            {where}
            ORDER BY b.created_at DESC
        """
        with self._connect() as connection:
            rows = connection.execute(sql, values).fetchall()
        return [
            BatchSummary(
                batch_id=UUID(row["batch_id"]), batch_code=row["batch_code"],
                pdf_filename=row["original_name"], analysis_type=row["analysis_type"],
                review_status=row["review_status"], workplace=row["workplace"],
                year=row["year"], period=row["period"],
                analysis_no_start=row["analysis_no_start"],
                analysis_no_end=row["analysis_no_end"],
                parser_version=row["parser_version"],
                correction_count=row["correction_count"], export_count=row["export_count"],
            )
            for row in rows
        ]

    def get_batch_detail(self, batch_id: UUID) -> AnalysisBatch:
        with self._connect() as connection:
            batch_row = connection.execute(
                """
                SELECT b.*, f.original_name, f.full_path, f.file_hash, f.file_size,
                       f.page_count
                FROM analysis_batches b
                JOIN source_files f ON f.batch_id = b.batch_id
                WHERE b.batch_id = ?
                """,
                (str(batch_id),),
            ).fetchone()
            if batch_row is None:
                raise RecordNotFoundError(f"분석 배치를 찾을 수 없습니다: {batch_id}")
            sample_rows = connection.execute(
                "SELECT * FROM samples WHERE batch_id = ? ORDER BY page_no",
                (str(batch_id),),
            ).fetchall()
            samples: list[Sample] = []
            for row in sample_rows:
                peak_rows = connection.execute(
                    "SELECT * FROM peaks WHERE sample_id = ? ORDER BY peak_no",
                    (row["sample_id"],),
                ).fetchall()
                peaks = [self._peak_from_row(value) for value in peak_rows]
                samples.append(self._sample_from_row(row, peaks))
        source = SourceFile(
            original_name=batch_row["original_name"],
            full_path=Path(batch_row["full_path"]), file_hash=batch_row["file_hash"],
            file_size=batch_row["file_size"], page_count=batch_row["page_count"],
        )
        return AnalysisBatch(
            batch_code=batch_row["batch_code"], source_file=source,
            analysis_type=batch_row["analysis_type"],
            analysis_no_start=batch_row["analysis_no_start"],
            analysis_no_end=batch_row["analysis_no_end"],
            parser_name=batch_row["parser_name"], parser_version=batch_row["parser_version"],
            parser_layout_id=batch_row["parser_layout_id"],
            extracted_at=datetime.fromisoformat(batch_row["extracted_at"]), samples=samples,
            warning_count=batch_row["warning_count"],
            review_status=ReviewStatus(batch_row["review_status"]),
            workplace=batch_row["workplace"], year=batch_row["year"],
            period=batch_row["period"], device_id=batch_row["device_id"],
            analyst=batch_row["analyst"], batch_id=UUID(batch_row["batch_id"]),
        )

    @staticmethod
    def _peak_from_row(row: sqlite3.Row) -> Peak:
        return Peak(
            peak_no=row["peak_no"], retention_time=Decimal(row["retention_time"]),
            area_raw=row["area_raw"], height=row["height"],
            material_raw=row["material_raw"], material_standard=row["material_standard"],
            peak_group_no=row["peak_group_no"],
            include_for_excel=bool(row["include_for_excel"]),
            exclude_reason=_optional_enum(ExcludeReason, row["exclude_reason"]),
            source_page=row["source_page"], peak_id=UUID(row["peak_id"]),
        )

    @staticmethod
    def _sample_from_row(row: sqlite3.Row, peaks: list[Peak]) -> Sample:
        return Sample(
            page_no=row["page_no"], sample_name_raw=row["sample_name_raw"],
            sample_name_normalized=row["sample_name_normalized"],
            sample_type=SampleType(row["sample_type"]), data_filename=row["data_filename"],
            method_filename=row["method_filename"], batch_filename=row["batch_filename"],
            acquired_at=_optional_datetime(row["acquired_at"]),
            concentration_level=_optional_enum(
                ConcentrationLevel, row["concentration_level"]
            ),
            replicate_no=row["replicate_no"], worker_match_key=row["worker_match_key"],
            is_blank=bool(row["is_blank"]), peaks=peaks,
            sample_id=UUID(row["sample_id"]),
        )

    def add_peak_correction(
        self, command: AddPeakCorrectionCommand
    ) -> PeakCorrectionResult:
        if command.area_after < 0:
            raise ValueError("area_after cannot be negative")
        if not command.reason.strip():
            raise ValueError("수정 사유를 입력해야 합니다.")
        corrected_at = _now()
        correction_id = uuid4()
        with self._connect() as connection:
            peak = connection.execute(
                "SELECT area_raw FROM peaks WHERE peak_id = ?", (str(command.peak_id),)
            ).fetchone()
            if peak is None:
                raise RecordNotFoundError(f"Peak를 찾을 수 없습니다: {command.peak_id}")
            latest = connection.execute(
                """
                SELECT area_after, revision_no FROM peak_corrections
                WHERE peak_id = ? ORDER BY revision_no DESC LIMIT 1
                """,
                (str(command.peak_id),),
            ).fetchone()
            current_revision = latest["revision_no"] if latest else 0
            if command.expected_revision_no != current_revision:
                raise RevisionConflictError(
                    f"수정 충돌: 예상 revision {command.expected_revision_no}, "
                    f"현재 revision {current_revision}"
                )
            area_before = latest["area_after"] if latest else peak["area_raw"]
            revision_no = current_revision + 1
            connection.execute(
                """
                INSERT INTO peak_corrections (
                    correction_id, peak_id, area_before, area_after, reason,
                    corrected_at, device_id, revision_no
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    str(correction_id), str(command.peak_id), area_before,
                    command.area_after, command.reason.strip(), corrected_at.isoformat(),
                    command.device_id, revision_no,
                ),
            )
        correction = PeakCorrection(
            correction_id, command.peak_id, area_before, command.area_after,
            command.reason.strip(), corrected_at, command.device_id, revision_no,
        )
        return PeakCorrectionResult(correction)

    def list_peak_corrections(self, peak_id: UUID) -> list[PeakCorrection]:
        with self._connect() as connection:
            rows = connection.execute(
                """
                SELECT * FROM peak_corrections WHERE peak_id = ? ORDER BY revision_no
                """,
                (str(peak_id),),
            ).fetchall()
        return [
            PeakCorrection(
                correction_id=UUID(row["correction_id"]), peak_id=UUID(row["peak_id"]),
                area_before=row["area_before"], area_after=row["area_after"],
                reason=row["reason"], corrected_at=datetime.fromisoformat(row["corrected_at"]),
                device_id=row["device_id"], revision_no=row["revision_no"],
            )
            for row in rows
        ]

    def save_export_job(self, command: SaveExportJobCommand) -> ExportJobResult:
        if command.std_method not in {"A", "B"}:
            raise ValueError("STD 방식은 A 또는 B여야 합니다.")
        export_job_id = uuid4()
        try:
            with self._connect() as connection:
                connection.execute(
                    """
                    INSERT INTO export_jobs (
                        export_job_id, batch_id, template_path, output_path, std_method,
                        device_id, status, created_at
                    ) VALUES (?, ?, ?, ?, ?, ?, 'COMPLETED', ?)
                    """,
                    (
                        str(export_job_id), str(command.batch_id), command.template_path,
                        command.output_path, command.std_method, command.device_id,
                        _now().isoformat(),
                    ),
                )
        except sqlite3.IntegrityError as exc:
            raise RecordNotFoundError(
                f"내보내기 대상 배치를 찾을 수 없습니다: {command.batch_id}"
            ) from exc
        return ExportJobResult(export_job_id, True)
