from __future__ import annotations

from datetime import datetime
from pathlib import Path
import shutil
from uuid import UUID, uuid4

from honyu_app.application.preview_excel_export import PreviewExcelExportService
from honyu_app.config.paths import excel_work_dir, failed_exports_dir
from honyu_app.domain.commands import SaveExportJobCommand
from honyu_app.domain.enums import ExcelPreviewStatus, StdMethod
from honyu_app.domain.errors import (
    ExcelExportError,
    ExcelRecalculationError,
    WorkbookStructureError,
)
from honyu_app.domain.models import ExcelCellWrite, ExcelExportResult
from honyu_app.services.database_service import DatabaseService
from honyu_app.services.excel_export_services import (
    ExcelCellWriter,
    ExcelRecalculator,
    WorkbookValidator,
)


class CreateExcelExportService:
    def __init__(
        self,
        database: DatabaseService,
        preview_service: PreviewExcelExportService,
        writer: ExcelCellWriter,
        validator: WorkbookValidator,
        recalculator: ExcelRecalculator,
        *,
        recalculate_with_excel: bool = False,
    ) -> None:
        self._database = database
        self._preview_service = preview_service
        self._writer = writer
        self._validator = validator
        self._recalculator = recalculator
        self._recalculate_with_excel = recalculate_with_excel

    def create(
        self,
        batch_id: UUID,
        template_path: Path,
        output_path: Path,
        std_method: StdMethod | str,
        device_id: str,
    ) -> ExcelExportResult:
        template = Path(template_path).resolve()
        output = Path(output_path).resolve()
        self._validate_paths(template, output)

        preview = self._preview_service.preview(batch_id, template, std_method)
        if not preview.can_generate:
            details = "; ".join(issue.message for issue in preview.issues[:5])
            raise ExcelExportError(f"미리보기 오류가 있어 Excel을 생성할 수 없습니다: {details}")
        writes = [
            ExcelCellWrite(row.target_sheet, row.target_cell, row.applied_area)
            for row in preview.rows
            if row.status is ExcelPreviewStatus.MAPPED
            and row.target_sheet is not None
            and row.target_cell is not None
        ]
        if not writes:
            raise ExcelExportError("Excel에 입력할 Peak가 없습니다.")
        if len({(item.sheet, item.address.upper()) for item in writes}) != len(writes):
            raise ExcelExportError("같은 Excel 셀에 둘 이상의 Peak가 배정되었습니다.")

        work_directory = excel_work_dir()
        work_directory.mkdir(parents=True, exist_ok=True)
        partial = work_directory / f"honyu_export_{uuid4().hex}.xlsx"
        self._writer.write_copy(template, partial, writes)
        recalculated = False
        try:
            before = self._validator.validate(
                template, partial, writes, after_excel_recalculation=False
            )
            if not before.valid:
                failed = self._preserve_failure(partial, output, "검증실패")
                raise WorkbookStructureError(
                    "Excel 사전 구조 검증에 실패했습니다. "
                    f"점검용 파일: {failed}\n" + "\n".join(before.errors[:10])
                )

            if self._recalculate_with_excel:
                try:
                    self._recalculator.recalculate(partial)
                    recalculated = True
                except ExcelRecalculationError as exc:
                    if exc.code != "OFFICE_NOT_ACTIVATED":
                        failed = self._preserve_failure(partial, output, "재계산실패")
                        raise ExcelExportError(
                            f"Excel 전체 재계산에 실패했습니다. 점검용 파일: {failed}\n{exc}"
                        ) from exc
                except Exception as exc:
                    failed = self._preserve_failure(partial, output, "재계산실패")
                    raise ExcelExportError(
                        f"Excel 전체 재계산에 실패했습니다. 점검용 파일: {failed}\n{exc}"
                    ) from exc

            after = self._validator.validate(
                template,
                partial,
                writes,
                after_excel_recalculation=recalculated,
            )
            if not after.valid:
                failed = self._preserve_failure(partial, output, "검증실패")
                raise WorkbookStructureError(
                    "Excel 재계산 후 구조 검증에 실패했습니다. "
                    f"점검용 파일: {failed}\n" + "\n".join(after.errors[:10])
                )
            try:
                shutil.move(str(partial), str(output))
            except Exception:
                output.unlink(missing_ok=True)
                raise
            final = self._validator.validate(
                template,
                output,
                writes,
                after_excel_recalculation=recalculated,
            )
            if not final.valid:
                failed = self._preserve_failure(output, output, "최종검증실패")
                raise WorkbookStructureError(
                    "최종 Excel 입력값 검증에 실패했습니다. SUCCESS로 처리하지 않았습니다. "
                    f"점검용 파일: {failed}\n" + "\n".join(final.errors[:10])
                )
        except Exception:
            if partial.exists():
                partial.unlink()
            raise

        method = StdMethod(std_method).value
        job = self._database.save_export_job(
            SaveExportJobCommand(
                batch_id=batch_id,
                template_path=str(template),
                output_path=str(output),
                std_method=method,
                device_id=device_id.strip() or "UNKNOWN",
            )
        )
        return ExcelExportResult(
            output, len(writes), True, recalculated, job.export_job_id
        )

    @staticmethod
    def _validate_paths(template: Path, output: Path) -> None:
        if not template.is_file():
            raise ExcelExportError(f"원본 Excel 파일이 없습니다: {template}")
        if template.suffix.lower() != ".xlsx" or output.suffix.lower() != ".xlsx":
            raise ExcelExportError("원본과 결과 파일은 .xlsx 형식이어야 합니다.")
        if template == output:
            raise ExcelExportError("원본과 결과 파일 경로는 달라야 합니다.")
        if not output.parent.is_dir():
            raise ExcelExportError(f"결과 저장 폴더가 없습니다: {output.parent}")
        if output.exists():
            raise ExcelExportError(f"같은 이름의 결과 파일이 이미 있습니다: {output}")

    @staticmethod
    def _preserve_failure(partial: Path, output: Path, reason: str) -> Path:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        directory = failed_exports_dir()
        directory.mkdir(parents=True, exist_ok=True)
        failed = directory / f"{output.stem}_{reason}_{stamp}.xlsx"
        shutil.move(str(partial), str(failed))
        return failed
