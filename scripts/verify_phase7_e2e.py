from __future__ import annotations

from pathlib import Path
import shutil
from tempfile import TemporaryDirectory

from honyu_app.application.create_excel_export import CreateExcelExportService
from honyu_app.application.preview_excel_export import PreviewExcelExportService
from honyu_app.application.review_extraction import ReviewExtractionService
from honyu_app.domain.enums import ExcelPreviewStatus, StdMethod
from honyu_app.domain.queries import BatchSearchQuery
from honyu_app.infrastructure.database.mock_database_service import MockDatabaseService
from honyu_app.infrastructure.excel.excel_recalculator import ExcelComRecalculator
from honyu_app.infrastructure.excel.workbook_inspector import XlsxTemplateInspector
from honyu_app.infrastructure.excel.workbook_validator import XlsxWorkbookValidator
from honyu_app.infrastructure.excel.xml_cell_writer import XlsxXmlCellWriter
from honyu_app.infrastructure.pdf.labsolutions_parser import LabSolutionsParser


def main() -> int:
    project_parent = Path(__file__).resolve().parents[2]
    pdf_path = project_parent / "TEST" / "혼유 39-73 병합완료.pdf"
    template_path = project_parent / "TEST" / "(혼유) 틀.xlsx"
    with TemporaryDirectory(prefix="honyu_phase7_e2e_") as temp:
        temp_path = Path(temp)
        database = MockDatabaseService(temp_path / "workflow.db")
        batch = LabSolutionsParser().parse(
            pdf_path,
            analysis_type="혼유",
            analysis_no_start=39,
            analysis_no_end=73,
        )
        review = ReviewExtractionService(database)
        review.complete_review(batch)
        saved = review.save_batch(batch)

        inspector = XlsxTemplateInspector()
        preview_service = PreviewExcelExportService(database, inspector)
        preview = preview_service.preview(saved.batch_id, template_path, StdMethod.A)
        if not preview.can_generate or preview.mapped_count != 196:
            raise RuntimeError(
                f"미리보기 불일치: can_generate={preview.can_generate}, "
                f"mapped={preview.mapped_count}, errors={preview.error_count}"
            )
        output_path = temp_path / "phase7_e2e_result.xlsx"
        try:
            result = CreateExcelExportService(
                database,
                preview_service,
                XlsxXmlCellWriter(),
                XlsxWorkbookValidator(),
                ExcelComRecalculator(timeout_seconds=180),
            ).create(
                saved.batch_id,
                template_path,
                output_path,
                StdMethod.A,
                "E2E-TEST",
            )
        except Exception:
            diagnostics = Path(__file__).resolve().parents[1] / "work" / "e2e_diagnostics"
            diagnostics.mkdir(parents=True, exist_ok=True)
            for failed in temp_path.glob("*.xlsx"):
                shutil.copy2(failed, diagnostics / failed.name)
            raise

        first = next(
            row for row in preview.rows if row.status is ExcelPreviewStatus.MAPPED
        )
        exported = inspector.inspect(output_path)
        actual = exported.cell(first.target_sheet, first.target_cell).value
        summaries = database.search_batches(BatchSearchQuery())
        if actual != first.applied_area:
            raise RuntimeError(
                f"대표 입력값 불일치: {first.target_sheet}!{first.target_cell} "
                f"expected={first.applied_area}, actual={actual}"
            )
        if summaries[0].export_count != 1:
            raise RuntimeError(f"출력 이력 불일치: {summaries[0].export_count}")
        print(
            "E2E_OK "
            f"samples={len(batch.samples)} raw_peaks={sum(len(s.peaks) for s in batch.samples)} "
            f"mapped={result.mapped_cell_count} recalculated={result.recalculated} "
            f"validated={result.validation_passed} output_bytes={output_path.stat().st_size} "
            f"representative={first.target_sheet}!{first.target_cell}:{actual} "
            f"export_jobs={summaries[0].export_count}"
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
