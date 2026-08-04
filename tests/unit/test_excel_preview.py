from datetime import datetime, timezone
from decimal import Decimal
from pathlib import Path
import unittest
from uuid import uuid4

from honyu_app.application.preview_excel_export import PreviewExcelExportService
from honyu_app.domain.enums import (
    ConcentrationLevel,
    ExcelPreviewStatus,
    ExcludeReason,
    ReviewStatus,
    SampleType,
    StdMethod,
)
from honyu_app.domain.models import (
    AnalysisBatch,
    Peak,
    PeakCorrection,
    Sample,
    SourceFile,
)
from honyu_app.infrastructure.excel.workbook_inspector import XlsxTemplateInspector
from honyu_app.services.excel_template_service import (
    ExcelTemplateSnapshot,
    TemplateCell,
)


SHEETS = ("검량선", "area", "최종결과", "회수율", "STD제조")


class FakeDatabase:
    def __init__(self) -> None:
        self.corrections: dict[object, list[PeakCorrection]] = {}

    def list_peak_corrections(self, peak_id):
        return self.corrections.get(peak_id, [])

    def get_batch_detail(self, batch_id):
        raise AssertionError("preview_batch should be used in unit tests")


class FakeTemplateService:
    def __init__(self, snapshot: ExcelTemplateSnapshot) -> None:
        self.snapshot = snapshot

    def inspect(self, path: Path) -> ExcelTemplateSnapshot:
        return self.snapshot


def snapshot(*cells: TemplateCell) -> ExcelTemplateSnapshot:
    return ExcelTemplateSnapshot(
        Path("template.xlsx"),
        SHEETS,
        {(cell.sheet, cell.address): cell for cell in cells},
    )


def peak(number: int, area: int, material: str = "n-hexane") -> Peak:
    return Peak(
        peak_no=number,
        retention_time=Decimal(f"{number}.100"),
        area_raw=area,
        material_raw=material,
        material_standard=material,
        source_page=1,
    )


def batch(samples: list[Sample]) -> AnalysisBatch:
    return AnalysisBatch(
        batch_code="BATCH-PREVIEW",
        source_file=SourceFile("sample.pdf", Path("sample.pdf"), "a" * 64, 100, 1),
        analysis_type="혼유",
        analysis_no_start=1,
        analysis_no_end=10,
        parser_name="test",
        parser_version="1",
        parser_layout_id="layout",
        extracted_at=datetime.now(timezone.utc),
        samples=samples,
        review_status=ReviewStatus.SAVED,
    )


class ExcelPreviewServiceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.database = FakeDatabase()

    def service(self, template: ExcelTemplateSnapshot | None = None):
        return PreviewExcelExportService(
            self.database, FakeTemplateService(template or snapshot())
        )

    def test_std_method_a_and_b_use_the_same_excel_std5_row(self) -> None:
        std5 = Sample(5, "STD5", "STD5", SampleType.STD, replicate_no=5, peaks=[peak(1, 500)])
        std6 = Sample(6, "STD6", "STD6", SampleType.STD, replicate_no=6, peaks=[peak(1, 600)])
        source = batch([std5, std6])

        method_a = self.service().preview_batch(source, Path("template.xlsx"), StdMethod.A)
        method_b = self.service().preview_batch(source, Path("template.xlsx"), StdMethod.B)

        mapped_a = [row for row in method_a.rows if row.status is ExcelPreviewStatus.MAPPED]
        mapped_b = [row for row in method_b.rows if row.status is ExcelPreviewStatus.MAPPED]
        self.assertEqual([(row.sample_name, row.target_cell) for row in mapped_a], [("STD5", "F19")])
        self.assertEqual([(row.sample_name, row.target_cell) for row in mapped_b], [("STD6", "F19")])

    def test_dibk_uses_corrected_area_top_two_and_keeps_overflow_excluded(self) -> None:
        values = [peak(1, 100, "DIBK"), peak(2, 300, "DIBK"), peak(3, 200, "DIBK")]
        correction = PeakCorrection(
            uuid4(), values[0].peak_id, 100, 400, "재적분", datetime.now(timezone.utc), "PC", 1
        )
        self.database.corrections[values[0].peak_id] = [correction]
        sample = Sample(1, "STD1", "STD1", SampleType.STD, replicate_no=1, peaks=values)

        result = self.service().preview_batch(batch([sample]), Path("template.xlsx"), "A")
        mapped = [row for row in result.rows if row.status is ExcelPreviewStatus.MAPPED]
        excluded = [row for row in result.rows if row.status is ExcelPreviewStatus.EXCLUDED]

        self.assertEqual([(row.peak_no, row.applied_area, row.target_cell) for row in mapped], [(1, 400, "Z15"), (2, 300, "AA15")])
        self.assertEqual(excluded[0].peak_no, 3)
        self.assertEqual(excluded[0].exclude_reason, ExcludeReason.DIBK_AREA_NOT_TOP2.value)
        self.assertTrue(result.can_generate)

    def test_recovery_level_and_replicate_map_to_confirmed_cell(self) -> None:
        recovery = Sample(
            1, "저2", "저2", SampleType.RECOVERY,
            concentration_level=ConcentrationLevel.LOW,
            replicate_no=2,
            peaks=[peak(1, 123, "acetone")],
        )
        result = self.service().preview_batch(batch([recovery]), Path("template.xlsx"), "A")
        self.assertEqual(result.rows[0].target_sheet, "회수율")
        self.assertEqual(result.rows[0].target_cell, "C38")

    def test_worker_requires_exactly_one_analysis_number_suffix_match(self) -> None:
        worker = Sample(
            1, "245", "245", SampleType.NUMERIC,
            worker_match_key="245", peaks=[peak(1, 123)],
        )
        unique = snapshot(TemplateCell("area", "A37", True, "261-245", "string"))
        duplicate = snapshot(
            TemplateCell("area", "A37", True, "261-245", "string"),
            TemplateCell("area", "A38", True, "262-245", "string"),
        )

        okay = self.service(unique).preview_batch(batch([worker]), Path("template.xlsx"), "A")
        failed = self.service(duplicate).preview_batch(batch([worker]), Path("template.xlsx"), "A")

        self.assertEqual(okay.rows[0].target_cell, "F37")
        self.assertTrue(okay.can_generate)
        self.assertFalse(failed.can_generate)
        self.assertEqual(failed.issues[0].code, "WORKER_ROW_NOT_UNIQUE")

    def test_formula_target_is_a_blocking_error(self) -> None:
        template = snapshot(
            TemplateCell("area", "F15", True, 0, "formula", "SUM(A1:A2)", 1)
        )
        std = Sample(1, "STD1", "STD1", SampleType.STD, replicate_no=1, peaks=[peak(1, 100)])
        result = self.service(template).preview_batch(batch([std]), Path("template.xlsx"), "A")
        self.assertFalse(result.can_generate)
        self.assertEqual(result.rows[0].status, ExcelPreviewStatus.ERROR)
        self.assertEqual(result.issues[0].code, "TARGET_IS_FORMULA")

    def test_duplicate_single_material_peaks_are_a_target_collision(self) -> None:
        std = Sample(
            1, "STD1", "STD1", SampleType.STD, replicate_no=1,
            peaks=[peak(1, 100), peak(2, 200)],
        )
        result = self.service().preview_batch(batch([std]), Path("template.xlsx"), "A")
        self.assertFalse(result.can_generate)
        self.assertTrue(all(row.status is ExcelPreviewStatus.ERROR for row in result.rows))
        self.assertEqual(result.issues[0].code, "TARGET_COLLISION")

    def test_dibk_equal_area_uses_peak_number_as_tie_breaker(self) -> None:
        values = [peak(3, 100, "DIBK"), peak(1, 100, "DIBK"), peak(2, 100, "DIBK")]
        std = Sample(1, "STD1", "STD1", SampleType.STD, replicate_no=1, peaks=values)
        result = self.service().preview_batch(batch([std]), Path("template.xlsx"), "A")
        mapped = [row for row in result.rows if row.status is ExcelPreviewStatus.MAPPED]
        self.assertEqual([(row.peak_no, row.target_cell) for row in mapped], [(1, "Z15"), (2, "AA15")])


class XlsxTemplateInspectorTests(unittest.TestCase):
    def test_sample_template_exposes_confirmed_input_and_formula_cells(self) -> None:
        template = Path(__file__).parents[3] / "TEST" / "(혼유) 틀.xlsx"
        inspected = XlsxTemplateInspector().inspect(template)
        self.assertEqual(inspected.sheet_names, SHEETS)
        self.assertFalse(inspected.cell("area", "F15").has_formula)
        self.assertEqual(inspected.cell("area", "F15").value_type, "blank")
        self.assertTrue(inspected.cell("area", "R15").has_formula)
        self.assertIn("Z15+AA15", inspected.cell("area", "R15").formula)


if __name__ == "__main__":
    unittest.main()
