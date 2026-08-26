from __future__ import annotations

import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from zipfile import ZipFile

from honyu_app.application.create_excel_export import CreateExcelExportService
from honyu_app.application.preview_excel_export import PreviewExcelExportService
from honyu_app.application.review_extraction import ReviewExtractionService
from honyu_app.domain.enums import ExcelPreviewStatus
from honyu_app.infrastructure.database.mock_database_service import MockDatabaseService
from honyu_app.infrastructure.excel.workbook_inspector import XlsxTemplateInspector
from honyu_app.infrastructure.excel.workbook_validator import XlsxWorkbookValidator
from honyu_app.infrastructure.excel.xml_cell_writer import XlsxXmlCellWriter
from honyu_app.infrastructure.pdf.labsolutions_parser import LabSolutionsParser


CASES = (
    ("DMF,DMA", "(DMF) 287-296.pdf", "(DMF) 287-296.xlsx", 287, 296, 28, 33),
    ("스토다드솔벤트", "스토다드솔벤트 705-706.pdf", "스토다드솔벤트 705-706.xlsx", 705, 706, 44, 715),
    ("알콜4", "알콜(4) 422-476.pdf", "(알콜4) 422-476.xlsx", 422, 476, 63, 63),
    (
        "1,2-에폭시프로판(산화프로필렌)",
        "에폭시프로판 136,137,152,153@완료.pdf",
        "(1,2-에폭시프로판) 136-137, 152-153.xlsx",
        136,
        153,
        16,
        36,
    ),
    ("디클로로메탄(MC)", "디클로로메탄(MC) 560-561.pdf", "(디클로로메탄(MC)) 560-561.xlsx", 560, 561, 14, 30),
    ("메틸 n아밀케톤", "(메틸 n아밀케톤) 263-364.pdf", "(메틸 n아밀케톤) 263-364.xlsx", 263, 364, 14, 79),
    ("비닐아세테이트", "비닐아세테이트 656-657.pdf", "(비닐아세테이트) 656-657.xlsx", 656, 657, 14, 102),
    ("이소프로필 아세테이트", "이소프로필아세테이트 462-463.pdf", "(이소프로필 아세테이트) 462-463.xlsx", 462, 463, 16, 55),
    ("피리딘", "피리딘 448-449.pdf", "(피리딘) 448-449.xlsx", 448, 449, 14, 105),
)


def _candidate_directories() -> tuple[Path, ...]:
    configured = os.environ.get("HONYU_INTEGRATED_TEST_DIR")
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


ACTUAL_CASES = tuple(
    (analysis_type, _find_file(pdf), _find_file(xlsx), start, end, mapped, excluded)
    for analysis_type, pdf, xlsx, start, end, mapped, excluded in CASES
)
ALL_FILES_PRESENT = all(pdf.is_file() and xlsx.is_file() for _, pdf, xlsx, *_ in ACTUAL_CASES)


@unittest.skipUnless(ALL_FILES_PRESENT, "통합 프로필 실제 PDF/XLSX 9세트가 없습니다.")
class IntegratedMaterialProfileActualFileTests(unittest.TestCase):
    def test_actual_pdf_preview_and_export_for_all_profiles(self) -> None:
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            database = MockDatabaseService(root / "profiles.db")
            parser = LabSolutionsParser()
            review = ReviewExtractionService(database)
            preview_service = PreviewExcelExportService(database, XlsxTemplateInspector())
            export_service = CreateExcelExportService(
                database,
                preview_service,
                XlsxXmlCellWriter(),
                XlsxWorkbookValidator(),
                object(),
                recalculate_with_excel=False,
            )

            for analysis_type, pdf, template, start, end, expected_mapped, expected_excluded in ACTUAL_CASES:
                with self.subTest(analysis_type=analysis_type):
                    batch = parser.parse(
                        pdf,
                        analysis_type=analysis_type,
                        analysis_no_start=start,
                        analysis_no_end=end,
                    )
                    self.assertEqual(batch.warning_count, 0)
                    self.assertEqual(
                        sum(
                            peak.exclude_reason.value == "UNKNOWN_MATERIAL"
                            for sample in batch.samples
                            for peak in sample.peaks
                            if peak.exclude_reason is not None
                        ),
                        0,
                    )
                    review.complete_review(batch)
                    saved = review.save_batch(batch)
                    preview = preview_service.preview(saved.batch_id, template, "A")
                    self.assertTrue(preview.can_generate, preview.issues)
                    self.assertEqual(preview.error_count, 0)
                    self.assertEqual(preview.mapped_count, expected_mapped)
                    self.assertEqual(preview.excluded_count, expected_excluded)
                    self.assertEqual(
                        len(
                            {
                                (row.target_sheet, row.target_cell)
                                for row in preview.rows
                                if row.status is ExcelPreviewStatus.MAPPED
                            }
                        ),
                        expected_mapped,
                    )

                    output = root / f"{analysis_type.replace('/', '_')}_result.xlsx"
                    exported = export_service.create(
                        saved.batch_id, template, output, "A", "REGRESSION"
                    )
                    self.assertTrue(exported.validation_passed)
                    self.assertFalse(exported.recalculated)
                    self.assertEqual(exported.mapped_cell_count, expected_mapped)

                    before = XlsxTemplateInspector().inspect(template)
                    after = XlsxTemplateInspector().inspect(output)
                    mapped_values = {
                        (row.target_sheet, row.target_cell): row.applied_area
                        for row in preview.rows
                        if row.status is ExcelPreviewStatus.MAPPED
                    }
                    for key, expected_value in mapped_values.items():
                        self.assertEqual(after.cells[key].value, expected_value, key)
                        self.assertEqual(after.cells[key].value_type, "numeric", key)
                    self.assertEqual(
                        {key: cell.formula for key, cell in before.cells.items() if cell.has_formula},
                        {key: cell.formula for key, cell in after.cells.items() if cell.has_formula},
                    )
                    with ZipFile(template) as original, ZipFile(output) as generated:
                        for name in original.namelist():
                            if name.startswith(("xl/charts/", "xl/drawings/", "xl/media/")):
                                self.assertEqual(original.read(name), generated.read(name), name)

    def test_special_profile_rules_use_confirmed_actual_cells(self) -> None:
        with TemporaryDirectory() as temporary:
            database = MockDatabaseService(Path(temporary) / "special.db")
            preview_service = PreviewExcelExportService(database, XlsxTemplateInspector())
            parser = LabSolutionsParser()
            expected_cells = {
                "DMF,DMA": {("LOD(area입력)", "G7"): 13466, ("LOD(area입력)", "J7"): 17733},
                "스토다드솔벤트": {
                    ("LOD(area입력)", "F5"): 1109240,
                    ("LOD(area입력)", "G5"): 981002,
                    ("LOD(area입력)", "H5"): 1497,
                    ("LOD(area입력)", "E19"): 0,
                    ("LOD(area입력)", "E20"): 0,
                },
                "알콜4": {("area입력", "G5"): 34009, ("area입력", "P9"): 524758},
                "1,2-에폭시프로판(산화프로필렌)": {
                    ("결과입력(area입력)", "F14"): 6554,
                    ("결과입력(area입력)", "F26"): 1436,
                },
                "디클로로메탄(MC)": {("LOD(area입력)", "F4"): 4095},
                "메틸 n아밀케톤": {("LOD(area입력)", "E4"): 14458},
                "비닐아세테이트": {("LOD(area입력)", "F4"): 7634},
                "이소프로필 아세테이트": {("LOD(area입력)", "E20"): 1051},
                "피리딘": {("LOD(area입력)", "F4"): 29575},
            }
            for analysis_type, pdf, template, start, end, *_ in ACTUAL_CASES:
                with self.subTest(analysis_type=analysis_type):
                    batch = parser.parse(
                        pdf,
                        analysis_type=analysis_type,
                        analysis_no_start=start,
                        analysis_no_end=end,
                    )
                    preview = preview_service.preview_batch(batch, template, "A")
                    values = {
                        (row.target_sheet, row.target_cell): row.applied_area
                        for row in preview.rows
                        if row.status is ExcelPreviewStatus.MAPPED
                    }
                    for key, expected in expected_cells[analysis_type].items():
                        self.assertEqual(values.get(key), expected, key)

            stoddard = ACTUAL_CASES[1]
            workbook = XlsxTemplateInspector().inspect(stoddard[2])
            self.assertTrue(workbook.cell("LOD(area입력)", "E5").has_formula)
            self.assertEqual(1109240 - 981002 - 1497, 126741)


if __name__ == "__main__":
    unittest.main()
