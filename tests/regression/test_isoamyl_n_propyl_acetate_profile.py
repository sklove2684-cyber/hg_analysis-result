from __future__ import annotations

import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from honyu_app.application.create_excel_export import CreateExcelExportService
from honyu_app.application.preview_excel_export import PreviewExcelExportService
from honyu_app.application.review_extraction import ReviewExtractionService
from honyu_app.domain.enums import ExcelPreviewStatus
from honyu_app.infrastructure.database.mock_database_service import MockDatabaseService
from honyu_app.infrastructure.excel.workbook_inspector import XlsxTemplateInspector
from honyu_app.infrastructure.excel.workbook_validator import XlsxWorkbookValidator
from honyu_app.infrastructure.excel.xml_cell_writer import XlsxXmlCellWriter
from honyu_app.infrastructure.pdf.labsolutions_parser import LabSolutionsParser


ACTUAL_DIR = Path(
    os.environ.get(
        "HONYU_ACETATE_127_130_TEST_DIR",
        r"\\172.30.1.100\data\분석결과(사업장별)★\양세경\09.02\아세테이트 127-130",
    )
)
PDF = ACTUAL_DIR / "아세테이트 127-130@완료.pdf"
TEMPLATE = ACTUAL_DIR / "(이소아밀,n-프로필 아세테이트) 127-130.xlsx"
ANALYSIS_TYPE = "이소아밀,n-프로필 아세테이트"


@unittest.skipUnless(
    PDF.is_file() and TEMPLATE.is_file(),
    "아세테이트 127-130 실제 PDF/XLSX가 없습니다.",
)
class IsoamylNPropylAcetateActualFileTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.batch = LabSolutionsParser().parse(
            PDF,
            analysis_type=ANALYSIS_TYPE,
            analysis_no_start=127,
            analysis_no_end=130,
        )

    def _preview(self, method: str):
        database = type(
            "NoCorrectionDatabase",
            (),
            {"list_peak_corrections": lambda self, _peak_id: []},
        )()
        return PreviewExcelExportService(
            database, XlsxTemplateInspector()
        ).preview_batch(self.batch, TEMPLATE, method)

    def test_actual_aliases_std_methods_and_recovery_values(self) -> None:
        self.assertEqual(len(self.batch.samples), 33)
        self.assertEqual(sum(len(sample.peaks) for sample in self.batch.samples), 105)
        self.assertFalse(
            any(
                peak.exclude_reason is not None
                and peak.exclude_reason.value == "UNKNOWN_MATERIAL"
                for sample in self.batch.samples
                for peak in sample.peaks
            )
        )
        first_standards = {}
        for sample in self.batch.samples:
            if sample.sample_name_normalized.startswith("STD"):
                first_standards.setdefault(sample.replicate_no, sample)
        expected_std = {
            1: (("4.811", 19368), ("7.266", 17191)),
            2: (("4.812", 38058), ("7.267", 32683)),
            3: (("4.812", 78497), ("7.267", 66969)),
            4: (("4.813", 160315), ("7.268", 135868)),
            5: (("4.813", 336203), ("7.270", 285896)),
            6: (("4.814", 423724), ("7.271", 360320)),
        }
        for replicate, expected in expected_std.items():
            peaks = [
                (str(peak.retention_time), peak.area_raw)
                for peak in first_standards[replicate].peaks
                if peak.material_standard in {
                    "n-프로필 아세테이트", "이소아밀 아세테이트"
                }
            ]
            self.assertEqual(tuple(peaks), expected, replicate)

        expected_last_std = {
            "A": {"J9": 336203, "G9": 285896},
            "B": {"J9": 423724, "G9": 360320},
        }
        for method in ("A", "B"):
            with self.subTest(method=method):
                preview = self._preview(method)
                mapped = {
                    row.target_cell: row.applied_area
                    for row in preview.rows
                    if row.status is ExcelPreviewStatus.MAPPED
                }
                self.assertTrue(preview.can_generate, preview.issues)
                self.assertEqual(preview.error_count, 0)
                self.assertEqual(preview.mapped_count, 28)
                self.assertEqual(preview.excluded_count, 77)
                self.assertEqual(mapped["J5"], 19368)
                self.assertEqual(mapped["G5"], 17191)
                self.assertEqual(
                    {cell: mapped[cell] for cell in ("J9", "G9")},
                    expected_last_std[method],
                )
                self.assertEqual(mapped["C28"], 40171)
                self.assertEqual(mapped["B28"], 34903)
                self.assertEqual(mapped["C36"], 302450)
                self.assertEqual(mapped["B36"], 258612)

    def test_actual_export_preserves_nd_and_workbook_structure(self) -> None:
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            database = MockDatabaseService(root / "acetate.db")
            review = ReviewExtractionService(database)
            batch = LabSolutionsParser().parse(
                PDF,
                analysis_type=ANALYSIS_TYPE,
                analysis_no_start=127,
                analysis_no_end=130,
            )
            review.complete_review(batch)
            saved = review.save_batch(batch)
            preview_service = PreviewExcelExportService(
                database, XlsxTemplateInspector()
            )
            output = root / "acetate-result.xlsx"
            result = CreateExcelExportService(
                database,
                preview_service,
                XlsxXmlCellWriter(),
                XlsxWorkbookValidator(),
                object(),
                recalculate_with_excel=False,
            ).create(saved.batch_id, TEMPLATE, output, "A", "ACETATE-REGRESSION")

            before = XlsxTemplateInspector().inspect(TEMPLATE)
            after = XlsxTemplateInspector().inspect(output)
            self.assertTrue(result.validation_passed)
            for row in range(20, 24):
                for column in ("F", "I"):
                    address = f"{column}{row}"
                    self.assertEqual(before.cell("LOD(area입력)", address).value, "N.D")
                    self.assertEqual(after.cell("LOD(area입력)", address).value, "N.D")
            self.assertEqual(
                {key: cell.formula for key, cell in before.cells.items() if cell.has_formula},
                {key: cell.formula for key, cell in after.cells.items() if cell.has_formula},
            )
            self.assertEqual(
                {key: cell.style_id for key, cell in before.cells.items()},
                {key: cell.style_id for key, cell in after.cells.items()},
            )
