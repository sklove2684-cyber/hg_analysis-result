from __future__ import annotations

from decimal import Decimal
import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import load_workbook

from honyu_app.application.preview_excel_export import PreviewExcelExportService
from honyu_app.infrastructure.excel.workbook_inspector import XlsxTemplateInspector
from honyu_app.infrastructure.excel.workbook_validator import XlsxWorkbookValidator
from honyu_app.infrastructure.excel.xml_cell_writer import XlsxXmlCellWriter
from honyu_app.infrastructure.pdf.analysis_type_detector import (
    detect_analysis_type_from_pdf_content,
)
from honyu_app.infrastructure.pdf.labsolutions_parser import LabSolutionsParser
from honyu_app.domain.models import ExcelCellWrite


ACTUAL_DIR = Path(os.environ.get(
    "HONYU_NAK_TEST_DIR",
    r"\\172.30.1.100\data\분석결과(사업장별)★\자동화프로그램@@\중금속",
))
PDF = ACTUAL_DIR / "Na,K 321-352.pdf"
TEMPLATE = ACTUAL_DIR / "Na,K 321-352.xlsx"


def _actual_files_available() -> bool:
    try:
        return PDF.is_file() and TEMPLATE.is_file()
    except OSError:
        return False


@unittest.skipUnless(_actual_files_available(), "Na,K 321-352 실제 PDF/XLSX가 필요합니다.")
class HeavyMetalNaKActualRegressionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.batch = LabSolutionsParser().parse(
            PDF,
            analysis_type="중금속",
            analysis_no_start=321,
            analysis_no_end=352,
        )
        cls.snapshot = XlsxTemplateInspector().inspect(TEMPLATE)

    def test_pdf_is_auto_detected_and_preserves_all_values_and_l_flags(self) -> None:
        self.assertEqual("중금속", detect_analysis_type_from_pdf_content(PDF))
        actual = {
            (item.element, item.level, item.replicate_no):
            (item.value, item.below_limit)
            for item in self.batch.heavy_metal_recovery_values
        }
        expected_values = {
            "Na": {
                "blank": ("0.0260", "0.0259", "0.0260"),
                "low": ("0.201", "0.203", "0.203"),
                "mid": ("0.382", "0.387", "0.382"),
                "high": ("0.773", "0.781", "0.781"),
            },
            "K": {
                "blank": ("0.0264", "0.0262", "0.0264"),
                "low": ("0.198", "0.199", "0.204"),
                "mid": ("0.388", "0.388", "0.393"),
                "high": ("0.769", "0.779", "0.781"),
            },
        }
        self.assertEqual(24, len(actual))
        for element, levels in expected_values.items():
            for level, values in levels.items():
                for replicate, value in enumerate(values, 1):
                    self.assertEqual(
                        (Decimal(value), level == "blank"),
                        actual[(element, level, replicate)],
                    )

    def test_structure_detection_maps_24_cells_without_errors(self) -> None:
        layout = PreviewExcelExportService._heavy_metal_na_k_layout(self.snapshot)
        self.assertEqual(("중금속", {"Na": 9, "K": 18}, "E", "F"), layout)
        result = PreviewExcelExportService._preview_heavy_metal(
            self.batch, TEMPLATE, self.snapshot
        )
        self.assertTrue(result.can_generate, result.issues)
        self.assertEqual(24, result.mapped_count)
        self.assertEqual(0, result.error_count)

    def test_result_has_24_exact_values_and_preserves_formula_cells(self) -> None:
        result = PreviewExcelExportService._preview_heavy_metal(
            self.batch, TEMPLATE, self.snapshot
        )
        writes = [
            ExcelCellWrite(row.target_sheet, row.target_cell, row.applied_area)
            for row in result.rows
        ]
        with TemporaryDirectory() as temporary:
            output = Path(temporary) / "Na,K 321-352_결과.xlsx"
            XlsxXmlCellWriter().write_copy(TEMPLATE, output, writes)
            validation = XlsxWorkbookValidator().validate(
                TEMPLATE, output, writes, after_excel_recalculation=False
            )
            self.assertTrue(validation.valid, validation.errors)
            workbook = load_workbook(output, data_only=False)
            try:
                sheet = workbook[result.rows[0].target_sheet]
                self.assertEqual(24, sum(
                    Decimal(str(sheet[row.target_cell].value))
                    == Decimal(str(row.applied_area))
                    for row in result.rows
                ))
                for address in ("G9", "H9", "I9", "J9", "G18", "H18", "I18", "J18"):
                    self.assertTrue(str(sheet[address].value).startswith("="), address)
            finally:
                workbook.close()


if __name__ == "__main__":
    unittest.main()
