from __future__ import annotations

from decimal import Decimal
from pathlib import Path
import os
import tempfile
import unittest

from honyu_app.application.preview_excel_export import PreviewExcelExportService
from honyu_app.domain.commands import SaveAnalysisBatchCommand
from honyu_app.domain.models import ExcelCellWrite
from honyu_app.infrastructure.database.mock_database_service import MockDatabaseService
from honyu_app.infrastructure.excel.workbook_inspector import XlsxTemplateInspector
from honyu_app.infrastructure.excel.workbook_validator import XlsxWorkbookValidator
from honyu_app.infrastructure.excel.xml_cell_writer import XlsxXmlCellWriter
from honyu_app.infrastructure.pdf.labsolutions_parser import LabSolutionsParser


ACTUAL_DIRECTORY = Path(os.environ.get(
    "HONYU_HEAVY_METAL_FIXTURE_DIR",
    r"\\172.30.1.100\data\분석결과(사업장별)★\자동화프로그램@@\중금속",
))
PDF = ACTUAL_DIRECTORY / "240-281.pdf"
XLSX = ACTUAL_DIRECTORY / "240-281.xlsx"


def _actual_files_available() -> bool:
    try:
        return PDF.is_file() and XLSX.is_file()
    except OSError:
        return False


EXPECTED = {
    "blank": (
        ("-0.0140", "-0.0081", "-0.0031", "-0.0052", "-0.0047", "-0.0037", "-0.0031", "-0.0137", "-0.0144", "-0.0035"),
        ("-0.0144", "-0.0082", "-0.0032", "-0.0055", "-0.0047", "-0.0039", "-0.0035", "-0.0144", "-0.0146", "-0.0033"),
        ("-0.0147", "-0.0082", "-0.0034", "-0.0054", "-0.0051", "-0.0040", "-0.0034", "-0.0144", "-0.0147", "-0.0034"),
    ),
    "low": (
        ("0.193", "0.198", "0.509", "0.0968", "0.197", "0.503", "0.0997", "0.493", "0.191", "0.102"),
        ("0.191", "0.196", "0.510", "0.0959", "0.194", "0.499", "0.0991", "0.488", "0.190", "0.101"),
        ("0.190", "0.196", "0.510", "0.0950", "0.193", "0.499", "0.0991", "0.487", "0.189", "0.100"),
    ),
    "mid": (
        ("0.589", "0.598", "1.52", "0.296", "0.589", "1.49", "0.296", "1.48", "0.591", "0.297"),
        ("0.582", "0.592", "1.52", "0.293", "0.586", "1.49", "0.295", "1.47", "0.588", "0.294"),
        ("0.581", "0.593", "1.52", "0.291", "0.581", "1.49", "0.296", "1.48", "0.587", "0.293"),
    ),
    "high": (
        ("1.18", "1.19", "3.01", "0.597", "1.18", "2.98", "0.596", "2.97", "1.19", "0.593"),
        ("1.18", "1.19", "3.03", "0.593", "1.17", "2.99", "0.596", "2.97", "1.19", "0.592"),
        ("1.17", "1.19", "3.01", "0.588", "1.16", "2.98", "0.595", "2.96", "1.19", "0.588"),
    ),
}
ELEMENTS = ("Fe", "Mn", "Al", "Cr", "Sn", "Ti", "Cu", "Zr", "Zn", "Ni")


@unittest.skipUnless(_actual_files_available(), "중금속 240-281 실제 PDF/XLSX가 필요합니다.")
class HeavyMetalActualRegressionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.batch = LabSolutionsParser().parse(
            PDF, analysis_type="중금속", analysis_no_start=240, analysis_no_end=281
        )

    def test_all_120_pdf_values_and_l_flags(self):
        actual = {(v.level, v.replicate_no, v.element): v for v in self.batch.heavy_metal_recovery_values}
        self.assertEqual(120, len(actual))
        for level, replicates in EXPECTED.items():
            for replicate, row in enumerate(replicates, 1):
                for element, expected in zip(ELEMENTS, row):
                    self.assertEqual(Decimal(expected), actual[level, replicate, element].value)
        l_keys = {(v.replicate_no, v.element) for v in self.batch.heavy_metal_recovery_values
                  if v.level == "blank" and v.below_limit}
        self.assertEqual({(r, e) for r in (1, 2, 3) for e in ("Al", "Sn", "Cu", "Zr")}, l_keys)

    def test_db_preview_and_final_xlsx_are_exact(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            db = MockDatabaseService(root / "db.sqlite")
            db.save_analysis_batch(SaveAnalysisBatchCommand(self.batch))
            loaded = db.get_batch_detail(self.batch.batch_id)
            self.assertEqual(120, len(loaded.heavy_metal_recovery_values))
            preview = PreviewExcelExportService(db, XlsxTemplateInspector()).preview(
                self.batch.batch_id, XLSX, "A"
            )
            self.assertTrue(preview.can_generate, preview.issues)
            self.assertEqual(120, preview.mapped_count)
            self.assertEqual(120, len({(r.target_sheet, r.target_cell) for r in preview.rows}))
            writes = [ExcelCellWrite(r.target_sheet, r.target_cell, r.applied_area) for r in preview.rows]
            output = root / "result.xlsx"
            XlsxXmlCellWriter().write_copy(XLSX, output, writes)
            validation = XlsxWorkbookValidator().validate(
                XLSX, output, writes, after_excel_recalculation=False
            )
            self.assertTrue(validation.valid, validation.errors)
            final = XlsxTemplateInspector().inspect(output)
            for write in writes:
                self.assertEqual(float(write.value), final.cell(write.sheet, write.address).value)
            self.assertEqual(("LOD(고온물질)", "회수율", "분석결과"), final.sheet_names)


if __name__ == "__main__":
    unittest.main()
