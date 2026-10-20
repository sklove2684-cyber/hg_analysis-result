import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from honyu_app.application.create_excel_export import CreateExcelExportService
from honyu_app.application.preview_excel_export import PreviewExcelExportService
from honyu_app.domain.commands import SaveAnalysisBatchCommand
from honyu_app.infrastructure.database.mock_database_service import MockDatabaseService
from honyu_app.infrastructure.excel.workbook_inspector import XlsxTemplateInspector
from honyu_app.infrastructure.excel.workbook_validator import XlsxWorkbookValidator
from honyu_app.infrastructure.excel.xml_cell_writer import XlsxXmlCellWriter
from honyu_app.infrastructure.pdf.labsolutions_parser import LabSolutionsParser

ROOT = Path(os.environ.get("HONYU_HYDROXIDE_TEST_DIR", str(
    Path.home() / "Desktop/넷하드/분석결과(사업장별)★/중금속")))
TEMPLATE = ROOT / "0. 틀-수산화.xlsx"
PDF = ROOT / "Na,K 321-352.pdf"


def available():
    try:
        return TEMPLATE.is_file() and PDF.is_file()
    except OSError:
        return False


def assert_only_inputs_changed(test, original, output, preview):
    with ZipFile(original) as before, ZipFile(output) as after:
        test.assertIsNone(after.testzip())
        test.assertEqual(set(before.namelist()), set(after.namelist()))
        parts = {name: before.read(name) for name in before.namelist()}
        sheet_part = XlsxXmlCellWriter._sheet_parts(parts)["230106"]
        changed = {name for name in before.namelist() if before.read(name) != after.read(name)}
        test.assertEqual({sheet_part}, changed)
        ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        targets = {row.target_cell for row in preview.rows}
        roots = [ET.fromstring(archive.read(sheet_part)) for archive in (before, after)]
        for root in roots:
            for row in root.findall("m:sheetData/m:row", ns):
                for cell in list(row):
                    if cell.attrib.get("r") in targets:
                        row.remove(cell)
        test.assertEqual(ET.tostring(roots[0]), ET.tostring(roots[1]))


@unittest.skipUnless(available(), "Actual hydroxide template and Na/K PDF are unavailable")
class HydroxideActualRegressionTests(unittest.TestCase):
    def test_actual_pdf_db_preview_creation_changes_only_24_inputs(self):
        batch = LabSolutionsParser().parse(PDF, analysis_type="중금속",
                                            analysis_no_start=321, analysis_no_end=352)
        self.assertEqual({1}, {v.group_no for v in batch.heavy_metal_recovery_values})
        with TemporaryDirectory() as directory:
            database = MockDatabaseService(Path(directory) / "actual.db")
            saved = database.save_analysis_batch(SaveAnalysisBatchCommand(batch))
            service = PreviewExcelExportService(database, XlsxTemplateInspector())
            preview = service.preview(saved.batch_id, TEMPLATE, "A")
            self.assertTrue(preview.can_generate, preview.issues)
            self.assertEqual(0, preview.error_count)
            self.assertEqual(24, preview.mapped_count)
            self.assertEqual({"230106"}, {row.target_sheet for row in preview.rows})
            expected = {f"E{row}" for row in (61, 62, 63, 70, 71, 72)}
            expected.update(f"F{row}" for row in range(61, 79))
            self.assertEqual(expected, {row.target_cell for row in preview.rows})

            class NoRecalculation:
                def recalculate(self, path):
                    raise AssertionError("Recovery-only export must preserve other sheets and caches")

            output = Path(directory) / "result.xlsx"
            created = CreateExcelExportService(
                database, service, XlsxXmlCellWriter(), XlsxWorkbookValidator(),
                NoRecalculation(), recalculate_with_excel=True,
            ).create(saved.batch_id, TEMPLATE, output, "A", "REGRESSION")
            self.assertTrue(created.validation_passed)
            self.assertFalse(created.recalculated)
            final = XlsxTemplateInspector().inspect(output)
            for row in preview.rows:
                self.assertEqual(float(row.applied_area),
                                 final.cell(row.target_sheet, row.target_cell).value)
            assert_only_inputs_changed(self, TEMPLATE, output, preview)
