from pathlib import Path
import tempfile
import unittest
from decimal import Decimal
from uuid import uuid4
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from honyu_app.application.create_excel_export import CreateExcelExportService
from honyu_app.domain.enums import ExcelPreviewStatus, SampleType
from honyu_app.domain.errors import ExcelExportError
from honyu_app.domain.models import (
    ExcelCellWrite,
    ExcelPreviewResult,
    ExcelPreviewRow,
)
from honyu_app.domain.results import ExportJobResult
from honyu_app.infrastructure.excel.workbook_inspector import XlsxTemplateInspector
from honyu_app.infrastructure.excel.workbook_validator import XlsxWorkbookValidator
from honyu_app.infrastructure.excel.xml_cell_writer import MAIN, NS, XlsxXmlCellWriter


TEMPLATE = Path(__file__).parents[3] / "TEST" / "(혼유) 틀.xlsx"


@unittest.skipUnless(TEMPLATE.is_file(), "승인된 샘플 Excel이 없습니다.")
class XlsxXmlExportTests(unittest.TestCase):
    def setUp(self) -> None:
        self.writer = XlsxXmlCellWriter()
        self.validator = XlsxWorkbookValidator()
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.output = Path(self.temp.name) / "result.xlsx"
        self.writes = [
            ExcelCellWrite("area", "F15", 123),
            ExcelCellWrite("area", "Z15", 456),
            ExcelCellWrite("area", "AA15", 789),
            ExcelCellWrite("area", "Z37", 321),
            ExcelCellWrite("회수율", "U37", 111),
        ]

    def test_writer_changes_only_approved_numeric_cells(self) -> None:
        before = XlsxTemplateInspector().inspect(TEMPLATE)
        self.writer.write_copy(TEMPLATE, self.output, self.writes)
        after = XlsxTemplateInspector().inspect(self.output)

        self.assertEqual(after.cell("area", "F15").value, 123)
        self.assertEqual(after.cell("area", "Z37").value, 321)
        self.assertEqual(after.cell("회수율", "U37").value, 111)
        self.assertEqual(
            after.cell("area", "R15").formula,
            before.cell("area", "R15").formula,
        )
        self.assertEqual(
            after.cell("area", "R37").formula,
            before.cell("area", "R37").formula,
        )
        validation = self.validator.validate(
            TEMPLATE, self.output, self.writes, after_excel_recalculation=False
        )
        self.assertTrue(validation.valid, validation.errors)

        with ZipFile(self.output) as archive:
            workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        calc = workbook.find(f"{{{MAIN}}}calcPr")
        self.assertEqual(calc.attrib["calcMode"], "auto")
        self.assertEqual(calc.attrib["fullCalcOnLoad"], "1")
        self.assertEqual(calc.attrib["forceFullCalc"], "1")

    def test_formula_target_is_refused(self) -> None:
        with self.assertRaises(ExcelExportError):
            self.writer.write_copy(
                TEMPLATE,
                self.output,
                [ExcelCellWrite("area", "R15", 999)],
            )
        self.assertFalse(self.output.exists())

    def test_validator_detects_unapproved_value_change(self) -> None:
        self.writer.write_copy(TEMPLATE, self.output, self.writes)
        tampered = Path(self.temp.name) / "tampered.xlsx"
        with ZipFile(self.output, "r") as source, ZipFile(tampered, "w") as target:
            sheet_parts = XlsxXmlCellWriter._sheet_parts(
                {name: source.read(name) for name in source.namelist()}
            )
            for info in source.infolist():
                data = source.read(info.filename)
                if info.filename == sheet_parts["area"]:
                    root = ET.fromstring(data)
                    cell = root.find(".//m:c[@r='R15']", NS)
                    formula = cell.find("m:f", NS)
                    formula.text = "Z15+AA15+1"
                    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                target.writestr(info, data)
        validation = self.validator.validate(
            TEMPLATE, tampered, self.writes, after_excel_recalculation=False
        )
        self.assertFalse(validation.valid)
        self.assertTrue(any("area!R15" in error for error in validation.errors))

    def test_creation_service_promotes_only_validated_result_and_saves_job(self) -> None:
        batch_id = uuid4()
        job_id = uuid4()

        class Preview:
            def preview(inner_self, actual_batch_id, template, method):
                self.assertEqual(actual_batch_id, batch_id)
                return ExcelPreviewResult(
                    template,
                    str(method),
                    rows=[
                        ExcelPreviewRow(
                            sample_name="STD1",
                            sample_type=SampleType.STD,
                            material="n-hexane",
                            peak_no=1,
                            retention_time=Decimal("1.0"),
                            area_raw=123,
                            applied_area=123,
                            target_sheet="area",
                            target_cell="F15",
                            status=ExcelPreviewStatus.MAPPED,
                        )
                    ],
                )

        class Database:
            command = None

            def save_export_job(inner_self, command):
                inner_self.command = command
                return ExportJobResult(job_id, True)

        class Recalculator:
            called = False

            def recalculate(inner_self, path):
                inner_self.called = True

        database = Database()
        recalculator = Recalculator()
        service = CreateExcelExportService(
            database, Preview(), self.writer, self.validator, recalculator
        )
        output = Path(self.temp.name) / "final.xlsx"
        result = service.create(batch_id, TEMPLATE, output, "A", "TEST-PC")

        self.assertTrue(output.is_file())
        self.assertTrue(recalculator.called)
        self.assertEqual(result.output_path, output.resolve())
        self.assertEqual(result.export_job_id, job_id)
        self.assertEqual(database.command.output_path, str(output.resolve()))

    def test_creation_service_refuses_existing_output(self) -> None:
        existing = Path(self.temp.name) / "existing.xlsx"
        existing.write_bytes(b"do not overwrite")
        service = CreateExcelExportService(None, None, None, None, None)
        with self.assertRaises(ExcelExportError):
            service.create(uuid4(), TEMPLATE, existing, "A", "TEST-PC")
        self.assertEqual(existing.read_bytes(), b"do not overwrite")


if __name__ == "__main__":
    unittest.main()
