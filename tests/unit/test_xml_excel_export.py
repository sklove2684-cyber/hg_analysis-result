from pathlib import Path
import base64
import subprocess
import tempfile
import unittest
from decimal import Decimal
from uuid import uuid4
from xml.etree import ElementTree as ET
from zipfile import ZipFile
from unittest.mock import patch

from honyu_app.application.create_excel_export import CreateExcelExportService
from honyu_app.domain.enums import ExcelPreviewStatus, SampleType
from honyu_app.domain.errors import ExcelExportError, ExcelRecalculationError
from honyu_app.domain.models import (
    ExcelCellWrite,
    ExcelPreviewResult,
    ExcelPreviewRow,
)
from honyu_app.domain.results import ExportJobResult
from honyu_app.infrastructure.excel.workbook_inspector import XlsxTemplateInspector
from honyu_app.infrastructure.excel.excel_recalculator import ExcelComRecalculator
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
            workbook_xml = archive.read("xl/workbook.xml")
            workbook = ET.fromstring(workbook_xml)
            sheet_parts = XlsxXmlCellWriter._sheet_parts(
                {name: archive.read(name) for name in archive.namelist()}
            )
            area_xml = archive.read(sheet_parts["area"])
        calc = workbook.find(f"{{{MAIN}}}calcPr")
        self.assertEqual(calc.attrib["calcMode"], "auto")
        self.assertEqual(calc.attrib["fullCalcOnLoad"], "1")
        self.assertEqual(calc.attrib["forceFullCalc"], "1")
        self.assertEqual(calc.attrib["calcOnSave"], "1")
        self.assertIn(b'xmlns:x15=', workbook_xml)
        self.assertIn(b'mc:Ignorable="x15"', workbook_xml)
        self.assertIn(b'xmlns:x14ac=', area_xml)
        self.assertIn(b'mc:Ignorable="x14ac"', area_xml)

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

    def test_validator_detects_missing_office_compatibility_namespace(self) -> None:
        self.writer.write_copy(TEMPLATE, self.output, self.writes)
        broken = Path(self.temp.name) / "missing-namespace.xlsx"
        with ZipFile(self.output, "r") as source, ZipFile(broken, "w") as target:
            for info in source.infolist():
                data = source.read(info.filename)
                if info.filename == "xl/workbook.xml":
                    data = data.replace(
                        b' xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"',
                        b"",
                    )
                target.writestr(info, data)
        validation = self.validator.validate(
            TEMPLATE, broken, self.writes, after_excel_recalculation=False
        )
        self.assertFalse(validation.valid)
        self.assertTrue(any("x15" in error for error in validation.errors))

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
            database, Preview(), self.writer, self.validator, recalculator,
            recalculate_with_excel=True,
        )
        output = Path(self.temp.name) / "final.xlsx"
        result = service.create(batch_id, TEMPLATE, output, "A", "TEST-PC")

        self.assertTrue(output.is_file())
        self.assertTrue(recalculator.called)
        self.assertTrue(result.recalculated)
        self.assertEqual(result.output_path, output.resolve())
        self.assertEqual(result.export_job_id, job_id)
        self.assertEqual(database.command.output_path, str(output.resolve()))

    def test_unlicensed_office_still_creates_a_valid_recalculation_pending_file(self) -> None:
        batch_id = uuid4()

        class Preview:
            def preview(inner_self, actual_batch_id, template, method):
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
            def save_export_job(inner_self, command):
                return ExportJobResult(uuid4(), True)

        class UnlicensedRecalculator:
            def recalculate(inner_self, path):
                raise ExcelRecalculationError(
                    "Office 제품 인증 필요", code="OFFICE_NOT_ACTIVATED"
                )

        output = Path(self.temp.name) / "unlicensed-result.xlsx"
        result = CreateExcelExportService(
            Database(), Preview(), self.writer, self.validator, UnlicensedRecalculator(),
            recalculate_with_excel=True,
        ).create(batch_id, TEMPLATE, output, "A", "TEST-PC")

        self.assertTrue(output.is_file())
        self.assertTrue(result.validation_passed)
        self.assertFalse(result.recalculated)
        self.assertEqual(XlsxTemplateInspector().inspect(output).cell("area", "F15").value, 123)

    def test_creation_service_refuses_existing_output(self) -> None:
        existing = Path(self.temp.name) / "existing.xlsx"
        existing.write_bytes(b"do not overwrite")
        service = CreateExcelExportService(None, None, None, None, None)
        with self.assertRaises(ExcelExportError):
            service.create(uuid4(), TEMPLATE, existing, "A", "TEST-PC")
        self.assertEqual(existing.read_bytes(), b"do not overwrite")


class ExcelRecalculatorErrorTests(unittest.TestCase):
    def test_default_creation_skips_com_and_validates_input_structure(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            template = root / "template.xlsx"
            template.write_bytes(b"template")
            output = root / "final.xlsx"
            work = root / "work"
            validation_modes = []

            class Preview:
                def preview(self, *_args):
                    return ExcelPreviewResult(
                        template,
                        "A",
                        rows=[ExcelPreviewRow(
                            sample_name="STD1",
                            sample_type=SampleType.STD,
                            material="n-hexane",
                            peak_no=1,
                            retention_time=Decimal("1.0"),
                            area_raw=123,
                            applied_area=123,
                            target_sheet="area입력",
                            target_cell="F15",
                            status=ExcelPreviewStatus.MAPPED,
                        )],
                    )

            class Writer:
                def write_copy(self, _template, partial, _writes):
                    partial.write_bytes(b"generated xlsx")

            class Validator:
                def validate(self, *_args, after_excel_recalculation):
                    validation_modes.append(after_excel_recalculation)
                    return type("Validation", (), {"valid": True, "errors": ()})()

            class Recalculator:
                def recalculate(self, _path):
                    raise AssertionError("COM recalculation must not run by default")

            class Database:
                def save_export_job(self, _command):
                    return ExportJobResult(uuid4(), True)

            with patch(
                "honyu_app.application.create_excel_export.excel_work_dir",
                return_value=work,
            ):
                result = CreateExcelExportService(
                    Database(), Preview(), Writer(), Validator(), Recalculator()
                ).create(uuid4(), template, output, "A", "TEST-PC")

            self.assertTrue(output.is_file())
            self.assertEqual(validation_modes, [False, False])
            self.assertFalse(result.recalculated)
            self.assertEqual(list(work.iterdir()), [])

    def test_failed_export_is_kept_out_of_final_output_directory(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            final_directory = root / "Desktop" / "분석프로그램"
            diagnostics = root / "diagnostics"
            final_directory.mkdir(parents=True)
            partial = root / "work" / "partial.xlsx"
            partial.parent.mkdir()
            partial.write_bytes(b"diagnostic workbook")
            output = final_directory / "result.xlsx"

            with patch(
                "honyu_app.application.create_excel_export.failed_exports_dir",
                return_value=diagnostics,
            ):
                failed = CreateExcelExportService._preserve_failure(
                    partial, output, "검증실패"
                )

            self.assertFalse(partial.exists())
            self.assertTrue(failed.is_file())
            self.assertEqual(failed.parent, diagnostics)
            self.assertEqual(list(final_directory.iterdir()), [])

    def test_formula_cache_comparison_detects_updated_results(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            before = Path(temp) / "before.xlsx"
            after = Path(temp) / "after.xlsx"
            sheet = (
                '<worksheet xmlns="http://schemas.openxmlformats.org/'
                'spreadsheetml/2006/main"><sheetData><row r="1">'
                '<c r="A1"><f>1+1</f><v>{value}</v></c>'
                '</row></sheetData></worksheet>'
            )
            with ZipFile(before, "w") as archive:
                archive.writestr("xl/worksheets/sheet1.xml", sheet.format(value="1"))
            with ZipFile(after, "w") as archive:
                archive.writestr("xl/worksheets/sheet1.xml", sheet.format(value="2"))

            old = ExcelComRecalculator._formula_cache_values(before)
            new = ExcelComRecalculator._formula_cache_values(after)

            self.assertEqual(len(old), 1)
            self.assertNotEqual(old, new)

    def test_recalculator_uses_an_owned_excel_instance(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            workbook = Path(temp) / "input.xlsx"
            workbook.write_bytes(b"test")
            failed = subprocess.CompletedProcess(
                args=[], returncode=1, stdout="", stderr="EXCEL_NOT_FOUND"
            )
            with patch(
                "honyu_app.infrastructure.excel.excel_recalculator.subprocess.run",
                return_value=failed,
            ) as run:
                with self.assertRaises(ExcelRecalculationError):
                    ExcelComRecalculator().recalculate(workbook)

            command = run.call_args.args[0]
            encoded = command[command.index("-EncodedCommand") + 1]
            script = base64.b64decode(encoded).decode("utf-16le")
            self.assertIn("New-Object -ComObject Excel.Application", script)
            self.assertIn("EXCEL_DEDICATED_INSTANCE_NOT_CREATED", script)
            self.assertIn("[ref]$excelProcessId", script)
            self.assertIn(
                "$existingExcelPids -contains [int]$excelProcessId", script
            )
            self.assertNotIn("[ref]$pid", script)
            self.assertIn("$excel.Workbooks.Open", script)
            self.assertIn("$book.ForceFullCalculation = $false", script)
            self.assertIn("$excel.Calculate()", script)
            self.assertNotIn("$book.Calculate()", script)
            self.assertNotIn("CalculateFullRebuild", script)
            self.assertIn("$worksheet.Calculate()", script)
            self.assertIn("$formulaCells.Calculate()", script)
            self.assertIn("$excel.CalculateFull()", script)
            self.assertIn("method=excel.Calculate()", script)
            self.assertIn("method=Worksheet.Calculate()", script)
            self.assertIn("method=Range.Calculate()", script)
            self.assertIn("method=excel.CalculateFull()", script)
            self.assertIn("$excel.Workbooks.Count -ne 1", script)
            self.assertIn("$originalCalculationMode = $excel.Calculation", script)
            self.assertIn("Write-Diagnostic 'calculation_mode_before'", script)
            self.assertIn("$excel.Calculation = -4105", script)
            self.assertIn("Write-WorkbookState 'automatic_calculation_enabled'", script)
            self.assertLess(
                script.index("$excel.Calculation = -4105"),
                script.index("$excel.Calculate()"),
            )
            self.assertLess(
                script.index("$book = $excel.Workbooks.Open"),
                script.index("$bootstrapBook.Close($false)"),
            )
            self.assertIn("$calculationWatch =", script)
            self.assertIn("Write-WorkbookState 'workbook_opened'", script)
            self.assertIn("Write-WorkbookState 'workbook_calculation_timeout'", script)
            for diagnostic_field in (
                "excelPid=", "workbook=", "calculation=", "calculationState=",
                "ready=", "readOnly=", "saved=", "externalLinks=",
                "connections=", "queries=", "circularReference=",
                "calculateBeforeSave=", "iteration=", "enableEvents=",
                "displayAlerts=", "screenUpdating=",
            ):
                self.assertIn(diagnostic_field, script)
            self.assertIn("$excel -ne $null -and $ownsExcel", script)
            self.assertNotIn("GetActiveObject", script)
            self.assertNotIn("EXCEL_ALREADY_RUNNING", script)
            self.assertNotIn("Stop-Process -Id $process.Id", script)

    def test_unlicensed_office_has_a_clear_error_message(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            workbook = Path(temp) / "input.xlsx"
            workbook.write_bytes(b"test")
            failed = subprocess.CompletedProcess(
                args=[], returncode=1, stdout="", stderr="OFFICE_NOT_ACTIVATED"
            )
            with patch(
                "honyu_app.infrastructure.excel.excel_recalculator.subprocess.run",
                return_value=failed,
            ):
                with self.assertRaisesRegex(ExcelRecalculationError, "제품 인증"):
                    ExcelComRecalculator().recalculate(workbook)


if __name__ == "__main__":
    unittest.main()
