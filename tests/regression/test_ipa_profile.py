from __future__ import annotations

import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from honyu_app.application.create_excel_export import CreateExcelExportService
from honyu_app.application.preview_excel_export import PreviewExcelExportService
from honyu_app.application.review_extraction import ReviewExtractionService
from honyu_app.domain.enums import ExcelPreviewStatus, ExcludeReason
from honyu_app.infrastructure.database.mock_database_service import MockDatabaseService
from honyu_app.infrastructure.excel.workbook_inspector import XlsxTemplateInspector
from honyu_app.infrastructure.excel.workbook_validator import XlsxWorkbookValidator
from honyu_app.infrastructure.excel.xml_cell_writer import XlsxXmlCellWriter
from honyu_app.infrastructure.pdf.labsolutions_parser import LabSolutionsParser


MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _candidate_directories() -> tuple[Path, ...]:
    configured = os.environ.get("HONYU_IPA_TEST_DIR")
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


def _merge_ranges(archive: ZipFile) -> dict[str, tuple[str, ...]]:
    result: dict[str, tuple[str, ...]] = {}
    for name in archive.namelist():
        if not name.startswith("xl/worksheets/sheet") or not name.endswith(".xml"):
            continue
        root = ET.fromstring(archive.read(name))
        result[name] = tuple(
            node.attrib["ref"]
            for node in root.findall(f".//{{{MAIN}}}mergeCell")
        )
    return result


PDF = _find_file("(IPA) 320-334.pdf")
TEMPLATE = _find_file("(IPA) 320-334.xlsx")
MIXTURE_TEMPLATE = _find_file("(혼유) 601-690.xlsx")
AREA_PDF = _find_file("IPA 120-167.pdf")
AREA_TEMPLATE = _find_file("(IPA) 120-167.xlsx")
IPA_168_DIR = Path(
    os.environ.get(
        "HONYU_IPA_168_213_TEST_DIR",
        r"\\172.30.1.100\data\분석결과(사업장별)★\양세경\09.01\IPA 168-213",
    )
)
IPA_168_PDF = IPA_168_DIR / "IPA 168-213@완료.pdf"
IPA_168_TEMPLATE = IPA_168_DIR / "(IPA) 168-213.xlsx"


@unittest.skipUnless(PDF.is_file() and TEMPLATE.is_file(), "IPA 실제 PDF/XLSX가 없습니다.")
class IpaActualFileTests(unittest.TestCase):
    def _saved_batch(self, database: MockDatabaseService):
        batch = LabSolutionsParser().parse(
            PDF,
            analysis_type="IPA",
            analysis_no_start=320,
            analysis_no_end=334,
        )
        self.assertEqual(batch.warning_count, 0)
        self.assertEqual(
            sum(
                peak.exclude_reason is not None
                and peak.exclude_reason.value == "UNKNOWN_MATERIAL"
                for sample in batch.samples
                for peak in sample.peaks
            ),
            0,
        )
        review = ReviewExtractionService(database)
        review.complete_review(batch)
        return review.save_batch(batch)

    def _preview(self, method: str):
        temporary = TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        database = MockDatabaseService(Path(temporary.name) / "ipa.db")
        saved = self._saved_batch(database)
        preview = PreviewExcelExportService(
            database, XlsxTemplateInspector()
        ).preview(saved.batch_id, TEMPLATE, method)
        return database, saved, preview

    def test_actual_pdf_maps_confirmed_cells_for_common_std_methods(self) -> None:
        expected_by_method = {
            "A": {"J5": 13122, "J6": 27075, "J7": 54873, "J8": 116486, "J9": 236389},
            "B": {"J5": 13122, "J6": 27075, "J7": 54873, "J8": 116486, "J9": 298792},
        }
        for method, expected_std in expected_by_method.items():
            with self.subTest(method=method):
                _database, _saved, preview = self._preview(method)
                mapped = {
                    (row.target_sheet, row.target_cell): row.applied_area
                    for row in preview.rows
                    if row.status is ExcelPreviewStatus.MAPPED
                }

                self.assertTrue(preview.can_generate, preview.issues)
                self.assertEqual(preview.error_count, 0)
                self.assertEqual(preview.mapped_count, 20)
                self.assertEqual(preview.excluded_count, 91)
                self.assertEqual(
                    {
                        address: mapped[("LOD(area입력)", address)]
                        for address in expected_std
                    },
                    expected_std,
                )
                self.assertEqual(mapped[("회수율", "B28")], 28737)
                self.assertEqual(mapped[("회수율", "B36")], 208976)
                self.assertEqual(mapped[("LOD(area입력)", "J21")], 39281)
                self.assertEqual(mapped[("LOD(area입력)", "J22")], 29474)
                self.assertEqual(mapped[("LOD(area입력)", "J23")], 1065)

    def test_actual_export_writes_numbers_and_preserves_workbook_structure(self) -> None:
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            database = MockDatabaseService(root / "ipa.db")
            saved = self._saved_batch(database)
            preview_service = PreviewExcelExportService(
                database, XlsxTemplateInspector()
            )
            output = root / "ipa-result.xlsx"
            result = CreateExcelExportService(
                database,
                preview_service,
                XlsxXmlCellWriter(),
                XlsxWorkbookValidator(),
                object(),
                recalculate_with_excel=False,
            ).create(saved.batch_id, TEMPLATE, output, "A", "IPA-REGRESSION")

            self.assertTrue(result.validation_passed)
            self.assertFalse(result.recalculated)
            self.assertEqual(result.mapped_cell_count, 20)
            before = XlsxTemplateInspector().inspect(TEMPLATE)
            after = XlsxTemplateInspector().inspect(output)
            for address, expected in {
                "J5": 13122,
                "J9": 236389,
                "J21": 39281,
                "J22": 29474,
                "J23": 1065,
            }.items():
                cell = after.cell("LOD(area입력)", address)
                self.assertEqual(cell.value, expected)
                self.assertEqual(cell.value_type, "numeric")
            for address, expected in {"B28": 28737, "B36": 208976}.items():
                cell = after.cell("회수율", address)
                self.assertEqual(cell.value, expected)
                self.assertEqual(cell.value_type, "numeric")
            self.assertEqual(
                {key: cell.formula for key, cell in before.cells.items() if cell.has_formula},
                {key: cell.formula for key, cell in after.cells.items() if cell.has_formula},
            )
            self.assertEqual(
                {key: cell.style_id for key, cell in before.cells.items()},
                {key: cell.style_id for key, cell in after.cells.items()},
            )
            with ZipFile(TEMPLATE) as original, ZipFile(output) as generated:
                self.assertEqual(_merge_ranges(original), _merge_ranges(generated))
                preserved = tuple(
                    name
                    for name in original.namelist()
                    if name.startswith(("xl/charts/", "xl/drawings/", "xl/media/"))
                )
                self.assertTrue(any(name.startswith("xl/charts/") for name in preserved))
                for name in preserved:
                    self.assertEqual(original.read(name), generated.read(name), name)

    @unittest.skipUnless(MIXTURE_TEMPLATE.is_file(), "비교용 혼유 Excel 양식이 없습니다.")
    def test_actual_mixture_workbook_is_rejected_as_profile_mismatch(self) -> None:
        with TemporaryDirectory() as temporary:
            database = MockDatabaseService(Path(temporary) / "ipa.db")
            saved = self._saved_batch(database)
            result = PreviewExcelExportService(
                database, XlsxTemplateInspector()
            ).preview(saved.batch_id, MIXTURE_TEMPLATE, "A")

        self.assertFalse(result.can_generate)
        self.assertEqual(result.issues[0].code, "TEMPLATE_PROFILE_MISMATCH")
        self.assertIn("IPA", result.issues[0].message)
        self.assertIn("혼유", result.issues[0].message)


@unittest.skipUnless(
    AREA_PDF.is_file() and AREA_TEMPLATE.is_file(),
    "기존 area형 IPA 실제 PDF/XLSX가 없습니다.",
)
class IpaAreaActualFileTests(unittest.TestCase):
    def _saved_batch(self, database: MockDatabaseService):
        batch = LabSolutionsParser().parse(
            AREA_PDF,
            analysis_type="IPA",
            analysis_no_start=120,
            analysis_no_end=167,
        )
        review = ReviewExtractionService(database)
        review.complete_review(batch)
        return review.save_batch(batch)

    def _preview(self, method: str):
        temporary = TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        database = MockDatabaseService(Path(temporary.name) / "ipa-area.db")
        saved = self._saved_batch(database)
        preview = PreviewExcelExportService(
            database, XlsxTemplateInspector()
        ).preview(saved.batch_id, AREA_TEMPLATE, method)
        return database, saved, preview

    def test_actual_area_workbook_maps_confirmed_cells_for_both_std_methods(self) -> None:
        expected_by_method = {
            "A": [16195, 32517, 65745, 132189, 274697],
            "B": [16195, 32517, 65745, 132189, 360598],
        }
        expected_recovery = [
            30718,
            30658,
            30475,
            97861,
            97862,
            97349,
            240974,
            240796,
            237157,
        ]
        for method in ("A", "B"):
            with self.subTest(method=method):
                _database, _saved, preview = self._preview(method)
                mapped = {
                    (row.target_sheet, row.target_cell): row.applied_area
                    for row in preview.rows
                    if row.status is ExcelPreviewStatus.MAPPED
                }

                self.assertTrue(preview.can_generate, preview.issues)
                self.assertEqual(preview.error_count, 0)
                self.assertEqual(
                    [mapped[("area", f"J{row}")] for row in range(5, 10)],
                    expected_by_method[method],
                )
                self.assertEqual(
                    [mapped[("회수율", f"B{row}")] for row in range(28, 37)],
                    expected_recovery,
                )
                self.assertFalse(
                    any(sheet == "area" and cell.startswith("K") for sheet, cell in mapped)
                )
                self.assertEqual(
                    len(
                        {
                            row.target_cell
                            for row in preview.rows
                            if row.status is ExcelPreviewStatus.MAPPED
                            and row.target_sheet == "area"
                            and row.target_cell is not None
                            and 20 <= int(row.target_cell[1:]) <= 46
                        }
                    ),
                    len(
                        {
                            row.sample_name
                            for row in preview.rows
                            if row.status is ExcelPreviewStatus.MAPPED
                            and row.target_sheet == "area"
                            and row.target_cell is not None
                            and 20 <= int(row.target_cell[1:]) <= 46
                        }
                    ),
                )
                self.assertTrue(
                    any(
                        row.applied_area == 1067
                        and row.status is ExcelPreviewStatus.EXCLUDED
                        for row in preview.rows
                    )
                )
                self.assertTrue(
                    any(
                        row.applied_area == 1171
                        and row.status is ExcelPreviewStatus.EXCLUDED
                        for row in preview.rows
                    )
                )
                late_std2 = [
                    row
                    for row in preview.rows
                    if row.sample_name.casefold() == "std2"
                    and row.exclude_reason == "DUPLICATE_STD_SET"
                ]
                self.assertTrue(late_std2)

    def test_actual_area_export_preserves_non_target_content(self) -> None:
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            database = MockDatabaseService(root / "ipa-area.db")
            saved = self._saved_batch(database)
            preview_service = PreviewExcelExportService(
                database, XlsxTemplateInspector()
            )
            preview = preview_service.preview(saved.batch_id, AREA_TEMPLATE, "A")
            output = root / "ipa-area-result.xlsx"
            result = CreateExcelExportService(
                database,
                preview_service,
                XlsxXmlCellWriter(),
                XlsxWorkbookValidator(),
                object(),
                recalculate_with_excel=False,
            ).create(saved.batch_id, AREA_TEMPLATE, output, "A", "IPA-AREA-REGRESSION")

            self.assertTrue(result.validation_passed)
            before = XlsxTemplateInspector().inspect(AREA_TEMPLATE)
            after = XlsxTemplateInspector().inspect(output)
            mapped_cells = {
                (row.target_sheet, row.target_cell)
                for row in preview.rows
                if row.status is ExcelPreviewStatus.MAPPED
            }
            for row in range(20, 47):
                before_j = before.cell("area", f"J{row}")
                after_j = after.cell("area", f"J{row}")
                if before_j.value == "N.D" and ("area", f"J{row}") not in mapped_cells:
                    self.assertEqual(after_j.value, "N.D")
                before_k = before.cell("area", f"K{row}")
                after_k = after.cell("area", f"K{row}")
                self.assertEqual(
                    (after_k.value, after_k.formula, after_k.style_id),
                    (before_k.value, before_k.formula, before_k.style_id),
                )
            for sheet in ("Sheet1", "Sheet2"):
                self.assertEqual(
                    {
                        key: (cell.value, cell.formula, cell.style_id)
                        for key, cell in before.cells.items()
                        if key[0] == sheet
                    },
                    {
                        key: (cell.value, cell.formula, cell.style_id)
                        for key, cell in after.cells.items()
                        if key[0] == sheet
                    },
                )
            self.assertEqual(
                {key: cell.formula for key, cell in before.cells.items() if cell.has_formula},
                {key: cell.formula for key, cell in after.cells.items() if cell.has_formula},
            )
            self.assertEqual(
                {key: cell.style_id for key, cell in before.cells.items()},
                {key: cell.style_id for key, cell in after.cells.items()},
            )
            with ZipFile(AREA_TEMPLATE) as original, ZipFile(output) as generated:
                self.assertEqual(_merge_ranges(original), _merge_ranges(generated))
                for name in original.namelist():
                    if name.startswith(("xl/charts/", "xl/drawings/", "xl/media/")):
                        self.assertEqual(original.read(name), generated.read(name), name)


@unittest.skipUnless(
    IPA_168_PDF.is_file() and IPA_168_TEMPLATE.is_file(),
    "IPA 168-213 실제 PDF/XLSX가 없습니다.",
)
class Ipa168213MissingWorkerRegressionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.temporary = TemporaryDirectory()
        cls.database = MockDatabaseService(Path(cls.temporary.name) / "ipa-168-213.db")
        batch = LabSolutionsParser().parse(
            IPA_168_PDF,
            analysis_type="IPA",
            analysis_no_start=168,
            analysis_no_end=213,
        )
        review = ReviewExtractionService(cls.database)
        review.complete_review(batch)
        cls.saved = review.save_batch(batch)
        cls.preview_service = PreviewExcelExportService(
            cls.database, XlsxTemplateInspector()
        )

    @classmethod
    def tearDownClass(cls) -> None:
        cls.temporary.cleanup()

    def test_missing_pdf_analysis_numbers_are_excluded_without_errors(self) -> None:
        preview = self.preview_service.preview(
            self.saved.batch_id, IPA_168_TEMPLATE, "A"
        )
        required_missing = {"196", "197", "198", "201", "203", "205", "206", "209"}
        missing_rows = [
            row
            for row in preview.rows
            if row.exclude_reason == ExcludeReason.WORKER_ROW_NOT_IN_TEMPLATE.value
        ]
        excluded_numbers = {
            number
            for row in missing_rows
            for number in required_missing
            if row.sample_name.startswith(number)
        }

        self.assertTrue(preview.can_generate, preview.issues)
        self.assertEqual(preview.error_count, 0)
        self.assertTrue(required_missing.issubset(excluded_numbers))
        self.assertTrue(missing_rows)
        self.assertTrue(
            all(row.status is ExcelPreviewStatus.EXCLUDED for row in missing_rows)
        )
        self.assertTrue(
            all(
                row.message == "PDF 분석번호가 Excel 양식에 없어 입력 제외"
                for row in missing_rows
            )
        )
        self.assertGreater(preview.mapped_count, 0)

    def test_actual_ipa_168_213_excel_creation_succeeds(self) -> None:
        output = Path(self.temporary.name) / "ipa-168-213-result.xlsx"
        result = CreateExcelExportService(
            self.database,
            self.preview_service,
            XlsxXmlCellWriter(),
            XlsxWorkbookValidator(),
            object(),
            recalculate_with_excel=False,
        ).create(
            self.saved.batch_id,
            IPA_168_TEMPLATE,
            output,
            "A",
            "IPA-168-213-REGRESSION",
        )

        self.assertTrue(result.validation_passed)
        self.assertTrue(output.is_file())


if __name__ == "__main__":
    unittest.main()
