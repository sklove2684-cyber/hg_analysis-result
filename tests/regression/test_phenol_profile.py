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
from honyu_app.application.sample_number_matching import extract_excel_analysis_number
from honyu_app.domain.enums import ExcelPreviewStatus, ExcludeReason
from honyu_app.infrastructure.database.mock_database_service import MockDatabaseService
from honyu_app.infrastructure.excel.workbook_inspector import XlsxTemplateInspector
from honyu_app.infrastructure.excel.workbook_validator import XlsxWorkbookValidator
from honyu_app.infrastructure.excel.xml_cell_writer import XlsxXmlCellWriter
from honyu_app.infrastructure.pdf.labsolutions_parser import LabSolutionsParser


MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _candidate_directories() -> tuple[Path, ...]:
    configured = os.environ.get("HONYU_PHENOL_TEST_DIR")
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


PDF = _find_file("(페놀) 256-305.pdf")
TEMPLATE = _find_file("(페놀) 256-305.xlsx")
MIXTURE_TEMPLATE = _find_file("(혼유) 601-690.xlsx")


@unittest.skipUnless(
    PDF.is_file() and TEMPLATE.is_file(), "페놀 실제 PDF/XLSX가 없습니다."
)
class PhenolActualFileTests(unittest.TestCase):
    def _saved_batch(self, database: MockDatabaseService):
        batch = LabSolutionsParser().parse(
            PDF,
            analysis_type="페놀",
            analysis_no_start=256,
            analysis_no_end=305,
        )
        self.assertEqual(batch.warning_count, 0)
        self.assertEqual(
            sum(
                peak.exclude_reason is ExcludeReason.UNKNOWN_MATERIAL
                for sample in batch.samples
                for peak in sample.peaks
            ),
            0,
        )
        methanol = [
            peak
            for sample in batch.samples
            for peak in sample.peaks
            if peak.material_standard == "Methanol"
        ]
        self.assertEqual(len(methanol), 53)
        self.assertFalse(any(peak.include_for_excel for peak in methanol))
        review = ReviewExtractionService(database)
        review.complete_review(batch)
        return review.save_batch(batch)

    def _preview(self, method: str):
        temporary = TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        database = MockDatabaseService(Path(temporary.name) / "phenol.db")
        saved = self._saved_batch(database)
        preview = PreviewExcelExportService(
            database, XlsxTemplateInspector()
        ).preview(saved.batch_id, TEMPLATE, method)
        return database, saved, preview

    def test_actual_template_structure_and_analysis_numbers(self) -> None:
        snapshot = XlsxTemplateInspector().inspect(TEMPLATE)

        self.assertEqual(
            snapshot.sheet_names,
            ("검량선", "LOD(area입력)", "회수율", "std"),
        )
        self.assertEqual(snapshot.cell("LOD(area입력)", "E2").value, "Phenol")
        analysis_numbers = tuple(
            number
            for row in range(19, 123)
            if (
                number := extract_excel_analysis_number(
                    snapshot.cell("LOD(area입력)", f"A{row}").value
                )
            )
            is not None
        )
        self.assertEqual(
            analysis_numbers,
            (
                "256", "257", "277", "278", "281", "282", "287", "288",
                "293", "294", "295", "296", "302", "303", "304", "305",
            ),
        )

    def test_actual_pdf_maps_confirmed_cells_for_common_std_methods(self) -> None:
        expected_by_method = {
            "A": {"F4": 32834, "F5": 64277, "F6": 128243, "F7": 262225, "F8": 539483},
            "B": {"F4": 32834, "F5": 64277, "F6": 128243, "F7": 262225, "F8": 688496},
        }
        expected_recovery = {
            "B28": 63822,
            "B29": 63993,
            "B30": 63644,
            "B31": 199622,
            "B32": 201507,
            "B33": 199485,
            "B34": 473548,
            "B35": 474074,
            "B36": 475858,
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
                self.assertEqual(preview.mapped_count, 14)
                self.assertEqual(preview.excluded_count, 54)
                self.assertEqual(
                    {
                        address: mapped[("LOD(area입력)", address)]
                        for address in expected_std
                    },
                    expected_std,
                )
                self.assertEqual(
                    {
                        address: mapped[("회수율", address)]
                        for address in expected_recovery
                    },
                    expected_recovery,
                )
                self.assertFalse(
                    any(
                        row.material == "Methanol"
                        and row.status is ExcelPreviewStatus.MAPPED
                        for row in preview.rows
                    )
                )

    def test_actual_export_writes_numbers_and_preserves_workbook_structure(self) -> None:
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            database = MockDatabaseService(root / "phenol.db")
            saved = self._saved_batch(database)
            preview_service = PreviewExcelExportService(
                database, XlsxTemplateInspector()
            )
            output = root / "phenol-result.xlsx"
            result = CreateExcelExportService(
                database,
                preview_service,
                XlsxXmlCellWriter(),
                XlsxWorkbookValidator(),
                object(),
                recalculate_with_excel=False,
            ).create(
                saved.batch_id,
                TEMPLATE,
                output,
                "A",
                "PHENOL-REGRESSION",
            )

            self.assertTrue(result.validation_passed)
            self.assertFalse(result.recalculated)
            self.assertEqual(result.mapped_cell_count, 14)
            before = XlsxTemplateInspector().inspect(TEMPLATE)
            after = XlsxTemplateInspector().inspect(output)
            for address, expected in {
                "F4": 32834,
                "F8": 539483,
            }.items():
                cell = after.cell("LOD(area입력)", address)
                self.assertEqual(cell.value, expected)
                self.assertEqual(cell.value_type, "numeric")
            for address, expected in {"B28": 63822, "B36": 475858}.items():
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

    @unittest.skipUnless(
        MIXTURE_TEMPLATE.is_file(), "비교용 혼유 Excel 양식이 없습니다."
    )
    def test_actual_mixture_workbook_is_rejected_as_profile_mismatch(self) -> None:
        with TemporaryDirectory() as temporary:
            database = MockDatabaseService(Path(temporary) / "phenol.db")
            saved = self._saved_batch(database)
            result = PreviewExcelExportService(
                database, XlsxTemplateInspector()
            ).preview(saved.batch_id, MIXTURE_TEMPLATE, "A")

        self.assertFalse(result.can_generate)
        self.assertEqual(result.issues[0].code, "TEMPLATE_PROFILE_MISMATCH")
        self.assertIn("페놀", result.issues[0].message)
        self.assertIn("혼유", result.issues[0].message)


if __name__ == "__main__":
    unittest.main()
