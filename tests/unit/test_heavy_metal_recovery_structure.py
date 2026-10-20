from dataclasses import replace
from pathlib import Path
import sqlite3
from tempfile import TemporaryDirectory
import unittest

from honyu_app.application.preview_excel_export import PreviewExcelExportService
from honyu_app.application.heavy_metal_recovery import find_recovery_blocks
from honyu_app.domain.commands import SaveAnalysisBatchCommand
from honyu_app.infrastructure.database.mock_database_service import MockDatabaseService
from honyu_app.services.excel_template_service import ExcelTemplateSnapshot, TemplateCell
from tests.unit.test_heavy_metal_dynamic_excel import batch_for, na_k_snapshot, dynamic_snapshot
from tests.unit import test_heavy_metal_dynamic_parser as parser_fixtures


def grouped_snapshot(sheet="상반기 틀", detection="검출농도"):
    source = na_k_snapshot(sheet)
    cells = dict(source.cells)
    # Nonessential headers may be absent or have any spacing/wording.
    for address in ("D8", "G8", "H8", "I8", "J8"):
        cells.pop((sheet, address))
    cells[(sheet, "F8")] = TemplateCell(sheet, "F8", True, detection, "string")
    cells[(sheet, "A9")] = TemplateCell(sheet, "A9", True, 1, "numeric")
    cells[(sheet, "A18")] = TemplateCell(sheet, "A18", True, 1, "numeric")
    for (sh, address), cell in list(cells.items()):
        col = "".join(filter(str.isalpha, address))
        row = int("".join(filter(str.isdigit, address)))
        value = 2 if col == "A" and row in (9, 18) else cell.value
        moved = f"{col}{row + 30}"
        cells[(sh, moved)] = replace(cell, address=moved, value=value)
    return ExcelTemplateSnapshot(Path("test.xlsx"), (sheet, "intensity"), cells, sheet)


def preview(batch, snapshot):
    return PreviewExcelExportService._preview_heavy_metal(batch, snapshot.path, snapshot)


class HeavyMetalRecoveryStructureTests(unittest.TestCase):
    def test_any_sheet_name_and_input_header_alias_work_without_other_headers(self):
        for sheet in ("회수율", "상반기 틀", "230106"):
            for detection in ("검출량", "검출농도"):
                with self.subTest(sheet=sheet, detection=detection):
                    source = grouped_snapshot(sheet, detection)
                    source.cells = {k: v for k, v in source.cells.items()
                                    if int("".join(filter(str.isdigit, k[1]))) < 30}
                    result = preview(batch_for(("Na", "K")), source)
                    self.assertTrue(result.can_generate, result.issues)
                    self.assertEqual(24, result.mapped_count)
                    self.assertEqual({sheet}, {r.target_sheet for r in result.rows})

    def test_g2_uses_second_region_found_by_structure(self):
        batch = batch_for(("Na", "K"))
        batch.heavy_metal_recovery_values = [replace(v, group_no=2)
                                            for v in batch.heavy_metal_recovery_values]
        result = preview(batch, grouped_snapshot())
        self.assertTrue(result.can_generate, result.issues)
        self.assertEqual({"E39", "E40", "E41"},
                         {r.target_cell for r in result.rows
                          if r.material == "Na" and r.sample_type.value == "RECOVERY_BLANK"})
        self.assertIn("F56", {r.target_cell for r in result.rows})

    def test_missing_g2_or_missing_group_in_pdf_never_guesses(self):
        source = grouped_snapshot()
        result = preview(batch_for(("Na", "K")), source)
        self.assertFalse(result.can_generate)
        self.assertEqual(0, result.mapped_count)
        batch = batch_for(("Na", "K"))
        batch.heavy_metal_recovery_values = [replace(v, group_no=3)
                                            for v in batch.heavy_metal_recovery_values]
        result = preview(batch, source)
        self.assertIn("HEAVY_METAL_ELEMENT_NOT_IN_TEMPLATE", [i.code for i in result.issues])

    def test_active_sheet_disambiguates_duplicate_tables_but_not_multiple_regions(self):
        first = grouped_snapshot("상반기 틀")
        second = grouped_snapshot("230106")
        source = ExcelTemplateSnapshot(Path("test.xlsx"), ("상반기 틀", "230106", "intensity"),
                                       {**first.cells, **second.cells}, "230106")
        batch = batch_for(("Na", "K"))
        batch.heavy_metal_recovery_values = [replace(v, group_no=1)
                                            for v in batch.heavy_metal_recovery_values]
        result = preview(batch, source)
        self.assertTrue(result.can_generate, result.issues)
        self.assertEqual({"230106"}, {r.target_sheet for r in result.rows})
        source.active_sheet = "intensity"
        self.assertFalse(preview(batch, source).can_generate)

    def test_general_metals_also_ignore_sheet_names_and_calculation_headers(self):
        source = dynamic_snapshot((("Fe", "A3", 10), ("Pb", "K12", 20)))
        source.sheet_names = ("임의 시트", "intensity")
        source.cells = {("임의 시트", a): replace(c, sheet="임의 시트")
                        for (_, a), c in source.cells.items()}
        for key, cell in list(source.cells.items()):
            if cell.value == "검출량": source.cells[key] = replace(cell, value="검출농도")
        result = preview(batch_for(("Fe", "Pb")), source)
        self.assertTrue(result.can_generate, result.issues)
        self.assertEqual(24, result.mapped_count)

    def test_input_formula_and_merged_cells_are_blocked(self):
        for formula in (True, False):
            source = na_k_snapshot()
            if formula:
                source.cells[("230106", "E9")] = TemplateCell("230106", "E9", True, None,
                                                               "formula", "=1")
            else:
                source.merged_ranges = (("230106", "E9:E10"),)
            result = preview(batch_for(("Na", "K")), source)
            self.assertFalse(result.can_generate)
            expected = "TARGET_CELL_HAS_FORMULA" if formula else "HEAVY_METAL_TARGET_NOT_UNIQUE"
            self.assertIn(expected, [i.code for i in result.issues])

    def test_incomplete_level_structure_has_no_input_block(self):
        source = na_k_snapshot()
        del source.cells[("230106", "C12")]
        result = preview(batch_for(("Na", "K")), source)
        self.assertIn("HEAVY_METAL_ELEMENT_NOT_IN_TEMPLATE", [i.code for i in result.issues])

    def test_g_column_survives_pdf_extraction_and_database_roundtrip(self):
        first = parser_fixtures.results_table(("Na", "K"))
        for index, row in enumerate(first): row.insert(1, "G#" if index == 0 else "1")
        second = parser_fixtures.results_table(("Na", "K"))[1:]
        for row in second: row.insert(1, "2")
        batch = parser_fixtures.HeavyMetalDynamicParserTests()._parse(first + second)
        self.assertEqual({1, 2}, {v.group_no for v in batch.heavy_metal_recovery_values})
        self.assertEqual(48, len(batch.heavy_metal_recovery_values))
        with TemporaryDirectory() as directory:
            database = MockDatabaseService(Path(directory) / "db.sqlite")
            saved = database.save_analysis_batch(SaveAnalysisBatchCommand(batch))
            loaded = database.get_batch_detail(saved.batch_id)
            self.assertCountEqual(batch.heavy_metal_recovery_values, loaded.heavy_metal_recovery_values)
            result = preview(loaded, grouped_snapshot())
            self.assertTrue(result.can_generate, result.issues)
            self.assertEqual(48, result.mapped_count)

    def test_pdf_with_no_recovery_rows_is_rejected(self):
        from honyu_app.domain.errors import ValidationError
        table = parser_fixtures.results_table(("Na", "K"))[:1]
        with self.assertRaisesRegex(ValidationError, "PDF_LAYOUT_MISMATCH"):
            parser_fixtures.HeavyMetalDynamicParserTests()._parse(table)

    def test_old_database_migration_preserves_values_and_defaults_group_to_unknown(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "old.sqlite"
            db = MockDatabaseService(path)
            batch = batch_for(("Na", "K"))
            batch.source_file = replace(batch.source_file, file_hash="a" * 64)
            saved = db.save_analysis_batch(SaveAnalysisBatchCommand(batch))
            with sqlite3.connect(path) as connection:
                sql = connection.execute("SELECT sql FROM sqlite_master WHERE name='heavy_metal_recovery_values'").fetchone()[0]
                old_sql = sql.replace("group_no INTEGER NOT NULL DEFAULT 0,", "").replace("element, group_no, level", "element, level")
                connection.execute("ALTER TABLE heavy_metal_recovery_values RENAME TO original_values")
                connection.execute(old_sql)
                columns = [r[1] for r in connection.execute("PRAGMA table_info(heavy_metal_recovery_values)")]
                names = ",".join(columns)
                connection.execute(f"INSERT INTO heavy_metal_recovery_values ({names}) SELECT {names} FROM original_values")
                connection.execute("DROP TABLE original_values")
            connection.close()
            loaded = MockDatabaseService(path).get_batch_detail(saved.batch_id)
            self.assertEqual(24, len(loaded.heavy_metal_recovery_values))
            self.assertTrue(all(v.group_no is None for v in loaded.heavy_metal_recovery_values))
