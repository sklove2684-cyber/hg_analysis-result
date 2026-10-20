from __future__ import annotations

from dataclasses import replace
from pathlib import Path
import os
from tempfile import TemporaryDirectory
import unittest

import pdfplumber

from honyu_app.application.heavy_metal_recovery import find_recovery_blocks
from honyu_app.application.preview_excel_export import PreviewExcelExportService
from honyu_app.domain.commands import SaveAnalysisBatchCommand
from honyu_app.infrastructure.database.mock_database_service import MockDatabaseService
from honyu_app.infrastructure.excel.workbook_inspector import XlsxTemplateInspector
from honyu_app.infrastructure.pdf.labsolutions_parser import LabSolutionsParser


ACTUAL_DIR = Path(os.environ.get(
    "HONYU_385_449_TEST_DIR",
    r"\\172.30.1.100\data\분석결과(사업장별)★\자동화프로그램@@\10.19\Na, K 385-449",
))
PDF = ACTUAL_DIR / "385-449.pdf"
TEMPLATE = Path(os.environ.get(
    "HONYU_385_449_TEMPLATE",
    str(Path.home() / "Desktop/넷하드/분석결과(사업장별)★/중금속/0. 틀-수산화.xlsx"),
))


def actual_files_available() -> bool:
    try:
        return PDF.is_file() and TEMPLATE.is_file()
    except OSError:
        return False


@unittest.skipUnless(actual_files_available(), "385-449 실제 PDF/XLSX가 필요합니다.")
class HeavyMetal385To449ActualRegressionTests(unittest.TestCase):
    def test_g1_survives_reextract_database_replacement_and_preview(self) -> None:
        with pdfplumber.open(PDF) as document:
            tables = document.pages[0].extract_tables()
        self.assertEqual("G#", tables[0][1][4])

        extracted = LabSolutionsParser().parse(
            PDF,
            analysis_type="중금속",
            analysis_no_start=385,
            analysis_no_end=449,
        )
        self.assertEqual(24, len(extracted.heavy_metal_recovery_values))
        self.assertEqual({"K", "Na"}, {
            value.element for value in extracted.heavy_metal_recovery_values
        })
        self.assertEqual({1}, {
            value.group_no for value in extracted.heavy_metal_recovery_values
        })

        with TemporaryDirectory() as directory:
            database = MockDatabaseService(Path(directory) / "actual.db")
            stale = replace(
                extracted,
                heavy_metal_recovery_values=[
                    replace(value, group_no=None)
                    for value in extracted.heavy_metal_recovery_values
                ],
            )
            saved = database.save_analysis_batch(SaveAnalysisBatchCommand(stale))
            replaced = database.replace_analysis_batch(
                saved.batch_id, SaveAnalysisBatchCommand(extracted)
            )
            loaded = database.get_batch_detail(replaced.batch_id)
            self.assertEqual({1}, {
                value.group_no for value in loaded.heavy_metal_recovery_values
            })

            inspector = XlsxTemplateInspector()
            snapshot = inspector.inspect(TEMPLATE)
            blocks = [
                block for block in find_recovery_blocks(snapshot)
                if block.element in {"Na", "K"}
            ]
            self.assertEqual({1, 2}, {block.group_no for block in blocks})

            preview = PreviewExcelExportService(database, inspector).preview(
                replaced.batch_id, TEMPLATE, "A"
            )
            self.assertTrue(preview.can_generate, preview.issues)
            self.assertEqual(24, preview.mapped_count)
            self.assertEqual(0, preview.error_count)
            self.assertEqual({"G#=1"}, {
                row.message.removesuffix(" L") for row in preview.rows
            })
            selected = [
                block for block in blocks
                if block.group_no == 1 and block.sheet == snapshot.active_sheet
            ]
            expected_targets = {
                (block.sheet, block.target(level, replicate))
                for block in selected
                for level in ("blank", "low", "mid", "high")
                for replicate in (1, 2, 3)
            }
            self.assertEqual(
                expected_targets,
                {(row.target_sheet, row.target_cell) for row in preview.rows},
            )


if __name__ == "__main__":
    unittest.main()
