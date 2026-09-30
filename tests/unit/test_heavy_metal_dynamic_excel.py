from datetime import datetime, timezone
from decimal import Decimal
from pathlib import Path
import unittest

from honyu_app.application.preview_excel_export import PreviewExcelExportService
from honyu_app.domain.models import (
    AnalysisBatch,
    HeavyMetalRecoveryValue,
    SourceFile,
)
from honyu_app.services.excel_template_service import ExcelTemplateSnapshot, TemplateCell


def batch_for(elements: tuple[str, ...]) -> AnalysisBatch:
    values = [
        HeavyMetalRecoveryValue(
            element=element,
            sample_name=("회수율-B" if level == "blank" else f"{label}{replicate}"),
            level=level,
            replicate_no=replicate,
            value=Decimal(f"{element_index + 1}.{replicate}"),
        )
        for element_index, element in enumerate(elements)
        for level, label in (("blank", ""), ("low", "저"), ("mid", "중"), ("high", "고"))
        for replicate in (1, 2, 3)
    ]
    return AnalysisBatch(
        batch_code="heavy-metal-test",
        source_file=SourceFile("1-2.pdf", Path("1-2.pdf"), "hash", 1, 1),
        analysis_type="중금속",
        analysis_no_start=1,
        analysis_no_end=2,
        parser_name="test",
        parser_version="1",
        parser_layout_id="dynamic",
        extracted_at=datetime.now(timezone.utc),
        heavy_metal_recovery_values=values,
    )


def dynamic_snapshot(blocks: tuple[tuple[str, str, int], ...]) -> ExcelTemplateSnapshot:
    cells: list[TemplateCell] = []
    for element, label_address, sequence in blocks:
        column_text = "".join(character for character in label_address if character.isalpha())
        base_row = int("".join(character for character in label_address if character.isdigit()))
        label_column = 0
        for character in column_text:
            label_column = label_column * 26 + ord(character) - ord("A") + 1

        def name(number: int) -> str:
            result = ""
            while number:
                number, remainder = divmod(number - 1, 26)
                result = chr(ord("A") + remainder) + result
            return result

        level_column = name(label_column + 1)
        blank_column = name(label_column + 3)
        recovery_column = name(label_column + 4)
        cells.extend((
            TemplateCell("회수율", label_address, True, f"{element}-{sequence}", "string"),
            TemplateCell("회수율", f"{level_column}{base_row}", True, "저", "string"),
            TemplateCell("회수율", f"{level_column}{base_row + 3}", True, "중", "string"),
            TemplateCell("회수율", f"{level_column}{base_row + 6}", True, "고", "string"),
            TemplateCell("회수율", f"{blank_column}2", True, "BLANK", "string"),
            TemplateCell("회수율", f"{recovery_column}2", True, "검출량", "string"),
        ))
    return ExcelTemplateSnapshot(
        Path("dynamic.xlsx"),
        ("LOD(고온물질)", "회수율", "분석결과"),
        {(cell.sheet, cell.address): cell for cell in cells},
    )


def na_k_snapshot(sheet: str = "230106") -> ExcelTemplateSnapshot:
    values = {
        "D8": "주입량", "E8": "BLANK", "F8": "검출량",
        "G8": "공시료 보정값", "H8": "평균", "I8": "회수율%",
        "J8": "평균 회수율%", "B9": "NaOH", "C9": "저",
        "C12": "중", "C15": "고", "B18": "KOH", "C18": "저",
        "C21": "중", "C24": "고",
    }
    cells = [
        TemplateCell(sheet, address, True, value, "string")
        for address, value in values.items()
    ]
    for row in range(9, 27):
        cells.extend((
            TemplateCell(sheet, f"E{row}", True, None, "blank"),
            TemplateCell(sheet, f"F{row}", True, None, "blank"),
            TemplateCell(sheet, f"G{row}", True, None, "formula", "=F1-E1"),
            TemplateCell(sheet, f"H{row}", True, None, "formula", "=AVERAGE(G1:G3)"),
            TemplateCell(sheet, f"I{row}", True, None, "formula", "=H1/D1*100"),
            TemplateCell(sheet, f"J{row}", True, None, "formula", "=AVERAGE(I1:I9)"),
        ))
    return ExcelTemplateSnapshot(
        Path("na-k.xlsx"), ("상반기 틀", sheet, "intensity"),
        {(cell.sheet, cell.address): cell for cell in cells},
    )


class HeavyMetalDynamicExcelTests(unittest.TestCase):
    def test_na_k_profile_is_detected_by_structure_and_maps_only_24_input_cells(self) -> None:
        result = PreviewExcelExportService._preview_heavy_metal(
            batch_for(("Na", "K")), Path("na-k.xlsx"), na_k_snapshot()
        )
        self.assertTrue(result.can_generate, result.issues)
        self.assertEqual(24, result.mapped_count)
        targets = {(row.material, row.target_cell) for row in result.rows}
        self.assertIn(("Na", "E9"), targets)
        self.assertIn(("Na", "F17"), targets)
        self.assertIn(("K", "E18"), targets)
        self.assertIn(("K", "F26"), targets)
        self.assertTrue(all(row.target_cell[0] in {"E", "F"} for row in result.rows))

    def test_na_k_profile_does_not_depend_on_sheet_name(self) -> None:
        result = PreviewExcelExportService._preview_heavy_metal(
            batch_for(("Na", "K")), Path("renamed.xlsx"),
            na_k_snapshot("회수율-날짜변경"),
        )
        self.assertTrue(result.can_generate, result.issues)
        self.assertEqual(
            {"회수율-날짜변경"}, {row.target_sheet for row in result.rows}
        )

    def test_na_k_profile_rejects_pdf_element_without_template_block(self) -> None:
        result = PreviewExcelExportService._preview_heavy_metal(
            batch_for(("Na", "K", "Pb")), Path("na-k.xlsx"), na_k_snapshot()
        )
        self.assertFalse(result.can_generate)
        self.assertEqual(
            "HEAVY_METAL_ELEMENT_NOT_IN_TEMPLATE", result.issues[0].code
        )
        self.assertIn("Pb", result.issues[0].message)

    def test_reordered_blocks_map_common_elements_and_leave_excel_only_untouched(self) -> None:
        pdf_elements = ("Fe", "Mn", "Al", "Cr", "Sn", "Ti", "Cu", "Zr", "Pb")
        blocks = (
            ("Pb", "A48", 10), ("Ni", "K39", 25), ("Fe", "A3", 1),
            ("Sn", "K3", 5), ("Mn", "A12", 2), ("Ti", "K12", 6),
            ("Al", "A21", 3), ("Cu", "K21", 7), ("Cr", "A30", 4),
            ("Zr", "K30", 8), ("Zn", "A39", 9),
        )
        result = PreviewExcelExportService._preview_heavy_metal(
            batch_for(pdf_elements), Path("dynamic.xlsx"), dynamic_snapshot(blocks)
        )
        self.assertTrue(result.can_generate, result.issues)
        self.assertEqual(108, result.mapped_count)
        self.assertEqual({"Zn", "Ni"}, {
            element
            for issue in result.issues
            if issue.code == "HEAVY_METAL_EXCEL_ONLY_ELEMENTS"
            for element in ("Zn", "Ni")
            if element in issue.message
        })
        self.assertFalse(any(row.material in {"Zn", "Ni"} for row in result.rows))
        self.assertEqual(12, sum(row.material == "Pb" for row in result.rows))

    def test_pdf_element_missing_from_template_blocks_generation(self) -> None:
        result = PreviewExcelExportService._preview_heavy_metal(
            batch_for(("Fe", "Cd")),
            Path("dynamic.xlsx"),
            dynamic_snapshot((("Fe", "A3", 1),)),
        )
        self.assertFalse(result.can_generate)
        self.assertEqual("HEAVY_METAL_ELEMENT_NOT_IN_TEMPLATE", result.issues[0].code)
        self.assertIn("Cd", result.issues[0].message)

    def test_non_contiguous_label_numbers_do_not_control_positions(self) -> None:
        result = PreviewExcelExportService._preview_heavy_metal(
            batch_for(("Fe", "Pb", "In")),
            Path("dynamic.xlsx"),
            dynamic_snapshot((("In", "U30", 25), ("Pb", "K12", 10), ("Fe", "A3", 1))),
        )
        self.assertTrue(result.can_generate, result.issues)
        self.assertIn("X30", {row.target_cell for row in result.rows if row.material == "In"})
        self.assertIn("N12", {row.target_cell for row in result.rows if row.material == "Pb"})


if __name__ == "__main__":
    unittest.main()
