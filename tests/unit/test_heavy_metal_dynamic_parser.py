from pathlib import Path
import tempfile
import unittest
from unittest.mock import MagicMock, patch

from honyu_app.domain.errors import ValidationError
from honyu_app.infrastructure.pdf.heavy_metal_results_parser import (
    HeavyMetalResultsParser,
)


RECOVERY_NAMES = (
    "회수율-B", "저1", "중1", "고1",
    "회수율-B", "저2", "중2", "고2",
    "회수율-B", "저3", "중3", "고3",
)


def results_table(elements: tuple[str, ...]) -> list[list[str]]:
    header = ["Sample Name", *(f"{element}\nQuant\nAverage" for element in elements)]
    rows = [
        [name, *(str(row + (column / 100)) for column in range(len(elements)))]
        for row, name in enumerate(RECOVERY_NAMES, 1)
    ]
    return [header, *rows]


class HeavyMetalDynamicParserTests(unittest.TestCase):
    def _parse(self, table: list[list[str]]):
        page = MagicMock()
        page.extract_tables.return_value = [table]
        document = MagicMock()
        document.pages = [page]
        context = MagicMock()
        context.__enter__.return_value = document
        with tempfile.TemporaryDirectory() as directory:
            pdf = Path(directory) / "1-2.pdf"
            pdf.write_bytes(b"fixture")
            with patch(
                "honyu_app.infrastructure.pdf.heavy_metal_results_parser.pdfplumber.open",
                return_value=context,
            ):
                return HeavyMetalResultsParser().parse(
                    pdf, analysis_type="중금속", analysis_no_start=1, analysis_no_end=2
                )

    def test_9_10_12_and_16_elements_use_dynamic_expected_counts(self) -> None:
        master = (
            "Fe", "Mn", "Al", "Cr", "Sn", "Ti", "Cu", "Zr", "Pb", "Zn",
            "Mg", "W", "Cd", "Co", "Sb", "Ba",
        )
        for count in (9, 10, 12, 16):
            with self.subTest(count=count):
                batch = self._parse(results_table(master[:count]))
                self.assertEqual(count * 12, len(batch.heavy_metal_recovery_values))
                self.assertEqual(set(master[:count]), {
                    value.element for value in batch.heavy_metal_recovery_values
                })

    def test_recovery_blank_name_variants(self) -> None:
        for name in ('회수율-B', '회수율B', '회수율 B', '회수율-BLANK', '회수율-blank', '회수율 BLANK', '회수율_blank'):
            with self.subTest(name=name):
                table = results_table(("Fe",))
                for index in (1, 5, 9):
                    table[index][0] = name
                batch = self._parse(table)
                blanks = [v for v in batch.heavy_metal_recovery_values if v.level == "blank"]
                self.assertEqual([1, 2, 3], [v.replicate_no for v in blanks])

    def test_reordered_elements_are_preserved_by_symbol(self) -> None:
        elements = ("Pb", "Fe", "Zr", "Mn", "Cu", "Al", "Ti", "Sn", "Cr")
        batch = self._parse(results_table(elements))
        self.assertEqual(elements, tuple(dict.fromkeys(
            value.element for value in batch.heavy_metal_recovery_values
        )))

    def test_group_header_variants_survive_list_of_results_title_prefix(self) -> None:
        for header in ("G#", "G #", "G\n#", "Group#", "Group #"):
            with self.subTest(header=header):
                table = results_table(("Na", "K"))
                table[0].insert(0, header)
                for row in table[1:]:
                    row.insert(0, "2")
                table.insert(0, ["List of Results", None, None, None])
                batch = self._parse(table)
                self.assertEqual(24, len(batch.heavy_metal_recovery_values))
                self.assertEqual(
                    {2}, {value.group_no for value in batch.heavy_metal_recovery_values}
                )

    def test_group_header_split_across_rows_is_detected(self) -> None:
        table = results_table(("Na", "K"))
        table[0].insert(0, "#")
        for row in table[1:]:
            row.insert(0, "1")
        table.insert(0, ["List of Results\nG", None, None, None])
        batch = self._parse(table)
        self.assertEqual(
            {1}, {value.group_no for value in batch.heavy_metal_recovery_values}
        )

    def test_duplicated_element_header_is_rejected(self) -> None:
        with self.assertRaisesRegex(ValidationError, "PDF_LAYOUT_MISMATCH"):
            self._parse(results_table(("Fe", "Mn", "Fe")))

    def test_one_missing_recovery_value_is_rejected(self) -> None:
        table = results_table(("Fe", "Mn", "Pb"))
        table[5][2] = ""
        with self.assertRaisesRegex(ValidationError, "PDF_LAYOUT_MISMATCH"):
            self._parse(table)


if __name__ == "__main__":
    unittest.main()
