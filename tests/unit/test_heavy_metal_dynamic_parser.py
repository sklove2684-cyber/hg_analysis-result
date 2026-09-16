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

    def test_reordered_elements_are_preserved_by_symbol(self) -> None:
        elements = ("Pb", "Fe", "Zr", "Mn", "Cu", "Al", "Ti", "Sn", "Cr")
        batch = self._parse(results_table(elements))
        self.assertEqual(elements, tuple(dict.fromkeys(
            value.element for value in batch.heavy_metal_recovery_values
        )))

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
