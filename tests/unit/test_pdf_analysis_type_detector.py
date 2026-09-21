import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from honyu_app.infrastructure.pdf.analysis_type_detector import (
    detect_analysis_type_from_pdf_content,
)


def heavy_metal_table(elements: tuple[str, ...]) -> list[list[str]]:
    header = ["Sample Name", *(f"{element}\nQuant\nAverage" for element in elements)]
    rows = [
        [name, *(["0.1"] * len(elements))]
        for name in ("회수율-B", "저1", "중1", "고1")
    ]
    return [header, *rows]


def gc_peak_table(materials: tuple[str, ...]) -> list[list[str]]:
    header = [
        "Peak#", "Ret. Time", "Area", "Height", "Conc.", "Unit", "Mark", "Name"
    ]
    rows = [
        [str(index), f"{index}.100", "100", "50", "", "", "", material]
        for index, material in enumerate(materials, 1)
    ]
    return [header, *rows, ["Total", "", "100", "", "", "", "", ""]]


class PdfAnalysisTypeContentDetectorTests(unittest.TestCase):
    @staticmethod
    def _detect(*pages: tuple[str, list[list[str]]]) -> str | None:
        page_mocks = []
        for text, table in pages:
            page = MagicMock()
            page.extract_text.return_value = text
            page.extract_tables.return_value = [table] if table else []
            page_mocks.append(page)
        document = MagicMock()
        document.pages = page_mocks
        context = MagicMock()
        context.__enter__.return_value = document
        with patch(
            "honyu_app.infrastructure.pdf.analysis_type_detector.pdfplumber.open",
            return_value=context,
        ):
            return detect_analysis_type_from_pdf_content(Path("123-150.pdf"))

    def test_dynamic_nine_element_layout_is_detected(self) -> None:
        elements = ("Fe", "Mn", "Al", "Cr", "Sn", "Ti", "Cu", "Zr", "Pb")
        self.assertEqual(
            "중금속",
            self._detect(("List of Results", heavy_metal_table(elements))),
        )

    def test_element_order_and_count_are_not_fixed(self) -> None:
        elements = (
            "In", "Pt", "Be", "As", "V", "Ba", "Sb", "Co", "Cd", "W", "Mg", "Fe"
        )
        self.assertEqual(
            "중금속",
            self._detect(("List of Results", heavy_metal_table(elements))),
        )

    def test_first_three_pages_are_scanned(self) -> None:
        table = heavy_metal_table(("Fe", "Mn", "Pb"))
        self.assertEqual(
            "중금속",
            self._detect(("cover", []), ("List of Results", table)),
        )
        self.assertIsNone(self._detect(
            ("cover", []), ("notes", []), ("other", []), ("List of Results", table)
        ))

    def test_numeric_named_gc_peak_table_is_not_misclassified(self) -> None:
        text = "<Sample Information>\nSample Name: 123\nPeak# R.Time Area Height Name"
        self.assertIsNone(self._detect((text, [])))

    def test_alcohol_two_is_detected_from_normalized_peak_materials(self) -> None:
        self.assertEqual(
            "(알콜2) IBA,1-BTOH",
            self._detect((
                "<Sample Information>", gc_peak_table(("IBA", "n-부탄올"))
            )),
        )

    def test_alcohol_four_unique_material_has_priority(self) -> None:
        for extra in ("IAA", "2-부탄올"):
            with self.subTest(extra=extra):
                self.assertEqual(
                    "알콜4",
                    self._detect((
                        "<Sample Information>",
                        gc_peak_table(("IBA", "1-BTOH", extra)),
                    )),
                )

    def test_incomplete_alcohol_or_unrelated_gc_is_not_misclassified(self) -> None:
        for materials in (("IBA",), ("n-BTOH",), ("IPA", "IAA")):
            with self.subTest(materials=materials):
                self.assertIsNone(self._detect((
                    "<Sample Information>", gc_peak_table(materials)
                )))

    def test_list_of_results_without_recovery_rows_is_not_enough(self) -> None:
        table = [[
            "Sample Name", "Fe\nQuant\nAverage",
            "Mn\nQuant\nAverage", "Pb\nQuant\nAverage",
        ]]
        self.assertIsNone(self._detect(("List of Results", table)))

    def test_fewer_than_three_element_columns_is_not_enough(self) -> None:
        table = heavy_metal_table(("Fe", "Pb"))
        self.assertIsNone(self._detect(("List of Results", table)))


if __name__ == "__main__":
    unittest.main()
