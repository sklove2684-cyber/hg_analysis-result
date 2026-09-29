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

    def test_dmf_is_detected_from_method_filename(self) -> None:
        self.assertEqual(
            "DMF,DMA",
            self._detect((
                "Sample Name : BLANK\nMethod Filename : DMF(1).gcm",
                gc_peak_table(()),
            )),
        )

    def test_dmf_is_detected_from_std_peak_material(self) -> None:
        self.assertEqual(
            "DMF,DMA",
            self._detect((
                "Sample Name : STD1\nMethod Filename : unknown.gcm",
                gc_peak_table(("DMF",)),
            )),
        )

    def test_dmf_in_non_std_peak_table_alone_is_not_enough(self) -> None:
        self.assertIsNone(self._detect((
            "Sample Name : 335\nMethod Filename : unknown.gcm",
            gc_peak_table(("DMF",)),
        )))

    def test_isoamyl_n_propyl_acetate_is_detected_from_std_materials(self) -> None:
        alias_pairs = (
            ("프로필아세테이트", "이소아밀아세테이트"),
            ("n-프로필 아세테이트", "초산이소아밀"),
            ("propyl acetate", "isoamyl acetate"),
        )
        for materials in alias_pairs:
            with self.subTest(materials=materials):
                self.assertEqual(
                    "이소아밀,n-프로필 아세테이트",
                    self._detect((
                        "Sample Name : STD1\nMethod Filename : unknown.gcm",
                        gc_peak_table(materials),
                    )),
                )

    def test_isoamyl_n_propyl_pair_has_priority_over_isopropyl_acetate(self) -> None:
        self.assertEqual(
            "이소아밀,n-프로필 아세테이트",
            self._detect((
                "Sample Name : STD1",
                gc_peak_table((
                    "n-propyl acetate",
                    "isoamyl acetate",
                    "isopropyl acetate",
                )),
            )),
        )

    def test_isopropyl_acetate_is_detected_from_std_material(self) -> None:
        for material in (
            "초산이소프로필",
            "이소프로필 아세테이트",
            "isopropyl acetate",
        ):
            with self.subTest(material=material):
                self.assertEqual(
                    "이소프로필 아세테이트",
                    self._detect((
                        "Sample Name : STD1",
                        gc_peak_table((material,)),
                    )),
                )

    def test_acetate_materials_outside_std_are_not_inferred(self) -> None:
        for materials in (
            ("n-propyl acetate", "isoamyl acetate"),
            ("isopropyl acetate",),
        ):
            with self.subTest(materials=materials):
                self.assertIsNone(self._detect((
                    "Sample Name : 346",
                    gc_peak_table(materials),
                )))

    def test_incomplete_or_unrelated_std_is_not_misclassified_as_acetate(self) -> None:
        for materials in (
            ("n-propyl acetate",),
            ("isoamyl acetate",),
            ("n-hexane", "acetone"),
        ):
            with self.subTest(materials=materials):
                self.assertIsNone(self._detect((
                    "Sample Name : STD1",
                    gc_peak_table(materials),
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
