import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from honyu_app.infrastructure.pdf.analysis_type_detector import (
    detect_analysis_type_from_pdf_content,
)


LEGACY_ELEMENT_HEADER = "Fe Mn Al Cr Sn Ti Cu Zr Zn Ni"
PB_ELEMENT_HEADER = "Fe Mn Al Cr Sn Ti Cu Zr Pb"


class PdfAnalysisTypeContentDetectorTests(unittest.TestCase):
    @staticmethod
    def _detect(*page_texts: str) -> str | None:
        pages = []
        for text in page_texts:
            page = MagicMock()
            page.extract_text.return_value = text
            pages.append(page)
        document = MagicMock()
        document.pages = pages
        context = MagicMock()
        context.__enter__.return_value = document
        with patch(
            "honyu_app.infrastructure.pdf.analysis_type_detector.pdfplumber.open",
            return_value=context,
        ):
            return detect_analysis_type_from_pdf_content(Path("123-150.pdf"))

    def test_split_heavy_metal_headers_are_detected(self) -> None:
        text = "\n".join((
            "List of Results", LEGACY_ELEMENT_HEADER,
            " ".join(["Quant"] * 10), " ".join(["Average"] * 10),
        ))
        self.assertEqual("중금속", self._detect(text))

    def test_pb_nine_element_layout_is_detected_without_sequence_dependency(self) -> None:
        text = "\n".join((
            "List of Results",
            "Fe Quant Average\nMn Quant Average\nAl Quant Average",
            "Cr Quant Average\nSn Quant Average\nTi Quant Average",
            "Cu Quant Average\nZr Quant Average\nPb Quant Average",
        ))
        self.assertEqual("중금속", self._detect(text))

    def test_first_three_pages_are_scanned(self) -> None:
        heavy_metal_page = "\n".join((
            "List of Results", PB_ELEMENT_HEADER,
            " ".join(["Quant"] * 9), " ".join(["Average"] * 9),
        ))
        self.assertEqual("중금속", self._detect("cover", heavy_metal_page))
        self.assertIsNone(self._detect("cover", "notes", "other", heavy_metal_page))

    def test_numeric_named_gc_peak_table_is_not_misclassified(self) -> None:
        text = "<Sample Information>\nSample Name: 123\nPeak# R.Time Area Height Name"
        self.assertIsNone(self._detect(text))

    def test_partial_element_header_is_not_enough(self) -> None:
        text = "List of Results\nFe Mn Al Cr\nQuant Quant Quant Quant\nAverage Average"
        self.assertIsNone(self._detect(text))

    def test_all_elements_without_result_table_structure_are_not_enough(self) -> None:
        text = f"List of Results\n{LEGACY_ELEMENT_HEADER}\nQuant Average"
        self.assertIsNone(self._detect(text))


if __name__ == "__main__":
    unittest.main()
