import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from honyu_app.infrastructure.pdf.analysis_type_detector import (
    detect_analysis_type_from_pdf_content,
)


ELEMENT_HEADER = "Fe Mn Al Cr Sn Ti Cu Zr Zn Ni"


class PdfAnalysisTypeContentDetectorTests(unittest.TestCase):
    @staticmethod
    def _detect(text: str) -> str | None:
        page = MagicMock()
        page.extract_text.return_value = text
        document = MagicMock()
        document.pages = [page]
        context = MagicMock()
        context.__enter__.return_value = document
        with patch(
            "honyu_app.infrastructure.pdf.analysis_type_detector.pdfplumber.open",
            return_value=context,
        ):
            return detect_analysis_type_from_pdf_content(Path("123-150.pdf"))

    def test_split_heavy_metal_headers_are_detected(self) -> None:
        text = "\n".join((
            "List of Results", ELEMENT_HEADER,
            " ".join(["Quant"] * 10), " ".join(["Average"] * 10),
        ))
        self.assertEqual("중금속", self._detect(text))

    def test_numeric_named_gc_peak_table_is_not_misclassified(self) -> None:
        text = "<Sample Information>\nSample Name: 123\nPeak# R.Time Area Height Name"
        self.assertIsNone(self._detect(text))

    def test_partial_element_header_is_not_enough(self) -> None:
        text = "List of Results\nFe Mn Al Cr\nQuant Quant Quant Quant\nAverage Average"
        self.assertIsNone(self._detect(text))


if __name__ == "__main__":
    unittest.main()
