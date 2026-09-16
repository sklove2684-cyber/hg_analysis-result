from __future__ import annotations

from pathlib import Path
import re

import pdfplumber
from pdfminer.pdfparser import PDFSyntaxError


HEAVY_METAL_ELEMENT_PROFILES = (
    ("Fe", "Mn", "Al", "Cr", "Sn", "Ti", "Cu", "Zr", "Zn", "Ni"),
    ("Fe", "Mn", "Al", "Cr", "Sn", "Ti", "Cu", "Zr", "Pb"),
)
MAX_CONTENT_SCAN_PAGES = 3


def detect_analysis_type_from_pdf_content(pdf_path: Path) -> str | None:
    """Inspect the first few pages for a distinctive heavy-metal results layout."""
    try:
        with pdfplumber.open(Path(pdf_path)) as pdf:
            if not pdf.pages:
                return None
            text = "\n".join(
                page.extract_text() or ""
                for page in pdf.pages[:MAX_CONTENT_SCAN_PAGES]
            )
    except (OSError, ValueError, PDFSyntaxError):
        return None

    compact = " ".join(text.split())
    if not re.search(r"\bList\s+of\s+Results\b", compact, re.IGNORECASE):
        return None

    quant_count = len(re.findall(r"\bQuant\b", compact, re.IGNORECASE))
    average_count = len(re.findall(r"\bAverage\b", compact, re.IGNORECASE))
    for elements in HEAVY_METAL_ELEMENT_PROFILES:
        elements_found = all(
            re.search(rf"\b{re.escape(element)}\b", compact) is not None
            for element in elements
        )
        if elements_found and quant_count >= len(elements) and average_count >= len(elements):
            return "중금속"
    return None
