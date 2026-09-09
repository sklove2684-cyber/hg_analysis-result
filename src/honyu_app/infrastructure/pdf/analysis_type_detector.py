from __future__ import annotations

from pathlib import Path
import re

import pdfplumber
from pdfminer.pdfparser import PDFSyntaxError


HEAVY_METAL_ELEMENTS = ("Fe", "Mn", "Al", "Cr", "Sn", "Ti", "Cu", "Zr", "Zn", "Ni")


def detect_analysis_type_from_pdf_content(pdf_path: Path) -> str | None:
    """Inspect only page one for a distinctive heavy-metal results layout."""
    try:
        with pdfplumber.open(Path(pdf_path)) as pdf:
            if not pdf.pages:
                return None
            text = pdf.pages[0].extract_text() or ""
    except (OSError, ValueError, PDFSyntaxError):
        return None

    compact = " ".join(text.split())
    if not re.search(r"\bList\s+of\s+Results\b", compact, re.IGNORECASE):
        return None
    header_sequence = r"\s+".join(re.escape(element) for element in HEAVY_METAL_ELEMENTS)
    if (
        re.search(rf"\b{header_sequence}\b", compact, re.IGNORECASE)
        and len(re.findall(r"\bQuant\b", compact, re.IGNORECASE)) >= 10
        and len(re.findall(r"\bAverage\b", compact, re.IGNORECASE)) >= 10
    ):
        return "중금속"
    return None
