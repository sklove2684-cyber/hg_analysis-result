from __future__ import annotations

from pathlib import Path
import re

import pdfplumber
from pdfminer.pdfparser import PDFSyntaxError

from honyu_app.config.analysis_types import infer_alcohol_analysis_type
from honyu_app.infrastructure.pdf.heavy_metal_layout import (
    find_heavy_metal_table_layout,
)
from honyu_app.infrastructure.pdf.labsolutions_parser import LabSolutionsParser

MAX_CONTENT_SCAN_PAGES = 3


def detect_analysis_type_from_pdf_content(pdf_path: Path) -> str | None:
    """Inspect a small PDF prefix for supported analysis-specific layouts."""
    try:
        with pdfplumber.open(Path(pdf_path)) as pdf:
            if not pdf.pages:
                return None
            pages = pdf.pages[:MAX_CONTENT_SCAN_PAGES]
            text = "\n".join(page.extract_text() or "" for page in pages)
            layouts_and_tables = []
            gc_materials: list[str] = []
            gc_parser = LabSolutionsParser()
            for page_no, page in enumerate(pages, 1):
                tables = page.extract_tables()
                gc_materials.extend(
                    gc_parser.recognized_materials_from_tables(tables, page_no)
                )
                for table in tables:
                    try:
                        layout = find_heavy_metal_table_layout(table, min_elements=3)
                    except ValueError:
                        continue
                    if layout is not None:
                        layouts_and_tables.append((layout, table))
    except (OSError, ValueError, PDFSyntaxError):
        return None

    compact = " ".join(text.split())
    if re.search(r"\bList\s+of\s+Results\b", compact, re.IGNORECASE):
        for layout, table in layouts_and_tables:
            sample_names = {
                (row[layout.sample_name_column] or "").strip()
                for row in table[layout.header_row + 1:]
                if len(row) > layout.sample_name_column
            }
            if (
                "회수율-B" in sample_names
                and any(re.fullmatch(r"저[123]", name) for name in sample_names)
                and any(re.fullmatch(r"중[123]", name) for name in sample_names)
                and any(re.fullmatch(r"고[123]", name) for name in sample_names)
            ):
                return "중금속"
    return infer_alcohol_analysis_type(tuple(gc_materials))
