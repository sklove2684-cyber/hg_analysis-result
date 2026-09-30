from __future__ import annotations

from pathlib import Path
import re

import pdfplumber
from pdfminer.pdfparser import PDFSyntaxError

from honyu_app.config.analysis_types import (
    infer_alcohol_analysis_type,
    infer_analysis_type,
)
from honyu_app.infrastructure.pdf.heavy_metal_layout import (
    find_heavy_metal_table_layout,
)
from honyu_app.infrastructure.pdf.labsolutions_parser import LabSolutionsParser

MAX_CONTENT_SCAN_PAGES = 3

ISOAMYL_N_PROPYL_ACETATE_MATERIALS = frozenset({
    "n-프로필 아세테이트",
    "이소아밀 아세테이트",
})
ISOPROPYL_ACETATE_MATERIAL = "Isopropyl acetate"


def detect_analysis_type_from_pdf_content(pdf_path: Path) -> str | None:
    """Inspect a small PDF prefix for supported analysis-specific layouts."""
    try:
        with pdfplumber.open(Path(pdf_path)) as pdf:
            if not pdf.pages:
                return None
            pages = pdf.pages[:MAX_CONTENT_SCAN_PAGES]
            page_texts = [page.extract_text() or "" for page in pages]
            text = "\n".join(page_texts)
            layouts_and_tables = []
            gc_materials: list[str] = []
            std_gc_materials: list[str] = []
            method_filenames: list[str] = []
            gc_parser = LabSolutionsParser()
            for page_no, (page, page_text) in enumerate(
                zip(pages, page_texts), 1
            ):
                method_filenames.extend(re.findall(
                    r"^\s*Method\s+Filename\s*:\s*(.+?)\s*$",
                    page_text,
                    re.IGNORECASE | re.MULTILINE,
                ))
                tables = page.extract_tables()
                page_materials = gc_parser.recognized_materials_from_tables(
                    tables, page_no
                )
                gc_materials.extend(page_materials)
                sample_match = re.search(
                    r"^\s*Sample\s+Name\s*:\s*(.+?)\s*$",
                    page_text,
                    re.IGNORECASE | re.MULTILINE,
                )
                if sample_match and re.fullmatch(
                    r"STD[1-6]", sample_match.group(1).strip(), re.IGNORECASE
                ):
                    std_gc_materials.extend(page_materials)
                for table in tables:
                    try:
                        layout = find_heavy_metal_table_layout(table, min_elements=2)
                    except ValueError:
                        continue
                    if layout is not None and (
                        len(layout.element_columns) >= 3
                        or {element for element, _ in layout.element_columns}
                        == {"Na", "K"}
                    ):
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
    if infer_analysis_type(
        "unknown.pdf", method_filenames=tuple(method_filenames)
    ) == "DMF,DMA":
        return "DMF,DMA"
    if "DMF" in std_gc_materials:
        return "DMF,DMA"
    std_materials = frozenset(std_gc_materials)
    if ISOAMYL_N_PROPYL_ACETATE_MATERIALS.issubset(std_materials):
        return "이소아밀,n-프로필 아세테이트"
    if ISOPROPYL_ACETATE_MATERIAL in std_materials:
        return "이소프로필 아세테이트"
    return infer_alcohol_analysis_type(tuple(gc_materials))
