from __future__ import annotations

import os
from pathlib import Path
import unittest

from honyu_app.config.analysis_types import infer_analysis_type
from honyu_app.infrastructure.pdf.labsolutions_parser import LabSolutionsParser


G3_ANALYSIS_TYPE = "(혼유-G3) 1,2-디클로로에틸렌,퍼클로로에틸렌,프로판,에탄"
ACTUAL_DIR = Path(
    os.environ.get(
        "HONYU_G3_152_153_TEST_DIR",
        r"\\172.30.1.100\data\분석결과(사업장별)★\양세경\09.02\혼유(G3-1,2디클로로에탄) 152-153",
    )
)
PDF = ACTUAL_DIR / "혼유(G3-1,2디클로로에탄) 152-153@완료.pdf"


@unittest.skipUnless(PDF.is_file(), "G3 152-153 실제 PDF가 없습니다.")
class G3FilenamePriorityActualFileTests(unittest.TestCase):
    def test_actual_filename_stays_g3_before_and_after_pdf_parsing(self) -> None:
        self.assertEqual(infer_analysis_type(PDF.name), G3_ANALYSIS_TYPE)
        batch = LabSolutionsParser().parse(
            PDF,
            analysis_type=G3_ANALYSIS_TYPE,
            analysis_no_start=152,
            analysis_no_end=153,
        )
        method_filenames = tuple(
            sample.method_filename for sample in batch.samples if sample.method_filename
        )
        materials = tuple(
            peak.material_standard or peak.material_raw or ""
            for sample in batch.samples
            for peak in sample.peaks
        )

        self.assertGreater(len(batch.samples), 0)
        self.assertEqual(batch.analysis_type, G3_ANALYSIS_TYPE)
        self.assertEqual(
            infer_analysis_type(PDF.name, method_filenames, materials),
            G3_ANALYSIS_TYPE,
        )
