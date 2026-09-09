from copy import deepcopy
from pathlib import Path
from tempfile import TemporaryDirectory
import os
import re
import unittest

from honyu_app.application.preview_excel_export import PreviewExcelExportService
from honyu_app.application.review_extraction import ReviewExtractionService
from honyu_app.application.sample_number_matching import classify_sample_number
from honyu_app.domain.enums import ExcelPreviewStatus, ExcludeReason, SampleType, StdMethod
from honyu_app.domain.models import ExcelCellWrite
from honyu_app.infrastructure.database.mock_database_service import MockDatabaseService
from honyu_app.infrastructure.excel.workbook_inspector import XlsxTemplateInspector
from honyu_app.infrastructure.excel.workbook_validator import XlsxWorkbookValidator
from honyu_app.infrastructure.excel.xml_cell_writer import XlsxXmlCellWriter
from honyu_app.infrastructure.pdf.labsolutions_parser import LabSolutionsParser


def _actual_file(name: str) -> Path:
    configured = os.environ.get("ALCOHOL_168_213_TEST_DIR")
    directories = (
        Path(configured) if configured else None,
        Path(
            r"\\172.30.1.100\data\분석결과(사업장별)★"
            r"\자동화프로그램@@\09.08\알콜 168-213"
        ),
        Path(__file__).resolve().parents[3] / "TEST" / "알콜 168-213",
    )
    return next(
        (
            directory / name
            for directory in directories
            if directory is not None and (directory / name).is_file()
        ),
        Path(),
    )


ACTUAL_PDF = _actual_file("알콜 168-213 병합.pdf")
ACTUAL_XLSX = _actual_file("(알콜2) IBA,1-BTOH 168-213.xlsx")

EXPECTED = {
    "168": {"F21": 4052, "I21": 6684},
    "170": {"F23": None, "I23": 1282},
    "172": {"F25": 5938, "I25": 1950},
    "173": {"F26": 4174, "I26": 3373},
    "175": {"F28": 27011, "I28": 42701},
    "177": {"F30": 5430, "I30": 1080},
    "179": {"F32": 3906, "I32": 2752},
    "183": {"F36": 3168, "I36": 3310},
    "188": {"F41": 3689, "I41": 82545},
    "197": {"F50": None, "I50": 1551},
    "198": {"F51": None, "I51": 1087},
    "200": {"F53": 2593, "I53": 2095},
    "201": {"F54": 25335, "I54": 25180},
    "203": {"F56": 3967, "I56": 1522},
    "207": {"F60": 9266, "I60": 6533},
    "209": {"F62": 19500, "I62": 28875},
    "210": {"F63": 13225, "I63": 7406},
    "211": {"F64": 3491, "I64": 2557},
}


@unittest.skipUnless(
    ACTUAL_PDF.is_file() and ACTUAL_XLSX.is_file(),
    "알콜 168-213 실제 PDF/Excel이 없습니다.",
)
class Alcohol168213ActualRegressionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.parsed = LabSolutionsParser().parse(
            ACTUAL_PDF,
            analysis_type="(알콜2) IBA,1-BTOH",
            analysis_no_start=168,
            analysis_no_end=213,
        )

    def test_pdf_db_preview_and_xlsx_match_all_18_confirmed_samples(self) -> None:
        with TemporaryDirectory() as temporary:
            database = MockDatabaseService(Path(temporary) / "mock.db")
            batch = deepcopy(self.parsed)
            review = ReviewExtractionService(database)
            review.complete_review(batch)
            saved = review.save_batch(batch)
            loaded = database.get_batch_detail(saved.batch_id)

            parsed_by_name = {
                sample.sample_name_raw: sample
                for sample in self.parsed.samples
                if sample.sample_type is SampleType.NUMERIC
            }
            loaded_by_name = {
                sample.sample_name_raw: sample
                for sample in loaded.samples
                if sample.sample_type is SampleType.NUMERIC
            }
            for name, parsed_sample in parsed_by_name.items():
                self.assertIn(name, loaded_by_name)
                self.assertEqual(
                    [
                        (
                            peak.peak_no,
                            peak.retention_time,
                            peak.area_raw,
                            peak.material_standard,
                        )
                        for peak in loaded_by_name[name].peaks
                    ],
                    [
                        (
                            peak.peak_no,
                            peak.retention_time,
                            peak.area_raw,
                            peak.material_standard,
                        )
                        for peak in parsed_sample.peaks
                    ],
                    name,
                )

            previews = {
                method: PreviewExcelExportService(
                    database, XlsxTemplateInspector()
                ).preview(saved.batch_id, ACTUAL_XLSX, method)
                for method in (StdMethod.A, StdMethod.B)
            }
            for method, preview in previews.items():
                with self.subTest(method=method):
                    self.assertTrue(preview.can_generate, preview.issues)
                    self.assertEqual(preview.error_count, 0)
                    mapped = {
                        (row.target_cell, row.applied_area)
                        for row in preview.rows
                        if row.sample_type is SampleType.NUMERIC
                        and row.status is ExcelPreviewStatus.MAPPED
                    }
                    for cells in EXPECTED.values():
                        for cell, area in cells.items():
                            if area is not None:
                                self.assertIn((cell, area), mapped)

                    wrong_non_analysis = [
                        row
                        for row in preview.rows
                        if row.sample_type is SampleType.NUMERIC
                        and row.exclude_reason
                        == ExcludeReason.NON_ANALYSIS_SAMPLE.value
                        and re.fullmatch(r"\d+-", row.sample_name)
                    ]
                    self.assertEqual(wrong_non_analysis, [])

            preview = previews[StdMethod.A]
            writes = [
                ExcelCellWrite(row.target_sheet, row.target_cell, row.applied_area)
                for row in preview.rows
                if row.status is ExcelPreviewStatus.MAPPED
            ]
            output = Path(temporary) / "alcohol-168-213-result.xlsx"
            XlsxXmlCellWriter().write_copy(ACTUAL_XLSX, output, writes)
            validation = XlsxWorkbookValidator().validate(
                ACTUAL_XLSX,
                output,
                writes,
                after_excel_recalculation=False,
            )
            self.assertTrue(validation.valid, validation.errors)

            before = XlsxTemplateInspector().inspect(ACTUAL_XLSX)
            after = XlsxTemplateInspector().inspect(output)
            for number, cells in EXPECTED.items():
                for cell, area in cells.items():
                    if area is None:
                        self.assertEqual(
                            after.cell("area입력", cell).value,
                            before.cell("area입력", cell).value,
                            (number, cell),
                        )
                    else:
                        self.assertEqual(
                            after.cell("area입력", cell).value,
                            area,
                            (number, cell),
                        )

    def test_old_regex_excluded_29_trailing_hyphen_workers_new_rule_excludes_zero(self) -> None:
        old_pattern = re.compile(r"^(?P<number>\d+)(?:-(?P<suffix>.+))?$")
        affected = [
            sample
            for sample in self.parsed.samples
            if sample.sample_type is SampleType.NUMERIC
            and re.fullmatch(r"\d+-", sample.sample_name_normalized)
            and not old_pattern.fullmatch(sample.sample_name_normalized)
        ]
        self.assertEqual(len(affected), 29)
        still_excluded = [
            sample
            for sample in affected
            if classify_sample_number(
                sample.sample_name_normalized,
                sample.sample_type,
                is_blank=sample.is_blank,
            ).analysis_number
            is None
        ]
        self.assertEqual(still_excluded, [])
