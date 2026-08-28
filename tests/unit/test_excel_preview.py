from datetime import datetime, timezone
from decimal import Decimal
from pathlib import Path
import unittest
from uuid import uuid4

from honyu_app.application.preview_excel_export import (
    ACN_PROFILE,
    ACETIC_ACID_PROFILE,
    BC_PROFILE,
    CELLOSOLVE_PROFILE,
    DIETHYL_ETHER_PROFILE,
    ETHYLENE_GLYCOL_PROFILE,
    G2_PROFILE,
    G3_PROFILE,
    IPA_PROFILE,
    ISOAMYL_N_PROPYL_ACETATE_PROFILE,
    MEK_PROFILE,
    METHANOL_PROFILE,
    ONE_COLUMN_PROFILE,
    PHENOL_PROFILE,
    PreviewExcelExportService,
)
from honyu_app.domain.enums import (
    ConcentrationLevel,
    ExcelPreviewStatus,
    ExcludeReason,
    ReviewStatus,
    SampleType,
    StdMethod,
)
from honyu_app.domain.models import (
    AnalysisBatch,
    Peak,
    PeakCorrection,
    Sample,
    SourceFile,
)
from honyu_app.infrastructure.excel.workbook_inspector import XlsxTemplateInspector
from honyu_app.services.excel_template_service import (
    ExcelTemplateSnapshot,
    TemplateCell,
)


SHEETS = ("검량선", "area", "최종결과", "회수율", "STD제조")
ONE_COLUMN_SHEETS = ("검량선", "area입력", "회수율", "STD제조", "Sheet1")
MEK_SHEETS = ("검량선", "LOD(area입력)", "회수율", "std")
ACN_SHEETS = ("검량선", "결과입력(area입력)", "회수율", "STD제조")
G2_SHEETS = ("검량선", "area입력", "회수율", "STD제조", "Sheet1")


class FakeDatabase:
    def __init__(self) -> None:
        self.corrections: dict[object, list[PeakCorrection]] = {}

    def list_peak_corrections(self, peak_id):
        return self.corrections.get(peak_id, [])

    def get_batch_detail(self, batch_id):
        raise AssertionError("preview_batch should be used in unit tests")


class FakeTemplateService:
    def __init__(self, snapshot: ExcelTemplateSnapshot) -> None:
        self.snapshot = snapshot

    def inspect(self, path: Path) -> ExcelTemplateSnapshot:
        return self.snapshot


def snapshot(*cells: TemplateCell) -> ExcelTemplateSnapshot:
    return ExcelTemplateSnapshot(
        Path("template.xlsx"),
        SHEETS,
        {(cell.sheet, cell.address): cell for cell in cells},
    )


def one_column_snapshot(*cells: TemplateCell) -> ExcelTemplateSnapshot:
    return ExcelTemplateSnapshot(
        Path("one-column-template.xlsx"),
        ONE_COLUMN_SHEETS,
        {(cell.sheet, cell.address): cell for cell in cells},
    )


def mek_snapshot(*cells: TemplateCell) -> ExcelTemplateSnapshot:
    defaults = tuple(
        TemplateCell("LOD(area입력)", f"A{row}", True, f"262-{number}", "string")
        for row, number in enumerate(range(74, 120), start=21)
    )
    return ExcelTemplateSnapshot(
        Path("mek-template.xlsx"),
        MEK_SHEETS,
        {(cell.sheet, cell.address): cell for cell in (*defaults, *cells)},
    )


def acetic_snapshot(*cells: TemplateCell) -> ExcelTemplateSnapshot:
    headers = (
        TemplateCell("LOD(area입력)", "E2", True, "Acetic acid", "string"),
        TemplateCell("LOD(area입력)", "A19", True, "261-637", "string"),
    )
    return ExcelTemplateSnapshot(
        Path("acetic-template.xlsx"),
        MEK_SHEETS,
        {(cell.sheet, cell.address): cell for cell in (*headers, *cells)},
    )


def ethylene_glycol_snapshot(*cells: TemplateCell) -> ExcelTemplateSnapshot:
    headers = (
        TemplateCell("LOD(area입력)", "E2", True, "Ethylene glycol", "string"),
        TemplateCell("LOD(area입력)", "A19", True, "261-599", "string"),
    )
    return ExcelTemplateSnapshot(
        Path("ethylene-glycol-template.xlsx"),
        MEK_SHEETS,
        {(cell.sheet, cell.address): cell for cell in (*headers, *cells)},
    )


def diethyl_ether_snapshot(*cells: TemplateCell) -> ExcelTemplateSnapshot:
    headers = (
        TemplateCell("LOD(area입력)", "E2", True, "Diethyl ether", "string"),
        TemplateCell("LOD(area입력)", "A19", True, "262-152", "string"),
        TemplateCell("LOD(area입력)", "A20", True, "262-153", "string"),
    )
    return ExcelTemplateSnapshot(
        Path("diethyl-ether-template.xlsx"),
        MEK_SHEETS,
        {(cell.sheet, cell.address): cell for cell in (*headers, *cells)},
    )


def ipa_snapshot(*cells: TemplateCell) -> ExcelTemplateSnapshot:
    headers = (
        TemplateCell("LOD(area입력)", "I3", True, "IPA", "string"),
        TemplateCell("LOD(area입력)", "J3", True, "area", "string"),
        TemplateCell("LOD(area입력)", "E20", True, "261-320", "string"),
    )
    return ExcelTemplateSnapshot(
        Path("ipa-template.xlsx"),
        MEK_SHEETS,
        {(cell.sheet, cell.address): cell for cell in (*headers, *cells)},
    )


def methanol_snapshot(*cells: TemplateCell) -> ExcelTemplateSnapshot:
    headers = (
        TemplateCell("LOD(area입력)", "E2", True, "Methanol", "string"),
        TemplateCell("LOD(area입력)", "A19", True, "261-237", "string"),
        TemplateCell("LOD(area입력)", "A20", True, "261-238", "string"),
    )
    return ExcelTemplateSnapshot(
        Path("methanol-template.xlsx"),
        MEK_SHEETS,
        {(cell.sheet, cell.address): cell for cell in (*headers, *cells)},
    )


def phenol_snapshot(*cells: TemplateCell) -> ExcelTemplateSnapshot:
    headers = (
        TemplateCell("LOD(area입력)", "E2", True, "Phenol", "string"),
        TemplateCell("LOD(area입력)", "A19", True, "261-256", "string"),
    )
    return ExcelTemplateSnapshot(
        Path("phenol-template.xlsx"),
        MEK_SHEETS,
        {(cell.sheet, cell.address): cell for cell in (*headers, *cells)},
    )


def acn_snapshot(*cells: TemplateCell) -> ExcelTemplateSnapshot:
    headers = (
        TemplateCell("결과입력(area입력)", "F2", True, "Acetonitrile", "string"),
        TemplateCell("결과입력(area입력)", "A26", True, "261-656", "string"),
    )
    return ExcelTemplateSnapshot(
        Path("acn-template.xlsx"),
        ACN_SHEETS,
        {(cell.sheet, cell.address): cell for cell in (*headers, *cells)},
    )


def bc_snapshot(*cells: TemplateCell) -> ExcelTemplateSnapshot:
    headers = (
        TemplateCell("결과입력(area입력)", "F2", True, "2-부톡시에탄올", "string"),
        TemplateCell("결과입력(area입력)", "A26", True, "261-588", "string"),
    )
    return ExcelTemplateSnapshot(
        Path("bc-template.xlsx"),
        ACN_SHEETS,
        {(cell.sheet, cell.address): cell for cell in (*headers, *cells)},
    )


def g2_snapshot(*cells: TemplateCell) -> ExcelTemplateSnapshot:
    headers = (
        TemplateCell("area입력", "F3", True, "THF", "string"),
        TemplateCell("area입력", "I3", True, "CFM", "string"),
        TemplateCell("area입력", "L3", True, "Benzene", "string"),
        TemplateCell("area입력", "O3", True, "Chlorobenzene", "string"),
        TemplateCell("area입력", "R3", True, "Carbon tetrachloride", "string"),
        TemplateCell("area입력", "A21", True, "261-655", "string"),
    )
    return ExcelTemplateSnapshot(
        Path("g2-template.xlsx"),
        G2_SHEETS,
        {(cell.sheet, cell.address): cell for cell in (*headers, *cells)},
    )


def acetate_snapshot(*cells: TemplateCell) -> ExcelTemplateSnapshot:
    headers = (
        TemplateCell("LOD(area입력)", "F3", True, "Isoamyl acetate", "string"),
        TemplateCell("LOD(area입력)", "I3", True, "n-Propyl acetae", "string"),
        TemplateCell("LOD(area입력)", "A20", True, "261-611", "string"),
    )
    return ExcelTemplateSnapshot(
        Path("acetate-template.xlsx"),
        MEK_SHEETS,
        {(cell.sheet, cell.address): cell for cell in (*headers, *cells)},
    )


def cellosolve_snapshot(*cells: TemplateCell) -> ExcelTemplateSnapshot:
    headers = (
        TemplateCell("area입력", "F3", True, "2-Butoxyethanol \n(EGBE)", "string"),
        TemplateCell("area입력", "I3", True, "2-Butoxyethyl acetate(EGBEA)", "string"),
        TemplateCell("area입력", "L3", True, "2-Ethoxy \nethanol(EGEE)", "string"),
        TemplateCell("area입력", "O3", True, "2-Ethoxy ethyl acetate(EGEEA)", "string"),
        TemplateCell("area입력", "A21", True, "681", "string"),
    )
    return ExcelTemplateSnapshot(
        Path("cellosolve-template.xlsx"),
        G2_SHEETS,
        {(cell.sheet, cell.address): cell for cell in (*headers, *cells)},
    )


def g3_snapshot(*cells: TemplateCell) -> ExcelTemplateSnapshot:
    headers = (
        TemplateCell("area입력", "F3", True, "1,2-\nDichloroethylene", "string"),
        TemplateCell("area입력", "I3", True, "Trichloroethylene", "string"),
        TemplateCell("area입력", "L3", True, "Tetrachloroethylene", "string"),
        TemplateCell("area입력", "O3", True, "1,2-\nDichloropropane", "string"),
        TemplateCell("area입력", "R3", True, "1,2-Dichloroethane", "string"),
        TemplateCell("area입력", "A21", True, "695", "string"),
    )
    return ExcelTemplateSnapshot(
        Path("g3-template.xlsx"),
        G2_SHEETS,
        {(cell.sheet, cell.address): cell for cell in (*headers, *cells)},
    )


def peak(number: int, area: int, material: str = "n-hexane") -> Peak:
    return Peak(
        peak_no=number,
        retention_time=Decimal(f"{number}.100"),
        area_raw=area,
        material_raw=material,
        material_standard=material,
        source_page=1,
    )


def batch(samples: list[Sample], analysis_type: str = "혼유") -> AnalysisBatch:
    return AnalysisBatch(
        batch_code="BATCH-PREVIEW",
        source_file=SourceFile("sample.pdf", Path("sample.pdf"), "a" * 64, 100, 1),
        analysis_type=analysis_type,
        analysis_no_start=1,
        analysis_no_end=10,
        parser_name="test",
        parser_version="1",
        parser_layout_id="layout",
        extracted_at=datetime.now(timezone.utc),
        samples=samples,
        review_status=ReviewStatus.SAVED,
    )


class ExcelPreviewServiceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.database = FakeDatabase()

    def service(self, template: ExcelTemplateSnapshot | None = None):
        return PreviewExcelExportService(
            self.database, FakeTemplateService(template or snapshot())
        )

    def test_std_method_a_and_b_use_the_same_excel_std5_row(self) -> None:
        std5 = Sample(5, "STD5", "STD5", SampleType.STD, replicate_no=5, peaks=[peak(1, 500)])
        std6 = Sample(6, "STD6", "STD6", SampleType.STD, replicate_no=6, peaks=[peak(1, 600)])
        source = batch([std5, std6])

        method_a = self.service().preview_batch(source, Path("template.xlsx"), StdMethod.A)
        method_b = self.service().preview_batch(source, Path("template.xlsx"), StdMethod.B)

        mapped_a = [row for row in method_a.rows if row.status is ExcelPreviewStatus.MAPPED]
        mapped_b = [row for row in method_b.rows if row.status is ExcelPreviewStatus.MAPPED]
        self.assertEqual([(row.sample_name, row.target_cell) for row in mapped_a], [("STD5", "F19")])
        self.assertEqual([(row.sample_name, row.target_cell) for row in mapped_b], [("STD6", "F19")])

    def test_first_complete_std_set_excludes_trailing_partial_recheck(self) -> None:
        standards = [
            Sample(page, f"STD{repeat}", f"STD{repeat}", SampleType.STD,
                   replicate_no=repeat, peaks=[peak(1, page * 100)])
            for page, repeat in enumerate((1, 2, 3, 4, 5, 1, 2), start=1)
        ]

        result = self.service().preview_batch(batch(standards), Path("template.xlsx"), "A")
        mapped = [row for row in result.rows if row.status is ExcelPreviewStatus.MAPPED]
        excluded = [row for row in result.rows if row.status is ExcelPreviewStatus.EXCLUDED]

        self.assertTrue(result.can_generate, result.issues)
        self.assertEqual([row.sample_name for row in mapped], [f"STD{n}" for n in range(1, 6)])
        self.assertEqual([row.sample_name for row in excluded], ["STD1", "STD2"])
        self.assertTrue(all(
            row.exclude_reason == ExcludeReason.DUPLICATE_STD_SET.value
            for row in excluded
        ))
        self.assertFalse(any(issue.code == "TARGET_COLLISION" for issue in result.issues))

    def test_two_complete_std_sets_require_review(self) -> None:
        standards = [
            Sample(page, f"STD{repeat}", f"STD{repeat}", SampleType.STD,
                   replicate_no=repeat, peaks=[peak(1, page * 100)])
            for page, repeat in enumerate((1, 2, 3, 4, 5, 1, 2, 3, 4, 5), start=1)
        ]

        result = self.service().preview_batch(batch(standards), Path("template.xlsx"), "A")

        self.assertFalse(result.can_generate)
        self.assertEqual(result.issues[0].code, "STD_SET_AMBIGUOUS")

    def test_partial_std_set_is_not_filled_in_implicitly(self) -> None:
        standards = [
            Sample(page, f"STD{repeat}", f"STD{repeat}", SampleType.STD,
                   replicate_no=repeat, peaks=[peak(1, page * 100)])
            for page, repeat in enumerate((1, 2, 3), start=1)
        ]

        result = self.service().preview_batch(batch(standards), Path("template.xlsx"), "A")

        self.assertFalse(result.can_generate)
        self.assertEqual(result.issues[0].code, "STD_SET_INCOMPLETE")

    def test_unregistered_analysis_type_has_clear_excel_profile_error(self) -> None:
        for analysis_type in ("PCE(테트라클로로에틸렌)",):
            with self.subTest(analysis_type=analysis_type):
                source = batch([])
                source.analysis_type = analysis_type
                result = self.service().preview_batch(
                    source, Path("template.xlsx"), "A"
                )

                self.assertFalse(result.can_generate)
                self.assertEqual(result.issues[0].code, "EXCEL_PROFILE_NOT_REGISTERED")
                self.assertIn(
                    "아직 등록되지 않은 Excel 양식/프로필",
                    result.issues[0].message,
                )
                self.assertIn(analysis_type, result.issues[0].message)

    def test_ipa_profile_uses_common_std_methods_and_runtime_std_rt(self) -> None:
        def ipa_peak(number: int, rt: str, area: int) -> Peak:
            return Peak(
                number,
                Decimal(rt),
                area,
                material_raw="IPA",
                material_standard="Isopropyl alcohol",
            )

        standards = [
            Sample(
                1,
                "STD1",
                "STD1",
                SampleType.STD,
                replicate_no=1,
                peaks=[ipa_peak(1, "3.720", 100), ipa_peak(2, "3.820", 9999)],
            ),
            *[
                Sample(
                    repeat,
                    f"STD{repeat}",
                    f"STD{repeat}",
                    SampleType.STD,
                    replicate_no=repeat,
                    peaks=[ipa_peak(1, "3.721", repeat * 100)],
                )
                for repeat in range(2, 7)
            ],
        ]
        recovery = Sample(
            10,
            "저1",
            "저1",
            SampleType.RECOVERY,
            concentration_level=ConcentrationLevel.LOW,
            replicate_no=1,
            peaks=[ipa_peak(1, "3.722", 500), ipa_peak(2, "3.820", 5000)],
        )
        worker = Sample(
            20,
            "320",
            "320",
            SampleType.NUMERIC,
            worker_match_key="320",
            peaks=[ipa_peak(1, "3.719", 600), ipa_peak(2, "3.820", 6000)],
        )
        source = batch([*standards, recovery, worker], "IPA")
        source.analysis_no_start = 320
        source.analysis_no_end = 334

        for method, expected_j9, mapped_std, excluded_std in (
            ("A", 500, "STD5", "STD6"),
            ("B", 600, "STD6", "STD5"),
        ):
            with self.subTest(method=method):
                result = self.service(ipa_snapshot()).preview_batch(
                    source, Path("ipa-template.xlsx"), method
                )
                mapped = {
                    (row.sample_name, row.target_sheet, row.target_cell): row.applied_area
                    for row in result.rows
                    if row.status is ExcelPreviewStatus.MAPPED
                }

                self.assertTrue(result.can_generate, result.issues)
                self.assertEqual(mapped[("STD1", "LOD(area입력)", "J5")], 100)
                self.assertEqual(mapped[("STD4", "LOD(area입력)", "J8")], 400)
                self.assertEqual(
                    mapped[(mapped_std, "LOD(area입력)", "J9")], expected_j9
                )
                self.assertEqual(mapped[("저1", "회수율", "B28")], 500)
                self.assertEqual(mapped[("320", "LOD(area입력)", "J20")], 600)
                excluded = [
                    row
                    for row in result.rows
                    if row.status is ExcelPreviewStatus.EXCLUDED
                ]
                self.assertTrue(any(row.sample_name == excluded_std for row in excluded))
                self.assertTrue(all(
                    row.applied_area not in {5000, 6000, 9999}
                    for row in result.rows
                    if row.status is ExcelPreviewStatus.MAPPED
                ))

    def test_methanol_profile_uses_common_std_methods_and_runtime_std_rt(self) -> None:
        def methanol_peak(number: int, rt: str, area: int) -> Peak:
            return Peak(
                number,
                Decimal(rt),
                area,
                material_raw="메탄올",
                material_standard="Methanol",
            )

        standards = [
            Sample(
                repeat,
                f"STD{repeat}",
                f"STD{repeat}",
                SampleType.STD,
                replicate_no=repeat,
                peaks=[methanol_peak(1, "2.032", repeat * 100)],
            )
            for repeat in range(1, 7)
        ]
        recovery = Sample(
            10,
            "저1",
            "저1",
            SampleType.RECOVERY,
            concentration_level=ConcentrationLevel.LOW,
            replicate_no=1,
            peaks=[
                methanol_peak(1, "2.031", 500),
                methanol_peak(2, "2.200", 5000),
            ],
        )
        worker = Sample(
            20,
            "237",
            "237",
            SampleType.NUMERIC,
            worker_match_key="237",
            peaks=[
                methanol_peak(1, "2.033", 600),
                methanol_peak(2, "2.180", 6000),
            ],
        )
        source = batch([*standards, recovery, worker], "메탄올A")
        source.analysis_no_start = 237
        source.analysis_no_end = 320

        for method, expected_f8, mapped_std, excluded_std in (
            ("A", 500, "STD5", "STD6"),
            ("B", 600, "STD6", "STD5"),
        ):
            with self.subTest(method=method):
                result = self.service(methanol_snapshot()).preview_batch(
                    source, Path("methanol-template.xlsx"), method
                )
                mapped = {
                    (row.sample_name, row.target_sheet, row.target_cell): row.applied_area
                    for row in result.rows
                    if row.status is ExcelPreviewStatus.MAPPED
                }

                self.assertTrue(result.can_generate, result.issues)
                self.assertEqual(mapped[("STD1", "LOD(area입력)", "F4")], 100)
                self.assertEqual(mapped[("STD4", "LOD(area입력)", "F7")], 400)
                self.assertEqual(
                    mapped[(mapped_std, "LOD(area입력)", "F8")], expected_f8
                )
                self.assertEqual(mapped[("저1", "회수율", "B28")], 500)
                self.assertEqual(mapped[("237", "LOD(area입력)", "F19")], 600)
                excluded = [
                    row
                    for row in result.rows
                    if row.status is ExcelPreviewStatus.EXCLUDED
                ]
                self.assertTrue(any(row.sample_name == excluded_std for row in excluded))
                self.assertTrue(all(
                    row.applied_area not in {5000, 6000}
                    for row in result.rows
                    if row.status is ExcelPreviewStatus.MAPPED
                ))

    def test_phenol_profile_uses_common_std_methods_and_runtime_std_rt(self) -> None:
        def phenol_peak(number: int, rt: str, area: int) -> Peak:
            return Peak(
                number,
                Decimal(rt),
                area,
                material_raw="페놀",
                material_standard="Phenol",
            )

        standards = [
            Sample(
                repeat,
                f"STD{repeat}",
                f"STD{repeat}",
                SampleType.STD,
                replicate_no=repeat,
                peaks=[phenol_peak(1, "2.754", repeat * 100)],
            )
            for repeat in range(1, 7)
        ]
        recovery = Sample(
            10,
            "저1",
            "저1",
            SampleType.RECOVERY,
            concentration_level=ConcentrationLevel.LOW,
            replicate_no=1,
            peaks=[phenol_peak(1, "2.755", 500)],
        )
        worker = Sample(
            20,
            "256",
            "256",
            SampleType.NUMERIC,
            worker_match_key="256",
            peaks=[
                phenol_peak(1, "2.756", 600),
                phenol_peak(2, "2.900", 6000),
            ],
        )
        source = batch([*standards, recovery, worker], "페놀")
        source.analysis_no_start = 256
        source.analysis_no_end = 305

        for method, expected_f8, mapped_std, excluded_std in (
            ("A", 500, "STD5", "STD6"),
            ("B", 600, "STD6", "STD5"),
        ):
            with self.subTest(method=method):
                result = self.service(phenol_snapshot()).preview_batch(
                    source, Path("phenol-template.xlsx"), method
                )
                mapped = {
                    (row.sample_name, row.target_sheet, row.target_cell): row.applied_area
                    for row in result.rows
                    if row.status is ExcelPreviewStatus.MAPPED
                }

                self.assertTrue(result.can_generate, result.issues)
                self.assertEqual(mapped[("STD1", "LOD(area입력)", "F4")], 100)
                self.assertEqual(mapped[("STD4", "LOD(area입력)", "F7")], 400)
                self.assertEqual(
                    mapped[(mapped_std, "LOD(area입력)", "F8")], expected_f8
                )
                self.assertEqual(mapped[("저1", "회수율", "B28")], 500)
                self.assertEqual(mapped[("256", "LOD(area입력)", "F19")], 600)
                excluded = [
                    row
                    for row in result.rows
                    if row.status is ExcelPreviewStatus.EXCLUDED
                ]
                self.assertTrue(any(row.sample_name == excluded_std for row in excluded))
                self.assertTrue(all(
                    row.applied_area != 6000
                    for row in result.rows
                    if row.status is ExcelPreviewStatus.MAPPED
                ))

    def test_phenol_rejects_mixture_workbook_with_clear_mismatch(self) -> None:
        result = self.service(snapshot()).preview_batch(
            batch([], "페놀"), Path("mixture-template.xlsx"), "A"
        )

        self.assertFalse(result.can_generate)
        self.assertEqual(result.issues[0].code, "TEMPLATE_PROFILE_MISMATCH")
        self.assertIn("페놀", result.issues[0].message)
        self.assertIn("혼유", result.issues[0].message)

    def test_methanol_rejects_mixture_workbook_with_clear_mismatch(self) -> None:
        result = self.service(snapshot()).preview_batch(
            batch([], "메탄올A"), Path("mixture-template.xlsx"), "A"
        )

        self.assertFalse(result.can_generate)
        self.assertEqual(result.issues[0].code, "TEMPLATE_PROFILE_MISMATCH")
        self.assertIn("메탄올A", result.issues[0].message)
        self.assertIn("혼유", result.issues[0].message)

    def test_ipa_rejects_mixture_workbook_with_clear_mismatch(self) -> None:
        result = self.service(snapshot()).preview_batch(
            batch([], "IPA"), Path("mixture-template.xlsx"), "A"
        )

        self.assertFalse(result.can_generate)
        self.assertEqual(result.issues[0].code, "TEMPLATE_PROFILE_MISMATCH")
        self.assertIn("IPA", result.issues[0].message)
        self.assertIn("혼유", result.issues[0].message)

    def test_ethylene_glycol_profile_maps_confirmed_cells_and_excludes_std6(self) -> None:
        source = batch([
            Sample(1, "STD1", "STD1", SampleType.STD, replicate_no=1,
                   peaks=[peak(1, 8587, "Ethylene glycol")]),
            Sample(2, "STD6", "STD6", SampleType.STD, replicate_no=6,
                   peaks=[peak(1, 209360, "Ethylene glycol")]),
            Sample(3, "저2", "저2", SampleType.RECOVERY,
                   concentration_level=ConcentrationLevel.LOW, replicate_no=2,
                   peaks=[peak(1, 18385, "Ethylene glycol")]),
            Sample(4, "599", "599", SampleType.NUMERIC, worker_match_key="599",
                   peaks=[peak(1, 12345, "Ethylene glycol")]),
        ])
        source.analysis_type = "에틸렌글리콜"
        source.analysis_no_start = 599
        source.analysis_no_end = 680

        result = self.service(ethylene_glycol_snapshot()).preview_batch(
            source, Path("ethylene-glycol-template.xlsx"), "A"
        )

        self.assertTrue(result.can_generate, result.issues)
        self.assertEqual(
            [(row.sample_name, row.target_sheet, row.target_cell)
             for row in result.rows if row.status is ExcelPreviewStatus.MAPPED],
            [
                ("STD1", "LOD(area입력)", "F4"),
                ("저2", "회수율", "B29"),
                ("599", "LOD(area입력)", "F19"),
            ],
        )
        std6 = next(row for row in result.rows if row.sample_name == "STD6")
        self.assertEqual(std6.status, ExcelPreviewStatus.EXCLUDED)
        self.assertEqual(std6.exclude_reason, "STD_METHOD_A_NOT_SELECTED")

    def test_diethyl_ether_profile_maps_confirmed_std_recovery_and_worker_cells(self) -> None:
        standards = [
            Sample(
                repeat,
                f"STD{repeat}",
                f"STD{repeat}",
                SampleType.STD,
                replicate_no=repeat,
                peaks=[Peak(
                    1,
                    Decimal("1.305"),
                    100 * (2 ** (repeat - 1)),
                    material_raw="디에틸에테르",
                    material_standard="Diethyl ether",
                )],
            )
            for repeat in range(1, 7)
        ]
        samples = [
            *standards,
            Sample(
                7,
                "저1",
                "저1",
                SampleType.RECOVERY,
                concentration_level=ConcentrationLevel.LOW,
                replicate_no=1,
                peaks=[peak(1, 700, "Diethyl ether")],
            ),
            Sample(
                8,
                "중2",
                "중2",
                SampleType.RECOVERY,
                concentration_level=ConcentrationLevel.MID,
                replicate_no=2,
                peaks=[peak(1, 800, "Diethyl ether")],
            ),
            Sample(
                9,
                "고3",
                "고3",
                SampleType.RECOVERY,
                concentration_level=ConcentrationLevel.HIGH,
                replicate_no=3,
                peaks=[peak(1, 900, "Diethyl ether")],
            ),
            Sample(
                10,
                "152-사업장",
                "152-사업장",
                SampleType.NUMERIC,
                worker_match_key="152",
                peaks=[peak(1, 1000, "Diethyl ether")],
            ),
            Sample(
                11,
                "153",
                "153",
                SampleType.NUMERIC,
                worker_match_key="153",
                peaks=[peak(1, 1100, "Diethyl ether")],
            ),
            Sample(
                12,
                "STD2",
                "STD2",
                SampleType.STD,
                replicate_no=2,
                peaks=[peak(1, 1200, "Diethyl ether")],
            ),
        ]

        result = self.service(diethyl_ether_snapshot()).preview_batch(
            batch(samples, "디에틸에테르"), Path("diethyl-ether-template.xlsx"), "A"
        )
        mapped = [row for row in result.rows if row.status is ExcelPreviewStatus.MAPPED]

        self.assertTrue(result.can_generate, result.issues)
        self.assertEqual(
            [(row.sample_name, row.target_sheet, row.target_cell) for row in mapped],
            [
                ("STD1", "LOD(area입력)", "F4"),
                ("STD2", "LOD(area입력)", "F5"),
                ("STD3", "LOD(area입력)", "F6"),
                ("STD4", "LOD(area입력)", "F7"),
                ("STD5", "LOD(area입력)", "F8"),
                ("저1", "회수율", "B28"),
                ("중2", "회수율", "B32"),
                ("고3", "회수율", "B36"),
                ("152-사업장", "LOD(area입력)", "F19"),
                ("153", "LOD(area입력)", "F20"),
            ],
        )
        std6 = next(row for row in result.rows if row.sample_name == "STD6")
        self.assertEqual(std6.status, ExcelPreviewStatus.EXCLUDED)
        self.assertEqual(std6.exclude_reason, "STD_METHOD_A_NOT_SELECTED")
        trailing = next(
            row
            for row in result.rows
            if row.sample_name == "STD2" and row.applied_area == 1200
        )
        self.assertEqual(trailing.exclude_reason, ExcludeReason.DUPLICATE_STD_SET.value)

    def test_diethyl_ether_std_method_b_uses_std1_to_std4_and_std6(self) -> None:
        standards = [
            Sample(
                repeat,
                f"STD{repeat}",
                f"STD{repeat}",
                SampleType.STD,
                replicate_no=repeat,
                peaks=[Peak(
                    1,
                    Decimal("1.305"),
                    100 * (2 ** (repeat - 1)),
                    material_raw="디에틸에테르",
                    material_standard="Diethyl ether",
                )],
            )
            for repeat in range(1, 7)
        ]

        result = self.service(diethyl_ether_snapshot()).preview_batch(
            batch(standards, "디에틸에테르"),
            Path("diethyl-ether-template.xlsx"),
            "B",
        )
        mapped = [row for row in result.rows if row.status is ExcelPreviewStatus.MAPPED]

        self.assertTrue(result.can_generate, result.issues)
        self.assertEqual(
            [(row.sample_name, row.target_cell) for row in mapped],
            [
                ("STD1", "F4"),
                ("STD2", "F5"),
                ("STD3", "F6"),
                ("STD4", "F7"),
                ("STD6", "F8"),
            ],
        )
        std5 = next(row for row in result.rows if row.sample_name == "STD5")
        self.assertEqual(std5.status, ExcelPreviewStatus.EXCLUDED)
        self.assertEqual(std5.exclude_reason, "STD_METHOD_B_NOT_SELECTED")

    def test_diethyl_ether_similar_std_areas_do_not_change_method_a_selection(self) -> None:
        areas = (100, 200, 202, 400, 800, 1200)
        standards = [
            Sample(
                repeat,
                f"STD{repeat}",
                f"STD{repeat}",
                SampleType.STD,
                replicate_no=repeat,
                peaks=[Peak(
                    1,
                    Decimal("1.305"),
                    areas[repeat - 1],
                    material_raw="디에틸에테르",
                    material_standard="Diethyl ether",
                )],
            )
            for repeat in range(1, 7)
        ]

        result = self.service(diethyl_ether_snapshot()).preview_batch(
            batch(standards, "디에틸에테르"),
            Path("diethyl-ether-template.xlsx"),
            "A",
        )
        mapped = [row for row in result.rows if row.status is ExcelPreviewStatus.MAPPED]

        self.assertTrue(result.can_generate, result.issues)
        self.assertEqual(
            [(row.sample_name, row.target_cell) for row in mapped],
            [
                ("STD1", "F4"),
                ("STD2", "F5"),
                ("STD3", "F6"),
                ("STD4", "F7"),
                ("STD5", "F8"),
            ],
        )
        self.assertEqual(result.error_count, 0)

    def test_diethyl_ether_single_peak_maps_using_observed_std_rt(self) -> None:
        standard = Sample(
            1,
            "STD1",
            "STD1",
            SampleType.STD,
            replicate_no=1,
            peaks=[Peak(1, Decimal("1.420"), 100, material_standard="Diethyl ether")],
        )
        worker = Sample(
            2,
            "152",
            "152",
            SampleType.NUMERIC,
            worker_match_key="152",
            peaks=[Peak(1, Decimal("1.500"), 123, material_standard="Diethyl ether")],
        )
        result = self.service(diethyl_ether_snapshot()).preview_batch(
            batch([standard, worker], "디에틸에테르"),
            Path("diethyl-ether-template.xlsx"),
            "A",
        )
        mapped = [
            row for row in result.rows
            if row.status is ExcelPreviewStatus.MAPPED and row.sample_name == "152"
        ]

        self.assertEqual([(row.peak_no, row.target_cell) for row in mapped], [(1, "F19")])

    def test_diethyl_ether_selects_closest_peak_to_observed_std_rt_not_largest_area(self) -> None:
        standard = Sample(
            1,
            "STD1",
            "STD1",
            SampleType.STD,
            replicate_no=1,
            peaks=[Peak(1, Decimal("1.250"), 100, material_standard="Diethyl ether")],
        )
        worker = Sample(
            2,
            "152",
            "152",
            SampleType.NUMERIC,
            worker_match_key="152",
            peaks=[
                Peak(1, Decimal("1.252"), 123, material_standard="Diethyl ether"),
                Peak(2, Decimal("1.305"), 9999, material_standard="Diethyl ether"),
            ],
        )
        result = self.service(diethyl_ether_snapshot()).preview_batch(
            batch([standard, worker], "디에틸에테르"),
            Path("diethyl-ether-template.xlsx"),
            "A",
        )
        mapped = [
            row for row in result.rows
            if row.status is ExcelPreviewStatus.MAPPED and row.sample_name == "152"
        ]
        residual = next(
            row for row in result.rows
            if row.sample_name == "152" and row.peak_no == 2
        )

        self.assertEqual([(row.peak_no, row.target_cell) for row in mapped], [(1, "F19")])
        self.assertEqual(
            residual.exclude_reason, ExcludeReason.MATERIAL_RT_NOT_CLOSEST.value
        )
        self.assertIn("기준 RT 1.250", residual.message or "")

    def test_diethyl_ether_without_std_rt_never_falls_back_to_largest_area(self) -> None:
        worker = Sample(
            1,
            "152",
            "152",
            SampleType.NUMERIC,
            worker_match_key="152",
            peaks=[
                Peak(1, Decimal("1.200"), 100, material_standard="Diethyl ether"),
                Peak(2, Decimal("1.400"), 9999, material_standard="Diethyl ether"),
            ],
        )

        result = self.service(diethyl_ether_snapshot()).preview_batch(
            batch([worker], "디에틸에테르"),
            Path("diethyl-ether-template.xlsx"),
            "A",
        )

        self.assertFalse(result.can_generate)
        self.assertFalse(any(
            row.status is ExcelPreviewStatus.MAPPED for row in result.rows
        ))
        self.assertTrue(all(
            row.status is ExcelPreviewStatus.ERROR for row in result.rows
        ))
        self.assertEqual(result.issues[0].code, "STD_TARGET_RT_NOT_FOUND")

    def test_diethyl_ether_rejects_mixture_workbook_with_clear_mismatch(self) -> None:
        source = batch([], "디에틸에테르")
        result = self.service(snapshot()).preview_batch(
            source, Path("mixture-template.xlsx"), "A"
        )

        self.assertFalse(result.can_generate)
        self.assertEqual(result.issues[0].code, "TEMPLATE_PROFILE_MISMATCH")
        self.assertIn("디에틸에테르", result.issues[0].message)
        self.assertIn("혼유", result.issues[0].message)

    def test_dibk_uses_corrected_area_top_two_and_keeps_overflow_excluded(self) -> None:
        values = [peak(1, 100, "DIBK"), peak(2, 300, "DIBK"), peak(3, 200, "DIBK")]
        correction = PeakCorrection(
            uuid4(), values[0].peak_id, 100, 400, "재적분", datetime.now(timezone.utc), "PC", 1
        )
        self.database.corrections[values[0].peak_id] = [correction]
        sample = Sample(1, "STD1", "STD1", SampleType.STD, replicate_no=1, peaks=values)

        result = self.service().preview_batch(batch([sample]), Path("template.xlsx"), "A")
        mapped = [row for row in result.rows if row.status is ExcelPreviewStatus.MAPPED]
        excluded = [row for row in result.rows if row.status is ExcelPreviewStatus.EXCLUDED]

        self.assertEqual([(row.peak_no, row.applied_area, row.target_cell) for row in mapped], [(1, 400, "Z15"), (2, 300, "AA15")])
        self.assertEqual(excluded[0].peak_no, 3)
        self.assertEqual(excluded[0].exclude_reason, ExcludeReason.DIBK_AREA_NOT_TOP2.value)
        self.assertTrue(result.can_generate)

    def test_recovery_level_and_replicate_map_to_confirmed_cell(self) -> None:
        recovery = Sample(
            1, "저2", "저2", SampleType.RECOVERY,
            concentration_level=ConcentrationLevel.LOW,
            replicate_no=2,
            peaks=[peak(1, 123, "acetone")],
        )
        result = self.service().preview_batch(batch([recovery]), Path("template.xlsx"), "A")
        self.assertEqual(result.rows[0].target_sheet, "회수율")
        self.assertEqual(result.rows[0].target_cell, "C38")

    def test_worker_requires_exactly_one_analysis_number_suffix_match(self) -> None:
        worker = Sample(
            1, "245", "245", SampleType.NUMERIC,
            worker_match_key="245", peaks=[peak(1, 123)],
        )
        source = batch([worker])
        source.analysis_no_start = 245
        source.analysis_no_end = 245
        unique = snapshot(TemplateCell("area", "A37", True, "261-245", "string"))
        duplicate = snapshot(
            TemplateCell("area", "A37", True, "261-245", "string"),
            TemplateCell("area", "A38", True, "262-245", "string"),
        )
        missing = snapshot()

        okay = self.service(unique).preview_batch(source, Path("template.xlsx"), "A")
        failed = self.service(duplicate).preview_batch(source, Path("template.xlsx"), "A")
        not_found = self.service(missing).preview_batch(
            source, Path("template.xlsx"), "A"
        )

        self.assertEqual(okay.rows[0].target_cell, "F37")
        self.assertTrue(okay.can_generate)
        self.assertFalse(failed.can_generate)
        self.assertEqual(failed.issues[0].code, "WORKER_ROW_NOT_UNIQUE")
        self.assertFalse(not_found.can_generate)
        self.assertEqual(not_found.issues[0].code, "WORKER_ROW_NOT_FOUND")

    def test_numeric_sample_outside_batch_and_excel_targets_is_excluded(self) -> None:
        source = batch([
            Sample(1, "1", "1", SampleType.NUMERIC, worker_match_key="1", peaks=[peak(1, 100)]),
            Sample(2, "4", "4", SampleType.NUMERIC, worker_match_key="4", peaks=[peak(1, 100)]),
            Sample(3, "22", "22", SampleType.NUMERIC, worker_match_key="22", peaks=[peak(1, 100)]),
        ])
        source.analysis_no_start = 74
        source.analysis_no_end = 119
        template = snapshot(
            TemplateCell("area", "A37", True, "262-53", "string"),
            TemplateCell("area", "A38", True, "262-54", "string"),
            TemplateCell("area", "A39", True, "262-74", "string"),
        )

        result = self.service(template).preview_batch(source, Path("template.xlsx"), "A")

        self.assertTrue(result.can_generate)
        self.assertEqual(result.mapped_count, 0)
        self.assertEqual(
            [(row.sample_name, row.exclude_reason) for row in result.rows],
            [
                ("1", ExcludeReason.NON_TARGET_SAMPLE.value),
                ("4", ExcludeReason.NON_TARGET_SAMPLE.value),
                ("22", ExcludeReason.NON_TARGET_SAMPLE.value),
            ],
        )

    def test_missing_excel_row_inside_batch_range_remains_error(self) -> None:
        source = batch([
            Sample(1, "84", "84", SampleType.NUMERIC, worker_match_key="84", peaks=[peak(1, 100)])
        ])
        source.analysis_no_start = 74
        source.analysis_no_end = 119

        result = self.service(snapshot()).preview_batch(source, Path("template.xlsx"), "A")

        self.assertFalse(result.can_generate)
        self.assertEqual(result.issues[0].code, "WORKER_ROW_NOT_FOUND")

    def test_missing_excel_row_with_only_unnamed_peak_is_excluded_not_error(self) -> None:
        unnamed = peak(1, 100)
        unnamed.material_raw = None
        unnamed.material_standard = None
        unnamed.include_for_excel = False
        unnamed.exclude_reason = ExcludeReason.UNNAMED_PEAK
        source = batch([
            Sample(1, "669-업체", "669-업체", SampleType.NUMERIC,
                   worker_match_key="669", peaks=[unnamed])
        ], "에틸렌글리콜")
        source.analysis_no_start = 599
        source.analysis_no_end = 680

        result = self.service(ethylene_glycol_snapshot()).preview_batch(
            source, Path("ethylene-glycol-template.xlsx"), "A"
        )

        self.assertTrue(result.can_generate, result.issues)
        self.assertEqual(result.error_count, 0)
        self.assertEqual(result.rows[0].exclude_reason, ExcludeReason.UNNAMED_PEAK.value)

    def test_common_sample_number_rules_map_real_numbers_and_exclude_qc(self) -> None:
        template = snapshot(
            TemplateCell("area", "A39", True, "262-84", "string"),
            TemplateCell("area", "A40", True, "262-85", "string"),
            TemplateCell("area", "A84", True, "262-119", "string"),
        )
        samples = [
            Sample(1, "84", "84", SampleType.NUMERIC, worker_match_key="84", peaks=[peak(1, 84)]),
            Sample(2, "85-업체명", "85-업체명", SampleType.NUMERIC, worker_match_key="85", peaks=[peak(1, 85)]),
            Sample(3, "119", "119", SampleType.NUMERIC, worker_match_key="119", peaks=[peak(1, 119)]),
            Sample(4, "BLANK", "BLANK", SampleType.BLANK, is_blank=True, peaks=[peak(1, 1)]),
            Sample(5, "B-control", "B-control", SampleType.UNKNOWN, peaks=[peak(1, 1)]),
            Sample(6, "0728bGCD-1", "0728bGCD-1", SampleType.NUMERIC, worker_match_key="0728", peaks=[peak(1, 1)]),
            Sample(7, "0803bGCGG-1", "0803bGCGG-1", SampleType.NUMERIC, worker_match_key="0803", peaks=[peak(1, 1)]),
        ]
        result = self.service(template).preview_batch(
            batch(samples), Path("template.xlsx"), "A"
        )
        mapped = [row for row in result.rows if row.status is ExcelPreviewStatus.MAPPED]
        excluded = [row for row in result.rows if row.status is ExcelPreviewStatus.EXCLUDED]

        self.assertTrue(result.can_generate, result.issues)
        self.assertEqual([row.target_cell for row in mapped], ["F39", "F40", "F84"])
        self.assertEqual(
            [row.exclude_reason for row in excluded],
            [
                ExcludeReason.BLANK_SAMPLE.value,
                ExcludeReason.QC_SAMPLE.value,
                ExcludeReason.QC_SAMPLE.value,
                ExcludeReason.QC_SAMPLE.value,
            ],
        )

    def test_formula_target_is_a_blocking_error(self) -> None:
        template = snapshot(
            TemplateCell("area", "F15", True, 0, "formula", "SUM(A1:A2)", 1)
        )
        std = Sample(1, "STD1", "STD1", SampleType.STD, replicate_no=1, peaks=[peak(1, 100)])
        result = self.service(template).preview_batch(batch([std]), Path("template.xlsx"), "A")
        self.assertFalse(result.can_generate)
        self.assertEqual(result.rows[0].status, ExcelPreviewStatus.ERROR)
        self.assertEqual(result.issues[0].code, "TARGET_IS_FORMULA")

    def test_duplicate_single_material_peaks_select_highest_area(self) -> None:
        std = Sample(
            1, "STD1", "STD1", SampleType.STD, replicate_no=1,
            peaks=[peak(1, 100), peak(2, 200)],
        )
        result = self.service().preview_batch(batch([std]), Path("template.xlsx"), "A")
        mapped = [row for row in result.rows if row.status is ExcelPreviewStatus.MAPPED]
        excluded = [row for row in result.rows if row.status is ExcelPreviewStatus.EXCLUDED]
        self.assertTrue(result.can_generate)
        self.assertEqual([(row.peak_no, row.target_cell) for row in mapped], [(2, "F15")])
        self.assertEqual(excluded[0].peak_no, 1)
        self.assertEqual(
            excluded[0].exclude_reason,
            ExcludeReason.MATERIAL_AREA_NOT_TOP1.value,
        )

    def test_one_column_duplicate_material_selects_peak_closest_to_target_rt(self) -> None:
        worker = Sample(
            1,
            "126-기존저장업체",
            "126-기존저장업체",
            SampleType.UNKNOWN,
            peaks=[
                Peak(7, Decimal("3.678"), 1097, material_raw="c-hexane", material_standard="c-hexane"),
                Peak(8, Decimal("3.816"), 10515, material_raw="c-hexane", material_standard="c-hexane"),
                Peak(9, Decimal("3.911"), 9546, material_raw="c-hexane", material_standard="c-hexane"),
            ],
        )
        template = one_column_snapshot(
            TemplateCell(ONE_COLUMN_PROFILE.area_sheet, "A25", True, "126", "string")
        )

        result = self.service(template).preview_batch(
            batch([worker], "1컬럼혼유"), Path("one-column-template.xlsx"), "A"
        )

        mapped = [row for row in result.rows if row.status is ExcelPreviewStatus.MAPPED]
        excluded = [row for row in result.rows if row.status is ExcelPreviewStatus.EXCLUDED]
        self.assertEqual([(row.peak_no, row.applied_area, row.target_cell) for row in mapped], [(7, 1097, "I25")])
        self.assertEqual(
            [row.exclude_reason for row in excluded],
            [
                ExcludeReason.MATERIAL_RT_NOT_CLOSEST.value,
                ExcludeReason.MATERIAL_RT_NOT_CLOSEST.value,
            ],
        )

    def test_one_column_template_maps_std_recovery_and_worker_cells(self) -> None:
        template = one_column_snapshot(
            TemplateCell("area입력", "A21", True, "39", "string")
        )
        samples = [
            Sample(
                1, "STD1", "STD1", SampleType.STD, replicate_no=1,
                peaks=[peak(1, 100, "methyl acetate")],
            ),
            Sample(
                2, "저2", "저2", SampleType.RECOVERY,
                concentration_level=ConcentrationLevel.LOW,
                replicate_no=2,
                peaks=[peak(1, 200, "c-hexane")],
            ),
            Sample(
                3, "39", "39", SampleType.NUMERIC, worker_match_key="39",
                peaks=[peak(1, 300, "n-heptane")],
            ),
        ]
        result = self.service(template).preview_batch(
            batch(samples, "1컬럼혼유"), Path("one-column-template.xlsx"), "A"
        )
        mapped = [row for row in result.rows if row.status is ExcelPreviewStatus.MAPPED]
        self.assertTrue(result.can_generate)
        self.assertEqual(
            [(row.target_sheet, row.target_cell) for row in mapped],
            [("area입력", "G5"), ("회수율", "C31"), ("area입력", "L21")],
        )

    def test_shared_four_column_alcohol_layout_uses_selected_analysis_profile(self) -> None:
        template = one_column_snapshot(
            TemplateCell("area입력", "F3", True, "IBA", "string"),
            TemplateCell("area입력", "I3", True, "1-BTOH", "string"),
            TemplateCell("area입력", "L3", True, "IAA", "string"),
            TemplateCell("area입력", "O3", True, "2-BTOH", "string"),
        )
        alcohol_2 = batch(
            [Sample(1, "STD1", "STD1", SampleType.STD, replicate_no=1,
                    peaks=[peak(1, 100, "IBA")])],
            "(알콜2) IBA,1-BTOH",
        )
        alcohol_4 = batch(
            [Sample(1, "STD1", "STD1", SampleType.STD, replicate_no=1,
                    peaks=[peak(1, 200, "IAA")])],
            "알콜4",
        )

        result_2 = self.service(template).preview_batch(
            alcohol_2, Path("(알콜2).xlsx"), "A"
        )
        result_4 = self.service(template).preview_batch(
            alcohol_4, Path("(알콜4).xlsx"), "A"
        )

        self.assertTrue(result_2.can_generate, result_2.issues)
        self.assertTrue(result_4.can_generate, result_4.issues)
        self.assertEqual(
            [(row.material, row.target_cell) for row in result_2.rows
             if row.status is ExcelPreviewStatus.MAPPED],
            [("IBA", "G5")],
        )
        self.assertEqual(
            [(row.material, row.target_cell) for row in result_4.rows
             if row.status is ExcelPreviewStatus.MAPPED],
            [("IAA", "M5")],
        )

    def test_stoddard_residual_sums_all_non_solvent_peaks(self) -> None:
        template = ExcelTemplateSnapshot(
            Path("stoddard.xlsx"),
            MEK_SHEETS,
            {
                ("LOD(area입력)", "D2"): TemplateCell(
                    "LOD(area입력)", "D2", True, "Stoddard solvent", "string"
                )
            },
        )
        std = Sample(
            1,
            "STD1",
            "STD1",
            SampleType.STD,
            replicate_no=1,
            total_area=1000,
            peaks=[
                Peak(1, Decimal("2.000"), 20, material_raw=None,
                     material_standard=None, include_for_excel=False,
                     exclude_reason=ExcludeReason.UNNAMED_PEAK, source_page=1),
                Peak(2, Decimal("2.620"), 700, material_raw=None,
                     material_standard=None, include_for_excel=False,
                     exclude_reason=ExcludeReason.UNNAMED_PEAK, source_page=1),
                Peak(3, Decimal("3.000"), 30, material_raw=None,
                     material_standard=None, include_for_excel=False,
                     exclude_reason=ExcludeReason.UNNAMED_PEAK, source_page=1),
                peak(4, 250, "Stoddard solvent"),
            ],
        )
        source = batch([std], "스토다드솔벤트")

        result = self.service(template).preview_batch(
            source, Path("stoddard.xlsx"), "A"
        )
        mapped = {
            row.target_cell: row.applied_area
            for row in result.rows
            if row.status is ExcelPreviewStatus.MAPPED
        }

        self.assertTrue(result.can_generate, result.issues)
        self.assertEqual(mapped, {"F4": 1000, "G4": 700, "H4": 50})

    @staticmethod
    def _stoddard_source(worker_total: int) -> AnalysisBatch:
        std = Sample(
            1,
            "STD1",
            "STD1",
            SampleType.STD,
            replicate_no=1,
            total_area=1000,
            peaks=[
                Peak(1, Decimal("2.620"), 700, material_raw=None,
                     material_standard=None, include_for_excel=False,
                     exclude_reason=ExcludeReason.UNNAMED_PEAK, source_page=1),
                Peak(2, Decimal("3.000"), 50, material_raw=None,
                     material_standard=None, include_for_excel=False,
                     exclude_reason=ExcludeReason.UNNAMED_PEAK, source_page=1),
            ],
        )
        worker = Sample(
            2,
            "1",
            "1",
            SampleType.NUMERIC,
            worker_match_key="1",
            total_area=worker_total,
            peaks=[
                Peak(1, Decimal("2.620"), 700, material_raw=None,
                     material_standard=None, include_for_excel=False,
                     exclude_reason=ExcludeReason.UNNAMED_PEAK, source_page=2),
                Peak(2, Decimal("3.000"), 50, material_raw=None,
                     material_standard=None, include_for_excel=False,
                     exclude_reason=ExcludeReason.UNNAMED_PEAK, source_page=2),
            ],
        )
        return batch([std, worker], "스토다드솔벤트")

    @staticmethod
    def _stoddard_worker_snapshot() -> ExcelTemplateSnapshot:
        return ExcelTemplateSnapshot(
            Path("stoddard.xlsx"),
            MEK_SHEETS,
            {
                ("LOD(area입력)", "D2"): TemplateCell(
                    "LOD(area입력)", "D2", True, "Stoddard solvent", "string"
                ),
                ("LOD(area입력)", "A19"): TemplateCell(
                    "LOD(area입력)", "A19", True, "261-1", "string"
                ),
                ("LOD(area입력)", "E19"): TemplateCell(
                    "LOD(area입력)", "E19", True, "N.D", "string"
                ),
                ("LOD(area입력)", "F19"): TemplateCell(
                    "LOD(area입력)", "F19", True, "#VALUE!", "formula", "E19*2"
                ),
            },
        )

    def test_stoddard_zero_worker_area_preserves_original_nd_cell(self) -> None:
        template = self._stoddard_worker_snapshot()
        result = self.service(template).preview_batch(
            self._stoddard_source(750), Path("stoddard.xlsx"), "A"
        )
        worker = next(
            row for row in result.rows
            if row.sample_name == "1" and row.material == "Stoddard solvent"
        )

        self.assertTrue(result.can_generate, result.issues)
        self.assertEqual(worker.applied_area, 0)
        self.assertEqual(worker.status, ExcelPreviewStatus.EXCLUDED)
        self.assertEqual(
            worker.exclude_reason, ExcludeReason.STODDARD_ND_PRESERVED.value
        )
        self.assertEqual(worker.target_cell, "E19")
        self.assertIn("N.D", worker.message)
        self.assertEqual(template.cell("LOD(area입력)", "E19").value, "N.D")
        self.assertTrue(template.cell("LOD(area입력)", "F19").has_formula)
        self.assertNotIn(
            ("LOD(area입력)", "E19", 0),
            [
                (row.target_sheet, row.target_cell, row.applied_area)
                for row in result.rows
                if row.status is ExcelPreviewStatus.MAPPED
            ],
        )

    def test_stoddard_positive_worker_area_is_still_mapped_as_number(self) -> None:
        result = self.service(self._stoddard_worker_snapshot()).preview_batch(
            self._stoddard_source(800), Path("stoddard.xlsx"), "A"
        )
        worker = next(
            row for row in result.rows
            if row.sample_name == "1" and row.material == "Stoddard solvent"
        )

        self.assertTrue(result.can_generate, result.issues)
        self.assertEqual(worker.applied_area, 50)
        self.assertEqual(worker.status, ExcelPreviewStatus.MAPPED)
        self.assertEqual(worker.target_cell, "E19")

    def test_stoddard_std2_total_cs2_residual_regression_is_unchanged(self) -> None:
        template = ExcelTemplateSnapshot(
            Path("stoddard.xlsx"),
            MEK_SHEETS,
            {
                ("LOD(area입력)", "D2"): TemplateCell(
                    "LOD(area입력)", "D2", True, "Stoddard solvent", "string"
                )
            },
        )
        std2 = Sample(
            1,
            "STD2",
            "STD2",
            SampleType.STD,
            replicate_no=2,
            total_area=1_109_240,
            peaks=[
                Peak(1, Decimal("2.620"), 981_002, material_raw=None,
                     material_standard=None, include_for_excel=False,
                     exclude_reason=ExcludeReason.UNNAMED_PEAK, source_page=1),
                Peak(2, Decimal("3.000"), 1_497, material_raw=None,
                     material_standard=None, include_for_excel=False,
                     exclude_reason=ExcludeReason.UNNAMED_PEAK, source_page=1),
            ],
        )
        result = self.service(template).preview_batch(
            batch([std2], "스토다드솔벤트"), Path("stoddard.xlsx"), "A"
        )
        mapped = {
            row.target_cell: row.applied_area
            for row in result.rows
            if row.status is ExcelPreviewStatus.MAPPED
        }

        self.assertTrue(result.can_generate, result.issues)
        self.assertEqual(
            mapped,
            {"F5": 1_109_240, "G5": 981_002, "H5": 1_497},
        )
        self.assertEqual(mapped["F5"] - mapped["G5"] - mapped["H5"], 126_741)

    def test_one_column_ignores_formula_rows_that_mirror_worker_numbers(self) -> None:
        template = one_column_snapshot(
            TemplateCell("area입력", "A23", True, "262-124", "string"),
            TemplateCell(
                "area입력", "A159", True, "262-124", "formula", "A23", 1
            ),
        )
        worker = Sample(
            1,
            "124",
            "124",
            SampleType.NUMERIC,
            worker_match_key="124",
            peaks=[peak(1, 300, "c-hexane")],
        )

        result = self.service(template).preview_batch(
            batch([worker], "1컬럼혼유"), Path("one-column-template.xlsx"), "A"
        )

        self.assertTrue(result.can_generate)
        self.assertEqual(result.rows[0].target_cell, "I23")
        self.assertFalse(result.issues)

    def test_mek_template_maps_std_recovery_and_worker_cells(self) -> None:
        samples = [
            *[
                Sample(
                    repeat, f"STD{repeat}", f"STD{repeat}", SampleType.STD,
                    replicate_no=repeat,
                    peaks=[peak(repeat, repeat * 100, "methyl ethyl ketone")],
                )
                for repeat in range(1, 6)
            ],
            *[
                Sample(
                    10 + offset, f"{label}{repeat}", f"{label}{repeat}",
                    SampleType.RECOVERY, concentration_level=level,
                    replicate_no=repeat,
                    peaks=[peak(offset, offset * 100, "methyl ethyl ketone")],
                )
                for offset, (label, level, repeat) in enumerate(
                    (
                        ("저", ConcentrationLevel.LOW, 1),
                        ("저", ConcentrationLevel.LOW, 2),
                        ("저", ConcentrationLevel.LOW, 3),
                        ("중", ConcentrationLevel.MID, 1),
                        ("중", ConcentrationLevel.MID, 2),
                        ("중", ConcentrationLevel.MID, 3),
                        ("고", ConcentrationLevel.HIGH, 1),
                        ("고", ConcentrationLevel.HIGH, 2),
                        ("고", ConcentrationLevel.HIGH, 3),
                    ),
                    start=1,
                )
            ],
            Sample(
                30, "74", "74", SampleType.NUMERIC, worker_match_key="74",
                peaks=[peak(1, 7400, "methyl ethyl ketone")],
            ),
            Sample(
                31, "119", "119", SampleType.NUMERIC, worker_match_key="119",
                peaks=[peak(1, 11900, "methyl ethyl ketone")],
            ),
        ]
        result = self.service(mek_snapshot()).preview_batch(
            batch(samples, "MEK"), Path("mek-template.xlsx"), "A"
        )
        mapped = [row for row in result.rows if row.status is ExcelPreviewStatus.MAPPED]

        self.assertTrue(result.can_generate, result.issues)
        self.assertEqual(
            [row.target_cell for row in mapped[:5]],
            ["F6", "F7", "F8", "F9", "F10"],
        )
        self.assertEqual(
            [row.target_cell for row in mapped[5:14]],
            ["B28", "B29", "B30", "B31", "B32", "B33", "B34", "B35", "B36"],
        )
        self.assertEqual(
            [(row.target_sheet, row.target_cell) for row in mapped[14:]],
            [("LOD(area입력)", "F21"), ("LOD(area입력)", "F66")],
        )

    def test_mek_duplicate_peaks_select_closest_to_std_retention_time(self) -> None:
        sample_90 = Sample(
            1, "90", "90", SampleType.NUMERIC, worker_match_key="90",
            peaks=[
                Peak(13, Decimal("3.747"), 2241, material_raw="MEK", material_standard="methyl ethyl ketone"),
                Peak(14, Decimal("3.930"), 8804, material_raw="MEK", material_standard="methyl ethyl ketone"),
                Peak(15, Decimal("4.111"), 3055, material_raw="MEK", material_standard="methyl ethyl ketone"),
            ],
        )
        sample_116 = Sample(
            2, "116", "116", SampleType.NUMERIC, worker_match_key="116",
            peaks=[
                Peak(12, Decimal("3.731"), 82446, material_raw="MEK", material_standard="methyl ethyl ketone"),
                Peak(13, Decimal("3.811"), 16329, material_raw="MEK", material_standard="methyl ethyl ketone"),
                Peak(14, Decimal("3.945"), 1630635, material_raw="MEK", material_standard="methyl ethyl ketone"),
                Peak(15, Decimal("4.125"), 380213, material_raw="MEK", material_standard="methyl ethyl ketone"),
            ],
        )
        result = self.service(mek_snapshot()).preview_batch(
            batch([sample_90, sample_116], "MEK"), Path("mek-template.xlsx"), "A"
        )
        mapped = [row for row in result.rows if row.status is ExcelPreviewStatus.MAPPED]
        residual = [row for row in result.rows if row.status is ExcelPreviewStatus.EXCLUDED]

        self.assertEqual(
            [(row.peak_no, row.applied_area, row.target_cell) for row in mapped],
            [(14, 8804, "F37"), (14, 1630635, "F63")],
        )
        self.assertTrue(all(
            row.exclude_reason == ExcludeReason.MATERIAL_RT_NOT_CLOSEST.value
            for row in residual
        ))

    def test_g2_template_maps_only_four_supported_materials(self) -> None:
        materials = ("THF", "CFM", "벤젠", "클로로벤젠")
        standards = [
            Sample(
                repeat,
                f"STD{repeat}",
                f"STD{repeat}",
                SampleType.STD,
                replicate_no=repeat,
                peaks=[peak(index, repeat * 100 + index, material) for index, material in enumerate(materials, 1)],
            )
            for repeat in range(1, 6)
        ]
        recovery = Sample(
            10,
            "저1",
            "저1",
            SampleType.RECOVERY,
            concentration_level=ConcentrationLevel.LOW,
            replicate_no=1,
            peaks=[peak(index, 1000 + index, material) for index, material in enumerate(materials, 1)],
        )
        worker = Sample(
            20,
            "655",
            "655",
            SampleType.NUMERIC,
            worker_match_key="655",
            peaks=[peak(index, 2000 + index, material) for index, material in enumerate(materials, 1)],
        )
        source = batch([*standards, recovery, worker])
        source.analysis_type = G2_PROFILE.name
        source.analysis_no_start = 655
        source.analysis_no_end = 686

        result = self.service(g2_snapshot()).preview_batch(
            source, Path("g2-template.xlsx"), "A"
        )
        mapped = [row for row in result.rows if row.status is ExcelPreviewStatus.MAPPED]

        self.assertTrue(result.can_generate, result.issues)
        self.assertEqual(
            [row.target_cell for row in mapped[:4]], ["G5", "J5", "M5", "P5"]
        )
        self.assertEqual(
            [row.target_cell for row in mapped[20:24]], ["B30", "C30", "D30", "E30"]
        )
        self.assertEqual(
            [row.target_cell for row in mapped[24:]], ["F21", "I21", "L21", "O21"]
        )
        self.assertFalse(any(row.target_cell and row.target_cell.startswith("R") for row in mapped))

    def test_isoamyl_n_propyl_acetate_profile_uses_confirmed_cells(self) -> None:
        materials = ("이소아밀 아세테이트", "n-프로필 아세테이트")
        standards = [
            Sample(
                repeat, f"STD{repeat}", f"STD{repeat}", SampleType.STD,
                replicate_no=repeat,
                peaks=[peak(index, repeat * 100 + index, material) for index, material in enumerate(materials, 1)],
            )
            for repeat in range(1, 6)
        ]
        recovery = Sample(
            10, "저1", "저1", SampleType.RECOVERY,
            concentration_level=ConcentrationLevel.LOW, replicate_no=1,
            peaks=[peak(index, 1000 + index, material) for index, material in enumerate(materials, 1)],
        )
        worker = Sample(
            20, "611", "611", SampleType.NUMERIC, worker_match_key="611",
            peaks=[peak(index, 2000 + index, material) for index, material in enumerate(materials, 1)],
        )
        source = batch([*standards, recovery, worker])
        source.analysis_type = ISOAMYL_N_PROPYL_ACETATE_PROFILE.name
        source.analysis_no_start = 611
        source.analysis_no_end = 631

        result = self.service(acetate_snapshot()).preview_batch(
            source, Path("acetate-template.xlsx"), "A"
        )
        mapped = [row for row in result.rows if row.status is ExcelPreviewStatus.MAPPED]

        self.assertTrue(result.can_generate, result.issues)
        self.assertEqual([row.target_cell for row in mapped[:2]], ["G5", "J5"])
        self.assertEqual([row.target_cell for row in mapped[10:12]], ["B28", "C28"])
        self.assertEqual([row.target_cell for row in mapped[12:]], ["F20", "I20"])

    def test_cellosolve_profile_uses_confirmed_cells_and_excludes_std6(self) -> None:
        materials = (
            "2-Butoxyethanol",
            "2-Butoxyethyl acetate",
            "2-Ethoxyethanol",
            "2-Ethoxyethyl acetate",
        )
        standards = [
            Sample(
                repeat, f"STD{repeat}", f"STD{repeat}", SampleType.STD,
                replicate_no=repeat,
                peaks=[peak(index, repeat * 100 + index, material) for index, material in enumerate(materials, 1)],
            )
            for repeat in range(1, 7)
        ]
        recovery = Sample(
            10, "저1", "저1", SampleType.RECOVERY,
            concentration_level=ConcentrationLevel.LOW, replicate_no=1,
            peaks=[peak(index, 1000 + index, material) for index, material in enumerate(materials, 1)],
        )
        worker = Sample(
            20, "681-대운", "681-대운", SampleType.NUMERIC, worker_match_key="681",
            peaks=[peak(index, 2000 + index, material) for index, material in enumerate(materials, 1)],
        )
        source = batch([*standards, recovery, worker])
        source.analysis_type = CELLOSOLVE_PROFILE.name
        source.analysis_no_start = 681
        source.analysis_no_end = 690

        result = self.service(cellosolve_snapshot()).preview_batch(
            source, Path("cellosolve-template.xlsx"), "A"
        )
        mapped = [row for row in result.rows if row.status is ExcelPreviewStatus.MAPPED]
        excluded_std6 = [
            row for row in result.rows
            if row.sample_name == "STD6" and row.status is ExcelPreviewStatus.EXCLUDED
        ]

        self.assertTrue(result.can_generate, result.issues)
        self.assertEqual([row.target_cell for row in mapped[:4]], ["G5", "J5", "M5", "P5"])
        self.assertEqual([row.target_cell for row in mapped[20:24]], ["B30", "C30", "D30", "E30"])
        self.assertEqual([row.target_cell for row in mapped[24:]], ["F21", "I21", "L21", "O21"])
        self.assertEqual(len(excluded_std6), 4)
        self.assertTrue(all(row.exclude_reason == "STD_METHOD_A_NOT_SELECTED" for row in excluded_std6))

    def test_acetic_acid_profile_uses_confirmed_single_material_cells(self) -> None:
        standards = [
            Sample(
                repeat, f"STD{repeat}", f"STD{repeat}", SampleType.STD,
                replicate_no=repeat,
                peaks=[peak(1, repeat * 100, "초산")],
            )
            for repeat in range(1, 6)
        ]
        recovery = Sample(
            10, "저1", "저1", SampleType.RECOVERY,
            concentration_level=ConcentrationLevel.LOW, replicate_no=1,
            peaks=[peak(1, 1000, "초산")],
        )
        worker = Sample(
            20, "637-비티젠", "637-비티젠", SampleType.NUMERIC,
            worker_match_key="637", peaks=[peak(1, 2000, "초산")],
        )
        source = batch([*standards, recovery, worker])
        source.analysis_type = ACETIC_ACID_PROFILE.name
        source.analysis_no_start = 637
        source.analysis_no_end = 666

        result = self.service(acetic_snapshot()).preview_batch(
            source, Path("acetic-template.xlsx"), "A"
        )
        mapped = [row for row in result.rows if row.status is ExcelPreviewStatus.MAPPED]

        self.assertTrue(result.can_generate, result.issues)
        self.assertEqual([row.target_cell for row in mapped[:5]], ["F4", "F5", "F6", "F7", "F8"])
        self.assertEqual(mapped[5].target_cell, "B28")
        self.assertEqual(mapped[6].target_cell, "F19")

    def test_acn_profile_uses_confirmed_single_material_cells(self) -> None:
        standards = [
            Sample(
                repeat, f"STD{repeat}", f"STD{repeat}", SampleType.STD,
                replicate_no=repeat,
                peaks=[peak(1, repeat * 100, "Acetonitrile")],
            )
            for repeat in range(1, 6)
        ]
        recovery = Sample(
            10, "저1", "저1", SampleType.RECOVERY,
            concentration_level=ConcentrationLevel.LOW, replicate_no=1,
            peaks=[peak(1, 1000, "Acetonitrile")],
        )
        worker = Sample(
            20, "656-비티젠", "656-비티젠", SampleType.NUMERIC,
            worker_match_key="656", peaks=[peak(1, 2000, "Acetonitrile")],
        )
        source = batch([*standards, recovery, worker])
        source.analysis_type = ACN_PROFILE.name
        source.analysis_no_start = 656
        source.analysis_no_end = 666

        result = self.service(acn_snapshot()).preview_batch(
            source, Path("acn-template.xlsx"), "A"
        )
        mapped = [row for row in result.rows if row.status is ExcelPreviewStatus.MAPPED]

        self.assertTrue(result.can_generate, result.issues)
        self.assertEqual([row.target_cell for row in mapped[:5]], ["F14", "F15", "F16", "F17", "F18"])
        self.assertEqual(mapped[5].target_cell, "D29")
        self.assertEqual(mapped[6].target_cell, "F26")

    def test_bc_profile_maps_only_confirmed_two_butoxyethanol_cells(self) -> None:
        standards = [
            Sample(
                repeat, f"STD{repeat}", f"STD{repeat}", SampleType.STD,
                replicate_no=repeat,
                peaks=[peak(1, repeat * 100, "2-Butoxyethanol")],
            )
            for repeat in range(1, 6)
        ]
        recovery = Sample(
            10, "저1", "저1", SampleType.RECOVERY,
            concentration_level=ConcentrationLevel.LOW, replicate_no=1,
            peaks=[peak(1, 1000, "2-Butoxyethanol")],
        )
        worker = Sample(
            20, "588-신성대", "588-신성대", SampleType.NUMERIC,
            worker_match_key="588", peaks=[peak(1, 2000, "2-Butoxyethanol")],
        )
        source = batch([*standards, recovery, worker])
        source.analysis_type = BC_PROFILE.name
        source.analysis_no_start = 570
        source.analysis_no_end = 588

        result = self.service(bc_snapshot()).preview_batch(
            source, Path("bc-template.xlsx"), "A"
        )
        mapped = [row for row in result.rows if row.status is ExcelPreviewStatus.MAPPED]

        self.assertTrue(result.can_generate, result.issues)
        self.assertEqual([row.target_cell for row in mapped[:5]], ["F14", "F15", "F16", "F17", "F18"])
        self.assertEqual(mapped[5].target_cell, "B29")
        self.assertEqual(mapped[6].target_cell, "F26")

    def test_g3_profile_uses_confirmed_cells_and_selects_tce_rt_cluster(self) -> None:
        materials = (
            ("1,2-Dichloroethylene", "3.698"),
            ("Trichloroethylene", "5.806"),
            ("Tetrachloroethylene", "6.437"),
            ("1,2-Dichloropropane", "7.005"),
            ("1,2-Dichloroethane", "7.675"),
        )

        def target_peaks(seed: int) -> list[Peak]:
            result = [
                Peak(index, Decimal(rt), seed + index, material_raw=material, material_standard=material)
                for index, (material, rt) in enumerate(materials, 1)
            ]
            result.append(
                Peak(
                    6,
                    Decimal("5.762"),
                    seed + 100,
                    material_raw="TCE",
                    material_standard="Trichloroethylene",
                )
            )
            return result

        standards = [
            Sample(
                repeat, f"STD{repeat}", f"STD{repeat}", SampleType.STD,
                replicate_no=repeat, peaks=target_peaks(repeat * 100),
            )
            for repeat in range(1, 6)
        ]
        recovery = Sample(
            10, "저1", "저1", SampleType.RECOVERY,
            concentration_level=ConcentrationLevel.LOW, replicate_no=1,
            peaks=target_peaks(1000),
        )
        worker = Sample(
            20, "695-제이앤지", "695-제이앤지", SampleType.NUMERIC,
            worker_match_key="695", peaks=target_peaks(2000),
        )
        source = batch([*standards, recovery, worker])
        source.analysis_type = G3_PROFILE.name
        source.analysis_no_start = 695
        source.analysis_no_end = 696

        result = self.service(g3_snapshot()).preview_batch(
            source, Path("g3-template.xlsx"), "A"
        )
        mapped = [row for row in result.rows if row.status is ExcelPreviewStatus.MAPPED]
        residual_tce = [
            row for row in result.rows
            if row.material == "Trichloroethylene"
            and row.status is ExcelPreviewStatus.EXCLUDED
        ]

        self.assertTrue(result.can_generate, result.issues)
        self.assertEqual([row.target_cell for row in mapped[:5]], ["G5", "J5", "M5", "P5", "S5"])
        self.assertEqual([row.target_cell for row in mapped[25:30]], ["B30", "C30", "D30", "E30", "F30"])
        self.assertEqual([row.target_cell for row in mapped[30:]], ["F21", "I21", "L21", "O21", "R21"])
        self.assertEqual(len(residual_tce), 7)
        self.assertTrue(all(row.exclude_reason == ExcludeReason.MATERIAL_RT_NOT_CLOSEST.value for row in residual_tce))

    def test_dibk_equal_area_uses_peak_number_as_tie_breaker(self) -> None:
        values = [peak(3, 100, "DIBK"), peak(1, 100, "DIBK"), peak(2, 100, "DIBK")]
        std = Sample(1, "STD1", "STD1", SampleType.STD, replicate_no=1, peaks=values)
        result = self.service().preview_batch(batch([std]), Path("template.xlsx"), "A")
        mapped = [row for row in result.rows if row.status is ExcelPreviewStatus.MAPPED]
        self.assertEqual([(row.peak_no, row.target_cell) for row in mapped], [(1, "Z15"), (2, "AA15")])


class XlsxTemplateInspectorTests(unittest.TestCase):
    def test_sample_template_exposes_confirmed_input_and_formula_cells(self) -> None:
        template = Path(__file__).parents[3] / "TEST" / "(혼유) 틀.xlsx"
        if not template.is_file():
            self.skipTest(f"실제 테스트 양식 없음: {template}")
        inspected = XlsxTemplateInspector().inspect(template)
        self.assertEqual(inspected.sheet_names, SHEETS)
        self.assertFalse(inspected.cell("area", "F15").has_formula)
        self.assertEqual(inspected.cell("area", "F15").value_type, "blank")
        self.assertTrue(inspected.cell("area", "R15").has_formula)
        self.assertIn("Z15+AA15", inspected.cell("area", "R15").formula)

    def test_one_column_template_exposes_confirmed_input_cells(self) -> None:
        template = Path(__file__).parents[3] / "TEST" / "(1컬럼혼유-틀).xlsx"
        if not template.is_file():
            self.skipTest(f"실제 테스트 양식 없음: {template}")
        inspected = XlsxTemplateInspector().inspect(template)
        self.assertEqual(inspected.sheet_names, ONE_COLUMN_SHEETS)
        for sheet, address in (
            ("area입력", "G5"),
            ("area입력", "J5"),
            ("area입력", "M5"),
            ("area입력", "P5"),
            ("회수율", "B30"),
            ("회수율", "E38"),
        ):
            with self.subTest(sheet=sheet, address=address):
                self.assertFalse(inspected.cell(sheet, address).has_formula)


if __name__ == "__main__":
    unittest.main()
