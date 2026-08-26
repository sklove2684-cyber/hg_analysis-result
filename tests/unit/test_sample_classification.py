import unittest
from decimal import Decimal

from honyu_app.domain.enums import ConcentrationLevel, ExcludeReason, SampleType
from honyu_app.domain.errors import ValidationError
from honyu_app.domain.models import Peak, Sample
from honyu_app.infrastructure.pdf.labsolutions_parser import LabSolutionsParser


class SampleClassificationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.parser = LabSolutionsParser()

    def test_std_replicate(self) -> None:
        sample_type, _, replicate, worker_key, is_blank = self.parser._classify_sample("STD6")
        self.assertEqual(sample_type, SampleType.STD)
        self.assertEqual(replicate, 6)
        self.assertIsNone(worker_key)
        self.assertFalse(is_blank)

    def test_recovery_level_and_replicate_with_space(self) -> None:
        sample_type, level, replicate, _, _ = self.parser._classify_sample("저 3")
        self.assertEqual(sample_type, SampleType.RECOVERY)
        self.assertEqual(level, ConcentrationLevel.LOW)
        self.assertEqual(replicate, 3)

    def test_numeric_worker_key(self) -> None:
        sample_type, _, _, worker_key, _ = self.parser._classify_sample("245")
        self.assertEqual(sample_type, SampleType.NUMERIC)
        self.assertEqual(worker_key, "245")

    def test_numeric_worker_key_with_workplace_suffix(self) -> None:
        sample_type, _, _, worker_key, _ = self.parser._classify_sample(
            "125-현대오일"
        )
        self.assertEqual(sample_type, SampleType.NUMERIC)
        self.assertEqual(worker_key, "125")

    def test_name_ending_b_is_recovery_blank(self) -> None:
        sample_type, _, _, _, is_blank = self.parser._classify_sample("회수율-B")
        self.assertEqual(sample_type, SampleType.RECOVERY_BLANK)
        self.assertTrue(is_blank)

    def test_total_only_table_is_accepted_only_as_a_continuation(self) -> None:
        total_only = [
            [["Total", None, "5394472", "1348209", None, None, None, None]]
        ]

        self.assertEqual(
            self.parser._find_peak_table(
                total_only, 40, allow_continuation=True
            ),
            ([], 5394472),
        )
        with self.assertRaisesRegex(ValidationError, "Peak Table"):
            self.parser._find_peak_table(total_only, 40)

    def test_headerless_peak_rows_are_accepted_as_a_continuation(self) -> None:
        continuation = [
            [
                ["27", "5.582", "2272", "442", "0.000", None, "TV", None],
                ["28", "5.721", "5125", "1753", "0.000", "ppm", "TV", "MIBK"],
                ["Total", None, "5394472", "1348209", None, None, None, None],
            ]
        ]

        rows, total_area = self.parser._find_peak_table(
            continuation, 40, allow_continuation=True
        )

        self.assertEqual([row[0] for row in rows], ["27", "28"])
        self.assertEqual(total_area, 5394472)

    @staticmethod
    def _named_sample(page: int, name: str, sample_type: SampleType, material: str) -> Sample:
        return Sample(
            page,
            name,
            name,
            sample_type,
            peaks=[
                Peak(
                    1,
                    Decimal("9.141"),
                    1000 + page,
                    material_raw=material,
                    material_standard=None,
                    include_for_excel=False,
                    exclude_reason=ExcludeReason.UNKNOWN_MATERIAL,
                    source_page=page,
                )
            ],
        )

    def test_registered_material_does_not_need_session_inference(self) -> None:
        samples = [
            self._named_sample(1, "STD1", SampleType.STD, "메틸 n아밀케톤"),
            self._named_sample(2, "STD2", SampleType.STD, "메틸 n아밀케톤"),
            self._named_sample(3, "저1", SampleType.RECOVERY, "메틸 n아밀케톤"),
        ]
        self.assertEqual(
            self.parser._infer_session_material_aliases("메틸 n아밀케톤", samples),
            {},
        )

    def test_pending_material_requires_repeated_std_confirmation(self) -> None:
        samples = [
            self._named_sample(1, "STD1", SampleType.STD, "메틸 n아밀케톤"),
            self._named_sample(2, "저1", SampleType.RECOVERY, "메틸 n아밀케톤"),
        ]
        self.assertEqual(
            self.parser._infer_session_material_aliases("메틸 n아밀케톤", samples), {}
        )

    def test_pending_complex_or_nonmatching_material_is_not_guessed(self) -> None:
        samples = [
            self._named_sample(1, "STD1", SampleType.STD, "DMF"),
            self._named_sample(2, "STD2", SampleType.STD, "DMF"),
        ]
        self.assertEqual(
            self.parser._infer_session_material_aliases("DMF,DMA", samples), {}
        )

    def test_dmf_dma_assigns_only_the_immediately_following_unnamed_peak(self) -> None:
        sample = Sample(
            1,
            "STD1",
            "STD1",
            SampleType.STD,
            replicate_no=1,
            peaks=[
                Peak(1, Decimal("2.700"), 100, material_raw="DMF",
                     material_standard="DMF", source_page=1),
                Peak(2, Decimal("3.200"), 200, material_raw=None,
                     material_standard=None, include_for_excel=False,
                     exclude_reason=ExcludeReason.UNNAMED_PEAK, source_page=1),
                Peak(3, Decimal("3.500"), 300, material_raw=None,
                     material_standard=None, include_for_excel=False,
                     exclude_reason=ExcludeReason.UNNAMED_PEAK, source_page=1),
            ],
        )

        self.parser._apply_analysis_special_rules(
            "DMF,DMA", [sample], {"DMF", "DMA", "CS2"}
        )

        self.assertEqual(sample.peaks[1].material_standard, "DMA")
        self.assertTrue(sample.peaks[1].include_for_excel)
        self.assertIsNone(sample.peaks[1].exclude_reason)
        self.assertIsNone(sample.peaks[2].material_standard)
        self.assertEqual(sample.peaks[2].exclude_reason, ExcludeReason.UNNAMED_PEAK)
