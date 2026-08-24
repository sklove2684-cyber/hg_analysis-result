import unittest
from decimal import Decimal

from honyu_app.domain.enums import ConcentrationLevel, ExcludeReason, SampleType
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

    def test_pending_single_material_exact_match_is_inferred_for_session(self) -> None:
        samples = [
            self._named_sample(1, "STD1", SampleType.STD, "메틸 n아밀케톤"),
            self._named_sample(2, "STD2", SampleType.STD, "메틸 n아밀케톤"),
            self._named_sample(3, "저1", SampleType.RECOVERY, "메틸 n아밀케톤"),
        ]
        aliases = self.parser._infer_session_material_aliases("메틸 n아밀케톤", samples)
        self.assertEqual(aliases, {"메틸n아밀케톤": "메틸 n아밀케톤"})

        self.parser._apply_session_material_aliases(
            samples, aliases, {"메틸 n아밀케톤", "CS2"}
        )
        self.assertTrue(
            all(sample.peaks[0].material_standard == "메틸 n아밀케톤" for sample in samples)
        )
        self.assertTrue(all(sample.peaks[0].exclude_reason is None for sample in samples))

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
