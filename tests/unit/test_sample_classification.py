import unittest

from honyu_app.domain.enums import ConcentrationLevel, SampleType
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

    def test_name_ending_b_is_recovery_blank(self) -> None:
        sample_type, _, _, _, is_blank = self.parser._classify_sample("회수율-B")
        self.assertEqual(sample_type, SampleType.RECOVERY_BLANK)
        self.assertTrue(is_blank)
