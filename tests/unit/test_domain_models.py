from decimal import Decimal
import unittest

from honyu_app.domain.models import Peak


class PeakTests(unittest.TestCase):
    def test_peak_preserves_integer_area(self) -> None:
        peak = Peak(peak_no=1, retention_time=Decimal("2.364"), area_raw=24350)
        self.assertEqual(peak.area_raw, 24350)
        self.assertIsInstance(peak.area_raw, int)

    def test_peak_rejects_negative_area(self) -> None:
        for area in (-1, -100):
            with self.subTest(area=area), self.assertRaisesRegex(ValueError, "area_raw"):
                Peak(peak_no=1, retention_time=Decimal("2.364"), area_raw=area)

    def test_peak_count_is_not_a_batch_invariant(self) -> None:
        peaks = [Peak(i, Decimal("1.0"), i) for i in range(1, 255)]
        self.assertEqual(len(peaks), 254)
