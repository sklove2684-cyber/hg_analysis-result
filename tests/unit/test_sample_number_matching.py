import unittest

from honyu_app.application.sample_number_matching import (
    classify_sample_number,
    extract_excel_analysis_number,
)
from honyu_app.domain.enums import ExcludeReason, SampleType


class SampleNumberMatchingTests(unittest.TestCase):
    def test_actual_analysis_sample_names(self) -> None:
        for raw, expected in (("84", "84"), ("85-업체명", "85"), ("119", "119")):
            with self.subTest(raw=raw):
                decision = classify_sample_number(raw, SampleType.NUMERIC)
                self.assertEqual(decision.analysis_number, expected)
                self.assertIsNone(decision.exclude_reason)

    def test_blank_qc_and_non_analysis_names_are_not_numbers(self) -> None:
        cases = (
            ("BLANK", SampleType.BLANK, ExcludeReason.BLANK_SAMPLE.value),
            ("B-control", SampleType.UNKNOWN, ExcludeReason.QC_SAMPLE.value),
            ("0728bGCD-1", SampleType.NUMERIC, ExcludeReason.QC_SAMPLE.value),
            ("0803bGCGG-1", SampleType.NUMERIC, ExcludeReason.QC_SAMPLE.value),
            ("관리시료", SampleType.UNKNOWN, ExcludeReason.NON_ANALYSIS_SAMPLE.value),
        )
        for raw, sample_type, expected_reason in cases:
            with self.subTest(raw=raw):
                decision = classify_sample_number(raw, sample_type)
                self.assertIsNone(decision.analysis_number)
                self.assertEqual(decision.exclude_reason, expected_reason)

    def test_excel_uses_only_final_analysis_number(self) -> None:
        self.assertEqual(extract_excel_analysis_number("262-84"), "84")
        self.assertEqual(extract_excel_analysis_number("262-119"), "119")
        self.assertIsNone(extract_excel_analysis_number("BLANK"))

    def test_sequence_exception_is_never_inferred_without_explicit_mapping(self) -> None:
        self.assertEqual(
            classify_sample_number("1", SampleType.NUMERIC).analysis_number, "1"
        )
        self.assertEqual(
            classify_sample_number(
                "1", SampleType.NUMERIC, explicit_sequence_overrides={"1": 86}
            ).analysis_number,
            "86",
        )

