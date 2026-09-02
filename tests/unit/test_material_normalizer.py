import unittest

from honyu_app.domain.enums import ExcludeReason, SampleType
from honyu_app.infrastructure.pdf.labsolutions_parser import LabSolutionsParser
from honyu_app.infrastructure.pdf.material_normalizer import MaterialNormalizer
from honyu_app.config.analysis_types import supported_canonical_names_for


class MaterialNormalizerTests(unittest.TestCase):
    def test_confirmed_aliases(self) -> None:
        normalizer = MaterialNormalizer()
        aliases = {
            "헥산": "n-hexane",
            "n-hexane": "n-hexane",
            "아세톤": "acetone",
            "acetone": "acetone",
            "E.A": "E.A",
            "MIBK": "MIBK",
            "Tol": "Toluene",
            "Toluene": "Toluene",
            "B.A": "B.A",
            "E.B": "E.B",
            "p": "p-xylene",
            "m": "m-xylene",
            "o": "o-xylene",
            "스티렌": "styrene",
            "시클로헥사논": "c-hexanone",
            "DIBK": "DIBK",
            "cs2": "CS2",
            "메틸아세테이트": "methyl acetate",
            "초산메틸": "methyl acetate",
            "시클로헥산": "c-hexane",
            "cyclohexane": "c-hexane",
            "헵탄": "n-heptane",
            "n-heptane": "n-heptane",
            "이소부틸아세테이트": "isobutyl acetate",
            "초산이소부틸": "isobutyl acetate",
            "IBA": "IBA",
            "n-BTOH": "n-BTOH",
            "MEK": "methyl ethyl ketone",
            "Methyl ethyl ketone": "methyl ethyl ketone",
            "메틸에틸케톤": "methyl ethyl ketone",
            "2-butanone": "methyl ethyl ketone",
            "THF": "THF",
            "CFM": "CFM",
            "벤젠": "벤젠",
            "클로로벤젠": "클로로벤젠",
            "초산프로필": "n-프로필 아세테이트",
            "초산이소아밀": "이소아밀 아세테이트",
            "이소아밀아세테이트": "이소아밀 아세테이트",
            "isoamyl acetate": "이소아밀 아세테이트",
            "프로필아세테이트": "n-프로필 아세테이트",
            "n-프로필아세테이트": "n-프로필 아세테이트",
            "propyl acetate": "n-프로필 아세테이트",
            "n-propyl acetate": "n-프로필 아세테이트",
            "부톡시에탄올(BC)": "2-Butoxyethanol",
            "부톡시에틸아세테이트": "2-Butoxyethyl acetate",
            "에톡시에탄올": "2-Ethoxyethanol",
            "에톡시에틸아세테이트": "2-Ethoxyethyl acetate",
            "1,2디클로로에틸렌": "1,2-Dichloroethylene",
            "TCE": "Trichloroethylene",
            "PCE": "Tetrachloroethylene",
            "1,2디클로로프로판": "1,2-Dichloropropane",
            "1,2디클로로에탄": "1,2-Dichloroethane",
            "IPA": "Isopropyl alcohol",
            "Isopropyl alcohol": "Isopropyl alcohol",
            "메탄올": "Methanol",
            "메탄올A": "Methanol",
            "ACN": "Acetonitrile",
            "페놀": "Phenol",
            "Phenol": "Phenol",
            "산화프로필렌": "Propylene oxide",
            "1,2-에폭시프로판": "Propylene oxide",
            "1,2에폭시프로판": "Propylene oxide",
            "디에틸에테르": "Diethyl ether",
            "테트라클로로에틸렌": "Tetrachloroethylene",
            "디클로로메탄": "Dichloromethane",
            "MC": "Dichloromethane",
            "M.C": "Dichloromethane",
            "Methylene chloride": "Dichloromethane",
            "메틸클로라이드": "Chloromethane",
            "비닐아세테이트": "Vinyl acetate",
            "이소프로필 아세테이트": "Isopropyl acetate",
            "초산": "초산",
            "피리딘": "Pyridine",
            "에틸렌글리콜": "Ethylene glycol",
            "Ethylene glycol": "Ethylene glycol",
            "E.G": "Ethylene glycol",
            "BC": "2-Butoxyethanol",
            "2-부톡시에탄올": "2-Butoxyethanol",
            "MeOH": "Methanol",
            "초산이소프로필": "Isopropyl acetate",
            "2-부탄올": "2-BTOH",
            "n-부탄올": "n-BTOH",
            "메틸 n-아밀케톤": "메틸 n-아밀케톤",
            "스토다드솔벤트": "Stoddard solvent",
        }
        for raw, expected in aliases.items():
            with self.subTest(raw=raw):
                self.assertEqual(normalizer.normalize(raw), expected)

    def test_unknown_name_is_not_guessed(self) -> None:
        self.assertIsNone(MaterialNormalizer().normalize("새로운 물질"))

    def test_carbon_tetrachloride_is_not_cfm_or_an_automation_material(self) -> None:
        normalizer = MaterialNormalizer()
        self.assertEqual(normalizer.normalize("CFM"), "CFM")
        self.assertEqual(normalizer.normalize("사염화탄소"), "Carbon tetrachloride")
        self.assertEqual(
            normalizer.normalize("carbon tetrachloride"), "Carbon tetrachloride"
        )
        self.assertIsNone(normalizer.normalize("tetrachloromethane"))
        self.assertIsNone(normalizer.normalize("CCl4"))
        reason = LabSolutionsParser._exclude_reason(
            "STD1", SampleType.STD, "사염화탄소", "Carbon tetrachloride",
            set(supported_canonical_names_for("(혼유-G2) THF,CFM,벤젠,클로로벤젠")) | {"CS2"},
        )
        self.assertIs(reason, ExcludeReason.MATERIAL_NOT_SUPPORTED_FOR_ANALYSIS)

    def test_dots_hyphens_case_and_spacing_are_normalized(self) -> None:
        normalizer = MaterialNormalizer()
        self.assertEqual(normalizer.normalize("  m. c  "), "Dichloromethane")
        self.assertEqual(normalizer.normalize("1,2에폭시프로판"), "Propylene oxide")
        self.assertEqual(normalizer.normalize("METHYLENE CHLORIDE"), "Dichloromethane")

    def test_known_material_outside_selected_analysis_is_excluded_without_unknown_warning(self) -> None:
        reason = LabSolutionsParser._exclude_reason(
            "611",
            SampleType.NUMERIC,
            "E.A",
            "E.A",
            {"n-프로필 아세테이트", "이소아밀 아세테이트", "CS2"},
        )
        self.assertIs(reason, ExcludeReason.MATERIAL_NOT_SUPPORTED_FOR_ANALYSIS)

    def test_bc_accepts_only_confirmed_target_and_rejects_known_solvents(self) -> None:
        parser = LabSolutionsParser()
        allowed = set(supported_canonical_names_for("B.C")) | {"CS2"}
        self.assertEqual(allowed, {"2-Butoxyethanol", "CS2"})
        self.assertIsNone(
            parser._exclude_reason(
                "STD1", SampleType.STD, "BC", "2-Butoxyethanol", allowed
            )
        )
        for raw, canonical in (("메탄올", "Methanol"), ("MC", "Dichloromethane")):
            with self.subTest(raw=raw):
                reason = parser._exclude_reason(
                    "STD1", SampleType.STD, raw, parser._normalizer.normalize(raw), allowed
                )
                self.assertEqual(parser._normalizer.normalize(raw), canonical)
                self.assertIs(
                    reason, ExcludeReason.MATERIAL_NOT_SUPPORTED_FOR_ANALYSIS
                )

    def test_registered_pce_is_eligible_for_pce_analysis(self) -> None:
        allowed = set(supported_canonical_names_for("PCE(테트라클로로에틸렌)")) | {"CS2"}
        canonical = MaterialNormalizer().normalize("PCE")
        self.assertEqual(canonical, "Tetrachloroethylene")
        self.assertIsNone(
            LabSolutionsParser._exclude_reason(
                "STD1", SampleType.STD, "PCE", canonical, allowed
            )
        )

    def test_cellosolve_excludes_other_known_registry_material(self) -> None:
        allowed = set(supported_canonical_names_for("셀로솔브")) | {"CS2"}
        canonical = MaterialNormalizer().normalize("메탄올")
        self.assertIs(
            LabSolutionsParser._exclude_reason(
                "STD1", SampleType.STD, "메탄올", canonical, allowed
            ),
            ExcludeReason.MATERIAL_NOT_SUPPORTED_FOR_ANALYSIS,
        )
