import unittest

from honyu_app.config.analysis_types import (
    ANALYSIS_TYPE_NAMES,
    AnalysisTypeDefinition,
    MaterialDefinition,
    G2_SUPPORTED_MATERIALS,
    analysis_type_display_name,
    analysis_type_key,
    excel_profile_key_for,
    has_excel_profile,
    infer_analysis_type,
    materials_pending_for,
    material_aliases,
    supported_materials_for,
    validate_analysis_type_registry,
)


EXPECTED_NEW_TYPES = (
    "(알콜2) IBA,1-BTOH",
    "IPA",
    "메탄올A",
    "ACN",
    "B.C",
    "DMF,DMA",
    "이소아밀,n-프로필 아세테이트",
    "페놀",
    "1,2-에폭시프로판(산화프로필렌)",
    "디에틸에테르",
    "THF(Diethylene oxide)",
    "PCE(테트라클로로에틸렌)",
    "디클로로메탄(MC)",
    "메틸 n아밀케톤",
    "메틸클로라이드(Chloromethane)",
    "비닐아세테이트",
    "셀로솔브",
    "알콜4",
    "스토다드솔벤트",
    "이소프로필 아세테이트",
    "초산",
    "피리딘",
    "에틸렌글리콜",
    "(혼유-G2) THF,CFM,벤젠,클로로벤젠",
    "(혼유-G3) 1,2-디클로로에틸렌,퍼클로로에틸렌,프로판,에탄",
)


class AnalysisTypeRegistryTests(unittest.TestCase):
    def test_only_ambiguous_analysis_types_remain_materials_pending(self) -> None:
        pending = {
            name for name in ANALYSIS_TYPE_NAMES if materials_pending_for(name)
        }
        self.assertEqual(
            pending,
            {
                "THF(Diethylene oxide)",
            },
        )

    def test_clear_pre_registered_materials_are_complete(self) -> None:
        expected = {
            "IPA": ("Isopropyl alcohol",),
            "메탄올A": ("Methanol",),
            "ACN": ("Acetonitrile",),
            "페놀": ("Phenol",),
            "1,2-에폭시프로판(산화프로필렌)": ("Propylene oxide",),
            "디에틸에테르": ("Diethyl ether",),
            "PCE(테트라클로로에틸렌)": ("Tetrachloroethylene",),
            "디클로로메탄(MC)": ("Dichloromethane",),
            "메틸클로라이드(Chloromethane)": ("Chloromethane",),
            "비닐아세테이트": ("Vinyl acetate",),
            "이소프로필 아세테이트": ("Isopropyl acetate",),
            "초산": ("초산",),
            "피리딘": ("Pyridine",),
            "에틸렌글리콜": ("Ethylene glycol",),
            "DMF,DMA": ("DMF", "DMA"),
            "메틸 n아밀케톤": ("메틸 n-아밀케톤",),
            "알콜4": ("IBA", "n-BTOH", "IAA", "2-BTOH"),
            "스토다드솔벤트": ("Stoddard solvent",),
        }
        for analysis_type, canonical_names in expected.items():
            with self.subTest(analysis_type=analysis_type):
                self.assertFalse(materials_pending_for(analysis_type))
                self.assertEqual(
                    tuple(item.canonical_name for item in supported_materials_for(analysis_type)),
                    canonical_names,
                )

    def test_ethylene_glycol_is_selectable_and_has_excel_profile(self) -> None:
        self.assertIn("에틸렌글리콜", ANALYSIS_TYPE_NAMES)
        self.assertFalse(materials_pending_for("에틸렌글리콜"))
        self.assertTrue(has_excel_profile("에틸렌글리콜"))
        self.assertEqual(
            supported_materials_for("에틸렌글리콜")[0].aliases,
            ("에틸렌글리콜", "Ethylene glycol", "E.G"),
        )
        self.assertEqual(
            infer_analysis_type("에틸렌글리콜 100-110.pdf"), "에틸렌글리콜"
        )

    def test_new_analysis_type_cannot_silently_omit_materials(self) -> None:
        with self.assertRaisesRegex(ValueError, "지원 물질이 없습니다"):
            validate_analysis_type_registry(
                (AnalysisTypeDefinition("new_type", "새 분석종류"),)
            )

    def test_normalized_alias_collision_is_rejected(self) -> None:
        definitions = (
            AnalysisTypeDefinition(
                "one", "one", (MaterialDefinition("A", ("M.C",)),)
            ),
            AnalysisTypeDefinition(
                "two", "two", (MaterialDefinition("B", ("MC",)),)
            ),
        )
        with self.assertRaisesRegex(ValueError, "중복 등록"):
            material_aliases(definitions, ())

    def test_g2_material_allow_list_excludes_carbon_tetrachloride(self) -> None:
        self.assertEqual(
            G2_SUPPORTED_MATERIALS,
            ("THF", "CFM", "벤젠", "클로로벤젠"),
        )
        self.assertNotIn("사염화탄소", G2_SUPPORTED_MATERIALS)

    def test_isoamyl_n_propyl_materials_and_aliases_live_in_registry(self) -> None:
        materials = supported_materials_for("이소아밀,n-프로필 아세테이트")
        self.assertEqual(
            [(item.canonical_name, item.aliases) for item in materials],
            [
                (
                    "n-프로필 아세테이트",
                    (
                        "프로필아세테이트",
                        "n-프로필아세테이트",
                        "n-프로필 아세테이트",
                        "초산프로필",
                        "propyl acetate",
                        "n-propyl acetate",
                    ),
                ),
                (
                    "이소아밀 아세테이트",
                    (
                        "이소아밀아세테이트",
                        "이소아밀 아세테이트",
                        "초산이소아밀",
                        "isoamyl acetate",
                    ),
                ),
            ],
        )

    def test_cellosolve_materials_and_aliases_live_in_registry(self) -> None:
        materials = supported_materials_for("셀로솔브")
        self.assertEqual(
            [(item.canonical_name, item.aliases) for item in materials],
            [
                ("2-Butoxyethanol", ("부톡시에탄올(BC)", "2-부톡시에탄올", "BC")),
                ("2-Butoxyethyl acetate", ("부톡시에틸아세테이트",)),
                ("2-Ethoxyethanol", ("에톡시에탄올",)),
                ("2-Ethoxyethyl acetate", ("에톡시에틸아세테이트",)),
            ],
        )

    def test_g3_materials_and_aliases_live_in_registry(self) -> None:
        materials = supported_materials_for(
            "(혼유-G3) 1,2-디클로로에틸렌,퍼클로로에틸렌,프로판,에탄"
        )
        self.assertEqual(
            [(item.canonical_name, item.aliases) for item in materials],
            [
                ("1,2-Dichloroethylene", ("1,2디클로로에틸렌",)),
                ("Trichloroethylene", ("tce",)),
                (
                    "Tetrachloroethylene",
                    ("pce", "테트라클로로에틸렌", "tetrachloroethylene"),
                ),
                ("1,2-Dichloropropane", ("1,2디클로로프로판",)),
                ("1,2-Dichloroethane", ("1,2디클로로에탄",)),
            ],
        )

    def test_registry_contains_existing_and_requested_types_once(self) -> None:
        self.assertEqual(ANALYSIS_TYPE_NAMES[:3], ("혼유", "1컬럼혼유", "MEK"))
        self.assertEqual(ANALYSIS_TYPE_NAMES[3:], EXPECTED_NEW_TYPES)
        self.assertNotIn("알콜", ANALYSIS_TYPE_NAMES)
        self.assertEqual(len(ANALYSIS_TYPE_NAMES), len(set(ANALYSIS_TYPE_NAMES)))

    def test_only_implemented_excel_profiles_are_marked_registered(self) -> None:
        for name in (
            "혼유",
            "1컬럼혼유",
            "MEK",
            "(알콜2) IBA,1-BTOH",
            "IPA",
            "메탄올A",
            "(혼유-G2) THF,CFM,벤젠,클로로벤젠",
            "이소아밀,n-프로필 아세테이트",
            "셀로솔브",
            "(혼유-G3) 1,2-디클로로에틸렌,퍼클로로에틸렌,프로판,에탄",
            "초산",
            "ACN",
            "에틸렌글리콜",
            "B.C",
            "디에틸에테르",
            "DMF,DMA",
            "알콜4",
            "스토다드솔벤트",
            "1,2-에폭시프로판(산화프로필렌)",
            "디클로로메탄(MC)",
            "메틸 n아밀케톤",
            "비닐아세테이트",
            "이소프로필 아세테이트",
            "피리딘",
            "페놀",
        ):
            self.assertTrue(has_excel_profile(name), name)
        for name in (
            item
            for item in EXPECTED_NEW_TYPES[1:]
            if item not in {
                "(혼유-G2) THF,CFM,벤젠,클로로벤젠",
                "이소아밀,n-프로필 아세테이트",
                "셀로솔브",
                "(혼유-G3) 1,2-디클로로에틸렌,퍼클로로에틸렌,프로판,에탄",
                "초산",
                "ACN",
                "IPA",
                "메탄올A",
                "에틸렌글리콜",
                "B.C",
                "디에틸에테르",
                "DMF,DMA",
                "알콜4",
                "스토다드솔벤트",
                "1,2-에폭시프로판(산화프로필렌)",
                "디클로로메탄(MC)",
                "메틸 n아밀케톤",
                "비닐아세테이트",
                "이소프로필 아세테이트",
                "피리딘",
                "페놀",
            }
        ):
            self.assertFalse(has_excel_profile(name), name)

    def test_display_names_and_internal_keys_are_separate(self) -> None:
        mappings = {
            "메탄올A": "methanol",
            "ACN": "acn",
            "에틸렌글리콜": "ethylene_glycol",
            "(알콜2) IBA,1-BTOH": "alcohol_2_iba_1_btoh",
            "(혼유-G2) THF,CFM,벤젠,클로로벤젠": "mixture_g2",
            "(혼유-G3) 1,2-디클로로에틸렌,퍼클로로에틸렌,프로판,에탄": "mixture_g3",
        }
        for display_name, key in mappings.items():
            self.assertEqual(analysis_type_key(display_name), key)
            self.assertEqual(analysis_type_display_name(key), display_name)

    def test_diethyl_ether_uses_dedicated_excel_profile_key(self) -> None:
        self.assertEqual(excel_profile_key_for("디에틸에테르"), "diethyl_ether")

    def test_ipa_uses_dedicated_excel_profile_key(self) -> None:
        self.assertEqual(excel_profile_key_for("IPA"), "ipa")

    def test_methanol_uses_dedicated_excel_profile_key(self) -> None:
        self.assertEqual(excel_profile_key_for("메탄올A"), "methanol")

    def test_phenol_uses_dedicated_excel_profile_key(self) -> None:
        self.assertEqual(excel_profile_key_for("페놀"), "phenol")

    def test_specific_detection_precedes_broad_family_detection(self) -> None:
        self.assertEqual(infer_analysis_type("알콜4 1-10.pdf"), "알콜4")
        self.assertEqual(
            infer_analysis_type("(혼유-G2)THF,CFM,벤젠,클로로벤젠 1-10.pdf"),
            "(혼유-G2) THF,CFM,벤젠,클로로벤젠",
        )
        self.assertEqual(
            infer_analysis_type("G3 혼유 695-696.pdf"),
            "(혼유-G3) 1,2-디클로로에틸렌,퍼클로로에틸렌,프로판,에탄",
        )
        self.assertEqual(infer_analysis_type("MEK 74-119.pdf"), "MEK")
        self.assertEqual(infer_analysis_type("메탄올 74-119.pdf"), "메탄올A")
        self.assertEqual(infer_analysis_type("ACN 74-119.pdf"), "ACN")
        self.assertEqual(infer_analysis_type("초산 489-530.pdf"), "초산")
        self.assertEqual(infer_analysis_type("acetic acid 489-530.pdf"), "초산")
        self.assertEqual(infer_analysis_type("(IPA) 320-334.pdf"), "IPA")
        self.assertEqual(infer_analysis_type("IPA 320-334.pdf"), "IPA")
        self.assertEqual(
            infer_analysis_type("unknown.pdf", method_filenames=("IPA",)), "IPA"
        )
        self.assertEqual(
            infer_analysis_type("MEK 74-119.pdf", materials=("IPA",)), "MEK"
        )
        self.assertEqual(
            infer_analysis_type("알콜(2) 74-119.pdf"), "(알콜2) IBA,1-BTOH"
        )
        self.assertEqual(
            infer_analysis_type("unknown.pdf", materials=("IBA", "1-BTOH")),
            "(알콜2) IBA,1-BTOH",
        )
        self.assertIsNone(infer_analysis_type("알콜 1-10.pdf"))
        self.assertEqual(infer_analysis_type("1컬럼혼유 120-130.pdf"), "1컬럼혼유")

    def test_ipa_token_detection_does_not_override_other_analysis_types(self) -> None:
        expected = {
            "MEK 74-119.pdf": "MEK",
            "ACN 656-666.pdf": "ACN",
            "초산 489-530.pdf": "초산",
            "알콜(2) 74-119.pdf": "(알콜2) IBA,1-BTOH",
            "1컬럼혼유 120-130.pdf": "1컬럼혼유",
            "G3 혼유 695-696.pdf": "(혼유-G3) 1,2-디클로로에틸렌,퍼클로로에틸렌,프로판,에탄",
        }
        for filename, analysis_type in expected.items():
            with self.subTest(filename=filename):
                self.assertEqual(infer_analysis_type(filename), analysis_type)
        self.assertIsNone(infer_analysis_type("isopropyl acetate 320-334.pdf"))

    def test_filename_analysis_type_has_priority_over_detected_materials(self) -> None:
        cases = (
            ("(페놀) 256-305.pdf", ("MeOH",), "페놀"),
            ("(메탄올)A 237-320.pdf", ("DMF",), "메탄올A"),
            ("초산 489-530.pdf", ("Formic acid",), "초산"),
            ("(IPA) 320-334.pdf", ("2-BTOH",), "IPA"),
        )
        for filename, materials, expected in cases:
            with self.subTest(filename=filename):
                self.assertEqual(
                    infer_analysis_type(filename, materials=materials), expected
                )

    def test_g2_g3_filename_markers_ignore_separators_and_override_internal_evidence(self) -> None:
        g2 = "(혼유-G2) THF,CFM,벤젠,클로로벤젠"
        g3 = "(혼유-G3) 1,2-디클로로에틸렌,퍼클로로에틸렌,프로판,에탄"
        cases = {
            "혼유-G2 결과.pdf": g2,
            "혼유(G2 결과).pdf": g2,
            "혼유(G2-THF).pdf": g2,
            "G2 혼유 결과.pdf": g2,
            "G2-혼유 결과.pdf": g2,
            "혼유 G2 결과.pdf": g2,
            "혼유-G3 결과.pdf": g3,
            "혼유(G3 결과).pdf": g3,
            "혼유(G3-결과).pdf": g3,
            "G3 혼유 결과.pdf": g3,
            "G3-혼유 결과.pdf": g3,
            "혼유 G3 결과.pdf": g3,
            "혼유(G3-1,2디클로로에탄) 152-153@완료.pdf": g3,
        }
        for filename, expected in cases.items():
            with self.subTest(filename=filename):
                self.assertEqual(
                    infer_analysis_type(
                        filename,
                        method_filenames=("일반 혼유.gcm",),
                        materials=("n-hexane", "DIBK"),
                    ),
                    expected,
                )
        self.assertEqual(infer_analysis_type("혼유 120-167.pdf"), "혼유")

    def test_internal_evidence_is_used_only_when_filename_is_unknown(self) -> None:
        self.assertEqual(
            infer_analysis_type(
                "unknown 256-305.pdf",
                method_filenames=("unknown.gcm",),
                materials=("Methanol",),
            ),
            "메탄올A",
        )
