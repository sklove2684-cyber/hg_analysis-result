import unittest

from honyu_app.infrastructure.pdf.material_normalizer import MaterialNormalizer


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
        }
        for raw, expected in aliases.items():
            with self.subTest(raw=raw):
                self.assertEqual(normalizer.normalize(raw), expected)

    def test_unknown_name_is_not_guessed(self) -> None:
        self.assertIsNone(MaterialNormalizer().normalize("새로운 물질"))
