from __future__ import annotations


INITIAL_MATERIAL_ALIASES: dict[str, str] = {
    "헥산": "n-hexane",
    "n-hexane": "n-hexane",
    "아세톤": "acetone",
    "acetone": "acetone",
    "e.a": "E.A",
    "mibk": "MIBK",
    "tol": "Toluene",
    "toluene": "Toluene",
    "b.a": "B.A",
    "e.b": "E.B",
    "p": "p-xylene",
    "p-xylene": "p-xylene",
    "m": "m-xylene",
    "m-xylene": "m-xylene",
    "o": "o-xylene",
    "o-xylene": "o-xylene",
    "스티렌": "styrene",
    "styrene": "styrene",
    "시클로헥사논": "c-hexanone",
    "c-hexanone": "c-hexanone",
    "dibk": "DIBK",
    "cs2": "CS2",
    "메틸아세테이트": "methyl acetate",
    "초산메틸": "methyl acetate",
    "methyl acetate": "methyl acetate",
    "methyl acatate": "methyl acetate",
    "시클로헥산": "c-hexane",
    "사이클로헥산": "c-hexane",
    "cyclohexane": "c-hexane",
    "c-hexane": "c-hexane",
    "헵탄": "n-heptane",
    "n-헵탄": "n-heptane",
    "heptane": "n-heptane",
    "n-heptane": "n-heptane",
    "이소부틸아세테이트": "isobutyl acetate",
    "초산이소부틸": "isobutyl acetate",
    "isobutyl acetate": "isobutyl acetate",
    "isobutyl acetae": "isobutyl acetate",
    "iba": "IBA",
    "n-btoh": "n-BTOH",
}


class MaterialNormalizer:
    def __init__(self, aliases: dict[str, str] | None = None) -> None:
        source = aliases or INITIAL_MATERIAL_ALIASES
        self._aliases = {self._key(alias): standard for alias, standard in source.items()}

    @staticmethod
    def _key(value: str) -> str:
        return " ".join(value.strip().split()).casefold()

    def normalize(self, raw_name: str | None) -> str | None:
        if not raw_name or not raw_name.strip():
            return None
        return self._aliases.get(self._key(raw_name))

    @property
    def standard_names(self) -> tuple[str, ...]:
        return tuple(sorted(set(self._aliases.values()), key=str.casefold))
