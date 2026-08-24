from __future__ import annotations

from honyu_app.config.analysis_types import material_aliases


INITIAL_MATERIAL_ALIASES: dict[str, str] = material_aliases()


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
