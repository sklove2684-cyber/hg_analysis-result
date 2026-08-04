from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Protocol


@dataclass(frozen=True, slots=True)
class TemplateCell:
    sheet: str
    address: str
    exists: bool
    value: object | None = None
    value_type: str = "missing"
    formula: str | None = None
    style_id: int | None = None

    @property
    def has_formula(self) -> bool:
        return self.formula is not None


@dataclass(slots=True)
class ExcelTemplateSnapshot:
    path: Path
    sheet_names: tuple[str, ...]
    cells: dict[tuple[str, str], TemplateCell] = field(default_factory=dict)

    def cell(self, sheet: str, address: str) -> TemplateCell:
        return self.cells.get(
            (sheet, address), TemplateCell(sheet=sheet, address=address, exists=False)
        )


class ExcelTemplateService(Protocol):
    def inspect(self, path: Path) -> ExcelTemplateSnapshot: ...
