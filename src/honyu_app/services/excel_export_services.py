from pathlib import Path
from typing import Protocol

from honyu_app.domain.models import ExcelCellWrite, WorkbookValidationResult


class ExcelCellWriter(Protocol):
    def write_copy(
        self,
        template_path: Path,
        output_path: Path,
        writes: list[ExcelCellWrite],
    ) -> None: ...


class ExcelRecalculator(Protocol):
    def recalculate(self, workbook_path: Path) -> None: ...


class WorkbookValidator(Protocol):
    def validate(
        self,
        original_path: Path,
        result_path: Path,
        writes: list[ExcelCellWrite],
        *,
        after_excel_recalculation: bool,
    ) -> WorkbookValidationResult: ...
