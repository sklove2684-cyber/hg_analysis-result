class HonyuError(Exception):
    """Base application error."""


class ValidationError(HonyuError):
    pass


class DuplicateSourceFileError(HonyuError):
    pass


class ParserUnavailableError(HonyuError):
    pass


class DatabaseUnavailableError(HonyuError):
    pass


class SharedFolderUnavailableError(HonyuError):
    pass


class SharedFolderPathError(HonyuError):
    pass


class RecordNotFoundError(HonyuError):
    pass


class RevisionConflictError(HonyuError):
    pass


class ExtractionCancelledError(HonyuError):
    pass


class ExcelExportError(HonyuError):
    pass


class ExcelRecalculationError(ExcelExportError):
    pass


class WorkbookStructureError(ExcelExportError):
    pass
