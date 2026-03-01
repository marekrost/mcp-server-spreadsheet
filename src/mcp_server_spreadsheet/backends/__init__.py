"""Spreadsheet backend factory — dispatches by file extension."""

from pathlib import Path

from .base import SpreadsheetWorkbook

SUPPORTED_EXTENSIONS = {".xlsx", ".csv", ".ods"}


def load_workbook(path: str) -> SpreadsheetWorkbook:
    """Load a spreadsheet file, choosing the backend by extension."""
    p = Path(path)
    if not p.exists():
        raise ValueError(f"File not found: {path}")
    ext = p.suffix.lower()
    if ext == ".xlsx":
        from .xlsx import XlsxWorkbook
        return XlsxWorkbook.load(path)
    if ext == ".csv":
        from .csv import CsvWorkbook
        return CsvWorkbook.load(path)
    if ext == ".ods":
        from .ods import OdsWorkbook
        return OdsWorkbook.load(path)
    raise ValueError(
        f"Unsupported file format: {ext!r}. "
        f"Supported: {', '.join(sorted(SUPPORTED_EXTENSIONS))}"
    )


def create_workbook(path: str, sheet_name: str | None = None) -> SpreadsheetWorkbook:
    """Create a new empty workbook in memory, choosing the backend by extension."""
    ext = Path(path).suffix.lower()
    if ext == ".xlsx":
        from .xlsx import XlsxWorkbook
        return XlsxWorkbook.create(sheet_name)
    if ext == ".csv":
        from .csv import CsvWorkbook
        return CsvWorkbook.create(sheet_name)
    if ext == ".ods":
        from .ods import OdsWorkbook
        return OdsWorkbook.create(sheet_name)
    raise ValueError(
        f"Unsupported file format: {ext!r}. "
        f"Supported: {', '.join(sorted(SUPPORTED_EXTENSIONS))}"
    )
