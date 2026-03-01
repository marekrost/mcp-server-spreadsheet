"""Abstract base classes and utilities for spreadsheet backends."""

from __future__ import annotations

import re
from abc import ABC, abstractmethod
from typing import Any, Iterator


# ---------------------------------------------------------------------------
# Cell-reference utilities (no openpyxl dependency)
# ---------------------------------------------------------------------------

def column_index_from_string(col: str) -> int:
    """Convert column letter(s) to 1-based index. 'A'→1, 'Z'→26, 'AA'→27."""
    col = col.upper()
    result = 0
    for ch in col:
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result


def get_column_letter(idx: int) -> str:
    """Convert 1-based column index to letter(s). 1→'A', 26→'Z', 27→'AA'."""
    if idx < 1:
        raise ValueError(f"Column index must be >= 1, got {idx}")
    result = []
    while idx:
        idx, rem = divmod(idx - 1, 26)
        result.append(chr(rem + ord("A")))
    return "".join(reversed(result))


def parse_cell(cell: str) -> tuple[int, int]:
    """Parse cell reference like 'B3' or '$B$3' into (row, col) 1-based ints."""
    cell = cell.replace("$", "")
    m = re.match(r"^([A-Za-z]+)(\d+)$", cell)
    if not m:
        raise ValueError(f"Invalid cell reference: {cell!r}")
    col = column_index_from_string(m.group(1))
    row = int(m.group(2))
    return row, col


def parse_range(range_str: str) -> tuple[int, int, int, int]:
    """Parse range like 'A1:D10' into (min_col, min_row, max_col, max_row)."""
    range_str = range_str.replace("$", "")
    m = re.match(r"^([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)$", range_str)
    if not m:
        raise ValueError(f"Invalid range: {range_str!r}")
    min_col = column_index_from_string(m.group(1))
    min_row = int(m.group(2))
    max_col = column_index_from_string(m.group(3))
    max_row = int(m.group(4))
    return min_col, min_row, max_col, max_row


def coerce_value(value: Any) -> Any:
    """Coerce string values: numeric strings → numbers, '=' prefix → kept as-is."""
    if value is None or not isinstance(value, str):
        return value
    if value.startswith("="):
        return value
    try:
        if "." in value or "e" in value.lower():
            return float(value)
        return int(value)
    except ValueError:
        return value


# ---------------------------------------------------------------------------
# Abstract Sheet
# ---------------------------------------------------------------------------

class SpreadsheetSheet(ABC):
    """Abstract interface for a single sheet in a spreadsheet workbook."""

    @property
    @abstractmethod
    def title(self) -> str: ...

    @title.setter
    @abstractmethod
    def title(self, value: str) -> None: ...

    @property
    @abstractmethod
    def max_row(self) -> int:
        """1-based index of the last used row, or 0 if empty."""
        ...

    @property
    @abstractmethod
    def max_column(self) -> int:
        """1-based index of the last used column, or 0 if empty."""
        ...

    @abstractmethod
    def cell_value(self, row: int, col: int) -> Any:
        """Get the value at (row, col), both 1-based. Returns None for empty."""
        ...

    @abstractmethod
    def set_cell(self, row: int, col: int, value: Any) -> None:
        """Set the value at (row, col), both 1-based."""
        ...

    @abstractmethod
    def iter_rows(
        self,
        *,
        min_row: int | None = None,
        max_row: int | None = None,
        min_col: int | None = None,
        max_col: int | None = None,
    ) -> Iterator[tuple]:
        """Yield tuples of cell values for the given rectangular region.

        Bounds are 1-based and inclusive. None means "use sheet extent".
        """
        ...

    @abstractmethod
    def append(self, values: list) -> None:
        """Append a row of values after the last used row."""
        ...

    @abstractmethod
    def insert_rows(self, row: int, count: int) -> None:
        """Insert *count* blank rows before *row* (1-based)."""
        ...

    @abstractmethod
    def delete_rows(self, row: int, count: int) -> None:
        """Delete *count* rows starting at *row* (1-based)."""
        ...

    @abstractmethod
    def insert_cols(self, col: int, count: int) -> None:
        """Insert *count* blank columns before *col* (1-based)."""
        ...

    @abstractmethod
    def delete_cols(self, col: int, count: int) -> None:
        """Delete *count* columns starting at *col* (1-based)."""
        ...


# ---------------------------------------------------------------------------
# Abstract Workbook
# ---------------------------------------------------------------------------

class SpreadsheetWorkbook(ABC):
    """Abstract interface for a spreadsheet workbook (one or more sheets)."""

    @property
    @abstractmethod
    def sheetnames(self) -> list[str]: ...

    @property
    @abstractmethod
    def worksheets(self) -> list[SpreadsheetSheet]: ...

    @abstractmethod
    def get_sheet(self, name: str) -> SpreadsheetSheet:
        """Return the sheet with the given name, or raise ValueError."""
        ...

    @abstractmethod
    def create_sheet(
        self, title: str | None = None, index: int | None = None,
    ) -> SpreadsheetSheet:
        """Create a new sheet. May raise for single-sheet formats (CSV)."""
        ...

    @abstractmethod
    def delete_sheet(self, name: str) -> None:
        """Delete a sheet by name. May raise for single-sheet formats."""
        ...

    @abstractmethod
    def copy_sheet(self, source_name: str) -> SpreadsheetSheet:
        """Duplicate a sheet. May raise for single-sheet formats."""
        ...

    @abstractmethod
    def move_sheet(self, sheet: SpreadsheetSheet, offset: int) -> None:
        """Move a sheet by *offset* positions. May raise for single-sheet formats."""
        ...

    @abstractmethod
    def save(self, path: str) -> None:
        """Persist the workbook to *path* atomically."""
        ...

    @classmethod
    @abstractmethod
    def load(cls, path: str) -> SpreadsheetWorkbook:
        """Load an existing file from *path*."""
        ...

    @classmethod
    @abstractmethod
    def create(cls, sheet_name: str | None = None) -> SpreadsheetWorkbook:
        """Create a new empty workbook in memory."""
        ...
