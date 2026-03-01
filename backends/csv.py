"""CSV backend — treats a CSV file as a single-sheet workbook."""

from __future__ import annotations

import csv
import os
import tempfile
from pathlib import Path
from typing import Any, Iterator

from .base import SpreadsheetSheet, SpreadsheetWorkbook

_DEFAULT_SHEET = "default"


def _coerce_csv_value(value: str) -> Any:
    """Coerce a raw CSV string into int/float when possible."""
    if value == "":
        return None
    try:
        if "." in value or "e" in value.lower():
            return float(value)
        return int(value)
    except ValueError:
        return value


class CsvSheet(SpreadsheetSheet):
    """In-memory 2D grid representing the single sheet of a CSV file."""

    def __init__(self, rows: list[list[Any]], title: str = _DEFAULT_SHEET) -> None:
        self._rows = rows
        self._title = title

    @property
    def title(self) -> str:
        return self._title

    @title.setter
    def title(self, value: str) -> None:
        self._title = value

    @property
    def max_row(self) -> int:
        return len(self._rows)

    @property
    def max_column(self) -> int:
        if not self._rows:
            return 0
        return max(len(r) for r in self._rows)

    def _ensure_size(self, row: int, col: int) -> None:
        """Expand grid so that (row, col) is addressable (1-based)."""
        while len(self._rows) < row:
            self._rows.append([])
        target_row = self._rows[row - 1]
        while len(target_row) < col:
            target_row.append(None)

    def cell_value(self, row: int, col: int) -> Any:
        if row < 1 or row > len(self._rows):
            return None
        r = self._rows[row - 1]
        if col < 1 or col > len(r):
            return None
        return r[col - 1]

    def set_cell(self, row: int, col: int, value: Any) -> None:
        self._ensure_size(row, col)
        self._rows[row - 1][col - 1] = value

    def iter_rows(
        self,
        *,
        min_row: int | None = None,
        max_row: int | None = None,
        min_col: int | None = None,
        max_col: int | None = None,
    ) -> Iterator[tuple]:
        r1 = min_row or 1
        r2 = max_row or self.max_row
        c1 = min_col or 1
        c2 = max_col or self.max_column
        for r in range(r1, r2 + 1):
            row_vals: list[Any] = []
            for c in range(c1, c2 + 1):
                row_vals.append(self.cell_value(r, c))
            yield tuple(row_vals)

    def append(self, values: list) -> None:
        self._rows.append(list(values))

    def insert_rows(self, row: int, count: int) -> None:
        idx = row - 1
        for _ in range(count):
            self._rows.insert(idx, [])

    def delete_rows(self, row: int, count: int) -> None:
        idx = row - 1
        del self._rows[idx : idx + count]

    def insert_cols(self, col: int, count: int) -> None:
        idx = col - 1
        for r in self._rows:
            for _ in range(count):
                r.insert(idx, None)

    def delete_cols(self, col: int, count: int) -> None:
        idx = col - 1
        for r in self._rows:
            del r[idx : idx + count]


class CsvWorkbook(SpreadsheetWorkbook):
    """A CSV file exposed as a single-sheet workbook."""

    def __init__(self, sheet: CsvSheet) -> None:
        self._sheet = sheet

    @property
    def sheetnames(self) -> list[str]:
        return [self._sheet.title]

    @property
    def worksheets(self) -> list[SpreadsheetSheet]:
        return [self._sheet]

    def get_sheet(self, name: str) -> CsvSheet:
        if name != self._sheet.title:
            raise ValueError(
                f"Sheet not found: {name!r}. "
                f"CSV files have a single sheet: {self._sheet.title!r}"
            )
        return self._sheet

    def create_sheet(
        self, title: str | None = None, index: int | None = None,
    ) -> SpreadsheetSheet:
        raise ValueError("CSV files support only a single sheet")

    def delete_sheet(self, name: str) -> None:
        raise ValueError("CSV files support only a single sheet")

    def copy_sheet(self, source_name: str) -> SpreadsheetSheet:
        raise ValueError("CSV files support only a single sheet")

    def move_sheet(self, sheet: SpreadsheetSheet, offset: int) -> None:
        raise ValueError("CSV files support only a single sheet")

    def save(self, path: str) -> None:
        p = Path(path)
        fd, tmp = tempfile.mkstemp(suffix=".csv", dir=p.parent)
        os.close(fd)
        try:
            with open(tmp, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                for row in self._sheet._rows:
                    writer.writerow(
                        ["" if v is None else v for v in row]
                    )
            os.replace(tmp, p)
        except BaseException:
            os.unlink(tmp)
            raise

    @classmethod
    def load(cls, path: str) -> CsvWorkbook:
        p = Path(path)
        if not p.exists():
            raise ValueError(f"File not found: {path}")
        with open(p, newline="", encoding="utf-8") as f:
            reader = csv.reader(f)
            rows = [
                [_coerce_csv_value(cell) for cell in row]
                for row in reader
            ]
        return cls(CsvSheet(rows))

    @classmethod
    def create(cls, sheet_name: str | None = None) -> CsvWorkbook:
        title = sheet_name or _DEFAULT_SHEET
        return cls(CsvSheet([], title=title))
