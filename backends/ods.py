"""ODS backend using odfpy."""

from __future__ import annotations

import copy
import os
import tempfile
from pathlib import Path
from typing import Any, Iterator

from odf import text as odf_text
from odf.opendocument import load as _odf_load
from odf.opendocument import OpenDocumentSpreadsheet
from odf.table import Table, TableRow, TableCell, TableColumn

from .base import SpreadsheetSheet, SpreadsheetWorkbook

# odfpy's Table/TableRow/TableCell are factory functions, not classes.
# We identify elements by their qname tuples instead of isinstance().
_TABLE_QNAME = Table().qname
_ROW_QNAME = TableRow().qname
_CELL_QNAME = TableCell().qname
_TEXT_P_QNAME = odf_text.P().qname


def _is_table(node) -> bool:
    return getattr(node, "qname", None) == _TABLE_QNAME


def _is_row(node) -> bool:
    return getattr(node, "qname", None) == _ROW_QNAME


def _is_cell(node) -> bool:
    return getattr(node, "qname", None) == _CELL_QNAME


def _cell_value(cell) -> Any:
    """Extract a typed Python value from an ODS table cell."""
    value_type = cell.getAttribute("valuetype")
    if value_type is None:
        # May contain text with no explicit type
        t = _cell_text(cell)
        return t if t else None

    if value_type == "float":
        raw = cell.getAttribute("value")
        if raw is not None:
            f = float(raw)
            return int(f) if f == int(f) else f
        return None

    if value_type == "string":
        return _cell_text(cell)

    if value_type == "boolean":
        return cell.getAttribute("booleanvalue") == "true"

    if value_type == "date":
        return cell.getAttribute("datevalue")

    # percentage, currency, time — return as string
    return _cell_text(cell) or cell.getAttribute("value")


def _cell_text(cell) -> str | None:
    """Get the plain-text content of a cell."""
    parts: list[str] = []
    for node in cell.childNodes:
        if node.qname == _TEXT_P_QNAME:
            txt = ""
            for child in node.childNodes:
                if hasattr(child, "data"):
                    txt += child.data
                else:
                    # nested elements like <text:s/> (spaces)
                    txt += str(child)
            parts.append(txt)
    return "\n".join(parts) if parts else None


def _make_cell(value: Any) -> TableCell:
    """Create an ODS TableCell from a Python value."""
    if value is None:
        return TableCell()

    if isinstance(value, bool):
        cell = TableCell(valuetype="boolean", booleanvalue=str(value).lower())
        cell.addElement(odf_text.P(text=str(value)))
        return cell

    if isinstance(value, int):
        cell = TableCell(valuetype="float", value=str(value))
        cell.addElement(odf_text.P(text=str(value)))
        return cell

    if isinstance(value, float):
        cell = TableCell(valuetype="float", value=str(value))
        cell.addElement(odf_text.P(text=str(value)))
        return cell

    # everything else as string
    cell = TableCell(valuetype="string")
    cell.addElement(odf_text.P(text=str(value)))
    return cell


def _row_cells(row) -> list:
    """Expand repeated cells into an explicit list."""
    cells = []
    for child in row.childNodes:
        if not _is_cell(child):
            continue
        repeat = child.getAttribute("numbercolumnsrepeated")
        n = int(repeat) if repeat else 1
        # Limit repeats to avoid huge expansion of trailing empty cells
        cells.extend(child if i == 0 else copy.deepcopy(child) for i in range(min(n, 1024)))
    return cells


class OdsSheet(SpreadsheetSheet):
    """Wraps an ODS <table:table> element, with an in-memory cell grid
    for reliable random access and mutation."""

    def __init__(self, table, grid: list[list[Any]]) -> None:
        self._table = table
        self._grid = grid  # list of rows, each a list of values

    @property
    def title(self) -> str:
        return self._table.getAttribute("name")

    @title.setter
    def title(self, value: str) -> None:
        self._table.setAttribute("name", value)

    @property
    def max_row(self) -> int:
        return len(self._grid)

    @property
    def max_column(self) -> int:
        if not self._grid:
            return 0
        return max((len(r) for r in self._grid), default=0)

    def _ensure_size(self, row: int, col: int) -> None:
        while len(self._grid) < row:
            self._grid.append([])
        target = self._grid[row - 1]
        while len(target) < col:
            target.append(None)

    def cell_value(self, row: int, col: int) -> Any:
        if row < 1 or row > len(self._grid):
            return None
        r = self._grid[row - 1]
        if col < 1 or col > len(r):
            return None
        return r[col - 1]

    def set_cell(self, row: int, col: int, value: Any) -> None:
        self._ensure_size(row, col)
        self._grid[row - 1][col - 1] = value

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
            vals: list[Any] = []
            for c in range(c1, c2 + 1):
                vals.append(self.cell_value(r, c))
            yield tuple(vals)

    def append(self, values: list) -> None:
        self._grid.append(list(values))

    def insert_rows(self, row: int, count: int) -> None:
        idx = row - 1
        for _ in range(count):
            self._grid.insert(idx, [])

    def delete_rows(self, row: int, count: int) -> None:
        idx = row - 1
        del self._grid[idx : idx + count]

    def insert_cols(self, col: int, count: int) -> None:
        idx = col - 1
        for r in self._grid:
            for _ in range(count):
                r.insert(idx, None)

    def delete_cols(self, col: int, count: int) -> None:
        idx = col - 1
        for r in self._grid:
            del r[idx : idx + count]


def _grid_from_table(table) -> list[list[Any]]:
    """Read all rows/cells from an ODS table into a 2D Python list."""
    grid: list[list[Any]] = []
    for child in table.childNodes:
        if not _is_row(child):
            continue
        repeat = child.getAttribute("numberrowsrepeated")
        n = int(repeat) if repeat else 1
        cells = _row_cells(child)
        row_vals = [_cell_value(c) for c in cells]
        # Trim trailing Nones
        while row_vals and row_vals[-1] is None:
            row_vals.pop()
        # Only add non-empty rows (limit repeated empty rows)
        if row_vals:
            for _ in range(min(n, 1)):
                grid.append(list(row_vals))
        elif n == 1:
            grid.append([])
    # Trim trailing empty rows
    while grid and not grid[-1]:
        grid.pop()
    return grid


def _rebuild_table(table, grid: list[list[Any]]) -> None:
    """Replace the content of an ODS table element from a 2D grid."""
    # Remove existing rows (keep column definitions)
    to_remove = [ch for ch in table.childNodes if _is_row(ch)]
    for node in to_remove:
        table.removeChild(node)

    for row_data in grid:
        tr = TableRow()
        for val in row_data:
            tr.addElement(_make_cell(val))
        table.addElement(tr)


class OdsWorkbook(SpreadsheetWorkbook):
    """Wraps an ODS document."""

    def __init__(self, doc, sheets: list[OdsSheet]) -> None:
        self._doc = doc
        self._sheets = sheets

    @property
    def sheetnames(self) -> list[str]:
        return [s.title for s in self._sheets]

    @property
    def worksheets(self) -> list[SpreadsheetSheet]:
        return list(self._sheets)

    def get_sheet(self, name: str) -> OdsSheet:
        for s in self._sheets:
            if s.title == name:
                return s
        raise ValueError(
            f"Sheet not found: {name!r}. Available: {self.sheetnames}"
        )

    def create_sheet(
        self, title: str | None = None, index: int | None = None,
    ) -> OdsSheet:
        name = title or f"Sheet{len(self._sheets) + 1}"
        table = Table(name=name)
        sheet = OdsSheet(table, [])
        if index is not None:
            self._sheets.insert(index, sheet)
        else:
            self._sheets.append(sheet)
        self._doc.spreadsheet.addElement(table)
        return sheet

    def delete_sheet(self, name: str) -> None:
        for i, s in enumerate(self._sheets):
            if s.title == name:
                self._doc.spreadsheet.removeChild(s._table)
                del self._sheets[i]
                return
        raise ValueError(f"Sheet not found: {name!r}")

    def copy_sheet(self, source_name: str) -> OdsSheet:
        src = self.get_sheet(source_name)
        new_grid = [list(row) for row in src._grid]
        new_table = Table(name=f"{source_name} Copy")
        _rebuild_table(new_table, new_grid)
        self._doc.spreadsheet.addElement(new_table)
        new_sheet = OdsSheet(new_table, new_grid)
        self._sheets.append(new_sheet)
        return new_sheet

    def move_sheet(self, sheet: SpreadsheetSheet, offset: int) -> None:
        idx = next(
            (i for i, s in enumerate(self._sheets) if s.title == sheet.title),
            None,
        )
        if idx is None:
            raise ValueError(f"Sheet not found: {sheet.title!r}")
        new_idx = max(0, min(len(self._sheets) - 1, idx + offset))
        s = self._sheets.pop(idx)
        self._sheets.insert(new_idx, s)

    def save(self, path: str) -> None:
        # Sync grids back to ODF DOM before saving
        for sheet in self._sheets:
            _rebuild_table(sheet._table, sheet._grid)

        # Rebuild the spreadsheet element's table children in correct order
        tables = [ch for ch in self._doc.spreadsheet.childNodes if _is_table(ch)]
        for t in tables:
            self._doc.spreadsheet.removeChild(t)
        for s in self._sheets:
            self._doc.spreadsheet.addElement(s._table)

        p = Path(path)
        fd, tmp = tempfile.mkstemp(suffix=".ods", dir=p.parent)
        os.close(fd)
        try:
            self._doc.save(tmp)
            os.replace(tmp, p)
        except BaseException:
            os.unlink(tmp)
            raise

    @classmethod
    def load(cls, path: str) -> OdsWorkbook:
        p = Path(path)
        if not p.exists():
            raise ValueError(f"File not found: {path}")
        doc = _odf_load(str(p))
        sheets: list[OdsSheet] = []
        for table in doc.spreadsheet.getElementsByType(Table):
            grid = _grid_from_table(table)
            sheets.append(OdsSheet(table, grid))
        if not sheets:
            # ODS files should have at least one table
            table = Table(name="Sheet1")
            doc.spreadsheet.addElement(table)
            sheets.append(OdsSheet(table, []))
        return cls(doc, sheets)

    @classmethod
    def create(cls, sheet_name: str | None = None) -> OdsWorkbook:
        doc = OpenDocumentSpreadsheet()
        name = sheet_name or "Sheet1"
        table = Table(name=name)
        doc.spreadsheet.addElement(table)
        sheet = OdsSheet(table, [])
        return cls(doc, [sheet])
