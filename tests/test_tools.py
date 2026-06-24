"""End-to-end tests for every MCP tool exposed by the server.

Tools are imported from `mcp_server_spreadsheet.server` and called as
plain Python functions. Each test runs against .xlsx, .csv, and .ods
unless the operation doesn't apply.
"""
from __future__ import annotations

import os
from pathlib import Path

import pytest

from mcp_server_spreadsheet import server


# ---------------------------------------------------------------------------
# Workbook operations
# ---------------------------------------------------------------------------

def test_list_workbooks_finds_supported_files(workbook, tmp_path):
    (tmp_path / "ignored.txt").write_text("nope")
    found = server.list_workbooks(str(tmp_path))
    assert workbook in found
    assert all(f.endswith((".xlsx", ".csv", ".ods")) for f in found)


def test_list_workbooks_rejects_missing_dir(tmp_path):
    with pytest.raises(ValueError, match="Not a directory"):
        server.list_workbooks(str(tmp_path / "nope"))


def test_create_workbook_file(empty_path):
    result = server.create_workbook_file(empty_path)
    assert Path(result).exists()
    assert server.list_sheets(empty_path)  # at least one sheet


def test_create_workbook_file_rejects_existing(workbook):
    with pytest.raises(ValueError, match="already exists"):
        server.create_workbook_file(workbook)


def test_copy_workbook(workbook, tmp_path, fmt):
    dst = str(tmp_path / f"copy{fmt}")
    server.copy_workbook(workbook, dst)
    assert Path(dst).exists()
    assert server.read_cell(dst, "A1") == "name"


def test_copy_workbook_rejects_existing_destination(workbook, tmp_path, fmt):
    dst = str(tmp_path / f"copy{fmt}")
    server.copy_workbook(workbook, dst)
    with pytest.raises(ValueError, match="already exists"):
        server.copy_workbook(workbook, dst)


def test_copy_workbook_rejects_missing_source(tmp_path, fmt):
    src = str(tmp_path / f"missing{fmt}")
    dst = str(tmp_path / f"dst{fmt}")
    with pytest.raises(ValueError, match="Source not found"):
        server.copy_workbook(src, dst)


# ---------------------------------------------------------------------------
# Sheet operations
# ---------------------------------------------------------------------------

def test_list_sheets(workbook, fmt):
    sheets = server.list_sheets(workbook)
    if fmt == ".csv":
        assert sheets == ["default"]
    else:
        assert "People" in sheets and "Orders" in sheets


def test_add_sheet(workbook, skip_csv):
    name = server.add_sheet(workbook, name="Extra")
    assert name == "Extra"
    assert "Extra" in server.list_sheets(workbook)


def test_add_sheet_at_position(workbook, skip_csv):
    server.add_sheet(workbook, name="First", position=1)
    assert server.list_sheets(workbook)[0] == "First"


def test_rename_sheet(workbook, skip_csv):
    server.rename_sheet(workbook, "People", "Humans")
    assert "Humans" in server.list_sheets(workbook)
    assert "People" not in server.list_sheets(workbook)


def test_delete_sheet(workbook, skip_csv):
    server.delete_sheet(workbook, "Orders")
    assert "Orders" not in server.list_sheets(workbook)


def test_copy_sheet(workbook, skip_csv):
    new = server.copy_sheet(workbook, "People", new_name="PeopleCopy")
    assert new == "PeopleCopy"
    assert server.read_cell(workbook, "A1", sheet="PeopleCopy") == "name"


# ---------------------------------------------------------------------------
# Reading data
# ---------------------------------------------------------------------------

def test_read_sheet_full(workbook):
    rows = server.read_sheet(workbook)
    assert rows[0] == ["name", "age", "city"]
    assert rows[1][0] == "Alice"


def test_read_sheet_with_bounds(workbook):
    rows = server.read_sheet(
        workbook, start_row=2, end_row=3, start_column=1, end_column=2
    )
    assert rows == [["Alice", 30], ["Bob", 25]]


def test_read_cell(workbook):
    assert server.read_cell(workbook, "A1") == "name"
    assert server.read_cell(workbook, "B2") == 30


def test_read_range(workbook):
    rows = server.read_range(workbook, "A1:B2")
    assert rows == [["name", "age"], ["Alice", 30]]


def test_get_sheet_dimensions(workbook):
    dims = server.get_sheet_dimensions(workbook)
    assert dims == {"rows": 4, "columns": 3}


# ---------------------------------------------------------------------------
# Writing data
# ---------------------------------------------------------------------------

def test_write_cell(workbook):
    server.write_cell(workbook, "B2", 99)
    assert server.read_cell(workbook, "B2") == 99


def test_write_cell_coerces_numeric_string(workbook):
    server.write_cell(workbook, "B2", "42")
    assert server.read_cell(workbook, "B2") == 42


def test_write_range(workbook):
    server.write_range(workbook, "E1", [["x", "y"], [1, 2]])
    assert server.read_range(workbook, "E1:F2") == [["x", "y"], [1, 2]]


def test_append_rows(workbook):
    server.append_rows(workbook, [["Dave", 50, "Plzen"]])
    rows = server.read_sheet(workbook)
    assert rows[-1] == ["Dave", 50, "Plzen"]


def test_insert_rows_blank(workbook):
    before = server.read_sheet(workbook)
    server.insert_rows(workbook, row=2, count=1)
    after = server.read_sheet(workbook)
    assert len(after) == len(before) + 1
    # original row 2 (Alice) is now at row 3
    assert after[2][0] == "Alice"


def test_insert_rows_with_data(workbook):
    server.insert_rows(workbook, row=2, count=1, data=[["Zed", 1, "Z"]])
    assert server.read_sheet(workbook)[1] == ["Zed", 1, "Z"]


def test_delete_rows(workbook):
    server.delete_rows(workbook, row=2, count=1)
    rows = server.read_sheet(workbook)
    assert rows[1][0] == "Bob"  # Alice was at row 2


def test_clear_range(workbook):
    server.clear_range(workbook, "A2:C2")
    rows = server.read_sheet(workbook)
    assert rows[1] == [None, None, None]


def test_copy_range_same_sheet(workbook):
    server.copy_range(workbook, "A1:C1", "E1")
    assert server.read_range(workbook, "E1:G1") == [["name", "age", "city"]]


def test_copy_range_across_sheets(workbook, skip_csv):
    server.copy_range(
        workbook, "A1:C1", "A5", sheet="People", dest_sheet="Orders"
    )
    assert server.read_range(workbook, "A5:C5", sheet="Orders") == [
        ["name", "age", "city"]
    ]


# ---------------------------------------------------------------------------
# Column operations
# ---------------------------------------------------------------------------

def test_insert_columns(workbook):
    server.insert_columns(workbook, column=2, count=1)
    assert server.read_cell(workbook, "A1") == "name"
    assert server.read_cell(workbook, "B1") is None
    assert server.read_cell(workbook, "C1") == "age"


def test_delete_columns(workbook):
    server.delete_columns(workbook, column=2, count=1)
    assert server.read_cell(workbook, "A1") == "name"
    assert server.read_cell(workbook, "B1") == "city"


# ---------------------------------------------------------------------------
# Search
# ---------------------------------------------------------------------------

def test_search_sheet_literal(workbook):
    results = server.search_sheet(workbook, "Alice")
    assert any(r["cell"] == "A2" and r["value"] == "Alice" for r in results)


def test_search_sheet_regex(workbook):
    results = server.search_sheet(workbook, r"^Bo")
    assert len(results) == 1
    assert results[0]["value"] == "Bob"


def test_search_sheet_no_match(workbook):
    assert server.search_sheet(workbook, "nothing_matches_xyz") == []


# ---------------------------------------------------------------------------
# Table mode — SQL
# ---------------------------------------------------------------------------

def test_describe_table_single_sheet(workbook):
    desc = server.describe_table(workbook, sheet=server.list_sheets(workbook)[0])
    assert isinstance(desc, dict)
    names = [c["name"] for c in desc["columns"]]
    assert names == ["name", "age", "city"]
    assert desc["row_count"] == 3


def test_describe_table_all_sheets(workbook, skip_csv):
    desc = server.describe_table(workbook)
    assert isinstance(desc, list)
    titles = {d["sheet"] for d in desc}
    assert {"People", "Orders"} <= titles


def test_sql_query_select(workbook, fmt):
    table = "default" if fmt == ".csv" else "People"
    rows = server.sql_query(
        workbook, f'SELECT name FROM "{table}" WHERE age > 25 ORDER BY name'
    )
    assert [r["name"] for r in rows] == ["Alice", "Carol"]


def test_sql_query_rejects_mutation(workbook, fmt):
    table = "default" if fmt == ".csv" else "People"
    with pytest.raises(ValueError, match="only accepts SELECT"):
        server.sql_query(workbook, f'DELETE FROM "{table}"')


def test_sql_query_cross_sheet_join(workbook, skip_csv):
    rows = server.sql_query(
        workbook,
        'SELECT o.order_id, p.city FROM "Orders" o '
        'JOIN "People" p ON o.name = p.name ORDER BY o.order_id',
    )
    assert rows[0] == {"order_id": 1, "city": "Prague"}
    assert len(rows) == 3


def test_sql_execute_update(workbook, fmt):
    table = "default" if fmt == ".csv" else "People"
    result = server.sql_execute(
        workbook, f"UPDATE \"{table}\" SET city = 'X' WHERE name = 'Alice'"
    )
    assert result["affected_rows"] == 1
    rows = server.sql_query(
        workbook, f'SELECT city FROM "{table}" WHERE name = \'Alice\''
    )
    assert rows == [{"city": "X"}]


def test_sql_execute_delete(workbook, fmt):
    table = "default" if fmt == ".csv" else "People"
    server.sql_execute(workbook, f"DELETE FROM \"{table}\" WHERE name = 'Bob'")
    rows = server.sql_query(workbook, f'SELECT name FROM "{table}"')
    assert "Bob" not in [r["name"] for r in rows]


def test_sql_execute_insert(workbook, fmt):
    table = "default" if fmt == ".csv" else "People"
    server.sql_execute(
        workbook,
        f"INSERT INTO \"{table}\" VALUES ('Eve', 22, 'Liberec')",
    )
    rows = server.sql_query(workbook, f'SELECT name FROM "{table}"')
    assert "Eve" in [r["name"] for r in rows]


def _seed_offset_workbook(tmp_path, fmt):
    """Workbook where headers/data don't start at row 1.

    People: headers at row 3, data starts row 4.
    Orders (non-CSV): header at row 1, units row at 2, data starts row 3.
    """
    path = tmp_path / f"offset{fmt}"
    wb = server.create_workbook(str(path), sheet_name="People")
    ws = wb.worksheets[0]
    ws.append(["Quarterly People Report", None, None])
    ws.append([None, None, None])
    ws.append(["name", "age", "city"])
    ws.append(["Alice", 30, "Prague"])
    ws.append(["Bob", 25, "Brno"])
    ws.append(["Carol", 40, "Ostrava"])

    if fmt != ".csv":
        orders = wb.create_sheet(title="Orders")
        orders.append(["order_id", "name", "amount"])
        orders.append(["int", "text", "USD"])
        orders.append([1, "Alice", 100])
        orders.append([2, "Bob", 250])
        orders.append([3, "Alice", 75])

    wb.save(str(path))
    return str(path)


def test_describe_table_with_offset_header(tmp_path, fmt):
    path = _seed_offset_workbook(tmp_path, fmt)
    table = "default" if fmt == ".csv" else "People"
    desc = server.describe_table(path, sheet=table, header_row=3)
    assert [c["name"] for c in desc["columns"]] == ["name", "age", "city"]
    assert desc["row_count"] == 3


def test_sql_query_per_sheet_header_row(tmp_path, fmt):
    if fmt == ".csv":
        pytest.skip("single-sheet")
    path = _seed_offset_workbook(tmp_path, fmt)
    rows = server.sql_query(
        path,
        'SELECT o.order_id, p.city FROM "Orders" o '
        'JOIN "People" p ON o.name = p.name ORDER BY o.order_id',
        header_row={"People": 3, "Orders": 1},
        data_start_row={"Orders": 3},
    )
    assert rows[0] == {"order_id": 1, "city": "Prague"}
    assert len(rows) == 3


def test_sql_execute_update_with_offset(tmp_path, fmt):
    path = _seed_offset_workbook(tmp_path, fmt)
    table = "default" if fmt == ".csv" else "People"
    result = server.sql_execute(
        path,
        f"UPDATE \"{table}\" SET city = 'X' WHERE name = 'Alice'",
        header_row=3,
    )
    assert result["affected_rows"] == 1
    rows = server.sql_query(
        path,
        f'SELECT city FROM "{table}" WHERE name = \'Alice\'',
        header_row=3,
    )
    assert rows == [{"city": "X"}]


def test_sql_query_header_and_data_both_offset(tmp_path, fmt):
    """Header at row 4, data starts row 6 (gap of one units row at row 5)."""
    path = tmp_path / f"gap{fmt}"
    wb = server.create_workbook(str(path), sheet_name="People")
    ws = wb.worksheets[0]
    ws.append(["Report title", None, None])
    ws.append([None, None, None])
    ws.append(["generated 2026-06-24", None, None])
    ws.append(["name", "age", "city"])
    ws.append(["str", "int", "str"])
    ws.append(["Alice", 30, "Prague"])
    ws.append(["Bob", 25, "Brno"])
    wb.save(str(path))

    table = "default" if fmt == ".csv" else "People"
    desc = server.describe_table(
        str(path), sheet=table, header_row=4, data_start_row=6,
    )
    assert [c["name"] for c in desc["columns"]] == ["name", "age", "city"]
    assert desc["row_count"] == 2
    # The units row must not leak in as data
    assert desc["sample"][0] == {"name": "Alice", "age": 30, "city": "Prague"}

    rows = server.sql_query(
        str(path),
        f'SELECT name, age FROM "{table}" ORDER BY age',
        header_row=4,
        data_start_row=6,
    )
    assert rows == [{"name": "Bob", "age": 25}, {"name": "Alice", "age": 30}]


def test_sql_execute_offset_preserves_pre_header_rows(tmp_path, fmt):
    """Writing back after sql_execute must not clobber rows above header_row."""
    if fmt == ".csv":
        pytest.skip("CSV has no concept of preserved pre-header content on reload")
    path = _seed_offset_workbook(tmp_path, fmt)
    server.sql_execute(
        path,
        "DELETE FROM \"People\" WHERE name = 'Bob'",
        header_row=3,
    )
    wb = server.load_workbook(path)
    ws = wb.worksheets[0]
    assert ws.cell_value(1, 1) == "Quarterly People Report"
    assert ws.cell_value(3, 1) == "name"


# ---------------------------------------------------------------------------
# Quote-stripping fallback (workaround for buggy MCP clients, e.g. OpenCode)
# ---------------------------------------------------------------------------

def test_sheet_name_double_quoted_fallback(workbook, fmt):
    """A double-quoted sheet name resolves to the unquoted sheet when no
    literal match exists."""
    sheet_name = "default" if fmt == ".csv" else "People"
    val = server.read_cell(workbook, "A1", sheet=f'"{sheet_name}"')
    assert val == "name"


def test_sheet_name_single_quoted_fallback(workbook, fmt):
    sheet_name = "default" if fmt == ".csv" else "People"
    val = server.read_cell(workbook, "A1", sheet=f"'{sheet_name}'")
    assert val == "name"


def test_sheet_name_unquoted_still_works(workbook, fmt):
    sheet_name = "default" if fmt == ".csv" else "People"
    assert server.read_cell(workbook, "A1", sheet=sheet_name) == "name"


def test_sheet_name_missing_still_errors(workbook):
    with pytest.raises(ValueError):
        server.read_cell(workbook, "A1", sheet="NoSuchSheet")


def test_sheet_name_missing_quoted_still_errors(workbook):
    with pytest.raises(ValueError):
        server.read_cell(workbook, "A1", sheet='"NoSuchSheet"')


def test_exact_quoted_name_wins_over_fallback(tmp_path):
    """Security-critical: when a sheet literally named '"Summary"' exists,
    the fallback must NOT silently redirect to a different sheet named
    'Summary'. Exact match always wins."""
    path = tmp_path / "evil.xlsx"
    wb = server.create_workbook(str(path), sheet_name='"Summary"')
    wb.worksheets[0].append(["evil"])
    real = wb.create_sheet(title="Summary")
    real.append(["good"])
    wb.save(str(path))

    # The literal quoted name must read the evil sheet, not the good one.
    assert server.read_cell(str(path), "A1", sheet='"Summary"') == "evil"
    # The unquoted name reads the good sheet.
    assert server.read_cell(str(path), "A1", sheet="Summary") == "good"


def test_rename_sheet_quoted_fallback(workbook, skip_csv):
    server.rename_sheet(workbook, '"People"', "Humans")
    assert "Humans" in server.list_sheets(workbook)


def test_delete_sheet_quoted_fallback(workbook, skip_csv):
    server.delete_sheet(workbook, '"Orders"')
    assert "Orders" not in server.list_sheets(workbook)


def test_copy_sheet_quoted_fallback(workbook, skip_csv):
    new = server.copy_sheet(workbook, '"People"', new_name="PeopleCopy")
    assert new == "PeopleCopy"


def test_file_path_double_quoted_fallback(workbook):
    """A double-quoted file path resolves to the unquoted path when no
    literal file exists at the quoted path."""
    assert server.read_cell(f'"{workbook}"', "A1") == "name"


def test_file_path_unquoted_still_works(workbook):
    assert server.read_cell(workbook, "A1") == "name"


def test_file_path_missing_still_errors(tmp_path):
    with pytest.raises((ValueError, FileNotFoundError)):
        server.read_cell(str(tmp_path / "nope.xlsx"), "A1")


# ---------------------------------------------------------------------------
# Path restriction (MCP_SPREADSHEET_ROOT)
# ---------------------------------------------------------------------------

def test_root_unset_allows_any_path(tmp_path, monkeypatch):
    monkeypatch.delenv("MCP_SPREADSHEET_ROOT", raising=False)
    # Should not raise — passes through.
    server._check_path(str(tmp_path / "anywhere.xlsx"))


def test_root_set_allows_inside(tmp_path, monkeypatch):
    monkeypatch.setenv("MCP_SPREADSHEET_ROOT", str(tmp_path))
    server._check_path(str(tmp_path / "ok.xlsx"))


def test_root_set_rejects_outside(tmp_path, monkeypatch):
    monkeypatch.setenv("MCP_SPREADSHEET_ROOT", str(tmp_path))
    with pytest.raises(ValueError, match="outside the allowed root"):
        server._check_path("/etc/passwd")


def test_root_set_rejects_traversal(tmp_path, monkeypatch):
    monkeypatch.setenv("MCP_SPREADSHEET_ROOT", str(tmp_path))
    with pytest.raises(ValueError, match="outside the allowed root"):
        server._check_path(str(tmp_path / ".." / ".." / "etc" / "passwd"))


def test_root_enforced_by_tool(tmp_path, monkeypatch):
    monkeypatch.setenv("MCP_SPREADSHEET_ROOT", str(tmp_path))
    with pytest.raises(ValueError, match="outside the allowed root"):
        server.list_workbooks("/etc")
