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
