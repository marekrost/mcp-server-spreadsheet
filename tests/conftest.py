"""Shared fixtures for the tool test suite.

Each test gets a fresh workbook seeded with predictable data so write
operations don't leak between tests. Formats are parametrized via the
`fmt` fixture; tests that don't apply to CSV opt out with `skip_csv`.
"""
from __future__ import annotations

import os
from pathlib import Path

import pytest

from mcp_server_spreadsheet import server

SEED_PEOPLE = [
    ["name", "age", "city"],
    ["Alice", 30, "Prague"],
    ["Bob", 25, "Brno"],
    ["Carol", 40, "Ostrava"],
]

SEED_ORDERS = [
    ["order_id", "name", "amount"],
    [1, "Alice", 100],
    [2, "Bob", 250],
    [3, "Alice", 75],
]


@pytest.fixture(params=[".xlsx", ".csv", ".ods"])
def fmt(request) -> str:
    return request.param


@pytest.fixture
def skip_csv(fmt):
    if fmt == ".csv":
        pytest.skip("not applicable to single-sheet CSV")


@pytest.fixture(autouse=True)
def _clear_root_env(monkeypatch):
    monkeypatch.delenv("MCP_SPREADSHEET_ROOT", raising=False)


@pytest.fixture
def workbook(tmp_path: Path, fmt: str) -> str:
    """Create a workbook seeded with People (and Orders for non-CSV)."""
    path = tmp_path / f"book{fmt}"
    wb = server.create_workbook(str(path), sheet_name="People")
    ws = wb.worksheets[0]
    for row in SEED_PEOPLE:
        ws.append(list(row))

    if fmt != ".csv":
        orders = wb.create_sheet(title="Orders")
        for row in SEED_ORDERS:
            orders.append(list(row))

    wb.save(str(path))
    return str(path)


@pytest.fixture
def empty_path(tmp_path: Path, fmt: str) -> str:
    """A path for a file that doesn't exist yet."""
    return str(tmp_path / f"new{fmt}")
