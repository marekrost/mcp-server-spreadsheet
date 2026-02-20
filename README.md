# mcp-server-xlsx

Data-first MCP server for reading and writing `.xlsx` files.

## Key features

- **Dual mode** — cell-level workbook operations and a DuckDB-powered SQL query engine, interleaved freely on the same file.
- **Workbook essentials** — worksheets, rows, columns, cells, search.
- **Data-only** — preserves existing formatting but only reads and writes values.
- **Stateless** — every call specifies `file` and `sheet` explicitly; no handles or sessions.
- **Atomic saves** — writes go to a temp file, then `os.replace()` into the target path.
- **Formulas preserved** — returns formula strings (e.g. `=SUM(A1:A5)`), not cached results.
- **Type coercion on write** — numeric strings become numbers, `=` prefix becomes a formula, everything else is text.
- **SQL across sheets** — JOINs, GROUP BY, aggregates, subqueries via in-memory DuckDB; mutations write back to the file.

## Requirements

- Python 3.13+
- [uv](https://docs.astral.sh/uv/)

## Installation

```bash
uv sync
```

## Usage

### Standalone (stdio transport)

```bash
uv run main.py
```

### Claude Desktop

Add to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "mcp-server-xlsx": {
      "command": "uv",
      "args": ["run", "--directory", "/path/to/mcp-server-xlsx", "main.py"]
    }
  }
}
```

### Claude Code

Add to your `.mcp.json`:

```json
{
  "mcpServers": {
    "mcp-server-xlsx": {
      "command": "uv",
      "args": ["run", "--directory", "/path/to/mcp-server-xlsx", "main.py"]
    }
  }
}
```

## Tools

### Workbook Operations

| Tool | Description |
|---|---|
| `list_workbooks` | List all `.xlsx` files in a directory (non-recursive) |
| `create_workbook` | Create a new empty `.xlsx` file with an optional initial sheet name |
| `copy_workbook` | Copy an existing `.xlsx` file to a new path |

### Sheet Operations

| Tool | Description |
|---|---|
| `list_sheets` | List all sheet names in a workbook |
| `add_sheet` | Add a new sheet (optional name and position) |
| `rename_sheet` | Rename an existing sheet |
| `delete_sheet` | Delete a sheet by name |
| `copy_sheet` | Duplicate a sheet within a workbook (optional new name and position) |

### Reading Data

| Tool | Description |
|---|---|
| `read_sheet` | Read entire sheet as rows (optional row/column bounds) |
| `read_cell` | Read a single cell value, e.g. `B3` |
| `read_range` | Read a rectangular range, e.g. `A1:D10` |
| `get_sheet_dimensions` | Get row and column count of the used range |

### Writing Data

| Tool | Description |
|---|---|
| `write_cell` | Write a value to a single cell |
| `write_range` | Write a 2D array starting at a given cell |
| `append_rows` | Append rows after the last used row |
| `insert_rows` | Insert blank or pre-filled rows at a position (shifts rows down) |
| `delete_rows` | Delete rows by index (shifts rows up) |
| `clear_range` | Clear values in a range without removing rows/columns |
| `copy_range` | Copy a block of cells to another location (optionally to a different sheet) |

### Column Operations

| Tool | Description |
|---|---|
| `insert_columns` | Insert blank columns at a position |
| `delete_columns` | Delete columns by index |

### Search

| Tool | Description |
|---|---|
| `search_sheet` | Search for a value or regex pattern, returns matching cell references |

### Table Mode (SQL)

| Tool | Description |
|---|---|
| `describe_table` | Inspect column names, inferred types, row count, and sample values |
| `sql_query` | Execute a read-only SQL `SELECT` (supports JOINs across sheets, GROUP BY, aggregates, subqueries) |
| `sql_execute` | Execute `INSERT INTO`, `UPDATE`, or `DELETE FROM` — writes changes back to the file |

SQL examples:

```sql
-- Filter and sort
SELECT name, revenue FROM Sales WHERE status = 'Active' ORDER BY revenue DESC LIMIT 20

-- Cross-sheet JOIN
SELECT o.order_id, c.name FROM Orders o JOIN Customers c ON o.customer_id = c.id

-- Aggregate
SELECT department, COUNT(*) AS n, AVG(salary) AS avg FROM Employees GROUP BY department

-- Mutate
UPDATE Sales SET status = 'Closed' WHERE quarter = 'Q1' AND revenue < 1000
DELETE FROM Logs WHERE date < '2024-01-01'
```

Sheet names with spaces must be quoted: `SELECT * FROM "Q1 Sales"`.

## Common Parameters

Every sheet-level tool accepts:

| Parameter | Required | Description |
|---|---|---|
| `file` | yes | Path to the `.xlsx` file |
| `sheet` | no | Sheet name. Defaults to the first sheet in the workbook |

All row/column indices are **1-based** to match Excel conventions. Cell references use standard Excel notation (`A1`, `$B$2`).
