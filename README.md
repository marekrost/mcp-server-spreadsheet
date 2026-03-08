# mcp-server-spreadsheet

mcp-name: io.github.marekrost/mcp-server-spreadsheet

Data-first MCP server for reading and writing spreadsheet files (`.xlsx`, `.csv`, `.ods`).

## Key features

- **Multi-format** — works with Excel (`.xlsx`), CSV (`.csv`), and OpenDocument (`.ods`) files through a unified tool interface.
- **Dual mode** — cell-level workbook operations and a DuckDB-powered SQL query engine, interleaved freely on the same file.
- **Workbook essentials** — worksheets, rows, columns, cells, search.
- **Data-only** — preserves existing formatting but only reads and writes values.
- **Stateless** — every call specifies `file` and `sheet` explicitly; no handles or sessions.
- **Atomic saves** — writes go to a temp file, then `os.replace()` into the target path.
- **Type coercion on write** — numeric strings become numbers, everything else is text.
- **SQL across sheets** — JOINs, GROUP BY, aggregates, subqueries via in-memory DuckDB; mutations write back to the file.
- **CSV as single-sheet workbook** — CSV files are treated as a workbook with one sheet named `default`.

## Requirements

- Python 3.13+

## Installation

### From PyPI (recommended)

No local checkout needed — just configure your MCP client (see below).

### From source (for development)

```bash
git clone https://github.com/marekrost/mcp-server-spreadsheet.git
cd mcp-server-spreadsheet
uv sync
```

## Usage

### Claude Desktop

Add to your `claude_desktop_config.json`:

**Using PyPI (recommended):**

```json
{
  "mcpServers": {
    "mcp-server-spreadsheet": {
      "command": "uvx",
      "args": ["mcp-server-spreadsheet"]
    }
  }
}
```

**Using local source:**

```json
{
  "mcpServers": {
    "mcp-server-spreadsheet": {
      "command": "uv",
      "args": ["run", "--directory", "/path/to/mcp-server-spreadsheet", "main.py"]
    }
  }
}
```

### Claude Code

Add to your `.mcp.json`:

**Using PyPI (recommended):**

```json
{
  "mcpServers": {
    "mcp-server-spreadsheet": {
      "command": "uvx",
      "args": ["mcp-server-spreadsheet"]
    }
  }
}
```

**Using local source:**

```json
{
  "mcpServers": {
    "mcp-server-spreadsheet": {
      "command": "uv",
      "args": ["run", "--directory", "/path/to/mcp-server-spreadsheet", "main.py"]
    }
  }
}
```

### Standalone (stdio transport)

```bash
# PyPI
uvx mcp-server-spreadsheet

# Local source
uv run main.py
```

## Format notes

| Format | Sheets | Formulas | Types |
|---|---|---|---|
| `.xlsx` | Multiple | Preserved as strings | Native (int, float, date, bool) |
| `.ods` | Multiple | Not preserved | Native (int, float, date, bool) |
| `.csv` | Single (`default`) | N/A | Inferred on load (int, float, text) |

Sheet management tools (`add_sheet`, `delete_sheet`, `copy_sheet`) raise an error for CSV files.

## Tools

### Workbook Operations

| Tool | Description |
|---|---|
| `list_workbooks` | List all spreadsheet files in a directory (non-recursive) |
| `create_workbook_file` | Create a new empty spreadsheet file (format by extension) |
| `copy_workbook` | Copy an existing file to a new path |

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
| `file` | yes | Path to the spreadsheet file (.xlsx, .csv, or .ods) |
| `sheet` | no | Sheet name. Defaults to the first sheet in the workbook |

All row/column indices are **1-based**. Cell references use A1 notation (`A1`, `$B$2`).
