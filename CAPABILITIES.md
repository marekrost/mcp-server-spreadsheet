# mcp-server-xlsx — Proposed Capabilities

MCP server for CSV-like operations on `.xlsx` files.
Preserves existing row/cell styling — only reads and writes **data values**.

The server is **stateless from the caller's perspective.** Every tool call
explicitly specifies the target `file` (path) and, where applicable, `sheet`
(name). There are no handles, no "open/close" lifecycle, and no "active"
context to manage — the AI agent orchestrates which file and sheet to work
with across calls.

Internally the server may cache recently-used workbooks in memory for
performance, but this is transparent to the caller.

---

## Workbook Operations

| Tool | Description |
|---|---|
| `list_workbooks` | List all `.xlsx` files in a given directory (non-recursive) |
| `create_workbook` | Create a new empty `.xlsx` file at the given path with an optional initial sheet name. Returns the path |
| `copy_workbook` | Copy an existing `.xlsx` file to a new path |

## Sheet Operations

| Tool | Description |
|---|---|
| `list_sheets` | List all sheet names in `file` |
| `add_sheet` | Add a new sheet to `file` (optional name and position) |
| `rename_sheet` | Rename an existing sheet in `file` |
| `delete_sheet` | Delete a sheet by name from `file` |
| `copy_sheet` | Duplicate a sheet within `file` (optional new name and position) |

## Reading Data

| Tool | Description |
|---|---|
| `read_sheet` | Read entire sheet as a list of rows (optional row/column bounds) |
| `read_cell` | Read value of a single cell, e.g. `B3` |
| `read_range` | Read a rectangular range, e.g. `A1:D10` |
| `get_sheet_dimensions` | Return row count and column count of the used range |

## Writing Data

| Tool | Description |
|---|---|
| `write_cell` | Write a value to a single cell |
| `write_range` | Write a 2D array of values starting at a given cell |
| `append_rows` | Append one or more rows after the last used row |
| `insert_rows` | Insert blank or pre-filled rows at a given position, shifts existing rows down |
| `delete_rows` | Delete rows by index range, shifts remaining rows up |
| `clear_range` | Clear cell values in a range without removing rows/columns |
| `copy_range` | Copy a block of cells to another location (optionally to a different `sheet`) |

## Column Operations

| Tool | Description |
|---|---|
| `insert_columns` | Insert blank columns at a given position |
| `delete_columns` | Delete columns by index range |

## Search / Lookup

| Tool | Description |
|---|---|
| `search_sheet` | Search for a value or regex pattern, return matching cell references |

---

## Table Mode (SQL)

Treats each sheet as a database table (header row = column names, data rows
below) and executes SQL via an embedded DuckDB engine.

| Tool | Description |
|---|---|
| `describe_table` | Return column names, inferred types, row count, and sample values (first 3 rows). Accepts `file`, optional `sheet` (omit to describe all sheets), optional `header_row` (default 1) |
| `sql_query` | Execute a read-only SQL `SELECT` statement. Returns results as a list of `{column: value}` objects |
| `sql_execute` | Execute a mutating SQL statement (`INSERT INTO`, `UPDATE`, `DELETE FROM`). Applies changes back to the Excel file and saves. Returns count of affected rows |

### SQL Examples

```sql
-- Filter and sort
SELECT name, revenue FROM Sales
WHERE status = 'Active' AND revenue > 50000
ORDER BY revenue DESC LIMIT 20

-- Aggregate with grouping
SELECT department, COUNT(*) AS headcount, AVG(salary) AS avg_salary
FROM Employees
GROUP BY department
HAVING COUNT(*) > 5

-- Cross-sheet JOIN
SELECT o.order_id, o.total, c.name, c.email
FROM Orders o
JOIN Customers c ON o.customer_id = c.id
WHERE o.total > 1000

-- Bulk update
UPDATE Sales SET status = 'Closed'
WHERE quarter = 'Q1' AND revenue < 1000

-- Conditional insert from another sheet
INSERT INTO Summary (category, total)
SELECT category, SUM(amount)
FROM Transactions
GROUP BY category

-- Delete matching rows
DELETE FROM Logs WHERE date < '2024-01-01'
```

Sheet names with spaces or special characters must be quoted: `SELECT * FROM "Q1 Sales"`.

---

## Common Parameters

Every sheet-level tool (Reading, Writing, Column, Search) accepts:

| Parameter | Required | Description |
|---|---|---|
| `file` | **yes** | Path to the `.xlsx` file |
| `sheet` | no | Sheet name. Defaults to the **first sheet** in the workbook |

Sheet Operations tools require `file` but address sheets by name in their own parameters.

---

## Design Notes

- **Stateless API.** No handles, no open/close, no "active" workbook or sheet. Every call is self-contained.
- **Two modes, one file.** Free-edit (cell-level) and SQL tools operate on the same workbook and can be interleaved freely.
- **Styling is never modified.** All write operations set cell values only.
- **Auto-save.** Every mutating operation saves the workbook to disk before returning. Writes use an atomic temp-file-and-rename strategy to prevent corruption on crash.
- **Internal caching.** The server may keep recently-used workbooks in memory (LRU) to avoid redundant disk reads. The cache is invisible to callers and has no effect on correctness.
- Cell references use standard Excel notation (`A1`, `$B$2`).
- All row/column indices in tool parameters are **1-based** to match Excel conventions.
- **Data types — as-is inference.** Values are written using the same logic as Excel paste. JSON numbers → numeric cells. Strings starting with `=` → formulas. Strings that parse as numbers or dates → stored as the corresponding type. Everything else → text. Reading returns values in their native type.
- Prefer bulk operations (`write_range`, `append_rows`) over many individual `write_cell` calls to minimize file rewrites.
- **SQL mode** is best for sheets with a clean header+rows structure. Use free-edit for everything else.
- **SQL mutations** (`sql_execute`) rewrite the target sheet's data region after executing. This compacts away any empty rows within the data.
