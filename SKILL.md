---
name: skill-access-explorer
description: "Use this skill ANY TIME a .accdb or .mdb file is mentioned — even casually. This includes schema exploration, form/report extraction, VBA analysis, migration to PostgreSQL/Supabase, or ANY task that touches Access database structure. Do NOT attempt Access work from memory — always load this skill first."
---

# Access Explorer

Extract structural metadata from Microsoft Access databases into JSON for analysis, migration, or web UI generation.

## Platform requirement

Windows only. Requires:
- `pyodbc` + Microsoft Access ODBC Driver (*.mdb, *.accdb)
- `pywin32` (win32com) for forms/VBA/reports — optional, ODBC works without it

## Workflow

1. **Determine scope**: schema only / full extraction / migration / VBA only
2. **Choose engine**: ODBC for schema+data, COM for forms/VBA/reports/ribbons
3. **Connect** (see Connection patterns in `references/schema.md` and `references/forms.md`)
4. **Extract** needed layers into JSON
5. **Output** using the standard JSON structure (see bottom of this file)

## Two extraction engines

| Engine | What it provides | When to use |
|--------|-----------------|-------------|
| ODBC (pyodbc) | Tables, columns, types, PKs, FKs, indexes, row counts, sample data | Schema exploration, data migration, type mapping |
| COM (win32com) | Forms, reports, VBA code, queries, ribbons, DAO metadata | UI extraction, VBA analysis, full migration |

COM connection with error handling and fallback:

```python
try:
    app = win32com.client.Dispatch("Access.Application")
    app.Visible = False
    app.OpenCurrentDatabase(db_path)
except Exception as e:
    # Fallback to ODBC-only mode (no forms/VBA)
    print(f"COM unavailable: {e}. Falling back to ODBC schema only.")
```

## Reference files

- `references/schema.md` — ODBC connection, extraction methods, type mapping
- `references/forms.md` — COM forms, reports, queries, ribbons
- `references/vba.md` — VBA extraction, import/export, sync patterns

Read only the file(s) relevant to your current task.

## Key pitfalls

1. **MSysObjects permission denied** — use ODBC metadata methods or COM/DAO instead of querying MSysObjects directly
2. **PrimaryKeys() returns empty** — fallback to searching indexes for unique index named "PrimaryKey"
3. **COM dates have tzinfo** — always `com_date.replace(tzinfo=None)` before comparing with naive datetime
4. **Calculated fields** — cannot INSERT; exclude from column lists, Access recalculates automatically
5. **USysRibbons access** — may need multiple fallback methods (dbOpenTable, SQL, DAO)
6. **VBA module naming** — document modules use `Form_` / `Report_` prefix, standard modules don't
7. **DB locked by another process** — COM will fail if Access has the file open; close Access first or use ODBC-only mode
8. **.mde/.accde files** — compiled databases, VBA code is NOT extractable

## Out of scope

- Executing business logic inside Access (use VBA directly)
- Creating new tables/forms in Access (separate task)
- Working with .mde/.accde (compiled — VBA inaccessible)
- Runtime Access operations (this skill is for extraction and analysis only)

## JSON output structure

```json
{
  "database": "PRINTAGRAM_DB.accdb",
  "tables": {
    "Orders": {
      "columns": [
        { "name": "idx_Order", "type_name": "COUNTER", "pg_type": "serial", "nullable": false, "ordinal_position": 1 },
        { "name": "id_Order_Customer", "type_name": "INTEGER", "pg_type": "integer", "nullable": true, "ordinal_position": 2 },
        { "name": "Order_Date", "type_name": "DATETIME", "pg_type": "timestamp", "nullable": true, "ordinal_position": 3 },
        { "name": "To_Pay", "type_name": "CURRENCY", "pg_type": "numeric(19,4)", "nullable": true, "ordinal_position": 4 }
      ],
      "primary_keys": [ { "column_name": "idx_Order", "key_seq": 1 } ],
      "foreign_keys": [ { "fk_column": "id_Order_Customer", "pk_table": "Customers", "pk_column": "id_Customer" } ],
      "indexes": [ { "name": "idx_order_date", "unique": false, "columns": ["Order_Date"] } ],
      "row_count": 24510,
      "sample_data": [ { "idx_Order": "1", "id_Order_Customer": "42", "Order_Date": "2024-01-15", "To_Pay": "1500.00" } ]
    }
  },
  "forms": {
    "f_Orders_List": {
      "record_source": "qs_Orders_List", "default_view": "Form", "ribbon_name": "RIBBON_OrdersList",
      "controls": { "header": [], "detail": [], "footer": [] },
      "subforms": [ { "name": "sfProducts", "source_object": "Form.sf_Order_Products", "link_child_fields": "idx_Order", "link_master_fields": "idx_Order" } ],
      "events": { "OnOpen": "[Event Procedure]", "BeforeUpdate": "[Event Procedure]" }
    }
  },
  "reports": {
    "r_Invoice": {
      "record_source": "qs_OrderInvoice", "sections": {}, "group_levels": [],
      "controls": {}, "vba_code": "...", "vba_dependencies": [], "vba_dependent_modules": {}
    }
  },
  "queries": { "qs_Orders_List": { "type": "Select", "sql": "SELECT ... FROM Orders" } },
  "ribbons": { "RIBBON_OrdersList": "<customUI>...</customUI>" }
}
```
