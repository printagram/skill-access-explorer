# Schema Extraction (ODBC)

## Connection

```python
import pyodbc

# Driver detection with fallback
drivers = ["Microsoft Access Driver (*.mdb, *.accdb)", "Microsoft Access Driver (*.mdb)"]
available = pyodbc.drivers()
driver = next((d for d in drivers if d in available), None)
if not driver:
    driver = next((d for d in available if "access" in d.lower()), None)

conn_str = f"DRIVER={{{driver}}};DBQ={db_path};"
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

# Always close when done
# cursor.close()
# conn.close()
```

## Extraction methods

| Layer | ODBC method | Key fields |
|-------|------------|------------|
| Tables | `cursor.tables(tableType="TABLE")` | table_name |
| Columns | `cursor.columns(table=name)` | column_name, type_name, column_size, decimal_digits, nullable, ordinal_position |
| Primary keys | `cursor.primaryKeys(table=name)` | column_name, key_seq, pk_name |
| Foreign keys | `cursor.foreignKeys(foreignTable=name)` | fk_column, pk_table, pk_column, fk_name |
| Indexes | `cursor.statistics(table=name)` | index_name, non_unique, column_name |
| Row count | `SELECT COUNT(*) FROM [table]` | integer |
| Sample data | `SELECT TOP 3 * FROM [table]` | dict per row |

Filter out system objects: skip names starting with `MSys` or `~`.

## Primary key fallback

`cursor.primaryKeys()` may return empty on some Access ODBC drivers. Fallback: search indexes for a UNIQUE index named "PrimaryKey":

```python
pks = list(cursor.primaryKeys(table=table_name))
if not pks:
    for idx in indexes:
        if idx["unique"] and idx["name"].lower() == "primarykey":
            pks = [{"column_name": c, "key_seq": i+1} for i, c in enumerate(idx["columns"])]
            break
```

## SQL injection prevention

```python
safe_name = table_name.replace("]", "]]")
cursor.execute(f"SELECT COUNT(*) FROM [{safe_name}]")
```

## Binary data in JSON

```python
if isinstance(val, bytes):
    record[col] = f"<binary {len(val)} bytes>"
else:
    record[col] = str(val)
```

## Access-to-PostgreSQL type mapping

| Access type | PostgreSQL type | Notes |
|-------------|----------------|-------|
| COUNTER/AUTOINCREMENT | serial | Auto-increment PK |
| VARCHAR/CHAR | varchar(N) if N<=255, else text | |
| LONGCHAR/MEMO | text | |
| INTEGER/SMALLINT | integer/smallint | |
| DOUBLE/FLOAT | double precision | If column name matches money pattern -> numeric(19,4) |
| CURRENCY | numeric(19,4) | Fixed precision |
| DECIMAL/NUMERIC | numeric(p,s) | From column metadata |
| DATETIME | timestamp | |
| DATE/TIME | date/time | |
| BIT/YESNO | boolean | |
| GUID | uuid | |
| LONGBINARY/IMAGE/OLE | bytea | |

Money column heuristic — force `numeric(19,4)` if column name matches:
`amount|price|cost|sum|total|balance|payment|fee|discount|tax|salary|wage|credit|debit`
