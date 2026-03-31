# VBA Code Extraction

## VBE access via COM

```python
vb_project = app.VBE.ActiveVBProject
components = vb_project.VBComponents

for i in range(components.Count):
    try:
        comp = components(i)
        name = comp.Name
        comp_type = comp.Type  # 1=Module, 2=Class, 3=UserForm, 100=Document
        if comp.CodeModule.CountOfLines > 0:
            code = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
    except Exception as e:
        print(f"  Skipping component {i}: {e}")
        continue
```

**Pitfall:** If the database has VBA compile errors, `VBComponents` iteration may fail silently or throw on individual components. Always wrap each component access in try/except and continue on error.

## Component types

| Type ID | Name | Extension | Description |
|---------|------|-----------|-------------|
| 1 | Standard Module | .bas | Global functions/procedures |
| 2 | Class Module | .cls | Class definitions |
| 3 | UserForm | — | User forms (not typically exported) |
| 100 | Document Module | .bas | Form/Report code-behind (Form_*, Report_*) |

## Export VBA to files

Write each component to .bas/.cls. File name = component name. Skip 0-line components.

```python
code_module = comp.CodeModule
line_count = code_module.CountOfLines
code = code_module.Lines(1, line_count)
with open(f"{comp.Name}.bas", "w", encoding="utf-8") as f:
    f.write(code)
```

## Import VBA from files

Replace existing code (delete + add, not append):

```python
comp = components.Item(comp_name)
comp.CodeModule.DeleteLines(1, comp.CodeModule.CountOfLines)
comp.CodeModule.AddFromString(new_code)
```

Component must already exist in the database.

## Dependency analysis

Parse VBA code to find cross-module calls:

1. Find `Call FuncName(...)` and standalone `FuncName(...)` patterns
2. Match against `Function X` / `Sub X` definitions in all modules
3. Store dependent module code alongside the calling module

```python
# Pattern: find dependencies for a report's VBA code
for comp_name_prefix in [f"Report_{report_name}", report_name]:
    try:
        component = vb_project.VBComponents(comp_name_prefix)
        code_module = component.CodeModule
        if code_module.CountOfLines > 0:
            vba_code = code_module.Lines(1, code_module.CountOfLines)
    except Exception:
        continue
```

## UUID generation in VBA (sync fields)

```vba
Private Declare PtrSafe Function CoCreateGuid Lib "ole32" (pguid As GUID_TYPE)
Private Declare PtrSafe Function StringFromGUID2 Lib "ole32" (rguid As GUID_TYPE, ByVal lpsz As LongPtr, ByVal cchMax As Long)

Public Function NewSyncId() As String
    Dim g As GUID_TYPE
    CoCreateGuid g
    buf = Space$(40)
    StringFromGUID2 g, StrPtr(buf), 40
    NewSyncId = LCase(Mid(buf, 2, 36))
End Function
```

## Sync field pattern (4 columns per table)

| Field | Type | Default | Role |
|-------|------|---------|------|
| sync_id | Text(36) | NewSyncId() | UUID |
| updated_at | DateTime | Now() | Last modification |
| sync_status | Text(10) | "pending" | pending / synced |
| sync_source | Text(10) | "access" | access / supabase |

## Data transfer patterns

### Insert with auto-increment PK

```python
def insert_record(cursor, table, columns, record):
    cols = ", ".join(f"[{c}]" for c in columns)
    placeholders = ", ".join("?" for _ in columns)
    values = [record[c] for c in columns]
    cursor.execute(f"INSERT INTO [{table}] ({cols}) VALUES ({placeholders})", values)
    cursor.execute("SELECT @@IDENTITY")
    return cursor.fetchone()[0]
```

### Transfer missing records between databases

```python
source_ids = {row[0] for row in source_cur.execute("SELECT [PK] FROM [Table]").fetchall()}
target_ids = {row[0] for row in target_cur.execute("SELECT [PK] FROM [Table]").fetchall()}
missing = source_ids - target_ids
```

### Calculated fields

Access tables may have calculated fields. INSERT fails with "field not updateable" — exclude that column from INSERT list. Access recalculates automatically.
