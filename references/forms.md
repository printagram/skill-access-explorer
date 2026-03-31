# Forms & Reports Extraction (COM)

## COM connection

See main SKILL.md for connection pattern with error handling.

## Safe property access

COM objects may throw on deleted/hidden properties:

```python
def safe_getattr(obj, name, default=None):
    try:
        val = getattr(obj, name, default)
        return default if val in ("", None) else val
    except Exception:
        return default
```

## DAO engine (metadata without Access UI)

```python
engine = win32com.client.Dispatch("DAO.DBEngine.120")
db = engine.OpenDatabase(db_path)
# db.TableDefs, db.QueryDefs
# td.DateCreated, td.LastUpdated — pywintypes.datetime with tzinfo
# IMPORTANT: strip tzinfo: com_date.replace(tzinfo=None)
db.Close()
```

## Forms

Open in Design mode:

```python
app.DoCmd.OpenForm(form_name, 1)  # acDesign = 1
form = app.Forms(form_name)
```

### Form properties

RecordSource, Filter, FilterOn, OrderBy, OrderByOn, DefaultView, AllowAdditions, AllowDeletions, AllowEdits, DataEntry, NavigationButtons, RecordSelectors, RibbonName, Width, InsideHeight.

### Form events

OnOpen, OnClose, OnLoad, OnUnload, BeforeUpdate, AfterUpdate, OnCurrent, BeforeInsert, AfterInsert, BeforeDelConfirm, AfterDelConfirm, OnTimer.

### Controls

Iterate `form.Controls`, group by section (0=detail, 1=header, 2=footer).

| Type ID | Type | Extra properties |
|---------|------|------------------|
| 100 | Label | Caption |
| 104 | CommandButton | Caption, OnClick |
| 109 | TextBox | ControlSource, Format, DefaultValue, InputMask |
| 110 | ListBox | RowSource, RowSourceType, ColumnCount, ColumnWidths, BoundColumn |
| 111 | ComboBox | RowSource, RowSourceType, ColumnCount, ColumnWidths, BoundColumn |
| 112 | Subform | SourceObject, LinkChildFields, LinkMasterFields |
| 123 | TabControl | Pages collection |
| 124 | Page | Caption (tab label) |

Common properties: name, type, type_id, control_source, left, top, width, height, visible, enabled, locked, label, events.

### Subform detection

Type 112 links parent-child via LinkChildFields / LinkMasterFields. Strip "Form." prefix from SourceObject.

## Reports

```python
app.DoCmd.OpenReport(report_name, 5)  # acDesign = 5
report = app.Reports(report_name)
```

### Sections

0=Detail, 1=ReportHeader, 2=ReportFooter, 3=PageHeader, 4=PageFooter, 5+=GroupHeader/GroupFooter pairs.

Per section: height, visible, back_color, can_grow, can_shrink.

### Group levels

```python
for i in range(report.GroupLevel.Count):
    gl = report.GroupLevel(i)
    # control_source, group_on, group_interval, group_header, group_footer, sort_order, keep_together
```

### Report control typography

Extra vs forms: FontName, FontSize, FontBold, FontItalic, FontUnderline, TextAlign, ForeColor, BackColor, BorderStyle, CanGrow, CanShrink, RunningSum, HideDuplicates.

## Queries

```python
db = app.CurrentDb()
for i in range(db.QueryDefs.Count):
    qd = db.QueryDefs(i)
    # qd.Name, qd.SQL, qd.Type
```

Types: 0=Select, 1=Crosstab, 2=Delete, 3=Update, 4=Append, 5=MakeTable, 6=DDL, 7=SPT, 8=Union.

Skip names starting with `~`.

## Ribbons

Stored in hidden USysRibbons table (RibbonName, RibbonXml):

```python
db = app.CurrentDb()
# Multiple fallback methods
rs = db.OpenRecordset("USysRibbons", 0)                       # dbOpenTable
rs = db.OpenRecordset("SELECT ... FROM [USysRibbons]", 1)     # SQL
rs = db.TableDefs("USysRibbons").OpenRecordset()               # DAO
```

Form's RibbonName property links to ribbon definition.

Ribbon VBA callbacks pattern:
- `rib_*_OnLoad(ribbon As IRibbonUI)` — init
- `btn_*_Action(control As IRibbonControl)` — buttons
- `edt_*_OnChange(control, strText)` — text fields
- `gobjRibbon_*` — global IRibbonUI ref for InvalidateControl
