# skill-access-explorer

A Claude Code skill for extracting and analyzing Microsoft Access databases — schema, forms, reports, VBA code, queries, and ribbons. Built from production experience migrating a real business system (13+ years on Access) to PostgreSQL/Supabase.

## Problem

Working with Access databases in 2025 means dealing with:

- No standard way to extract schema programmatically
- Forms and reports locked inside proprietary binary format
- VBA code scattered across modules with undocumented dependencies
- No straightforward migration path to modern databases
- COM/ODBC quirks that aren't documented anywhere

## Solution

`access-explorer` gives Claude structured extraction patterns for every layer of an Access database, with real-world pitfalls from production use.

| Layer | Engine | What you get |
|-------|--------|-------------|
| Schema | ODBC | Tables, columns, types, PKs, FKs, indexes, row counts |
| Forms | COM | Controls, layout, events, subforms, ribbon links |
| Reports | COM | Sections, group levels, typography, VBA code |
| VBA | COM/VBE | All modules, dependencies, import/export |
| Queries | DAO | SQL, query types |
| Ribbons | DAO | XML definitions linked to forms |

## Requirements

- Windows only
- Python: `pyodbc`, `pywin32`
- Microsoft Access ODBC Driver (*.mdb, *.accdb)
- Microsoft Access installed (for COM/forms/VBA extraction)

ODBC-only mode works without Access installed — covers schema and data.

## Installation

### Global (all projects)
```bash
mkdir -p ~/.claude/skills/skill-access-explorer
cp -r SKILL.md references ~/.claude/skills/skill-access-explorer/
```

### Per-project
```bash
mkdir -p .claude/skills/skill-access-explorer
cp -r SKILL.md references .claude/skills/skill-access-explorer/
```

## Usage

Once installed, Claude Code automatically loads this skill when you mention `.accdb`, `.mdb`, or Access database tasks:
```
Extract the schema from our production database
Analyze all forms in PRINTAGRAM_DB.accdb
Export VBA modules from the database
Migrate Access table structure to Supabase
```

## Output

All extraction outputs to structured JSON:
```json
{
  "database": "PRINTAGRAM_DB.accdb",
  "tables": {
    "Orders": {
      "columns": ["..."],
      "primary_keys": ["..."],
      "foreign_keys": ["..."],
      "row_count": 24510
    }
  },
  "forms": { },
  "reports": { },
  "queries": { },
  "ribbons": { }
}
```

## Key pitfalls covered

- `MSysObjects` permission denied — use ODBC metadata instead
- `PrimaryKeys()` returns empty — fallback to index scan
- COM dates carry `tzinfo` — strip before comparison
- Calculated fields fail on INSERT — exclude from column lists
- VBA compile errors break component iteration — wrap in try/except
- `.mde/.accde` files — VBA not extractable
- DB locked by another process — COM fails, ODBC-only fallback

## File structure
```
skill-access-explorer/
├── SKILL.md              <- main skill file
└── references/
    ├── schema.md         <- ODBC extraction patterns + type mapping
    ├── forms.md          <- COM forms, reports, queries, ribbons
    └── vba.md            <- VBA extraction, sync patterns, UUID generation
```

## Access to PostgreSQL type mapping

| Access | PostgreSQL |
|--------|-----------|
| COUNTER | serial |
| VARCHAR/CHAR | varchar(N) / text |
| LONGCHAR/MEMO | text |
| CURRENCY | numeric(19,4) |
| DATETIME | timestamp |
| BIT/YESNO | boolean |
| GUID | uuid |

## Background

Built while migrating Printagram — a print and advertising company operating on Access since 2012 — to a modern Python + Supabase stack. The patterns here come from extracting 40+ tables, 30+ forms, and years of VBA business logic from a live production database.

## License

MIT
