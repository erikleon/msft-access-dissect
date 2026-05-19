# access-dissect

Reverse-engineer, document, and modernize Microsoft Access applications.

## Overview

`access-dissect` is a Windows-native Python tool that extracts everything from an Access `.accdb`/`.mdb` file into a structured catalog, then renders it as documentation and migration artifacts.

**Pipeline:**
1. **Extract** → structured JSON/YAML catalog (requires Windows + Access)
2. **Render** → Markdown docs, HTML report, SQL DDL (runs anywhere)
3. **Analyze** → dependency graph, complexity scores, modernization recommendations

## Requirements

- Windows (extraction requires COM automation)
- Python 3.11+ (64-bit recommended)
- Microsoft Access or [Access Runtime](https://support.microsoft.com/en-us/office/download-and-install-microsoft-access-runtime) installed
- [64-bit ACE redistributable](https://www.microsoft.com/en-us/download/details.aspx?id=54920) if using 64-bit Python without full Office

## Installation

```bash
pip install access-dissect
```

## Usage

```bash
# Quick info without full extraction
access-dissect info myapp.accdb

# Extract to catalog (schema + VBA + forms, no data)
access-dissect extract myapp.accdb

# Extract with data profiling
access-dissect extract myapp.accdb --scope full

# Render catalog to docs
access-dissect render myapp_catalog.json --format markdown --format html --format sql

# Full pipeline: extract + all renders + analysis
access-dissect run myapp.accdb --scope full --output-dir ./myapp_docs/

# Analyze dependencies and generate recommendations
access-dissect analyze myapp_catalog.json
```

## Scope Modes

| Mode | What's extracted |
|------|-----------------|
| `structure` (default) | Schema, queries, forms, reports, VBA — no data |
| `migration` | Structure + row counts + INSERT scripts |
| `full` | Migration + data profiling (null rates, cardinality, type hints) |

## Output Formats

| Format | Description |
|--------|-------------|
| JSON/YAML | Machine-readable catalog (intermediate format) |
| Markdown | One `.md` file per object, cross-linked, git-friendly |
| HTML | Single self-contained report with Mermaid ER diagrams |
| SQL | DDL scripts (PostgreSQL, SQLite, or SQL Server dialect) |

## Known Limitations (v1)

- **Windows only** for extraction (render/analyze runs cross-platform from a catalog)
- **VBA project passwords**: Cannot extract code without the unlocked database
- **Embedded macros**: Action list not fully decoded in v1
- **ActiveX controls**: ProgIDs cataloged but not translated
- **Linked table traversal**: External sources cataloged by connection string only
- **VBA → Python translation**: Out of scope for v1; code is extracted and documented

## Architecture

See the [plan file](../.claude/plans/create-a-plan-for-playful-panda.md) for the full architectural design.
