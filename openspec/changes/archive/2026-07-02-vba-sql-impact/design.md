# Technical Design: VBA SQL Impact Analysis (vba-sql-impact)

## 1. Architecture & Design Decisions

Rather than modifying the database schema or introducing heavy AST-based SQL parsers to the core TypeScript codebase (Option B), this design employs an **Agent-Side Dynamic Skill** (Option A). The skill is located at `.agents/skills/vba-sql-impact/SKILL.md`.

At invocation time, the agent reads query configurations, source code, and form layout files directly from the workspace and uses regex-based extraction combined with LLM semantic reasoning to map out the lineage.

---

## 2. Data Flow & Execution Steps

### 2.1 Tracing VBA Callers
The agent will trace references to the query name (e.g. `qryGetRiesgos`) inside VBA modules:
1. **Query Indexer**: Query the SQLite database using `codegraph_explore` to look for references, stubs, or calls matching the query name.
2. **Fallback File Reading**: Open VBA modules (`.cls`/`.bas` files) and perform case-insensitive regex sweeps for:
   - `OpenRecordset\s*\(\s*"([^"]+)"`
   - `QueryDefs\s*\(\s*"([^"]+)"`
3. Record the matching file, line number, and context string.

### 2.2 Form and Report Bindings Extraction
Access form and report layouts are stored in `.form.txt` and `.report.txt` files.
1. **Scan Workspace**: Locate layout files matching the query name or referenced tables.
2. **Scan Properties**: Read each layout file line-by-line and extract bindings:
   - **RecordSource** (on the `Begin Form` or `Begin Report` container):
     - Pattern: `RecordSource\s*=\s*"([^"]+)"`
   - **RowSource** (on controls like `ListBox` or `ComboBox`):
     - Pattern: `RowSource\s*=\s*"([^"]+)"`
3. **Trace Control Context**: Track parent control blocks (e.g., `Begin ComboBox` -> `Name = "cboUsuario"`) to attribute the `RowSource` to the specific control.

### 2.3 SQL Alias Resolution & Lineage Mapping
To trace column lineage, the agent will analyze the saved query definition file `queries/<QueryName>.sql`:
1. **Read Definition**: Retrieve the raw SQL text of the query.
2. **Resolve Table Aliases**:
   - Scan `FROM` and `JOIN` clauses to extract tables and their aliases.
   - Patterns:
     - Explicit: `FROM\s+([a-zA-Z0-9_]+)\s+AS\s+([a-zA-Z0-9_]+)`
     - Implicit: `FROM\s+([a-zA-Z0-9_]+)\s+([a-zA-Z0-9_]+)`
3. **Map Columns**:
   - Parse the `SELECT` list and JOIN conditions for column references using aliases (e.g., `r.estado`).
   - Resolve each alias back to its base table (e.g., if `tblRiesgos AS r` is present, map `r.estado` to `tblRiesgos.estado`).
   - If a column reference is unaliased (e.g. `nombre`), check it against target tables; if ambiguous, log a warning.

---

## 3. Output Payload Schema

The skill yields a structured JSON output mapping the impact:

```json
{
  "query_name": "qryGetRiesgos",
  "callers": [
    {
      "file": "src/modules/ModRiesgos.bas",
      "line": 23,
      "context": "db.OpenRecordset(\"qryGetRiesgos\", dbOpenSnapshot)"
    }
  ],
  "form_bindings": [
    {
      "file": "src/forms/frmRiesgos.form.txt",
      "control": "Form",
      "property": "RecordSource",
      "target": "qryGetRiesgos"
    },
    {
      "file": "src/forms/frmRiesgos.form.txt",
      "control": "cboUsuario",
      "property": "RowSource",
      "target": "tblUsuarios"
    }
  ],
  "tables_touched": [
    "tblRiesgos",
    "tblUsuarios"
  ],
  "lineage": [
    {
      "source": "r.estado",
      "resolved": "tblRiesgos.estado"
    },
    {
      "source": "u.nombre",
      "resolved": "tblUsuarios.nombre"
    }
  ],
  "downstream_impact": {
    "queries": ["qryGetRiesgos"],
    "forms": ["frmRiesgos"],
    "vba_callers": ["ModRiesgos.bas"]
  },
  "warnings": []
}
```
