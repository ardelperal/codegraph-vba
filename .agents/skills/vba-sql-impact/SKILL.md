---
name: vba-sql-impact
description: Traces VBA references, form bindings, and SQL table/column alias lineage for MS Access impact analysis.
---

# Skill: VBA SQL Impact Analysis (vba-sql-impact)

Use this skill to assess the impact of changes in Access/VBA database tables and saved queries (QueryDefs) on VBA modules and form layout bindings.

## 1. Analysis Steps

### Step 1: Trace VBA Callers
1. Search the index using `codegraph_explore` or grep for the query/table name.
2. Read VBA module files (`.cls`, `.bas`, `.frm`) and run `traceVbaCallers` from `src/utils/sql-impact-helpers` to locate line numbers where `OpenRecordset` or `QueryDefs` references the target query.

### Step 2: Extract Form and Report Bindings
1. Scan the workspace for layout files ending in `.form.txt` or `.report.txt`.
2. For each layout file, read the content and execute `extractFormBindings` to extract:
   - `RecordSource` (bound to the form itself)
   - `RowSource` (bound to controls like `ComboBox` or `ListBox`)
3. Match targets to see if they reference the query/table being analyzed.

### Step 3: Resolve SQL Alias & Lineage
1. Read the saved query SQL definition file (typically in `queries/<QueryName>.sql`).
2. Run `resolveSqlLineage` to:
   - Identify tables and aliases in FROM and JOIN clauses.
   - Map qualified column references (e.g. `r.estado`) to their base tables (e.g. `tblRiesgos.estado`).

### Step 4: Compute Downstream Impact & Warnings
1. Compile the list of queries, forms/reports, and VBA callers affected.
2. If a table schema changes, verify which columns/queries/forms are impacted downstream.
3. Emit warnings for unaliased columns or ambiguous references.

## 2. Output Schema Format

Output the findings as a JSON payload conforming to the following structure:

```json
{
  "query_name": "<TargetQueryName>",
  "callers": [
    {
      "file": "<FilePath>",
      "line": <LineNumber>,
      "context": "<MatchingLineContent>"
    }
  ],
  "form_bindings": [
    {
      "file": "<FilePath>",
      "control": "<ControlNameOrForm>",
      "property": "RecordSource|RowSource",
      "target": "<TargetQueryOrTable>"
    }
  ],
  "tables_touched": [
    "<TableName>"
  ],
  "lineage": [
    {
      "source": "<Alias.Column>",
      "resolved": "<TableName.Column>"
    }
  ],
  "downstream_impact": {
    "queries": ["<AffectedQueryName>"],
    "forms": ["<AffectedFormName>"],
    "vba_callers": ["<AffectedVbaFileName>"]
  },
  "warnings": [
    "<WarningMessage>"
  ]
}
```
