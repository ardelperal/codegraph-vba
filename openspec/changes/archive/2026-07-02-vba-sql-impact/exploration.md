# Exploration Report: VBA SQL Impact Analysis

This report explores the technical feasibility and implementation paths for the `vba-sql-impact` capability, which aims to trace VBA code callers, form bindings, and database tables/columns touched by saved queries (Access QueryDefs represented as `.sql` files).

---

## 1. Codebase Investigation Findings

### 1.1 Saved Query Representation
Access saved queries (QueryDefs) are exported by Dysflow as raw `.sql` files under the `queries/` directory, accompanied by a `queries.json` manifest.
- **Nodes**: Inside `src/extraction/sql-query-extractor.ts`, these files are parsed. A file named `queries/Consulta.sql` produces:
  - A `file` node tracking the `.sql` file path.
  - A `query` node named after the file basename (e.g., `Consulta`), with `kind: 'query'` and `language: 'sql'`.
  - Synthetic table placeholder nodes with `kind: 'class'` (e.g., `tblRiesgos`).
- **Edges**: The extractor uses a regex pattern (`TABLE_RE`) to find table references following keywords like `FROM`, `JOIN`, `INTO`, and `UPDATE`. It emits a `references` edge from the `query` node to the table placeholder node, tagged with `metadata.synthesizedBy = 'sql-query-table'`.
- **Limitations**: The current extraction does not extract database columns, select expressions, alias mappings, or where clauses.

### 1.2 VBA SQL Table References
In `src/extraction/vba-extractor.ts`:
- The extractor detects inline SQL strings (e.g., `m_SQL = "SELECT ..."` and `Set rs = Db.OpenRecordset(m_SQL)`).
- It extracts table names from these SQL strings using a similar regex (`SQL_TABLE_RE`).
- It emits a `references` edge from the VBA module to the synthetic table (`class`) node, tagged with `metadata.synthesizedBy = 'vba-sql-table'`.
- No column-level dependencies are captured.

### 1.3 Form Properties & Bindings
In `src/extraction/vba-form-extractor.ts`:
- Form layout files (`.form.txt` and `.report.txt`) are scanned for control declarations (like `TextBox`, `ListBox`, `ComboBox`, `CommandButton`).
- It emits a `form-layout` node per file and binds it to its sibling class module via an `UnresolvedReference` with `metadata.synthesizedBy = 'vba-form-binding'`.
- **Limitations**: Essential Access form data-source properties such as `RecordSource` (specifying the query or table bound to the form itself) and `RowSource` (specifying queries or table data sources bound to dropdowns/lists) are **not currently parsed or stored** in the database.

### 1.4 SQL Parsing Infrastructure
- There are **no SQL parser dependencies** in the project's `package.json`.
- Current parsing is entirely heuristic-based, relying on regular expressions to identify table names. There is no existing facility for AST-based SQL analysis or alias resolution.

---

## 2. Technical Options Evaluated

### Option A: Agent-Side Custom Skill (`.agents/skills/vba-sql-impact/SKILL.md`)
Implement the impact analyzer as a dynamic agent-side skill that utilizes the existing `codegraph_explore` MCP tool and source file reading.

- **How it works**:
  1. The agent queries `codegraph_explore` to obtain the structure of callers (identifying literal `OpenRecordset` calls and query name references in the database) and basic table references.
  2. The agent reads the target `.sql` saved query file or local `.cls`/`.form.txt` files directly from the workspace.
  3. The agent leverages its reasoning capabilities and regex patterns to parse the SQL locally:
     - Detects column names in the `SELECT` list.
     - Resolves table aliases (e.g., `FROM tblRiesgos r WHERE r.estado` -> resolves `r` to `tblRiesgos` and extracts columns `id`, `descripcion`, `estado`).
     - Maps form controls and their bindings (`RecordSource`, `RowSource`) by parsing form configuration files (`.form.txt` / `.report.txt`) on the fly.
  4. Generates the structured impact report as a JSON payload conforming to the spec contract.
- **Pros**:
  - Decoupled from the indexer core. No need to modify the SQLite schema, parser rules, or indexer runtime.
  - Zero performance impact on indexer startup or re-indexing runs.
  - Leverages agent flexibility to handle highly non-standard Access SQL dialects and local aliases without maintaining a complex SQL parser in TypeScript.
- **Cons**:
  - Relies on the agent to parse files on demand, which could be slightly slower if the number of files to read is very large (mitigated by only reading files related to the queried symbol).

### Option B: Core Indexer & Database Schema Changes
Add comprehensive column-level indexing and form property extraction directly into the core indexer (`codegraph-vba`).

- **How it works**:
  1. Modify `schema.sql` to support column nodes, query-column references, and control properties.
  2. Integrate a robust SQL parser library (e.g., a node SQL parser or a custom Tree-sitter grammar for MS Access SQL) into `package.json` and the indexer pipeline.
  3. Update `VbaFormExtractor` to parse and index `RecordSource` and `RowSource` properties, storing them as nodes/edges in the database.
  4. Update `VbaExtractor` and `SqlQueryExtractor` to index every column-to-table lookup.
- **Pros**:
  - Column and form-property information is fully queryable in SQLite via standard SQL queries or MCP tools without reading source files.
- **Cons**:
  - Significant code churn in the core parser and DB layers.
  - Access SQL dialect is notoriously quirky and poorly supported by standard npm SQL parsers, leading to high parser error rates or the need to write and maintain a custom parser.
  - Increases database size and slows down the indexing process.

---

## 3. Recommendation

We strongly recommend **Option A (Dynamic Agent Skill)**. 

### Rationale:
1. **Decoupled Architecture**: It prevents bloating the core indexer with complex SQL alias resolution and MS Access-specific dialect parsing logic.
2. **Dynamic Alias Resolution**: Resolving SQL aliases (e.g., matching `r.estado` back to `tblRiesgos.estado`) is highly contextual and complex for static indexers but natural and precise for LLM agents when they have access to the target query source code.
3. **No Schema Expansion Needed**: Avoids adding new node/edge tables or property columns to the database schema for a single phase-1 capability.
4. **Resilience to Dialects**: When SQL parsing fails due to quirky Microsoft Access SQL syntax, the agent can fallback gracefully and extract meaningful table/column lists using regex and semantic understanding, which is much harder to implement robustly in a static TypeScript parser.

---

## 4. Proposed Implementation Plan for Option A

1. **Create the Skill File**: Write the step-by-step guide at `.agents/skills/vba-sql-impact/SKILL.md` directing the agent on how to retrieve the files, parse column references, resolve aliases, and format the output.
2. **Query Database**: Use `codegraph_explore` to resolve:
   - VBA caller modules referring to the query name.
   - Sibling form modules and controls.
3. **Read and Parse Files**:
   - Retrieve query definition text from `queries/<QueryName>.sql`.
   - Parse form file `.form.txt`/`.report.txt` properties directly to find `RecordSource` and `RowSource` bindings.
4. **Resolve SQL Column & Table References**:
   - Implement simple regex/semantic parsing inside the skill to capture columns and aliases.
5. **Format & Return JSON**: Emit the structured output payload as defined in the spec.
