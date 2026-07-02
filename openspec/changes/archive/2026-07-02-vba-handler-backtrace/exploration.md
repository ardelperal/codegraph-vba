## Exploration: vba-handler-backtrace

This report explores the implementation of the `vba-handler-backtrace` capability, analyzing the current state of VBA modeling in the codebase and comparing architectural options for graph traversal.

### 1. Codebase Investigation

A detailed audit of the VBA extractor (`src/extraction/vba-extractor.ts`) and post-extraction resolver (`src/resolution/index.ts`) reveals how relevant concepts are currently modeled in the SQLite database schema (`src/db/schema.sql`).

#### A. Subroutine Calls
Subroutine, function, and property calls are modeled in the database graph via `calls` edges:
- **Same-File Calls:** For bare calls (e.g., `MySub(...)` or statement-form `MySub arg1`), the extractor performs a same-file lookup using `this.findFunctionNodeByName`. If resolved, it directly emits a `calls` edge from the caller procedure's node to the target procedure's node.
- **Qualified Cross-Module/Class Calls:** For qualified calls (e.g., `Receiver.Member(...)` or `Receiver.Member arg1`), the extractor synthesizes a stub `function` node locally (to pass the per-file edge filter) and tags it with `metadata.stub = true`. It emits a heuristic `calls` edge to this stub.
- **Post-Extraction Name Resolution:** During the resolution pass (`resolveVbaCallStubs` in `src/resolution/index.ts`), stubs are repointed to their unique real targets when uniquely resolvable:
  1. *Class-typed objects:* Resolved by exact `qualifiedName` match (e.g., `${ClassName}.${proc}`).
  2. *`.bas`-qualified calls:* Resolved by splitting the stub's qualified name and narrowing bare-member candidates to the `.bas` file where the sibling `module` node name matches the `receiver` name.
- **Unresolved Stubs:** If the call target is ambiguous (2+ candidates) or unresolvable, the stub and its edges are left untouched in the database.

#### B. Form Control Events
Form control events in Access are modeled as follows:
- **Handler Subroutines:** If a procedure in a Form code-behind (`Form_*.cls`) has a name matching the Access naming convention `<ControlName>_<EventName>` (e.g., `cmdSave_Click`), it is split on the last underscore.
- **Control Stubs:** The extractor synthesizes a stub `form-instance-control` node with the name `<ControlName>` and a file path corresponding to the sibling `.form.txt` file (deterministic ID).
- **Edges:** It emits an `event-handler` edge from the procedure's function node (`source`) to the stub `form-instance-control` node (`target`), carrying `metadata.eventName` (e.g., `"Click"`).
- **Overwriting:** At index time, when the sibling `.form.txt` is processed by the `VbaFormExtractor`, it emits the real `form-instance-control` node with the same ID, overwriting the stub (preserving properties and layout information).

#### C. DAO Calls & SQL Extraction
DAO calls and SQL queries are extracted as follows:
- **Detection:** The extractor uses `SQL_WRAPPERS` or `SQL_VAR_EXEC_RE` regexes (matching `DoCmd.RunSQL`, `db.OpenRecordset`, `db.Execute`, `getdb().Execute`, and variable-based calls).
- **String Accumulation:** SQL string concatenations (including self-referential `sql = sql & "..."`) are tracked and accumulated in an in-memory map (`localVarTypeMap`).
- **Table References:** The collected SQL string is scanned using `SQL_TABLE_RE` to find `FROM`, `INTO`, or `UPDATE` table names.
- **Edges:** For each table found, a synthetic `class` node representing the table is emitted. A `references` edge from the containing module/class node (not the procedure node) to the synthetic table node is created with `metadata.synthesizedBy = 'vba-sql-table'`.
- **Limitation:** The actual SQL query strings or hints are **not** stored in the database nodes or edge metadata.

#### D. UDT Type References (`type_refs`)
User-Defined Types (UDTs) are modeled as follows:
- **Type Declarations:** `Type...End Type` declarations become `type` and `type_member` nodes connected by `type-member` edges.
- **Variable Declarations:** Local variables declared via `Dim/Private/Public` extract unqualified or qualified type names (ignoring primitives/keywords) and emit a generic `references` edge from the containing module/class node to the type node with `metadata.synthesizedBy = 'vba-name-resolution'`.
- **Limitation on Subroutine Parameters/Variables:** Subroutine parameters are **not** parsed or modeled by the extractor (the regex `PROC_RE` only captures the header up to the name). Additionally, local variables do not get their own nodes or `type_of` edges. Therefore, UDT type references (`type_refs`) are **not** stored or modeled on parameters or variables in the database.

---

### 2. Architectural Approaches

Two options were evaluated to implement the `vba-handler-backtrace` capability:

#### Option A: Dynamic Agent Skill (Recommended)
Resolve the backtrace dynamically at query time using the custom agent skill `vba-handler-backtrace` by executing SQLite queries via `codegraph_explore` (or direct query interface) to traverse the call graph up to the specified `depth`.
- **Mechanism:**
  - Locate the starting `form-instance-control` or handler sub node.
  - Traverse incoming/outgoing edges (`event-handler`, `calls`, `references` / `vba-sql-table`) in-memory using agent-side graph traversal logic.
  - For SQL hints: Since the SQL string is not in the DB, the skill fallback-reads the source file at the DAO call line to extract the literal string or variable, reconstructs concatenations, and yields the first 200 characters.
  - For UDT parameters: Since subroutine parameters are not in the DB, the skill fallback-reads the procedure's signature line and parses parameter type names, resolving UDT type references on the fly.
  - Handle cycle detection, truncation, and warnings in agent-side JS logic.
- **Pros:**
  - Keeps the core indexer clean and decoupled from ad-hoc backtrace heuristics.
  - Agent can easily perform file-system fallback reads to grab SQL queries and parameter type signatures without modifying the DB schema.
  - Traversal rules, cycle detection, and warning formats can be adjusted instantly without rebuilding or re-releasing the indexer binaries.
- **Cons:**
  - Relies on the agent performing file-system fallback reads for SQL string contents and parameter signatures, which increases token usage slightly.

#### Option B: Indexer Core/Endpoint
Add a dedicated endpoint or core command to the TypeScript codebase for executing backtraces.
- **Mechanism:**
  - Implement a new API endpoint/CLI command that performs recursive SQL queries (e.g., SQLite CTEs or recursive JS traversal) on the indexer side.
- **Pros:**
  - Faster query execution directly in memory on the indexer side.
- **Cons:**
  - Bloats the indexer core with domain-specific logic.
  - The indexer does not currently store SQL strings or parameter types, meaning we would need to either:
    1. Expand the DB schema to store complete SQL query strings and parameter lists (massive effort, migration needed).
    2. Make the indexer read local files during query execution, which breaks separation of concerns.

---

### 3. Comparison and Recommendation

| Metric | Option A: Dynamic Agent Skill | Option B: Indexer Core/Endpoint |
| :--- | :--- | :--- |
| **Indexer Cleanliness** | **High** (No core code modifications) | Low (Bloats core with custom graph queries) |
| **DB Schema Stability** | **High** (No migrations needed) | Low (Requires schema changes for SQL/params) |
| **Development Effort** | **Low** (Agent-level JS logic) | High (Core TypeScript & CLI additions) |
| **Flexibility** | **High** (Easy to tweak traversal rules) | Low (Requires compilation and updates) |
| **Query Latency** | **Fast (<2s)** | Extremely Fast (<100ms) |

#### Recommendation
We strongly recommend **Option A (Dynamic Agent Skill)**.
The dynamic skill approach keeps the indexer core clean and stable. Since the DB schema does not store SQL query strings or subroutine parameter UDTs, the agent-side skill can easily perform light fallback file-reads on the few files along the resolved trace path. This provides maximum flexibility for cycle detection, max-depth handling, and warning formatting without introducing heavy migrations or bloat into the indexer core.

---

### 4. Risks & Mitigations

- **Out-of-Sync File Snapshot:** Reading file contents for SQL query strings or parameters relies on source files matching the DB index.
  *Mitigation:* Check node `updatedAt` / file `modifiedAt` or issue a warning if the file changed after the index was built.
- **Depth Explosion & Cycles:** Mutual recursion or deep call graphs could lead to out-of-memory or timeout errors.
  *Mitigation:* Strict cycle detection (visited set) and depth capping (default 5, configurable) as defined in the spec.
