# Delta for vba-code-extraction

## MODIFIED Requirements

### Requirement: Call Sites Emit Edges

The system MUST emit a `calls` edge for every call expression inside a procedure body, subject to:

- **Same-file bare calls** (`Bar(...)`/`Bar arg`): emitted only when `Bar` resolves in the same file. No `synthesizedBy`.
- **Qualified paren-form calls** (`Foo.Bar(...)`): always emit heuristic edge, `provenance: 'heuristic'`, `metadata.synthesizedBy = 'vba-name-resolution'`.
- **Qualified statement-form calls** (`Foo.Bar arg`): emit heuristic edge ONLY when `Foo` is a file-local variable (`Dim`/`Private`/`Public`/`WithEvents`) whose declared type is a simple, non-qualified, non-primitive identifier (candidate project class). Otherwise silent — no edge, no error.
- **String-literal masking**: call patterns inside `"..."` never produce edges.
- **Post-extraction stub resolution (NEW)**: after full-project `indexAll()`, every qualified-statement edge above and every `.bas`-qualified paren edge (`modUtils.Foo(...)`) MUST repoint from its stub target to the REAL node when uniquely resolvable — class-typed receiver: real method via `localVarTypeMap`'s resolved type (`matchByQualifiedName`); `.bas`-qualified: real node with matching bare `qualifiedName` in that module. Zero/ambiguous matches keep the stub target with `metadata.stub = true` — never dropped, never a crash. Re-running `indexAll()` unchanged MUST NOT duplicate the repointed edge or any node.
- **CI**: post-repoint scenarios MUST be verifiable in `extraction-vba.test.ts` and/or `extraction-vba-realfixtures.test.ts` (Windows CI VBA regression subset).

(Previously: qualified stub edges were permanent dead ends; no post-extraction repoint pass existed.)

#### Scenario: Same-file call emits plain calls edge
- GIVEN `Sub Outer()` calls `Inner`, `Sub Inner()` in the same file
- WHEN extracted THEN a `calls` edge `Outer`→`Inner` with no `synthesizedBy`

#### Scenario: Cross-module qualified call uses synthesizedBy
- GIVEN `.bas` calling `modHelpers.CalcTotal(...)`, `modHelpers` not in file
- WHEN extracted THEN edge to `modHelpers.CalcTotal`, `provenance === 'heuristic'`

#### Scenario: Qualified statement call on declared project-class variable emits stub edge
- GIVEN `Dim m_Op As ARAuditoriaOperaciones` then `m_Op.Eliminar p_Error`
- WHEN extracted THEN heuristic `calls` edge is emitted (pre-resolution target may be a stub)

#### Scenario: Unresolvable call is silent
- GIVEN `UnknownExternal.Whatever` with `UnknownExternal` undeclared
- WHEN extracted THEN no edge targeting `UnknownExternal.*`; no throw

#### Scenario: Qualified statement call on DAO runtime variable is silent
- GIVEN `Dim rcdDatos As DAO.Recordset` then `rcdDatos.AddNew`
- WHEN extracted THEN no heuristic edge (qualified declared type)

#### Scenario: Cross-file class-typed stub resolves to real method node
- GIVEN `ARAuditoriaOperaciones.cls` defines `Public Sub Eliminar(...)`; another `.cls` has `Dim m_Op As ARAuditoriaOperaciones` and `m_Op.Eliminar p_Error`
- WHEN the project is fully extracted via `indexAll()` and resolution runs
- THEN the edge targets the real `Eliminar` node owned by `ARAuditoriaOperaciones`; `metadata.stub` is absent/false

#### Scenario: Cross-file .bas-qualified call resolves to real bare-name node
- GIVEN `modUtils.bas` defines `Public Function Foo()`; another `.bas` calls `modUtils.Foo(...)`
- WHEN `indexAll()` + resolution runs THEN the edge targets `modUtils`'s real `Foo` node

#### Scenario: Same-file .bas-qualified self-call resolves
- GIVEN `modUtils.bas` itself calls `modUtils.Foo(...)` where `Foo` is declared in that same file
- WHEN `indexAll()` + resolution runs THEN the edge targets the real same-file `Foo` node

#### Scenario: Ambiguous or unmatched stub keeps stub metadata
- GIVEN a class-typed call whose type resolves to two classes with a same-named method (ambiguous), OR a `.bas`-qualified call whose member has no matching real node (unmatched)
- WHEN `indexAll()` + resolution runs
- THEN the edge is retained, still points at its stub target, `metadata.stub === true`; resolution does not throw

#### Scenario: Re-index is idempotent for repointed edges
- GIVEN a project already extracted once with a repointed cross-file class-typed edge
- WHEN `indexAll()` runs again unchanged
- THEN node count and edge count for that call site are unchanged (no duplicates)

### Requirement: SQL String Table References

The system MUST scan string literals passed to `DoCmd.RunSQL`, `CurrentDb.OpenRecordset`, `CurrentDb.Execute`, and `db.Execute` for table names following `FROM`, `INTO`, `UPDATE`, emitting one `references` edge per table with `metadata.synthesizedBy = 'vba-sql-table'`. MUST NOT match inside `'...`/`Rem ...` comments.

- **Accumulation (NEW)**: when a local variable's RHS begins, case-insensitively, with `<varName> &` (e.g. `sql = sql & "..."`, `Sql = sql & "..."`), the assignment ACCUMULATES: tracked SQL text includes every string-literal fragment from all prior self-referential assignments plus the new one, in order, up to the wrapper call. Every `FROM`/`INTO`/`UPDATE` table across ALL fragments gets a `references` edge.
- **Reset (NEW)**: a non-self-referential reassignment (RHS not starting with `<varName> &`, case-insensitive) MUST reset tracking — replace, not append.
- **CI**: accumulation/reset scenarios MUST be verifiable in `extraction-vba.test.ts` and/or `extraction-vba-realfixtures.test.ts` (Windows CI VBA regression subset).

(Previously: `trackSqlVariableAssignment` always overwrote tracked text on reassignment, silently dropping earlier fragments' tables — typically the initial `FROM <table>`.)

#### Scenario: FROM clause resolves table
- GIVEN `DoCmd.RunSQL "SELECT * FROM tblCustomers"` THEN edge to `tblCustomers`

#### Scenario: UPDATE statement resolves table
- GIVEN `CurrentDb.Execute "UPDATE tblOrders SET Status = 1"` THEN edge to `tblOrders`

#### Scenario: INTO clause resolves table
- GIVEN `DoCmd.RunSQL "INSERT INTO tblAudit (Id) VALUES (1)"` THEN edge to `tblAudit`

#### Scenario: SQL inside VBA comment does not match
- GIVEN `' DoCmd.RunSQL "SELECT * FROM tblFake"` THEN no edge to `tblFake`

#### Scenario: Two-fragment self-referential concatenation accumulates
- GIVEN `sql = "SELECT * FROM tblA"` then `sql = sql & " WHERE x=1"` then `db.Execute sql`
- WHEN extracted THEN `references` edges exist for `tblA` (from fragment 1, not lost)

#### Scenario: Three-plus fragment accumulation
- GIVEN `sql = "SELECT * FROM tblA"`, `sql = sql & " INTO tblB"`, `sql = sql & " ..."`, then `db.Execute sql`
- WHEN extracted THEN `references` edges exist for BOTH `tblA` and `tblB`

#### Scenario: Fresh reassignment after use resets tracking
- GIVEN `sql = "SELECT * FROM tblA"`, `db.Execute sql`, then later `sql = "UPDATE tblC SET x=1"` (fresh, not self-referential), `db.Execute sql`
- WHEN extracted THEN the second `db.Execute` yields an edge to `tblC` only, not `tblA`

#### Scenario: Case-insensitive self-reference is detected
- GIVEN `Sql = "SELECT * FROM tblA"` then `Sql = sql & " WHERE x=1"` (mixed case) then `db.Execute Sql`
- WHEN extracted THEN accumulation applies (edge to `tblA` retained), treated as self-referential despite case mismatch
