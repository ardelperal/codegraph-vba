# vba-code-extraction

## Purpose

Indexes Dysflow-exported `.bas`/`.cls` VBA source into codegraph. Emits `module`/class nodes per file, `function` nodes per `Sub`/`Function`/`Property`, and the `calls`/`contains`/`implements`/`references` edges between them. Cross-module calls, qualified type references, `WithEvents` listeners, and SQL table references inside string literals emit synthesized edges tagged with `metadata.synthesizedBy` so `codegraph_explore` can render provenance inline. `.form.txt`/`.report.txt` are out of scope here (see `vba-form-ui-extraction`).

## Requirements

### Requirement: Procedure Declarations In Standard Modules

For each `.bas` source the system MUST emit one `module` node plus one `function` node per `Sub`/`Function`/`Property` declaration. The function node MUST carry `metadata.visibility` set to the declared keyword (`'Public'`, `'Private'`, `'Friend'`, `'Static'`) or `'Public'` when no visibility keyword is present.

#### Scenario: Public Sub in .bas

- GIVEN a `.bas` source containing `Public Sub SaveRecord()` ... `End Sub`
- WHEN the extractor processes the source with filePath `src/modules/modRepo.bas`
- THEN it emits one `module` node for the file
- AND one `function` node named `SaveRecord` with `metadata.visibility === 'Public'`

#### Scenario: Private Function in .bas

- GIVEN a `.bas` source containing `Private Function CalcTotal() As Long` ... `End Function`
- WHEN the extractor processes the source
- THEN it emits one `function` node named `CalcTotal` with `metadata.visibility === 'Private'`

#### Scenario: Property declaration in .bas

- GIVEN a `.bas` source containing `Property Get Name() As String` ... `End Property`
- WHEN the extractor processes the source
- THEN it emits one `function` node named `Name` with `metadata.visibility === 'Public'`

### Requirement: Procedure Declarations In Class Modules

For each `.cls` source the system MUST emit one `class` node plus one `function` node per `Sub`/`Function`/`Property`, with a `contains` edge from the class node to each procedure.

#### Scenario: Method in .cls

- GIVEN a `.cls` source containing `Public Function Calc() As Long` ... `End Function`
- WHEN the extractor processes the source with filePath `src/classes/CalcEngine.cls`
- THEN it emits one `class` node and one `function` node named `Calc` with `metadata.visibility === 'Public'`
- AND one `contains` edge from the class node to the `Calc` function node

### Requirement: Class Initializer Marker

When a `.cls` declares `Sub New()` (either `Public Sub New()` or `Private Sub New()`), the system MUST set `metadata.hasClassInitializer = true` and `metadata.initializerName = 'New'` on the class node. When no `Sub New` is present, `hasClassInitializer` MUST be absent or `false` and `initializerName` MUST be absent.

#### Scenario: Public Sub New sets marker

- GIVEN a `.cls` source containing `Public Sub New()` ... `End Sub`
- WHEN the extractor processes the source
- THEN the class node has `metadata.hasClassInitializer === true` and `metadata.initializerName === 'New'`

#### Scenario: Private Sub New sets marker

- GIVEN a `.cls` source containing `Private Sub New()` ... `End Sub`
- WHEN the extractor processes the source
- THEN the class node has `metadata.hasClassInitializer === true`

#### Scenario: Missing Sub New leaves marker unset

- GIVEN a `.cls` source containing `Public Function DoWork()` ... `End Function` and no `Sub New`
- WHEN the extractor processes the source
- THEN the class node's `metadata` does NOT contain `hasClassInitializer` or `initializerName`

### Requirement: Call Sites Emit Edges

The system MUST emit a `calls` edge for every call expression `Foo.Bar(...)` or `Bar(...)` inside a procedure body. Same-file-resolvable calls MUST NOT carry `synthesizedBy`. Calls whose receiver resolves outside the file MUST carry `provenance: 'heuristic'` and `metadata.synthesizedBy = 'vba-name-resolution'`. Unresolvable receivers MUST NOT emit an edge and MUST NOT raise an error.

#### Scenario: Same-file call emits plain calls edge

- GIVEN a `.bas` where `Sub Outer()` calls `Inner` and `Sub Inner()` is also defined in the same file
- WHEN the extractor processes the source
- THEN it emits a `calls` edge from `Outer` to `Inner` with no `metadata.synthesizedBy`

#### Scenario: Cross-module qualified call uses synthesizedBy

- GIVEN a `.bas` calling `modHelpers.CalcTotal` where `modHelpers` is not defined in the file
- WHEN the extractor processes the source
- THEN it emits a `calls` edge to `modHelpers.CalcTotal` with `provenance === 'heuristic'` and `metadata.synthesizedBy === 'vba-name-resolution'`

#### Scenario: Unresolvable call is silent

- GIVEN a `.bas` containing `UnknownExternal.Whatever` inside a procedure body
- WHEN the extractor processes the source
- THEN no edge whose target starts with `UnknownExternal` is emitted
- AND the extractor returns without throwing

### Requirement: Implements Edges

The system MUST emit one `implements` edge from the class node to each interface listed in `Implements IFoo` declarations.

#### Scenario: Implements IFoo emits edge

- GIVEN a `.cls` containing `Implements IFoo` at module scope
- WHEN the extractor processes the source
- THEN it emits one `implements` edge from the class node to a node named `IFoo`

### Requirement: Qualified Dim Type References

The system MUST emit a `references` edge with `metadata.synthesizedBy = 'vba-name-resolution'` for `Dim x As Foo.Bar`, targeting the outer type `Foo`. Unqualified `Dim x As Long` MUST NOT emit a `references` edge.

#### Scenario: Qualified Dim references outer type

- GIVEN a `.bas` containing `Dim m_Calc As CalcEngine.Helper`
- WHEN the extractor processes the source
- THEN it emits a `references` edge to `CalcEngine` with `metadata.synthesizedBy === 'vba-name-resolution'`

#### Scenario: Unqualified Dim does not emit edge

- GIVEN a `.bas` containing `Dim m_Count As Long`
- WHEN the extractor processes the source
- THEN no `references` edge is emitted from that `Dim`

### Requirement: WithEvents References

The system MUST emit a `references` edge with `metadata.synthesizedBy = 'vba-withevents'` for each `WithEvents m_X As Form_Foo`, targeting `Form_Foo`.

#### Scenario: WithEvents emits synthesized reference

- GIVEN a `.cls` containing `WithEvents m_Form As Form_Main`
- WHEN the extractor processes the source
- THEN it emits a `references` edge from the class node to `Form_Main` with `metadata.synthesizedBy === 'vba-withevents'`

### Requirement: SQL String Table References

The system MUST scan string literals passed to `DoCmd.RunSQL`, `CurrentDb.OpenRecordset`, `CurrentDb.Execute`, and `db.Execute` for table names following the SQL keywords `FROM`, `INTO`, and `UPDATE`, emitting one `references` edge per discovered table with `metadata.synthesizedBy = 'vba-sql-table'`. The system MUST NOT match table names inside string literals that appear inside `'...` VBA comments or `Rem ...` lines.

#### Scenario: FROM clause resolves table

- GIVEN a `.bas` containing `DoCmd.RunSQL "SELECT * FROM tblCustomers"`
- WHEN the extractor processes the source
- THEN it emits a `references` edge to `tblCustomers` with `metadata.synthesizedBy === 'vba-sql-table'`

#### Scenario: UPDATE statement resolves table

- GIVEN a `.bas` containing `CurrentDb.Execute "UPDATE tblOrders SET Status = 1"`
- WHEN the extractor processes the source
- THEN it emits a `references` edge to `tblOrders` with `metadata.synthesizedBy === 'vba-sql-table'`

#### Scenario: INTO clause resolves table

- GIVEN a `.bas` containing `DoCmd.RunSQL "INSERT INTO tblAudit (Id) VALUES (1)"`
- WHEN the extractor processes the source
- THEN it emits a `references` edge to `tblAudit` with `metadata.synthesizedBy === 'vba-sql-table'`

#### Scenario: SQL inside VBA comment does not match

- GIVEN a `.bas` containing `' DoCmd.RunSQL "SELECT * FROM tblFake"` followed by an unrelated `Sub DoWork() ... End Sub`
- WHEN the extractor processes the source
- THEN no `references` edge to `tblFake` is emitted

### Requirement: .form.txt Produces Zero Code Nodes

When invoked on a `.form.txt` source (regardless of content), the system MUST emit zero `function`, zero non-form `module`, and zero `class` nodes. The form-side behavior lives in `vba-form-ui-extraction`.

#### Scenario: .form.txt input is rejected by this extractor

- GIVEN a `.form.txt` source containing the literal text `Sub Form_Load() End Sub`
- WHEN the extractor processes the source
- THEN the result contains zero `function`, zero `class`, and zero `module` nodes

### Requirement: Option Directives Emit Nothing

The system MUST NOT emit any node or edge for `Option Compare Database`, `Option Explicit`, `Option Base`, or any other `Option` directive.

#### Scenario: Option directives are inert

- GIVEN a `.bas` whose entire body is `Option Explicit` followed by `Option Compare Database`
- WHEN the extractor processes the source
- THEN the result contains zero nodes and zero edges

### Requirement: Module Display Name From VB_Name

When the first non-empty line is `Attribute VB_Name = "Name"`, the system MUST use `Name` as the display name for the `module` (`.bas`) or `class` (`.cls`) node. Otherwise the system MUST fall back to the file basename without extension.

#### Scenario: VB_Name attribute is used

- GIVEN a `.bas` whose first line is `Attribute VB_Name = "modHelpers"`
- WHEN the extractor processes the source with filePath `src/modules/something.bas`
- THEN the `module` node's `name` equals `'modHelpers'`

#### Scenario: Filename is used when VB_Name absent

- GIVEN a `.bas` with no `Attribute VB_Name` line
- WHEN the extractor processes the source with filePath `src/modules/modHelpers.bas`
- THEN the `module` node's `name` equals `'modHelpers'`