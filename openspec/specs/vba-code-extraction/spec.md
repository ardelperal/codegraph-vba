# vba-code-extraction

## Purpose

Indexes Dysflow-exported `.bas`/`.cls` VBA source into codegraph. Emits `module`/class nodes per file, `function` nodes per `Sub`/`Function`/`Property`, `enum`/`enum_member` nodes per `Enum` block, `constant` nodes per `Const` declaration, and the `calls`/`contains`/`implements`/`references` edges between them. Cross-module calls, qualified type references, `WithEvents` listeners, and SQL table references inside string literals emit synthesized edges tagged with `metadata.synthesizedBy` so `codegraph_explore` can render provenance inline. `.form.txt`/`.report.txt` are out of scope here (see `vba-form-ui-extraction`).

## Requirements

### Requirement: Procedure Declarations In Standard Modules

For each `.bas` source the system MUST emit one `module` node plus one `function` node per `Sub`/`Function`/`Property` declaration. The function node MUST carry `node.visibility` set to `'public'` (for `Public`, `Friend`, `Static`, or no visibility keyword — VBA's `Static` is a storage specifier, not a visibility modifier; `Friend` is not in the canonical `Node.visibility` enum and folds to `'public'` as the closest broader-than-private option) or `'private'` (for `Private`). The visibility value uses the canonical lowercase enum used by every other extractor in the project (`'public' | 'private' | 'protected' | 'internal'`).

#### Scenario: Public Sub in .bas

- GIVEN a `.bas` source containing `Public Sub SaveRecord()` ... `End Sub`
- WHEN the extractor processes the source with filePath `src/modules/modRepo.bas`
- THEN it emits one `module` node for the file
- AND one `function` node named `SaveRecord` with `node.visibility === 'public'`

#### Scenario: Private Function in .bas

- GIVEN a `.bas` source containing `Private Function CalcTotal() As Long` ... `End Function`
- WHEN the extractor processes the source
- THEN it emits one `function` node named `CalcTotal` with `node.visibility === 'private'`

#### Scenario: Property declaration in .bas

- GIVEN a `.bas` source containing `Property Get Name() As String` ... `End Property`
- WHEN the extractor processes the source
- THEN it emits one `function` node named `Name` with `node.visibility === 'public'`

### Requirement: Procedure Declarations In Class Modules

For each `.cls` source the system MUST emit one `class` node plus one `function` node per `Sub`/`Function`/`Property`, with a `contains` edge from the class node to each procedure.

#### Scenario: Method in .cls

- GIVEN a `.cls` source containing `Public Function Calc() As Long` ... `End Function`
- WHEN the extractor processes the source with filePath `src/classes/CalcEngine.cls`
- THEN it emits one `class` node and one `function` node named `Calc` with `node.visibility === 'public'`
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

The system MUST emit a `calls` edge for every call expression inside a procedure body, subject to the following rules:

- **Same-file bare calls** (`Bar(...)` or statement-form `Bar arg`): emitted only when `Bar` resolves to a procedure declared in the same file. These edges carry no `synthesizedBy`.
- **Qualified paren-form calls** (`Foo.Bar(...)`): always emit a heuristic `calls` edge with `provenance: 'heuristic'` and `metadata.synthesizedBy = 'vba-name-resolution'`.
- **Qualified statement-form calls** (`Foo.Bar arg`, no parens): emit a heuristic `calls` edge ONLY when `Foo` is a file-local variable declared via `Dim`/`Private`/`Public`/`WithEvents` and whose declared type is a **simple, non-qualified, non-primitive** identifier (a candidate project class). When `Foo` is not declared in the file, or its type is qualified (e.g. `DAO.Recordset`) or a primitive, the call is **silent** (no edge, no error) — satisfying "Unresolvable call is silent".
- **String-literal masking**: call patterns inside `"..."` string literals MUST NOT produce edges. The scanner masks string content before applying call-site patterns.

#### Scenario: Same-file call emits plain calls edge

- GIVEN a `.bas` where `Sub Outer()` calls `Inner` and `Sub Inner()` is also defined in the same file
- WHEN the extractor processes the source
- THEN it emits a `calls` edge from `Outer` to `Inner` with no `metadata.synthesizedBy`

#### Scenario: Cross-module qualified call uses synthesizedBy

- GIVEN a `.bas` calling `modHelpers.CalcTotal(...)` (paren form) where `modHelpers` is not defined in the file
- WHEN the extractor processes the source
- THEN it emits a `calls` edge to `modHelpers.CalcTotal` with `provenance === 'heuristic'` and `metadata.synthesizedBy === 'vba-name-resolution'`

#### Scenario: Qualified statement call on declared project-class variable emits edge

- GIVEN a `.cls` containing `Dim m_Op As ARAuditoriaOperaciones` and later `m_Op.Eliminar p_Error` (statement form, no parens)
- WHEN the extractor processes the source
- THEN it emits a heuristic `calls` edge from the enclosing procedure to `m_Op.Eliminar` with `provenance === 'heuristic'`

#### Scenario: Unresolvable call is silent

- GIVEN a `.bas` containing `UnknownExternal.Whatever` (statement form, no parens) inside a procedure body where `UnknownExternal` is not declared as a local variable
- WHEN the extractor processes the source
- THEN no edge whose target starts with `UnknownExternal` is emitted
- AND the extractor returns without throwing

#### Scenario: Qualified statement call on DAO runtime variable is silent

- GIVEN a `.cls` containing `Dim rcdDatos As DAO.Recordset` and later `rcdDatos.AddNew` (statement form)
- WHEN the extractor processes the source
- THEN no heuristic edge is emitted for `rcdDatos.AddNew` (qualified declared type → silent)

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

### Requirement: Enum Declarations

For each `Enum <Name>` ... `End Enum` block the system MUST emit one `enum` node named `<Name>` and one `enum_member` node per member line, with a `contains` edge from the `enum` node to each `enum_member` node and a `contains` edge from the module/class node to the `enum` node. The `enum` node MUST carry `node.visibility` folded from the declaration keyword (`Private` → `'private'`; `Public`/`Global`/`Friend`/none → `'public'`). Each `enum_member` node's `qualifiedName` MUST be `<EnumName>.<MemberName>` so members of distinct enums that share a member name remain distinguishable.

#### Scenario: Public Enum emits enum + member nodes

- GIVEN a `.bas` containing `Public Enum EnumTipoUsuario` / `Administrador = 1` / `Calidad = 2` / `End Enum`
- WHEN the extractor processes the source
- THEN it emits one `enum` node named `EnumTipoUsuario` with `node.visibility === 'public'`
- AND one `enum_member` node per member (`Administrador`, `Calidad`)
- AND a `contains` edge from the `enum` node to each `enum_member` node
- AND a `contains` edge from the `module` node to the `enum` node

#### Scenario: enum_member qualifiedName is enum-scoped

- GIVEN a `.bas` containing `Public Enum EnumTipoUsuario` / `Administrador = 1` / `End Enum`
- WHEN the extractor processes the source
- THEN the `enum_member` node named `Administrador` has `qualifiedName === 'EnumTipoUsuario.Administrador'`

### Requirement: Const Declarations

For each `Const` declaration line the system MUST emit one `constant` node per declared name (a multi-name line such as `Const A = 1, B = 2` emits one node per name), with a `contains` edge from the module/class node to each `constant` node. The `constant` node MUST carry `node.visibility` folded from the declaration keyword (`Private` → `'private'`; `Public`/`Global`/`Friend`/none → `'public'`).

#### Scenario: Public Const emits a constant node

- GIVEN a `.bas` containing `Public Const msoFileDialogOpen As Long = 1`
- WHEN the extractor processes the source
- THEN it emits one `constant` node named `msoFileDialogOpen` with `node.visibility === 'public'`
- AND a `contains` edge from the `module` node to that `constant` node

#### Scenario: multi-name Const line emits one node per name

- GIVEN a `.bas` containing `Const A = 1, B = 2, C = 3`
- WHEN the extractor processes the source
- THEN it emits three `constant` nodes named `A`, `B`, and `C`

### Requirement: .form.txt Produces Zero Code Nodes

When invoked on a `.form.txt` source (regardless of content), the system MUST emit zero `function`, zero non-form `module`, and zero `class` nodes. The form-side behavior lives in `vba-form-ui-extraction`.

#### Scenario: .form.txt input is rejected by this extractor

- GIVEN a `.form.txt` source containing the literal text `Sub Form_Load() End Sub`
- WHEN the extractor processes the source
- THEN the result contains zero `function`, zero `class`, and zero `module` nodes

### Requirement: Option Directives Emit Nothing

The system MUST NOT emit any node or edge for `Option Compare Database`, `Option Explicit`, `Option Base`, or any other `Option` directive. A file whose only declarations are `Option` directives (no procedures, enums, constants, or other symbols) MUST emit zero symbol nodes — the lazy module/class node is suppressed. (A file with `Enum`/`Const` but no procedures DOES emit a module/class node — see Enum/Const Declarations.)

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

### Requirement: Bitwise-Precise Conditional-Compilation

To support accurate extraction of VBA code containing conditional compilation directives, the preprocessor MUST parse and evaluate conditional compilation expressions using bitwise-precise evaluation, signed 32-bit integer arithmetic, and VBA's `-1` (True) / `0` (False) truthiness semantics. The evaluation MUST NOT use unsafe JavaScript `eval` or `Function` execution.

The preprocessor MUST support:
- Logical/bitwise operators: `And`, `Or`, `Not`, `Xor` (case-insensitive, evaluated as bitwise operations on signed 32-bit integers, without short-circuiting).
- Comparison operators: `=`, `<>`, `<`, `>`, `<=`, `>=` (returning `-1` for true, `0` for false).
- Hardcoded constants: `True` -> `-1`, `False` -> `0`, `VBA7` -> `-1`, `Win64` -> `-1`, `Win32` -> `-1`, `Win16` -> `0`, `Mac` -> `0`.
- File-scoped constants declared with `#Const`.
- Unresolved/undefined identifiers falling back to `0`.
- Out-of-grammar/unhandled syntax falling back safely to false (evaluating to `0` / blanking the inactive branches).

#### Scenario: Bitwise AND evaluates to false

- GIVEN a VBA source file containing:
  ```vba
  #Const FLAGS = 3
  #If (FLAGS And 4) Then
  Public Sub ActiveSub()
  End Sub
  #Else
  Public Sub InactiveSub()
  End Sub
  #End If
  ```
- WHEN the preprocessor evaluates the conditional expression `(FLAGS And 4)`
- THEN `(3 And 4)` evaluates to `0`
- AND the `#If` branch is blanked (deactivated) while the `#Else` branch is kept active.

#### Scenario: Bitwise NOT

- GIVEN a VBA source file containing:
  ```vba
  #Const FLAGS = 0
  #If Not FLAGS Then
  Public Sub ActiveSub()
  End Sub
  #End If
  ```
- WHEN the preprocessor evaluates the conditional expression `Not FLAGS`
- THEN `Not 0` evaluates to `-1`
- AND the `#If` branch is kept active.

#### Scenario: Bitwise XOR

- GIVEN a VBA source file containing:
  ```vba
  #Const FLAGS = 1
  #If FLAGS Xor 1 Then
  Public Sub ActiveSub()
  End Sub
  #Else
  Public Sub InactiveSub()
  End Sub
  #End If
  ```
- WHEN the preprocessor evaluates the conditional expression `FLAGS Xor 1`
- THEN `1 Xor 1` evaluates to `0`
- AND the `#If` branch is blanked while the `#Else` branch is kept active.

#### Scenario: Comparisons return -1 and 0

- GIVEN a VBA source file containing:
  ```vba
  #If Win64 = -1 Then
  Public Sub Win64Sub()
  End Sub
  #End If
  ```
- WHEN the preprocessor evaluates `Win64 = -1` (where `Win64` is pre-defined to `-1`)
- THEN it performs a comparison and returns `-1` (true)
- AND the `#If` branch is kept active.

#### Scenario: Unhandled or out-of-grammar syntax fallback

- GIVEN a VBA source file containing invalid conditional expression syntax:
  ```vba
  #If FLAGS InvalidSyntax @@@ Then
  Public Sub InvalidSub()
  End Sub
  #End If
  ```
- WHEN the preprocessor parses the expression
- THEN it falls back safely to `0` (false)
- AND the `#If` branch is blanked.

