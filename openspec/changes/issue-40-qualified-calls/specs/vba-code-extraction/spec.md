# Delta for vba-code-extraction

## MODIFIED Requirements

### Requirement: Call Sites Emit Edges

The system MUST emit a `calls` edge for every call expression inside a procedure body, subject to the following rules:

- **Same-file bare calls** (`Bar(...)` or statement-form `Bar arg`): emitted only when `Bar` resolves to a procedure declared in the same file. These edges carry no `synthesizedBy`.
- **Qualified paren-form calls** (`Foo.Bar(...)`): emit a heuristic `calls` edge only when `Foo` is eligible for qualified-call processing.
- **Qualified statement-form calls** (`Foo.Bar arg`, no parens): emit a heuristic `calls` edge only when `Foo` is eligible for qualified-call processing.
- **Qualified-call processing**: `Foo` is eligible when it is a declared local project class variable, or when it is not declared as a local variable at all. If `Foo` is declared locally and its type is not a project class, the call is silent.
- **String-literal masking**: call patterns inside `"..."` string literals MUST NOT produce edges. The scanner masks string content before applying call-site patterns.

(Previously: qualified paren-form calls were always emitted heuristically, while qualified statement-form calls were only emitted when the receiver was a declared local project class variable. External/primitive local receivers could still create dead-end paren-form stubs, and module-name receivers in statement form were inconsistently dropped.)

#### Scenario: Same-file call emits plain calls edge

- GIVEN a `.bas` where `Sub Outer()` calls `Inner` and `Sub Inner()` is also defined in the same file
- WHEN the extractor processes the source
- THEN it emits a `calls` edge from `Outer` to `Inner` with no `metadata.synthesizedBy`

#### Scenario: Cross-module qualified call uses synthesizedBy

- GIVEN a `.bas` calling `modHelpers.CalcTotal(...)` where `modHelpers` is not defined in the file
- WHEN the extractor processes the source
- THEN it emits a heuristic `calls` edge to `modHelpers.CalcTotal`
- AND `metadata.synthesizedBy === 'vba-name-resolution'`

#### Scenario: Qualified statement call on declared project-class variable emits edge

- GIVEN a `.cls` containing `Dim m_Op As ARAuditoriaOperaciones` and later `m_Op.Eliminar p_Error`
- WHEN the extractor processes the source
- THEN it emits a heuristic `calls` edge from the enclosing procedure to `ARAuditoriaOperaciones.Eliminar`

#### Scenario: Qualified statement call on module name emits edge

- GIVEN a `.bas` containing `modUtils.Foo arg` inside a procedure body where `modUtils` is not declared as a local variable
- WHEN the extractor processes the source
- THEN it emits a heuristic `calls` edge to `modUtils.Foo`

#### Scenario: Qualified statement call on DAO runtime variable is silent

- GIVEN a `.cls` containing `Dim rcdDatos As DAO.Recordset` and later `rcdDatos.AddNew`
- WHEN the extractor processes the source
- THEN no heuristic edge is emitted for `rcdDatos.AddNew`

#### Scenario: Qualified paren-form call on primitive local variable is silent

- GIVEN a `.cls` containing `Dim nCount As Long` and later `nCount.ToString()`
- WHEN the extractor processes the source
- THEN no heuristic edge is emitted for `nCount.ToString()`
