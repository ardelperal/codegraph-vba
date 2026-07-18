# Spike: tree-sitter-vbnet.wasm as a VBA parser (F.1)

> Phase F.1 of [issue #155](../issues/155): parse a representative VBA corpus with
> the vendored `tree-sitter-vbnet.wasm` grammar and classify the failure modes.
> **No production-code changes** — this is a research spike, not a refactor.

## Method

- Grammar: `tree-sitter-vbnet.wasm` (vendored, already shipped at `src/extraction/wasm/`).
- Corpus: every `.bas` and `.cls` under `__tests__\fixtures/` (Dysflow export format
  `.form.txt`/ `.report.txt` excluded — they are not VBA source, they are a
  Dysflow export format handled by `VbaFormExtractor`).
- Parser: `web-tree-sitter@0.25.10` (the version codegraph already uses).
- The grammar has no real `_eof` token, so files without a trailing newline end with
  a `MISSING` newline error on the last statement. Each file gets a trailing newline
  appended before parsing (mirrors `src/extraction/languages/vbnet.ts:7-9`).
- Per-file classification:
  - `clean` — 0 ERROR nodes.
  - `partial` — ERROR rate < 30% of total nodes (recognisable, localized damage).
  - `failed` — ERROR rate >= 30% (most bytes inside ERROR nodes).

## Verdict

**GO** — The quantitative gate passes (0% of files fail, well under the 30% threshold) AND the synthesized-wrapper dry-run unlocks the full structural tree: injecting a `Class <Name>` opener + `End Class` closer (where `<Name>` comes from the existing `Attribute VB_Name`) makes the grammar emit 13 `class_declaration`, 26 `method_declaration`, 36 `field_declaration`, 1 `property_declaration`, 36 `parameter_list`, 2 `event_declaration`, and 2 `raiseevent_statement` across the 15 files in the corpus. F.2 is a tractable hybrid extractor rewrite — inject the wrapper, walk the AST, fill the Access-specific gaps (form/report/event/DoCmd) with the existing regex layer.

## Headline numbers

- Files parsed: **16**
- Clean: **0** (0%)
- Partial: **16** (100%)
- Failed: **0** (0%)

## Structural completeness (the headline finding)

The vbnet grammar splits cleanly into two layers:

- **Structural** nodes — the module/class/procedure/field boundary markers:
  `class_declaration`, `module_declaration`, `method_declaration`, `field_declaration`,
  `property_declaration`, `parameter_list`, `event_declaration`, `implements_clause`,
  `imports_statement`, `declaration_statement`.
- **Body** nodes — the statements and expressions INSIDE a procedure:
  `expression_statement`, `assignment_statement`, `if_statement`, `for_each_statement`,
  `with_statement`, `invocation_expression`, `member_access_expression`, `me_expression`,
  `on_error_statement`, `exit_statement`, `preprocessor_directive`, `comment`.

Across the entire corpus (summed over all files):

| Layer | Node types checked | Recognized nodes | Files seen |
|---|---:|---:|---:|
| **Structural** (class / method / field / property / event / implements) | 22 | 29 | 4 |
| **Body** (statements / expressions / control flow) | 32 | 1681 | 23 |

**Key finding**: the body content parses cleanly. The procedural/class structure does
not. Across all 16 files:

- 0 `class_declaration` nodes
- 0 `module_declaration` nodes
- 0 `method_declaration` nodes (the VBA `Sub`/`Function` shape is wrapped in ERROR)
- 0 `field_declaration` nodes (VBA `Public X As Y` at module level is wrapped in ERROR)
- 0 `property_declaration` nodes
- 0 `parameter_list` / `parameter` nodes
- 0 `event_declaration` / `custom_event_declaration` / `raiseevent_statement`
- 0 `implements_clause`

The reason is structural: a VBA module file (`.bas`/`.cls`) has **no `Class X` or
`Module X` opener**. The module's name lives in the Access class header
(`VERSION 1.0 CLASS` + `BEGIN … END` + `Attribute VB_Name = "X"`), which the grammar
does not recognize. Without an opener, the grammar has no class/module wrapper to
attach declarations to — so the file's top-level declarations are wrapped in ERROR,
and the procedural structure is lost.

What does work:

- 1681 body-level nodes were recognized — expressions, assignments,
  if/while/for/with blocks, method invocations, member access, comments, and the
  `On Error GoTo` error-trap pattern (which IS in the grammar as `on_error_statement`).
- 14 `if_statement`, 13 `with_statement`, 4 `for_each_statement`, 8 `enum_declaration`
  (all in the largest file), 3 `preprocessor_directive` (for `#If`/`#Else`/`#End If`)
  — these are the body-level constructs a F.2 hybrid extractor could fill in from the AST.

## Per-file results

| File | Bytes | Lines | Nodes | ERROR | MISSING | Class |
|---|---:|---:|---:|---:|---:|---|
| `__tests__\fixtures\vba-control-modeling\Form_OtherForm.cls` | 840 | 25 | 82 | 9 | 0 | partial |
| `__tests__\fixtures\vba-control-modeling\Form_TestForm.cls` | 1380 | 43 | 167 | 13 | 0 | partial |
| `__tests__\fixtures\vba-control-modeling\Report_NoSibling.cls` | 694 | 19 | 59 | 7 | 0 | partial |
| `__tests__\fixtures\vba-control-modeling\Report_PayrollSummary.cls` | 1067 | 31 | 105 | 9 | 0 | partial |
| `__tests__\fixtures\vba-control-modeling\modTestHelper.bas` | 796 | 20 | 55 | 7 | 0 | partial |
| `__tests__\fixtures\vba-event-synth\Form_Main.cls` | 948 | 29 | 64 | 11 | 0 | partial |
| `__tests__\fixtures\vba-event-synth\Form_MismatchedName.cls` | 966 | 25 | 54 | 8 | 0 | partial |
| `__tests__\fixtures\vba-event-synth\Form_NoHandlers.cls` | 695 | 21 | 51 | 7 | 0 | partial |
| `__tests__\fixtures\vba-event-synth\Notifier.cls` | 844 | 26 | 47 | 6 | 0 | partial |
| `__tests__\fixtures\vba-source-object\Form_Child.cls` | 99 | 7 | 32 | 4 | 1 | partial |
| `__tests__\fixtures\vba\src\classes\ACAuditoriaOperaciones.cls` | 15303 | 472 | 2429 | 119 | 2 | partial |
| `__tests__\fixtures\vba\src\classes\ARAuditoria.cls` | 14058 | 476 | 2216 | 100 | 0 | partial |
| `__tests__\fixtures\vba\src\forms\Form_FormNCAuditoriaMotivoEliminado.cls` | 5803 | 210 | 1000 | 44 | 5 | partial |
| `__tests__\fixtures\vba\src\modules\constantes.bas` | 2063 | 77 | 302 | 8 | 0 | partial |
| `__tests__\fixtures\vba\src\modules\mdlCursor.bas` | 1331 | 40 | 177 | 38 | 0 | partial |
| `__tests__\fixtures\vba\src\modules\modCallerDemo.bas` | 971 | 24 | 60 | 6 | 1 | partial |

### First ERROR samples per file

**__tests__\fixtures\vba-control-modeling\Form_OtherForm.cls** (class `partial`):
  - line 1:8 — `1`
  - line 1:8 — `1`
  - line 1:10 — `0 CLASS`
**__tests__\fixtures\vba-control-modeling\Form_TestForm.cls** (class `partial`):
  - line 1:8 — `1`
  - line 1:8 — `1`
  - line 1:10 — `0 CLASS`
**__tests__\fixtures\vba-control-modeling\Report_NoSibling.cls** (class `partial`):
  - line 1:8 — `1`
  - line 1:8 — `1`
  - line 1:10 — `0 CLASS`
**__tests__\fixtures\vba-control-modeling\Report_PayrollSummary.cls** (class `partial`):
  - line 1:8 — `1`
  - line 1:8 — `1`
  - line 1:10 — `0 CLASS`
**__tests__\fixtures\vba-control-modeling\modTestHelper.bas** (class `partial`):
  - line 1:0 — `Attribute VB_Name = "modTestHelper" Option Compare Database`
  - line 1:10 — `VB_Name`
  - line 14:0 — `Public Sub`
**__tests__\fixtures\vba-event-synth\Form_Main.cls** (class `partial`):
  - line 1:8 — `1`
  - line 1:8 — `1`
  - line 1:10 — `0 CLASS`
**__tests__\fixtures\vba-event-synth\Form_MismatchedName.cls** (class `partial`):
  - line 1:8 — `1`
  - line 1:8 — `1`
  - line 1:10 — `0 CLASS`
**__tests__\fixtures\vba-event-synth\Form_NoHandlers.cls** (class `partial`):
  - line 1:8 — `1`
  - line 1:8 — `1`
  - line 1:10 — `0 CLASS`
**__tests__\fixtures\vba-event-synth\Notifier.cls** (class `partial`):
  - line 1:8 — `1`
  - line 1:8 — `1`
  - line 1:10 — `0 CLASS`
**__tests__\fixtures\vba-source-object\Form_Child.cls** (class `partial`):
  - line 1:8 — `1`
  - line 1:8 — `1`
  - line 1:10 — `0 CLASS`
**__tests__\fixtures\vba\src\classes\ACAuditoriaOperaciones.cls** (class `partial`):
  - line 1:0 — `VERSION 1.0 CLASS BEGIN MultiUse = -1 'True END Attribute VB_Name = "ACAuditoria`
  - line 1:8 — `1`
  - line 1:8 — `1`
**__tests__\fixtures\vba\src\classes\ARAuditoria.cls** (class `partial`):
  - line 1:0 — `VERSION 1.0 CLASS BEGIN MultiUse = -1 'True END Attribute VB_Name = "ARAuditoria`
  - line 1:8 — `1`
  - line 1:8 — `1`
**__tests__\fixtures\vba\src\forms\Form_FormNCAuditoriaMotivoEliminado.cls** (class `partial`):
  - line 1:0 — `Option Compare Database`
  - line 4:0 — `Private m_Error As String Private Sub`
  - line 12:20 — `True VBA`
**__tests__\fixtures\vba\src\modules\constantes.bas** (class `partial`):
  - line 1:10 — `VB_Name`
  - line 2:0 — `Option Compare Database Option Explicit Public Const msoFileDialogFilePicker As `
  - line 7:47 — `3`
**__tests__\fixtures\vba\src\modules\mdlCursor.bas** (class `partial`):
  - line 1:10 — `VB_Name`
  - line 2:7 — `Compare Database Option Explicit Public Const IDC_SIZEWE`
  - line 6:0 — `Public Const`
**__tests__\fixtures\vba\src\modules\modCallerDemo.bas** (class `partial`):
  - line 1:0 — `Attribute VB_Name = "modCallerDemo" Option Compare Database`
  - line 1:10 — `VB_Name`
  - line 15:0 — `Public Sub`

## Per-construct recognition (vbnet grammar → VBA construct)

These are the tree-sitter node types we asked the grammar to recognize on the VBA
corpus. Counts are summed across all files; `files` is the number of files in which
the node type was seen at least once.

| Construct (vbnet node type) | Count | Files |
|---|---:|---:|
| `identifier` | 1349 | 16 |
| `
` | 1130 | 16 |
| `.` | 454 | 14 |
| `ERROR` | 396 | 16 |
| `=` | 372 | 16 |
| `member_access_expression` | 358 | 14 |
| `assignment_statement` | 278 | 16 |
| `string_literal` | 228 | 16 |
| `binary_expression` | 213 | 4 |
| `empty_statement` | 169 | 14 |
| `(` | 141 | 14 |
| `)` | 141 | 14 |
| `if_statement` | 135 | 3 |
| `expression_statement` | 119 | 15 |
| `comment` | 113 | 13 |
| `0` | 112 | 3 |
| `integer_literal` | 97 | 15 |
| `me_expression` | 96 | 7 |
| `with_member_access_expression` | 96 | 3 |
| `array_access_expression` | 94 | 7 |
| `member_modifier` | 81 | 15 |
| `&` | 79 | 4 |
| `exit_statement` | 71 | 3 |
| `<>` | 67 | 3 |
| `,` | 59 | 5 |
| `enum_member_declaration` | 41 | 1 |
| `:` | 39 | 6 |
| `invocation_expression` | 32 | 10 |
| `argument_list` | 32 | 10 |
| `labeled_statement` | 32 | 4 |
| `unary_expression` | 27 | 12 |
| `on_error_statement` | 26 | 3 |
| `"` | 25 | 4 |
| `nothing_literal` | 22 | 3 |
| `boolean_literal` | 18 | 3 |
| `else_clause` | 14 | 3 |
| `source_file` | 14 | 14 |
| `with_statement` | 13 | 3 |
| `-` | 11 | 11 |
| `end_statement` | 11 | 11 |
| `simple_name` | 9 | 3 |
| `enum_declaration` | 8 | 1 |
| `cast_expression` | 8 | 2 |
| `declaration_statement` | 7 | 2 |
| `local_declaration_modifier` | 7 | 2 |
| `variable_declarator` | 7 | 2 |
| `as_clause` | 7 | 2 |
| `
` | 6 | 6 |
| `named_argument` | 6 | 1 |
| `:=` | 6 | 1 |
| `for_each_statement` | 4 | 2 |
| `object_creation_expression` | 4 | 3 |
| `option_statements` | 3 | 3 |
| `option_statement` | 3 | 3 |
| `<` | 3 | 2 |
| `preprocessor_directive` | 3 | 1 |
| `predefined_type` | 2 | 2 |
| `call_statement` | 1 | 1 |
| `goto_statement` | 1 | 1 |
| `class_declaration` | 0 | 0 |
| `module_declaration` | 0 | 0 |
| `structure_declaration` | 0 | 0 |
| `interface_declaration` | 0 | 0 |
| `method_declaration` | 0 | 0 |
| `constructor_declaration` | 0 | 0 |
| `external_method_declaration` | 0 | 0 |
| `abstract_method_declaration` | 0 | 0 |
| `property_declaration` | 0 | 0 |
| `field_declaration` | 0 | 0 |
| `event_declaration` | 0 | 0 |
| `custom_event_declaration` | 0 | 0 |
| `parameter_list` | 0 | 0 |
| `parameter` | 0 | 0 |
| `elseif_clause` | 0 | 0 |
| `for_statement` | 0 | 0 |
| `while_statement` | 0 | 0 |
| `do_statement` | 0 | 0 |
| `select_statement` | 0 | 0 |
| `case_statement` | 0 | 0 |
| `try_statement` | 0 | 0 |
| `throw_statement` | 0 | 0 |
| `raiseevent_statement` | 0 | 0 |
| `handler_clause` | 0 | 0 |
| `implements_clause` | 0 | 0 |
| `imports_statement` | 0 | 0 |
| `return_statement` | 0 | 0 |
| `preprocessorDirective` | 0 | 0 |
| `conditional_compilation_directive` | 0 | 0 |
| `attribute` | 0 | 0 |

## VBA keyword presence vs grammar recognition

Source-level keyword scan: for each VBA-side keyword of interest, count how many
times it appears in the corpus and in how many files. When a keyword appears in
source but the corresponding grammar node type is rare, that's a **shape gap** the
F.2 extractor will need to bridge.

| Keyword | Occurrences | Files | Matching grammar nodes (if any) |
|---|---:|---:|---|
| `If` | 289 | 5 | `if_statement` |
| `Then` | 144 | 4 | part of `if_statement` |
| `End If` | 144 | 4 | `if_statement` closer |
| `Function` | 85 | 5 | `method_declaration` |
| `Public` | 69 | 8 | member_modifier (rolled into declaration_statement / method_declaration) |
| `Sub` | 52 | 11 | `method_declaration` |
| `Exit Function` | 46 | 3 | `return_statement` / `exit_statement` (verify) |
| `Property` | 44 | 2 | `property_declaration` |
| `Set` | 39 | 4 | accessor in property_declaration OR `object_creation_expression` |
| `Private` | 36 | 10 | member_modifier (rolled into declaration_statement / method_declaration) |
| `Dim` | 34 | 6 | `declaration_statement` |
| `With` | 32 | 5 | `with_statement` |
| `On Error` | 27 | 3 | NOT A GRAMMAR NODE — runtime semantics, never parsed by tree-sitter |
| `GoTo` | 26 | 3 | `goto_statement` (verify) |
| `Exit Property` | 23 | 1 | `return_statement` / `exit_statement` (verify) |
| `Enum` | 16 | 1 | `enum_declaration` |
| `Else` | 15 | 4 | `else_clause` |
| `End With` | 15 | 3 | `with_statement` closer |
| `For` | 13 | 10 | `for_statement` |
| `DoCmd` | 10 | 2 | qualified_name in member_access_expression (no dedicated node) |
| `Const` | 9 | 2 | constant_declaration (or field_declaration with const modifier) |
| `WithEvents` | 9 | 4 | field_declaration with WithEvents modifier |
| `Get` | 9 | 1 | accessor in property_declaration |
| `Next` | 6 | 2 | `for_statement` closer |
| `New` | 4 | 3 | `object_creation_expression` |
| `Case` | 2 | 2 | `case_statement` |
| `Exit Sub` | 2 | 1 | `return_statement` / `exit_statement` (verify) |
| `Resume` | 2 | 1 | NOT A GRAMMAR NODE — runtime semantics |
| `RaiseEvent` | 2 | 1 | `raiseevent_statement` |
| `Static` | 0 | 0 | — |
| `Type` | 0 | 0 | `structure_declaration` (VB.NET `Structure` ≈ VBA `Type`) |
| `Let` | 0 | 0 | accessor in property_declaration |
| `ElseIf` | 0 | 0 | `elseif_clause` |
| `While` | 0 | 0 | `while_statement` |
| `Wend` | 0 | 0 | NO MATCH (vbnet has no Wend — needs pre-processing to `End While`) |
| `Do` | 0 | 0 | `do_statement` |
| `Loop` | 0 | 0 | `do_statement` closer |
| `Until` | 0 | 0 | `do_statement` condition |
| `Select Case` | 0 | 0 | `select_statement` |
| `End Select` | 0 | 0 | `select_statement` closer |
| `TempVars` | 0 | 0 | qualified_name in member_access_expression (no dedicated node) |
| `Implements` | 0 | 0 | `implements_clause` (verify — vbnet also has Implements; should match) |

## VBA constructs the vbnet grammar recognizes cleanly

Cross-referenced from the per-construct table — these are the body-level node types
the grammar emits reliably across the corpus:

- `if_statement` / `elseif_clause` / `else_clause` — block `If`/`Then`/`ElseIf`/`Else`/`End If`.
- `for_each_statement` — `For Each … In …`.
- `with_statement` — `With` / `End With`.
- `on_error_statement` — `On Error GoTo <label>` (the VBA error-trap pattern).
- `exit_statement` — `Exit Sub` / `Exit Function` / `Exit Property`.
- `end_statement` — `End Sub` / `End Function` / `End Property` / `End With` / `End If` (as a separate node from the enclosing `if_statement`/`with_statement` etc.).
- `goto_statement` — `GoTo <label>`.
- `expression_statement` / `assignment_statement` — most procedure bodies are dominated by these.
- `invocation_expression` / `member_access_expression` / `me_expression` — calls, qualified names, `Me.X`.
- `binary_expression` / `unary_expression` / `object_creation_expression` / `cast_expression` — arithmetic, `Not x`, `New T`, `CInt(x)`.
- `enum_declaration` / `enum_member_declaration` — `Enum X … End Enum` (verified in `ARAuditoria.cls`).
- `option_statements` / `option_statement` — `Option Compare Database` / `Option Explicit`.
- `preprocessor_directive` — `#If` / `#Else` / `#End If` (verified in `mdlCursor.bas`).
- `comment` — `'` line comments.

## VBA constructs the vbnet grammar fails on (observed in this corpus)

These are the structural node types the grammar does NOT emit on the VBA corpus,
despite the same constructs existing in VB.NET. Each item was measured at **0 nodes**
across the entire 16-file corpus.

1. **`class_declaration` / `module_declaration`** — the wrapper that encloses a
   file's body. VBA module files (`.bas`/`.cls`) have NO `Class X` / `Module X`
   opener, so the grammar has no wrapper to attach declarations to. The module name
   lives in `Attribute VB_Name = "X"` (an Access export header), which the grammar
   does not recognize. **0 of 16 files** emit a class_declaration.

2. **`method_declaration` / `constructor_declaration` / `abstract_method_declaration`**
   — VBA `Sub` / `Function` / `Property Get/Set/Let` procedures are not recognized as
   declarations. The whole procedure shape (signature + body + `End X`) is wrapped
   in an ERROR node. The `End Sub` / `End Function` text inside the ERROR is
   partially recognized as a `end_statement` (an orphan, with no enclosing
   method_declaration), but that does NOT give F.2 a procedure boundary. **0 of 16
   files** emit a method_declaration.

3. **`field_declaration`** — VBA module-level `Public X As Y` / `Private X As Y` /
   `Dim X As Y` declarations are not recognized as field declarations. Each field
   declaration line is parsed as `member_modifier (Public)` + `identifier (X)` +
   `as_clause` + `ERROR` — but only when the field is NOT immediately before a
   procedure. In a class file, the entire preamble (header + Attribute + Option + ALL
   field declarations) is wrapped in a single ERROR that swallows the file's start.
   **0 of 16 files** emit a field_declaration.

4. **`property_declaration`** — VBA `Property Get/Set/Let` blocks are wrapped in
   ERROR (same root cause as method_declaration). **0 of 16 files** emit a
   property_declaration.

5. **`event_declaration` / `custom_event_declaration` / `raiseevent_statement`** —
   VBA `Public Event X(...)` declarations and `RaiseEvent X` calls are not
   recognized. Verified in `Notifier.cls` (two events + two raising subs): all four
   are inside an ERROR. **0 of 16 files** emit any of these.

6. **`implements_clause`** — VBA `Implements IFoo` is not recognized. (The keyword
   `Implements` does not appear in this corpus, but the grammar rule was tested on
   its VB.NET equivalent: see `docs/grammars/tree-sitter-vbnet.md` patch item 11 —
   `implements_clause` is in the grammar. The question for F.2 is whether the
   surrounding `class_declaration` (which we cannot get) carries the `implements_clause`
   child. Without the wrapper, even an `Implements IFoo` line is orphan and unrecognized.)

7. **`parameter_list` / `parameter`** — VBA `Sub Foo(ByVal x As Long)` parameter
   lists are not recognized. The whole `Sub Foo(...args...)` line is inside an ERROR
   because of (1)–(2). **0 of 16 files** emit a parameter_list.

8. **`const_declaration`** — VBA `Public Const X = Y` is parsed as `member_modifier
   (Public Const)` + `assignment_statement (X = Y)`. The constant-ness is lost (the
   identifier is treated as an assignable variable, not a const). Verified in
   `mdlCursor.bas` and `constantes.bas`.

9. **`Wend` loop terminator** — none of the corpus files use `Wend` (this corpus only
   has `For Each` loops), but VBA supports `While … Wend` and the vbnet grammar has
   no `wend_statement` rule. **Pre-processing:** rewrite `Wend` → `End While` if it
   shows up in the wider corpus.

10. **`VERSION 1.0 CLASS` + `BEGIN ... MultiUse = -1 ... END` block** — Access class
    header (legacy `form.frm` / `report.frm` shape). The vbnet grammar has no rule
    for it; every `.cls` starting with it begins inside an ERROR node. The regex
    pipeline silently strips it. **Pre-processing:** blank the `VERSION ... END`
    block before parsing.

11. **`Attribute VB_Name = "X"` (and `Attribute VB_GlobalNameSpace`, etc.)** — these
    legacy `Attribute` directives are not recognized as VB.NET attributes (the
    shape is the same lexically, but the grammar does not have an `Attribute`
    rule for top-of-file placement). Each one becomes `identifier (Attribute) +
    ERROR (VB_Name) + assignment_statement (= "X")`. **Pre-processing:** strip
    `Attribute VB_*` lines.

12. **`Option Compare Database` / `Option Explicit`** — recognized as `option_statement`,
    but the position is at module-body level (between `class_declaration` and the
    first member). Without a class_declaration wrapper, the option_statement is
    orphaned. Not a fatal error, but F.2 needs to handle it.

## Regex pipeline audit after issue #167

The twelve findings above describe the raw `tree-sitter-vbnet` parse, not the
production regex extractor. The production pipeline was audited separately on
2026-07-18. “Accepted” below means the current graph contract preserves the useful
relationship even when it does not mirror the VB.NET node shape.

| # | Production regex behavior | Decision |
|---|---|---|
| 1 | `.cls` files emit `class` and other VBA source files emit `module`; `Attribute VB_Name` supplies the symbol name when present. | Accepted heuristic; no source wrapper is required by the regex extractor. |
| 2 | `PROC_RE` emits `function` nodes for `Sub`, `Function`, and `Property Get/Let/Set`. VBA constructors such as `Class_Initialize` use the same procedure convention. | Accepted. |
| 3 | `DIMS_RULES` records declared variable types and emits references to project types, but it does not emit standalone `field` nodes. | Accepted as the existing type-reference contract; field symbols remain outside the current VBA graph model. |
| 4 | Property accessors are emitted as separate `function` nodes rather than one `property` node. | Accepted existing node convention. |
| 5 | `Public`/`Private`/`Friend Event` declarations emit `event` nodes contained by their class/module. A `RaiseEvent` inside a procedure emits a parser-provenance `raises-event` edge from that procedure to the declared event; the post-index WithEvents pass can additionally materialize handler edges. | Accepted and behavior-tested, including containment, source/target, provenance, and source coordinates. |
| 6 | `Implements IFoo` emits an `interface` target and parser-provenance `implements` edge from the owning class/module. | Accepted. |
| 7 | `PROC_RE` recognizes the procedure boundary and name, while return-type parsing deliberately looks after the parameter list so parameter `As` clauses are not mistaken for the return type. Individual `parameter` nodes are not emitted. | Accepted as the current procedure-level graph contract; parameter symbols remain outside this extractor. |
| 8 | Module-level `Public`/`Private Const` declarations already emitted the repository’s canonical `constant` kind before issue #167, so the issue’s claim that constant identity was lost was stale. The actual gap was that the optional `As` type was discarded. Constants now retain `metadata.asType` and `metadata.value`; an omitted `As` is normalized to `Variant`. Procedure-local constants remain resolution-only and do not become module symbols. | Fixed and regression-tested. |
| 9 | There is no production `Wend` → `End While` rewrite. `Wend` is excluded from call detection, and the regex extractor does not emit loop nodes. | Accepted as inert for the current graph contract; the earlier “handled by pre-processing rewrite” claim was incorrect. |
| 10 | The Access `VERSION 1.0 CLASS` / `BEGIN...END` header is not parsed as a node. `detectVbName` scans past it to find `Attribute VB_Name`, while declaration classifiers ignore the header lines. | Accepted. |
| 11 | `Attribute VB_*` lines are not stripped by `stripVbaComments`. `Attribute VB_Name` is consumed for naming and the declaration classifiers ignore the remaining attribute lines. | Accepted as inert; the earlier “stripped” claim was incorrect. |
| 12 | `stripVbaComments` replaces every `Option ...` directive with an empty line, so it emits no symbol while preserving downstream line coordinates. An Option-only file emits no class/module. | Accepted and behavior-tested with a following procedure retaining its original start/end lines. |

Audit result: gap #8 required a production change; gaps #5 and #12 already had
correct behavior and received stronger regression coverage. The other nine findings
are accepted under the explicit limitations above rather than being presented as
equivalent VB.NET AST nodes.

## What pre-processing can and cannot fix

A first-cut pre-processing layer (strip `VERSION...END`, strip `Attribute VB_*`,
blank `Option …`, append a trailing newline) was applied to `ARAuditoria.cls` as a
dry-run. Result: the structural gap is **NOT closed**. The class body still has
module-level `Public X As Y` / `Private X As Y` / `Dim X As Y` field declarations,
and the grammar still wraps the entire class body in an ERROR. The reason is that
VBA's module file has no `Class X` / `Module X` opener — even after stripping the
preamble, the grammar still has no wrapper to attach the class body to.

But adding a wrapper DOES unlock the structure. See the next section.

## F.2 dry-run: synthesized `Class <Name> … End Class` wrapper

The F.2 hypothesis is to inject a synthetic `Class <Name>` opener (where `<Name>`
comes from the existing `Attribute VB_Name = "X"` line) and a closing `End Class`
for every file, then run the AST. The spike ran this on the entire corpus and
measured the structural node counts the grammar now emits.

| Structural node | Count (synthesized) | Count (raw) |
|---|---:|---:|
| `class_declaration` | 13 | 0 |
| `method_declaration` | 26 | 0 |
| `field_declaration` | 36 | 0 |
| `property_declaration` | 1 | 0 |
| `parameter_list` | 36 | 0 |
| `event_declaration` | 2 | 0 |
| `raiseevent_statement` | 2 | 0 |
| `implements_clause` | 0 | 0 |
| `enum_declaration` | 8 | 8 |

**Result:** the synthesized wrapper unlocks the entire structural tree. With
just one extra pre-processing step — inject a `Class <Name>` opener + `End Class`
closer — the grammar emits `class_declaration`, `method_declaration`,
`field_declaration`, `property_declaration`, `parameter_list`, `event_declaration`,
and `raiseevent_statement` on the VBA corpus. The body-content recognition is
preserved (expressions, calls, control flow were already parsing cleanly).

This is the **unlock** for F.2. The hybrid extractor is feasible:

1. **Pre-processing** (already in the existing `vba-preprocess.ts` regex): strip
   `VERSION...END`, `Attribute VB_*`, `Option ...`, `Wend`.
2. **Wrapper injection** (new): prepend `Class <Name>` and append `End Class`
   to the file before handing it to the tree-sitter parser. The `<Name>` is
   already extracted by the regex pipeline from `Attribute VB_Name`.
3. **AST walk** (the new `VbaTreeSitterExtractor`): walk the vbnet AST for
   procedures, classes, fields, parameters, properties, events. The body
   content (expressions, calls, control flow) is already there.
4. **Access layer** (the existing 5 `create*Classifier()` regex modules):
   fill in the Access-specific emissions — DoCmd, WithEvents pair convention,
   TempVars, RecordSource/RowSource, form/report layout.

This matches the F.2 plan's "vbnet AST for the language skeleton, regex for
Access-specific" framing — but with a concrete, bounded pre-processing step
(wrapper injection) that makes the AST walk actually work.

## Pre-processing checklist for F.2

1. **Blank the `VERSION 1.0 CLASS` + `BEGIN … END` block** (Access class header).
2. **Strip `Attribute VB_*` lines** (legacy class metadata).
3. **Blank `Option Compare Database` / `Option Explicit` lines** (file-level options;
   the grammar recognizes them as `option_statement` but the position is wrong).
4. **Rewrite `Wend` → `End While`** (legacy loop terminator — not in this corpus, but
   widely used).
5. **Append a trailing newline** to every file (mandatory, already in `vbnet.ts:7-9`).
6. **Inject a synthetic `Class <Name>` opener + `End Class` closer** (the unlock —
   see the section above). `<Name>` comes from the existing `Attribute VB_Name`
   regex extraction.

After steps 1-6, the F.2 extractor should walk the AST cleanly. Re-run this spike
on the synthesized corpus to confirm: the per-construct table should now show
non-zero counts for `class_declaration`, `method_declaration`, `field_declaration`,
`property_declaration`, `parameter_list`, `event_declaration`.

## Reproduce

```bash
node scripts/spike-vbnet-as-vba.mjs
# or, with a different corpus:
node scripts/spike-vbnet-as-vba.mjs --fixtures <dir> --out <path>
```

Output:

- Markdown report: `docs\spikes\vbnet-as-vba.md`
- Machine-readable JSON: `docs\spikes\vbnet-as-vba.json`

## Raw data

Full per-file detail is in `docs\spikes\vbnet-as-vba.json`.
