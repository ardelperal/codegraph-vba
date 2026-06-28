# Tasks: vba-extractor — VBA / Access Language Support (regex extractors)

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | ~1,560 (prod ~560 + tests ~1,000) |
| 400-line budget risk | Low |
| Chained PRs recommended | No |
| Suggested split | Single PR (user override on review budget) |
| Delivery strategy | single-pr-default |
| Chain strategy | pending |

Decision needed before apply: No
Chained PRs recommended: No
Chain strategy: pending
400-line budget risk: Low

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | T1: wiring + detect helper + types | PR 1 | Self-contained; no extractor logic |
| 2 | T2: pre-processing helpers | PR 1 | Pure functions, fully TDD'd |
| 3 | T3: VbaExtractor | PR 1 | Core extractor, 18 scenarios |
| 4 | T4: VbaFormExtractor | PR 1 | Form UI extractor, 6 scenarios |
| 5 | T5: server-instructions + E2E regression | PR 1 | Agent guidance + regression gate |

## Phase 1: T1 — Wiring, Types, Extension Detection

**Tests first (`__tests__/detect-vba-form-file.test.ts`):**

- [ ] T1.1 Write 3–5 vitest cases for `detectVbaFormFile()`:
  - `Form_Main.form.txt` → true
  - `Report_Orders.report.txt` → true
  - `Form_Main.form` (single segment) → false
  - `Document.txt` (non-VBA) → false
  - `Form_Main.form.txt` with backslash paths → true (Windows)
  - **RED**: `npm test -- __tests__/detect-vba-form-file.test.ts` → 3–5 failing cases

**Implementation:**

- [ ] T1.2 Add `'vba'` to `LANGUAGES` tuple in `src/types.ts` (before `'unknown'`)
- [ ] T1.3 Add `'vba'` to `EXTENSION_MAP` for `.bas`, `.cls`, `.frm`, `.dsr` in `src/extraction/grammars.ts`
- [ ] T1.4 Add `detectVbaFormFile(filePath: string): boolean` helper to `src/extraction/grammars.ts` (checks `.form.txt` / `.report.txt` two-segment extension)
- [ ] T1.5 Call `detectVbaFormFile()` BEFORE `path.extname()` in `detectLanguage()` — if true, return `'vba'`
- [ ] T1.6 Call `detectVbaFormFile()` in `isSourceFile()` — same pre-check pattern
- [ ] T1.7 Add `'vba'` to `isLanguageSupported()` → `true` and `isGrammarLoaded()` → `true` in `grammars.ts`
- [ ] T1.8 Add `'vba': 'VBA / Access'` to `getLanguageDisplayName()` in `grammars.ts`
- [ ] T1.9 Add `'vba'` to `getSupportedLanguages()` return array in `grammars.ts`
- [ ] T1.10 Import and dispatch both classes in `extractFromSource()` in `src/extraction/tree-sitter.ts` (two new branches: `VbaExtractor` for `.bas/.cls/.frm/.dsr`, `VbaFormExtractor` for `.form.txt/.report.txt`)

**GREEN**: `npm test -- __tests__/detect-vba-form-file.test.ts` → all pass; `npm test` → no regressions

**Commit**: `feat(vba): wire VBA language into types, grammars, and extractFromSource dispatch` — SDD: vba-extractor

**Verification**: `npm run build && npm test` (no regressions in existing tests)

---

## Phase 2: T2 — Pre-processing Helpers

**Tests first (`__tests__/extraction-vba-preprocess.test.ts`):**

- [ ] T2.1 Write 12 vitest cases for `joinLineContinuations`:
  - `Sub Foo()\n  DoCmd.RunSQL _\n    "SELECT * FROM tbl"` → single joined line
  - `Debug.Print _\n  "hello"` → joined
  - No continuation → unchanged
  - Multiple continuations in one statement
  - Continuation inside a string literal (must NOT join across string boundary)
  - Empty string after continuation
  - **RED**: `npm test -- __tests__/extraction-vba-preprocess.test.ts -t "joinLineContinuations"` → 12 failing

- [ ] T2.2 Write 12 vitest cases for `stripVbaComments`:
  - `Dim x As Long ' comment` → comment stripped, code preserved
  - `Rem LegacyComment` → stripped
  - `Debug.Print "It's a test"` → string with `'` inside preserved
  - SQL inside `'...` comment → NOT stripped (precedes comment removal)
  - `Option Explicit` → stripped
  - Multi-line: `'` on one line, code on next → only comment line affected
  - **RED**: `npm test -- __tests__/extraction-vba-preprocess.test.ts -t "stripVbaComments"` → 12 failing

- [ ] T2.3 Write 12 vitest cases for `extractStringLiterals`:
  - `DoCmd.RunSQL "SELECT * FROM tbl"` → returns one literal `{text, line, column}`
  - Multiple literals in one statement → returns all
  - No literals → returns `[]`
  - Literal with escaped `"` inside (`""`) → preserved
  - Literal spanning `&` concatenation → each segment separate
  - Literal inside a comment → still returned (comment already stripped before this runs)
  - **RED**: `npm test -- __tests__/extraction-vba-preprocess.test.ts -t "extractStringLiterals"` → 12 failing

**Implementation:**

- [ ] T2.4 Create `src/extraction/vba-preprocess.ts` with `joinLineContinuations(src: string): string` — joins lines ending with `_`
- [ ] T2.5 Add `stripVbaComments(src: string): string` — character-by-character walk, handles both `'` and `Rem `, skips content inside double-quoted strings
- [ ] T2.6 Add `extractStringLiterals(src: string): Array<{text: string, line: number, column: number}>` — walks source, records every `"..."` span
- [ ] T2.7 Export all three from `vba-preprocess.ts` for unit testing

**GREEN**: All 36 tests pass (`npm test -- __tests__/extraction-vba-preprocess.test.ts`)

**Commit**: `feat(vba): add pre-processing pipeline (joinLineContinuations, stripVbaComments, extractStringLiterals)` — SDD: vba-extractor — TDD: 36 passing preprocess tests

**Verification**: `npm test -- __tests__/extraction-vba-preprocess.test.ts` → 36/36 green

---

## Phase 3: T3 — VbaExtractor

**Tests first (`__tests__/extraction-vba.test.ts`):**

Write 18 vitest scenarios grouped by regex (TDD red→green per group, loop tight):

**Group A — Module/Class Nodes + VB_Name:**

- [ ] T3.1 `.bas` with `Public Sub Foo()` → emits `module` + `function` node; visibility = 'Public' — **RED**
- [ ] T3.2 `.bas` with `Private Function Calc()` → visibility = 'Private' — **RED**
- [ ] T3.3 `.bas` with `Property Get Name()` → emits `function` node named `Name` — **RED**
- [ ] T3.4 `.cls` with `Public Sub Foo()` → emits `class` node + `function` node + `contains` edge class→function — **RED**
- [ ] T3.5 No `Attribute VB_Name` → module/class name = file basename — **RED**
- [ ] T3.6 `Attribute VB_Name = "modHelpers"` → module name = `modHelpers` — **RED**

**Group B — Sub New marker:**

- [ ] T3.7 `.cls` with `Public Sub New()` → class node has `metadata.hasClassInitializer = true`, `metadata.initializerName = 'New'` — **RED**
- [ ] T3.8 `.cls` with `Private Sub New()` → same marker set — **RED**
- [ ] T3.9 `.cls` without `Sub New` → no `hasClassInitializer` in metadata — **RED**

**Group C — Call edges:**

- [ ] T3.10 `Sub Outer()` calls `Inner()` (both in same `.bas`) → `calls` edge Outer→Inner, no `synthesizedBy` — **RED**
- [ ] T3.11 `modHelpers.CalcTotal` where `modHelpers` not in file → `calls` edge with `provenance: 'heuristic'` + `synthesizedBy: 'vba-name-resolution'` — **RED**
- [ ] T3.12 `UnknownExternal.Whatever` → silent (no edge emitted, no error) — **RED**

**Group D — Implements, Dim, WithEvents:**

- [ ] T3.13 `.cls` with `Implements IFoo` → `implements` edge class→IFoo — **RED**
- [ ] T3.14 `Dim m_Calc As CalcEngine.Helper` → `references` edge to `CalcEngine` with `synthesizedBy: 'vba-name-resolution'` — **RED**
- [ ] T3.15 `Dim m_Count As Long` → no `references` edge — **RED**
- [ ] T3.16 `WithEvents m_Form As Form_Main` → `references` edge with `synthesizedBy: 'vba-withevents'` — **RED**

**Group E — SQL in strings:**

- [ ] T3.17 `DoCmd.RunSQL "SELECT * FROM tblCustomers"` → `references` edge to `tblCustomers` with `synthesizedBy: 'vba-sql-table'` — **RED**
- [ ] T3.18 `' DoCmd.RunSQL "SELECT * FROM tblFake"` (VBA comment) → no edge to `tblFake` (comment stripped before SQL scan) — **RED**

**Implementation:**

- [ ] T3.19 Create `src/extraction/vba-extractor.ts` with `VbaExtractor` class (constructor + `extract()` → `ExtractionResult`):
  - File node (kind `file`)
  - Pre-process pipeline: `joinLineContinuations` → `stripVbaComments` → `extractStringLiterals`
  - `Attribute VB_Name` parse → module (`.bas`) or class (`.cls`) node
  - R1/R2 sweep: `Sub`/`Function`/`Property` regex → `function` nodes with `metadata.visibility`
  - R3: detect `Sub New` in `.cls` → set `hasClassInitializer`/`initializerName` on class node
  - R4: call-site regex → `calls` edges (same-file plain, cross-module + `synthesizedBy: 'vba-name-resolution'`)
  - R5: `Implements` regex → `implements` edges
  - R6: qualified `Dim As` regex → `references` edges with `synthesizedBy: 'vba-name-resolution'`
  - R7: `WithEvents` regex → `references` edges with `synthesizedBy: 'vba-withevents'`
  - R8: iterate `extractStringLiterals`, scan SQL wrappers → `references` edges with `synthesizedBy: 'vba-sql-table'`
  - R9: if `filePath.endsWith('.form.txt')` or `filePath.endsWith('.report.txt')` → return empty result
  - R10: skip `Option` directive lines before R1 sweep
  - R11: `generateNodeId()` with qualified name including `$`/`%` suffix for overloads

**GREEN**: All 18 scenarios pass (`npm test -- __tests__/extraction-vba.test.ts`)

**Commit**: `feat(vba): implement VbaExtractor for .bas/.cls with 18 spec scenarios` — SDD: vba-extractor — TDD: 18/18 green

**Verification**: `npm test -- __tests__/extraction-vba.test.ts` → 18/18 green; `npm run build` → no TS errors

---

## Phase 4: T4 — VbaFormExtractor

**Tests first (`__tests__/extraction-vba-form.test.ts`):**

- [ ] T4.1 `Form_Main.form.txt` with `Attribute VB_Name = "Form_Main"` + one TextBox → emits `module` node named `Form_Main` + `property` node with `metadata.controlType = 'TextBox'` + `references` edge to sibling class node — **RED**
- [ ] T4.2 `Form_Main.form.txt` no VB_Name → module name = `Form_Main` (basename) — **RED**
- [ ] T4.3 `Form_Main.form.txt` with TextBox + CommandButton → two `property` nodes with correct `controlType` values — **RED**
- [ ] T4.4 `Report_Orders.report.txt` with TextBox → behaves identically to form — **RED**
- [ ] T4.5 `Form_Main.form.txt` containing literal `Sub Form_Load()` text → zero `function`/`class` nodes emitted — **RED**
- [ ] T4.6 Empty `.form.txt` → zero `property` nodes, one `module` node, one `references` edge — **RED**

**Implementation:**

- [ ] T4.7 Create `src/extraction/vba-form-extractor.ts` with `VbaFormExtractor` class:
  - Import `stripVbaComments` from `vba-preprocess.ts` (form text has no SQL, no procedure bodies)
  - File node (kind `file`)
  - `stripVbaComments()` on source (no line-continuation needed; `.form.txt` has none)
  - `Attribute VB_Name` parse → `module` node (NOT `class` — form modules are `module`)
  - `references` edge from module → sibling `.cls` basename (R12/R14)
  - Control-block regex: `Begin\s+(\w+)\s*$` + next-line type read → `property` node with `metadata.controlType`
  - R15: never run procedure sweep — return empty `function`/`class` nodes

**GREEN**: All 6 scenarios pass (`npm test -- __tests__/extraction-vba-form.test.ts`)

**Commit**: `feat(vba): implement VbaFormExtractor for .form.txt/.report.txt with 6 spec scenarios` — SDD: vba-extractor — TDD: 6/6 green

**Verification**: `npm test -- __tests__/extraction-vba-form.test.ts` → 6/6 green; `npm run build` → no TS errors

---

## Phase 5: T5 — Agent Guidance + E2E Regression

**Implementation:**

- [ ] T5.1 Add one paragraph to `src/mcp/server-instructions.ts` after the "Limitations" section (or as a new section):
  > **VBA / Access** — Dysflow exports Access/VBA source as `.bas`/`.cls`/`.form.txt`/`.report.txt`. Codegraph extracts `.bas`/`.cls` as `module`/`class`/`function` nodes (with `calls`, `implements`, `references` edges and five kinds of heuristic edges tagged `synthesizedBy`). `.form.txt`/`.report.txt` are extracted as `module` + `property` nodes only — **no** `function`/`sub`/`class` nodes come from form files; the sibling `.cls` holds the canonical code. Pass `projectPath` to a codegraph index that includes VBA files.
- [ ] T5.2 Run E2E regression (per obs #14702): index the main codegraph checkout with the new fork build, compare non-VBA node/edge counts against upstream baseline — diff must be zero for non-VBA languages.
- [ ] T5.3 Verify: `npm run build && npm test` → all green

**Commit**: `docs(vba): add server-instructions paragraph for VBA/Access; add E2E regression` — SDD: vba-extractor

**Verification**: `npm run build && npm test` → 100% green; E2E regression: non-VBA counts match upstream baseline

---

## Implementation commits

| Commit | Work unit | SDD tasks | Verification | Access sync |
|---|---|---|---|---|
| `76e7454` | T1 wiring | T1.x | `npm test -- detect-vba-form-file` green (8/8); `npm test` no regressions | N/A |
| `60146e9` | T2 helpers | T2.x | `npm test -- extraction-vba-preprocess` green (36/36) | N/A |
| `176a667` | T3 VbaExtractor | T3.x | `npm test -- extraction-vba` green (22/22) | N/A |
| `1ba73f9` | T4 VbaFormExtractor | T4.x | `npm test -- extraction-vba-form` green (7/7) | N/A |
| `d189b50` | T5 server-instructions + E2E | T5.x | `npm test -- extraction-vba-e2e` green (2/2); E2E indexes codegraph_main with 0 VBA files | N/A |