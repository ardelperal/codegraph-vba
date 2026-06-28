# Design: VBA / Access Language Support (regex extractors)

## Technical Approach

Two new regex-based extractors — `VbaExtractor` (`.bas`/`.cls`/legacy `.frm`/`.dsr`) and `VbaFormExtractor` (`.form.txt`/`.report.txt`) — wired into `extractFromSource()` next to the existing `LiquidExtractor` / `DfmExtractor` dispatch chain. A shared pure-function pre-processing pipeline (`joinLineContinuations` → `stripVbaComments` → `extractStringLiterals`) runs **once** at the top of `VbaExtractor.extract()` so every downstream regex sees logical single-line statements and never matches text inside VBA comments. The pipeline is testable atom-by-atom; every spec scenario is one vitest case. Both extractors mirror the established `DfmExtractor` shape (constructor + `extract()`, `nodes`/`edges`/`errors` arrays), so `extractFromSource()` plumbing is unchanged.

## Architecture Decisions

### Decision: Pre-processing pipeline before regex

**Choice**: Three pure helpers run in order at the top of `VbaExtractor.extract()`. `VbaFormExtractor` reuses only `stripVbaComments` (form text has no SQL and no procedure bodies).
**Alternatives considered**: (a) per-regex inline string-stripping — duplicates logic, drifts. (b) tokenizer — overkill for 28 scenarios.
**Rationale**: VBA's `_` line-continuation is mandatory (real Dysflow-emitted code spans multi-line calls/Dims constantly). Stripping comments BEFORE regex runs also satisfies the SQL-in-comment spec scenario naturally. Each helper has its own vitest block (table below).

### Decision: `.form.txt` / `.report.txt` extension wiring

**Choice**: Add a `detectVbaFormFile()` helper invoked BEFORE `path.extname()` in both `detectLanguage()` and `isSourceFile()`. `.form.txt` and `.report.txt` are **two-segment** extensions that `path.extname()` collapses to `.txt` — verified empirically (`node -e "console.log(path.extname('Form_Main.form.txt'))"` → `.txt`). A naive `EXTENSION_MAP['.txt'] = 'vba'` would wrongly map every `.txt` file.
**Alternatives considered**: Rename Dysflow output to `.form.vba` / `.report.vba` (single-segment) — blocked; Dysflow owns that contract.
**Rationale**: Two-segment extensions are rare in codegraph (`.cshtml`/`.razor` are single-segment), so a single helper mirrors the existing `isShopifyLiquidJson()` precedent (`grammars.ts:143`).

### Decision: Sub New → metadata marker, NOT an edge

**Choice**: `Sub New()` (Public or Private) sets `metadata.hasClassInitializer = true` and `metadata.initializerName = 'New'` on the class node. The `Sub New` itself is still emitted as a `function` node with a `contains` edge.
**Alternatives considered**: `instantiates` edge from class → `Sub New` (semantically wrong: `Sub New` is the class, not its target).
**Rationale**: A class is its own initializer; the marker is metadata, not a graph hop. Matches how `extractor's metadata.hasClassInitializer` would surface in `codegraph_explore`.

### Decision: Overloaded-name disambiguation (Foo vs Foo$)

**Choice**: `generateNodeId(filePath, 'function', qualifiedName, line)` where `qualifiedName` includes the signature suffix when present (`Foo` vs `Foo$` vs `Foo%`) — preserving the literal character the source wrote. Two procedures with the same bare name on different lines get distinct IDs (line is in the hash).
**Alternatives considered**: Stripping `$`/`%` suffixes (loses the actual symbol).
**Rationale**: Same hashing pattern upstream uses for Swift overloaded methods. The `buildFlowFromNamedSymbols` query-side overload-aware matching already handles the rare `Foo` / `Foo$` ambiguity by qualified-name co-naming.

## Data Flow

```
extractFromSource(filePath, source, 'vba')
  ├─ detectLanguage() resolves '.bas'/'.cls'/'.frm'/'.dsr' → 'vba'
  │   OR detectVbaFormFile() pre-check → 'vba' for '.form.txt'/'.report.txt'
  │
  ├─ VbaExtractor branch (.bas/.cls/.frm/.dsr):
  │     source → joinLineContinuations() → stripVbaComments() → extractStringLiterals()
  │     → attribute VB_Name parse → module/class node
  │     → per-requirement regex sweep → emits nodes/edges + synthesizeBy metadata
  │
  └─ VbaFormExtractor branch (.form.txt/.report.txt):
        source → stripVbaComments()
        → attribute VB_Name parse → module node + single references edge to .cls sibling
        → control-block regex sweep → property nodes with metadata.controlType
```

## File Changes

| File | Action | Description |
|------|--------|-------------|
| `src/extraction/vba-extractor.ts` | Create | `VbaExtractor` class + 3 pure pre-processing helpers (exported for unit tests) |
| `src/extraction/vba-form-extractor.ts` | Create | `VbaFormExtractor` class; reuses `stripVbaComments` from vba-extractor |
| `src/extraction/tree-sitter.ts` | Modify | Add 2 dispatch branches in `extractFromSource()`; import both classes |
| `src/extraction/grammars.ts` | Modify | `'vba'` to `LANGUAGES`; `.bas/.cls/.frm/.dsr` to `EXTENSION_MAP`; `detectVbaFormFile()` helper; `isLanguageSupported('vba')` and `isGrammarLoaded('vba')` return true; `getLanguageDisplayName` adds `'VBA / Access'`; `getSupportedLanguages` adds `'vba'` |
| `src/types.ts` | Modify | `'vba'` to `LANGUAGES` tuple before `'unknown'` |
| `src/mcp/server-instructions.ts` | Modify | One-paragraph addition to "Supported Languages" — explicit `.cls`/`.form.txt` split warning |
| `__tests__/extraction-vba.test.ts` | Create | 18 scenarios from `vba-code-extraction` spec (9 requirements × ~2 each) |
| `__tests__/extraction-vba-form.test.ts` | Create | 6 scenarios from `vba-form-ui-extraction` spec (3 requirements × 2) |
| `__tests__/extraction-vba-preprocess.test.ts` | Create | 12 unit tests for the 3 pre-processing helpers (incl. adversarial `_` joins, `Rem` comments, `It''s` quoted-string handling) |

## Pre-processing pipeline (pure functions)

| Helper | Input | Output | Spec scenarios it gates |
|---|---|---|---|
| `joinLineContinuations(src)` | Raw source | Source with `_`-suffixed lines joined to their predecessor | All 28 (multi-line Sub/Function bodies, multi-line Dims, multi-line SQL strings) |
| `stripVbaComments(src)` | Joined source | Same with `'`-to-EOL and `Rem `-prefixed lines removed (only outside string literals) | SQL-inside-comment scenario, Rem comment edge case |
| `extractStringLiterals(src)` | Comment-stripped source | Array of `{ text, line, column }` for SQL parser; source unchanged | All 3 SQL scenarios + the comment-stripping scenario |

Each helper is `< 30 LOC` and tested in isolation in `extraction-vba-preprocess.test.ts` BEFORE the extractor tests are written (TDD red → green).

## Per-requirement regex strategy

| # | Spec requirement | Regex (post-preprocess) | `synthesizedBy` |
|---|---|---|---|
| R1 | Sub/Function/Property in .bas | `^\s*((?:Public\|Private\|Friend\|Static)\s+)?(?:Static\s+)?(Sub\|Function\|Property\s+(?:Get\|Let\|Set))\s+([A-Za-z_]\w*)` | — (plain `function` node + `metadata.visibility`) |
| R2 | Methods in .cls | Same regex; emit `contains` edge class → function | — |
| R3 | Sub New marker | Match R1, then check group[3] === `'New'` and file is `.cls` | — (metadata only, no edge) |
| R4 | Call sites | `(?<!\w)([A-Za-z_]\w*)(?:\.([A-Za-z_]\w*))?\s*\(` inside procedure bodies (non-greedy receiver) | none (same-file) or `'vba-name-resolution'` (qualified receiver) |
| R5 | Implements | `^\s*Implements\s+([A-Za-z_]\w*)` | `'vba-name-resolution'` |
| R6 | Qualified Dim | `^\s*(?:Dim\|Private\|Public)\s+\w+\s+As\s+([A-Za-z_]\w*)\.[A-Za-z_]\w*` → outer type | `'vba-name-resolution'` |
| R7 | WithEvents | `^\s*(?:Dim\|Private\|Public)\s+(\w+)\s+As\s+(New\s+)?([A-Za-z_]\w*)` with `WithEvents` prefix | `'vba-withevents'` |
| R8 | SQL in strings | Iterate `extractStringLiterals`; per string: `\b(?:FROM\|INTO\|UPDATE)\s+(\[?\w+\]?)` | `'vba-sql-table'` |
| R9 | .form.txt rejection | VbaExtractor returns empty result if `filePath.endsWith('.form.txt')` or `.report.txt` | — |
| R10 | Option directives | Skip lines matching `^\s*Option\s+\w+` BEFORE R1 sweep | — |
| R11 | VB_Name | `^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"` (first non-empty line) | — |
| R12 | Form module node | Single `module` node from VB_Name or basename; emit `references` edge to sibling `.cls` | `'vba-form-binding'` |
| R13 | Control → property | Match `^\s*Begin\s+(\w+)\s*$` then read control type from next `^\s*(\w+)\s*$` block | — (`metadata.controlType`) |
| R14 | Report = form | Identical dispatch to form, filePath only check differs | same as R12/R13 |
| R15 | Zero code from form | VbaFormExtractor NEVER runs R1/R2/R4/R6/R7/R8 sweeps | — |

**Anti-catastrophic-backtracking**: every regex uses non-greedy quantifiers for body text, anchored at line start where possible, and runs AFTER comment-stripping (smaller input). The SQL regex is a tight `\b(FROM|INTO|UPDATE)\s+(\w+)` — no backtracking risk.

## Naming and naming collisions

- **Same-named Sub overloads**: `Sub Foo()` and `Sub Foo$()` get distinct IDs because `generateNodeId` hashes `(filePath, 'function', qualifiedName, line)`. The qualified name preserves the literal suffix (`Foo$`). Mirror upstream Swift overload handling (`buildFlowFromNamedSymbols` per CLAUDE.md).
- **Module vs filename**: `modHelpers.bas` with `Attribute VB_Name = "modHelpers"` → node name `"modHelpers"`. Without VB_Name → node name = file basename. Verified by spec R11 scenarios.
- **Class vs form**: `Form_Main.cls` emits a `class` node named `Form_Main`. `Form_Main.form.txt` emits a `module` node named `Form_Main`. The `references` edge (R12) connects them at graph time. The two nodes are distinct kinds and won't collide on `name` alone.

## Module-level Attribute parsing

`^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"` matches the **first non-empty line** (per spec R11). Only `VB_Name` is captured; other `Attribute VB_GlobalNameSpace = False`, `Attribute VB_Creatable = False`, `Attribute VB_PredeclaredId = False`, `Attribute VB_Exposed = False` are recognized but **emit nothing** (they are compiler hints, not symbols). The regex is anchored to line start to avoid matching inside string literals (post-comment-strip, this is safe).

## Edge cases

- **Multi-line bodies**: handled by `joinLineContinuations` BEFORE regex sweeps. Adversarial test: `Sub Foo()\n  DoCmd.RunSQL _\n    "SELECT * FROM tblX" _\n    & " WHERE Id = 1"` → must collapse to one logical line AND match R8.
- **`Rem` comments**: `Rem ` (with space) is legacy syntax for `'`. `stripVbaComments` recognizes both. Test: `Rem this is a comment\nSub X(): End Sub` → comment removed; Sub parsed.
- **Quoted strings containing `'`**: `Debug.Print "It's a test"` → the `'` is INSIDE `"..."` and is NOT a comment marker. `stripVbaComments` walks strings character-by-character (mirrors `RazorExtractor.matchBrace()` at `razor-extractor.ts:192`). Test: `Sub X(): Debug.Print "It's here": End Sub` → string preserved intact.
- **`#Const` / `#If` preprocessor**: emit nothing. Single-line regex `^\s*#\w+` skips them silently. No test required (out of the 28 spec scenarios; deferred).
- **WithEvents mid-class**: emit one `references` edge per `WithEvents` declaration, NOT per assignment. Test: a class with two `WithEvents` lines → two edges, distinct targets.

## Testing Strategy

| Layer | What to Test | Approach |
|---|---|---|
| Unit (preprocess) | `joinLineContinuations`, `stripVbaComments`, `extractStringLiterals` | `__tests__/extraction-vba-preprocess.test.ts` — 12 cases incl. adversarial quoted-`'` and `Rem` |
| Unit (extractor) | Each of 18 spec scenarios | `__tests__/extraction-vba.test.ts` — one `it()` per scenario, red → green per regex |
| Unit (form) | Each of 6 form spec scenarios | `__tests__/extraction-vba-form.test.ts` — including R15 ("zero code from form") |
| Integration | `extractFromSource('modFoo.bas', source)` returns expected node/edge shape | Inline `expect(...)` calls in the same files; no separate integration file |
| E2E regression | Codegraph main checkout non-VBA counts match baseline | Deferred to `sdd-verify` (obs #14702); uses npm-published codegraph as baseline |

TDD sequence: write the failing `it()` → run vitest (red) → write the regex → green → refactor. Never batch.

## Migration / Rollout

No migration. The change is additive: new `vba` language token, new `LANGUAGES` entry, new extractor files. Existing non-VBA projects hit the dispatch else-branch and behave identically. The `.codegraph/` index in VBA projects gets the new nodes/edges on first re-index after upgrade (obs #14702 strategy).

## Explicitly out of scope (deferred — DO NOT pull into sdd-tasks)

- Legacy `.frm` / `.dsr` binary extraction (per proposal "Out of Scope"; the new extractor's `filePath.endsWith('.frm')` arm emits a stub file node only).
- `dsCreateClassModule` Access runtime call (proposal "Decisions baked in" chose `Sub New` only).
- Preprocessor block parsing (`#Const`, `#If`) — emit nothing.
- Tree-sitter grammar for VBA (Path A — rejected at explore).
- `codegraph_explore` overload-aware upgrades for VBA specifically — the generic `buildFlowFromNamedSymbols` handles it.
- Adding `pascal-form`-style dedicated rendering in `src/mcp/tools.ts` for `synthesizedBy === 'vba-*'` — the generic fallback at `tools.ts:1820` already renders them as `"<kebab> (dynamic dispatch)"`; sdd-apply may add nicer labels if user requests.

## Open Questions

- [ ] None blocking. The 15 spec requirements are fully covered. If `sdd-tasks` finds a scenario ambiguous (e.g., `New Clase(...)` instantiation edge vs call edge), escalate to the user before splitting.