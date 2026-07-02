# Technical Design: VBA API Declarations and Preprocessor Improvements

## Technical Approach

We will enhance the VBA extraction pipeline by implementing:
1. **Conditional Compilation Preprocessing**: Evaluate `#If ... #End If` directives in `vba-preprocess.ts` before stripping comments, replacing inactive branches and directives with empty strings `""` to preserve line indices.
2. **DLL API Declarations**: Parse `Declare [PtrSafe]` statements during procedure sweeps, generating `'function'` nodes with `{ isDeclare: true }` metadata.
3. **Custom DB Variable SQL Extraction**: Generalize SQL execution regexes and database wrappers to support arbitrary variables ending in `db` (case-insensitive).
4. **OpenForm Constant Resolution**: Cache local constants during enum/const sweeps, and resolve `DoCmd.OpenForm` arguments (both string literals and bare identifiers) to their values.

## Architectural Decisions

| Area | Option | Tradeoff | Decision |
| :--- | :--- | :--- | :--- |
| **DLL API Declarations** | A: Map to standard `'function'` nodes.<br>B: Introduce new `'dll_declare'` node kind. | A: Simplifies queries and reuses existing call resolution.<br>B: Requires client-side changes and updates to node schemas. | **Option A**: Emits `'function'` nodes with `startLine === endLine` and metadata `{ isDeclare: true }`. |
| **Conditional Compilation** | A: Parse conditions during AST sweeps.<br>B: Blank inactive branches before AST sweeps. | A: High complexity handling duplicates in sweeps.<br>B: Simplifies sweeps but requires a line-preserving preprocessor. | **Option B**: Strip inactive branches and directives, replacing them with empty lines to preserve line count parity. |
| **Custom SQL DB Variables** | A: Hardcode recognized variable names.<br>B: Match identifiers ending in `db` (case-insensitive). | A: Low risk of false positives, high maintenance.<br>B: Catches all custom DB variables but requires careful regex. | **Option B**: Use suffix-bounded regex to match any valid identifier suffix `db` (e.g. `p_db.Execute`). |
| **OpenForm Constants** | A: Resolve local constants only.<br>B: Perform multi-file constant indexing. | A: misses cross-module constants but simple and fast.<br>B: High complexity and indexing overhead. | **Option A**: Build a local constants map during extraction sweeps; fall back to the constant name when unresolved. |

## Data Flow Diagram

```
Source File
   │
   ▼
[joinLineContinuations]
   │
   ▼
[preprocessConditionalCompilation] ──(Inactive branches & directives replaced by "")
   │
   ▼
[stripVbaComments] ──────────────────(Comments & options replaced by "")
   │
   ▼
[AST Extraction Sweeps] ─────────────(Procedures, DLLs, Dims, Enums, Constants)
   │
   ▼
Graph Nodes & Edges Output
```

## Detailed Component Designs

### 1. Conditional Compilation Nesting & Evaluation
- **Nesting Stack**: Track nested `#If` groups using a stack of:
  `{ isCurrentlyActive: boolean, hasAnyBranchBeenActive: boolean }`.
  - Inside a block, lines are active only if every stack level has `isCurrentlyActive === true`.
  - Replace directives and inactive lines with empty strings `""` to preserve line indices.
- **Operator Mapping & Cleanup**:
  - Strip trailing `Then` keyword (case-insensitive).
  - Map case-insensitive VBA operators: `And`/`AND` $\rightarrow$ `&&`, `Or`/`OR` $\rightarrow$ `||`, `Not`/`NOT` $\rightarrow$ `!`, `=` $\rightarrow$ `===`, `<>` $\rightarrow$ `!==`.
  - Replace compilation constants (e.g., `VBA7` $\rightarrow$ `true`, `Win64` $\rightarrow$ `true`, `Mac` $\rightarrow$ `false`).
- **Sanitization Guard**: Before running `eval` / `Function`, check translated expression against:
  `/^(?:true|false|[0-9]+|&&|\|\||!|===|!==|\(|\)|\s)+$/`
  Reject any expression containing `.`, `[`, `]`, quotes, or backticks to prevent JS injection.

### 2. Constant Resolution & OpenForm Parsing
- **Constant Value Extraction**: In `sweepEnumsAndConsts`, parse constant definitions by scanning the body left-to-right (respecting string literals/escapes and not splitting on commas inside quotes) to support multi-constant lines like `Const A = "x", B = 1`. Cache name-value pairs in a `localConstants` map:
  - Unwrap outer double-quotes.
  - Collapse double-quote escapes (`""` $\rightarrow$ `"`).
- **OpenForm Match**: Match form name using:
  `/\bDoCmd\.OpenForm\s+("[^"]+"|\p{L}[\p{L}\p{N}_]*)/gu`
  - If argument starts with `"`, unwrap quotes to get form name.
  - If argument is an identifier, look up in `localConstants`, falling back to the identifier name.

### 3. Custom DB Variable Regex
- **Suffix Pattern**: Match variables with case-insensitive `*db` suffix using:
  `\b(?:\p{L}[\p{L}\p{N}_]*)?db\b` (prevents partial-word matching on `dbtable` or `db_test`).
- **SQL Exec / Wrapper Patterns**:
  - Variable Exec: `/\b(?:\p{L}[\p{L}\p{N}_]*)?db\b(?:\(\))?\.(?:OpenRecordset|Execute)\s*\(?\s*(\p{L}[\p{L}\p{N}_]*)\s*\)?/giu`
  - Inline Literals:
    - `/\b(?:\p{L}[\p{L}\p{N}_]*)?db\b(?:\(\))?\.OpenRecordset\s+"((?:[^"]|"")*)"/giu`
    - `/\b(?:\p{L}[\p{L}\p{N}_]*)?db\b(?:\(\))?\.Execute\s+"((?:[^"]|"")*)"/giu`

## Specific File Changes

| File | Changes |
| :--- | :--- |
| [vba-preprocess.ts](file:///C:/00repos/codigo/codegraph-vba/src/extraction/vba-preprocess.ts) | • Implement `preprocessConditionalCompilation(src: string): string` with the nesting stack, operator translator, and regex sanitization guard. |
| [vba-extractor.ts](file:///C:/00repos/codigo/codegraph-vba/src/extraction/vba-extractor.ts) | • Integrate `preprocessConditionalCompilation` in `extract()`.<br>• Match `Declare [PtrSafe]` in `sweepProcedures` via `DLL_DECLARE_RE`.<br>• Parse and cache constant values (supporting multi-constant lines and unwrapping quotes) in `localConstants` in `sweepEnumsAndConsts`.<br>• Update `scanOpenFormCalls` to use new `OPEN_FORM_RE` and resolve form names against `localConstants`.<br>• Update `SQL_VAR_EXEC_RE` and `SQL_WRAPPERS` to use the `*db` suffix regex. |
| [extraction-vba-preprocess.test.ts](file:///C:/00repos/codigo/codegraph-vba/__tests__/extraction-vba-preprocess.test.ts) | • Add unit tests for conditional compilation (line count parity, `#If Not Mac And Win64 Then` evaluation, nesting, sanitization rejection). |
| [extraction-vba.test.ts](file:///C:/00repos/codigo/codegraph-vba/__tests__/extraction-vba.test.ts) | • Add integration tests for DLL declarations extraction, `*db` SQL queries, multi-constant parsing, and OpenForm constant resolution. |

## Code Snippet Interface Contracts

```typescript
// In src/extraction/vba-preprocess.ts
export function preprocessConditionalCompilation(src: string): string;

// DLL Node Metadata Schema
export interface FunctionNodeMetadata {
  isDeclare?: boolean;
}
```

## Testing Strategy

1. **Preprocessor Unit Tests**:
   - Verify line count parity between input and output.
   - Test nested conditional directives and complex conditions like `#If Not Mac And Win64 Then`.
   - Verify sanitization rejects expressions containing unsafe characters.
 2. **Extractor Integration Tests**:
   - Extract DLL declarations as `'function'` nodes with `{ isDeclare: true }` metadata.
   - Extract SQL queries executed on variables ending in `db` (e.g., `p_db`, `m_Db`).
   - Extract multi-constant declarations and verify quote unwrapping.
   - Resolve `DoCmd.OpenForm` calls using defined constants and fallback cases.
3. **Vitest Verification**: Run `npx vitest run` to ensure all tests pass.

## Migration & Open Questions

- **Migration**: Metadata is backward-compatible. No database schema migration required.
- **Open Questions**: None.
