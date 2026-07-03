# Proposal: VBA API Declarations and Preprocessor Improvements

## Intent

Enhance the VBA extraction parser to support external DLL declarations, preprocessor conditional compilation, customized database SQL execution, and constant resolution in `DoCmd.OpenForm` calls.

## Scope

### In Scope
- Parse external DLL `Declare [PtrSafe]` statements as `'function'` nodes in `src/extraction/vba-extractor.ts`.
- Preprocess conditional compilation directives (`#If`, `#ElseIf`, `#Else`, `#End If`) in `src/extraction/vba-preprocess.ts` by blanking inactive lines.
- Extend SQL extraction to support custom database variable identifiers ending with `db` (case-insensitive) in `src/extraction/vba-extractor.ts`.
- Resolve constant identifiers to form names inside `DoCmd.OpenForm` calls using tracked constants.
- Update Vitest suites: `__tests__/extraction-vba.test.ts` and `__tests__/extraction-vba-preprocess.test.ts`.

### Out of Scope
- Supporting dynamic condition evaluation other than predefined flags (`VBA7 = true`, `Win64 = true`, `Mac = false`).
- Full constant expression evaluation (only single-literal/identifier values).

## Capabilities

### New Capabilities
- `vba-extraction-enhancements`: Emits `'function'` nodes for DLL API declarations, resolves constants in `DoCmd.OpenForm`, preprocesses conditional compilation, and maps SQL execution via arbitrary `*db` variables.

## Approach

Based on exploration findings:
1. **DLL Declarations**: Extract `Declare` statements as single-line procedures of kind `'sub'` or `'function'` with metadata `{ isDeclare: true }` in `sweepProcedures`.
2. **Conditional Compilation**: Add `preprocessConditionalCompilation` to `vba-preprocess.ts`. Evaluate conditions with `VBA7 = true`, `Win64 = true`, `Mac = false`, replacing inactive lines and directives with empty strings. Apply this before comments striping.
3. **Customized SQL DB**: Modify regex to match arbitrary variable names ending in `db` (e.g. `p_db`, `m_Db`, `g_db`, `db`) for `.OpenRecordset` and `.Execute`.
4. **OpenForm Constants**: Track constant definitions in `sweepEnumsAndConsts`. In `scanOpenFormCalls`, resolve variable arguments against `localConstants`.

## Affected Areas

| Area | Impact |
|------|--------|
| `src/extraction/vba-extractor.ts` | Modified |
| `src/extraction/vba-preprocess.ts` | Modified |
| `__tests__/extraction-vba.test.ts` | Modified |
| `__tests__/extraction-vba-preprocess.test.ts` | Modified |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Branch preprocessor line mismatch | Low | Replace ignored lines with empty string (`""`) to preserve line numbering |
| Custom `*db` regex false positives | Med | Restrict match to valid identifier characters preceding `db` |

## Rollback Plan

Revert the git commit. The changes are local to the VBA extractor modules and do not affect downstream parsers.

## Success Criteria

- [ ] DLL declarations extract as `'function'` procedures with correct visibility and metadata.
- [ ] Inactive preprocessor blocks are blanked out, preventing duplicate node declarations.
- [ ] SQL matches for variables like `p_db` or `m_Db` successfully emit `references` edges.
- [ ] `DoCmd.OpenForm CONST` resolves to form name or falls back to const name.
- [ ] Vitest suite passes without regression.
