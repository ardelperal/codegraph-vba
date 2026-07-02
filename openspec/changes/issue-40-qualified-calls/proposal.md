# Proposal: fix(vba): inconsistent gating of qualified calls (Issue #40)

## Intent

Resolve inconsistent behavior and parsing issues in VBA qualified calls (`Receiver.Member`). Paren form currently pollutes the database with dead-end stubs for external/primitive local variables. Statement form silently drops valid cross-module calls when the receiver is a module name.

## Scope

### In Scope
- Implement a unified gating function `shouldProcessQualifiedCall(receiver)` in `VbaExtractor`.
- Apply this gating rule to both qualified paren-form calls (`Receiver.Member(...)`) and qualified statement-form calls (`Receiver.Member args`).
- Gate qualified calls to only allow processing if:
  - The receiver is a declared local project class variable (`isLocalProjectClassVar` is true), OR
  - The receiver is NOT a declared local variable at all (meaning it is a candidate module name).
- Add new test coverage in `extraction-vba.test.ts` for:
  - Suppression of external/primitive local variables in paren form.
  - Parsing of cross-module calls in statement form when the receiver is a module name.

### Out of Scope
- Global module registry lookups at extraction time.
- Resolving external types or libraries dynamically.

## Capabilities

### New Capabilities
None

### Modified Capabilities
- `vba-code-extraction`: Update qualified call requirements to enforce consistent gating across paren and statement call forms, preventing external local variable stubs while preserving candidate module calls.

## Approach

1. Create a private method `shouldProcessQualifiedCall(receiver: string): boolean` in [vba-extractor.ts](file:///C:/00repos/codigo/codegraph-vba/src/extraction/vba-extractor.ts):
   - Returns `true` if `isLocalProjectClassVar(receiver)` is true.
   - Returns `false` if `localVarTypeMap.has(receiver.toLowerCase())` is true (but not a local project class).
   - Returns `true` otherwise.
2. In `scanCallSites`, for qualified paren-form calls, check `shouldProcessQualifiedCall(receiver)` before emitting the synthetic node and edge.
3. In `sweepCallsAndSql`, for qualified statement-form calls, replace `this.isLocalProjectClassVar(qualStmt.receiver)` with `this.shouldProcessQualifiedCall(qualStmt.receiver)`.
4. Update `openspec/specs/vba-code-extraction/spec.md` with the new unified rules.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `src/extraction/vba-extractor.ts` | Modified | Add helper and apply gating to paren/statement scanner. |
| `openspec/specs/vba-code-extraction/spec.md` | Modified | Update requirements to specify unified qualified call gating. |
| `__tests__/extraction-vba.test.ts` | Modified | Add tests for both forms. |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Receiver shadowing (local variable name collides with module name) | Low | Standard compiler shadowing behavior; shadowed receiver calls are skipped locally. |

## Rollback Plan

Revert git changes to the modified files.

## Dependencies

None

## Success Criteria

- [ ] External/primitive types (e.g., `DAO.Recordset`, `Long`) do not generate synthetic nodes or call edges in qualified paren form.
- [ ] Statement calls referencing module names not declared as local variables (e.g., `modUtils.Foo`) successfully generate heuristic call edges.
- [ ] All existing and new tests pass successfully under strict TDD rules.

## Proposal Question Round

### Proposed Product/Business Questions
1. Shadowing behavior: In VBA, a local variable could shadow a global module name. Should we support any diagnostics or warnings when a module name is shadowed, or is silent suppression the desired outcome?
2. Undeclared globals: The current approach treats any undeclared receiver (e.g. `ExcelApp`) as a candidate module name, which is left as a stub in the database. Are there specific global objects we should blacklist similar to `DoCmd`/`Me`?

### Key Assumptions
- All local variables of project class types will be correctly mapped in `localVarTypeMap`.
- If a receiver is not in `localVarTypeMap`, it is assumed to be a module name or global object.
