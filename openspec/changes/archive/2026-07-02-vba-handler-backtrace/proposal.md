# Proposal: VBA Handler Backtrace Custom Agent Skill

## Intent
Implement a custom agent skill `vba-handler-backtrace` that dynamically resolves event-handler backtraces, including called procedures, DAO SQL strings, and UDT parameter types.

## Scope

### In Scope
- Create the custom agent skill `vba-handler-backtrace` at `.agents/skills/vba-handler-backtrace/SKILL.md`.
- Dynamic query-time traversal of VBA call graphs from form control event-handler procedures to a configurable depth (default 5).
- On-the-fly parsing of source files along the trace path to extract:
  - SQL strings from DAO database calls (e.g. `.OpenRecordset`, `.Execute`) up to 200 characters.
  - User-Defined Type (UDT) parameter type names from procedure declarations.
- Robust cycle detection and truncation warning reporting.
- Vitest test suite for verification of code-under-test.

### Out of Scope
- Modifying the core indexer codebase or database schema.
- Comprehensive dataflow analysis or constant propagation outside local strings.

## Capabilities

### New Capabilities
- `vba-handler-backtrace`: Traces form control events back through subroutine calls to dependencies, extracting SQL query strings and UDT parameters.

## Approach
Implement Option A (Dynamic Agent Skill):
1. **Dynamic Traversal**: Starting from a Form control, find `event-handler` edges in the database, then recursively trace downstream `calls` and `references` using SQLite queries.
2. **SQL String Extraction**: For DAO `references` edges, read the corresponding file lines dynamically to extract full SQL strings, handling multiline string accumulation.
3. **UDT Parameters**: Parse the procedure signature directly from the source code during traversal to identify UDT parameter types, matching them against known types.
4. **Cycle/Capping**: Maintain a visited set to avoid infinite loops and enforce the maximum search depth.

## Affected Areas

| Area | Impact |
|------|--------|
| `.agents/skills/vba-handler-backtrace/SKILL.md` | Created |
| `__tests__/vba-handler-backtrace.test.ts` | Created |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| File out of sync with DB | Low | Validate file timestamps against the index snapshot |
| Heavy token usage on deep traces | Med | Cap max-depth, trace only relevant path files |

## Rollback Plan
Revert git changes. The capability is self-contained as a dynamic skill and has no database migration or core code impact.

## Success Criteria
- [ ] Trace resolves form control to its handler procedure and called helpers.
- [ ] DAO query strings are extracted and formatted as hints.
- [ ] Parameter UDT types are parsed and reported.
- [ ] Cycle detection identifies loops without infinite recursion.
- [ ] Vitest test suite passes.
