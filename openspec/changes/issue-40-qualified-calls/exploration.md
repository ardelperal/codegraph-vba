## Exploration: fix(vba): inconsistent gating of qualified calls (Issue #40)

### Current State
Today, qualified VBA calls (`Receiver.Member`) are parsed differently depending on whether they use parentheses (paren form) or not (statement form):
1. **Paren Form (`Receiver.Member(...)`)**: Matches via `CALL_RE` and unconditionally creates a synthetic stub node and `calls` edge. If the receiver is a local variable representing an external library or primitive (e.g. `rcdDatos` declared as `DAO.Recordset`), it still pollutes the database with dead-end stubs that cannot be resolved.
2. **Statement Form (`Receiver.Member args`)**: Matches via `detectQualifiedStatementCall` but is gated strictly by `this.isLocalProjectClassVar(receiver)`. If the receiver is a module name (e.g. `modUtils`), it is not in the local variable type map, causing the cross-module call to be silently dropped.

### Affected Areas
- [vba-extractor.ts](file:///C:/00repos/codigo/codegraph-vba/src/extraction/vba-extractor.ts) — Gating logic in `scanCallSites` (paren form) needs to skip stubs for declared local variables of non-project class types, and gating in `sweepCallsAndSql` (statement form) needs to allow calls where the receiver is a candidate module name (not present in `localVarTypeMap`).
- [extraction-vba.test.ts](file:///C:/00repos/codigo/codegraph-vba/__tests__/extraction-vba.test.ts) — Needs test coverage for paren form non-gating/pollution and statement form cross-module calls.

### Approaches
1. **Consistent Gating Check for Both Forms (Recommended)**
   Implement a consistent gating rule for qualified calls in both forms. A call `Receiver.Member` is parsed only if:
   - `this.isLocalProjectClassVar(receiver)` is `true` (it is a declared local project class variable).
   - OR `!this.localVarTypeMap.has(receiver.toLowerCase())` is `true` (it is not a declared local variable, thus a candidate module name).
   
   If `receiver` is in `localVarTypeMap` but `isLocalProjectClassVar(receiver)` is `false` (e.g. primitive/external type), the call is ignored.
   - **Pros**:
     - Prevents paren-form stub pollution from external variables (`DAO.Recordset`, etc.) and primitives.
     - Preserves cross-module statement calls (e.g. `modUtils.Foo args`).
     - Retains local, high-performance regex extraction without needing cross-file context during extraction.
   - **Cons**:
     - Cannot distinguish undeclared external globals (e.g. `ExcelApp.Workbooks`) from project modules at extraction time (these are safely left as stubs and remain unresolved in the DB, which is standard).
   - **Effort**: Low.

2. **Global Module Registry lookup at Extraction Time**
   Collect a global set of all module/class names in the codebase before extraction, and pass it to each `VbaExtractor` instance to explicitly white-list matching receivers.
   - **Pros**:
     - 100% precise; suppresses all external/unknown receiver stubs.
   - **Cons**:
     - High complexity; breaks the parallel, single-pass design of `parse-pool.ts` and complicates incremental syncs.
   - **Effort**: High.

### Recommendation
Proceed with **Approach 1**. It is extremely simple, fits the existing single-file regex architecture, and fully solves both paren pollution and statement drops.

### Risks
- Minor risk of receiver shadowing (if a module name collides with a declared local non-project variable, calls on the module in that scope will be skipped, which mirrors standard compiler shadowing anyway).

### Ready for Proposal
Yes. The orchestrator should proceed to the proposal phase.
