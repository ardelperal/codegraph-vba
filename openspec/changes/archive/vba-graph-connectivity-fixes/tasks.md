# Tasks: VBA Graph Connectivity Fixes (issues #12, #13)

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | ~600–700 (prod ~150–180: extractor+resolver+queries+wiring; tests ~380–450: 4 DB-primitive units, 4 SQL units, 6 e2e scenarios, 3–4 updated assertions; fixtures ~30–50) |
| 400-line budget risk | High |
| Chained PRs recommended | Yes |
| Suggested split | 3 work units — see below |
| Delivery strategy | single-pr-default (project `openspec/config.yaml`) |
| Chain strategy | pending — maintainer decides chain vs. `size:exception` |

Decision needed before apply: Yes
Chained PRs recommended: Yes
Chain strategy: pending
400-line budget risk: High

Note: `openspec/config.yaml` sets `review_budget_lines: not_enforced` and `pr_strategy: single-pr-default` for this fork, but the forecast is run per explicit request. If the maintainer accepts `size:exception` (consistent with the archived `vba-extractor` change), set `Chained PRs recommended: No` and proceed single-PR.

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | Phases 1–2: fixtures + DB primitives | PR 1 | Foundation, no behavior change yet |
| 2 | Phase 3: #13 SQL accumulation | PR 2 | Localized `trackSqlVariableAssignment`; must land before Unit 3 (shared file region) |
| 3 | Phases 4–7: stub tagging, resolver pass, wiring, e2e, cleanup | PR 3 | Depends on Units 1–2 |

## Phase 1: Fixture Creation

- [x] 1.1 Create `__tests__/fixtures/vba/src/modules/modCallerDemo.bas`: `Dim x As ACAuditoriaOperaciones` with **two** call sites `x.Registrar p_Error` (same target, for F1 duplicate-collapse) + one `.bas`-qualified paren call `mdlCursor.MouseCursor()` (module-scoped fallback target, real `Public Function` already in `mdlCursor.bas`).
- [x] 1.2 Update the "indexes 6 VBA files" assertion in `__tests__/extraction-vba-realfixtures.test.ts` (~line 55–60) to 7 (new `.bas` file added).

## Phase 2: DB Primitives (`src/db/queries.ts`)

New file `__tests__/db-vba-call-stub-queries.test.ts` (direct `DatabaseConnection`+`QueryBuilder` unit tests, pattern from `db-perf.test.ts`; supplementary — NOT the load-bearing coverage, which comes from Phase 6 e2e).

- [x] 2.1 RED: test `getVbaCallStubs()` returns only `kind='function' && language='vba' && metadata.stub===true` nodes (LIKE prefilter + JS check). **Design deviation**: `nodes` has no `metadata` column (only `edges` does — schema.sql has no such column, confirmed; `Node.metadata` was never persisted anywhere pre-existing in this codebase, including the `DoCmd.OpenForm` stub precedent). Implemented as a JOIN against `edges.metadata` (which DOES persist) instead — same LIKE-prefilter-then-JS-check shape, functionally equivalent, documented inline in `src/db/queries.ts::getVbaCallStubs`.
- [x] 2.2 GREEN: implement `getVbaCallStubs()`.
- [x] 2.3 RED: test `repointEdgeTarget(edgeId, newTargetId, metadataJson)` updates `target`+`metadata` in place, leaves other columns untouched.
- [x] 2.4 GREEN: implement `repointEdgeTarget()`. Also added `Edge.id?: number` (populated in `rowToEdge`) since edges have no natural unique key and the design's per-edge repoint needs the row id.
- [x] 2.5 RED: test `deleteEdgeById(id)` removes exactly that row.
- [x] 2.6 GREEN: implement `deleteEdgeById()`.
- [x] 2.7 RED: test `edgeExists(source, target, kind)` true/false cases.
- [x] 2.8 GREEN: implement `edgeExists()`.

**Commit**: `feat(vba): add DB primitives for call-stub resolution` — SDD: vba-graph-connectivity-fixes
**Verification**: `npx vitest run __tests__/db-vba-call-stub-queries.test.ts`

## Phase 3: #13 — SQL Accumulation (`vba-extractor.ts::trackSqlVariableAssignment`)

Sequenced BEFORE Phase 4/5 (shared file region). Tests in `__tests__/extraction-vba.test.ts` (CI-covered).

- [x] 3.1 RED: two-fragment self-referential concat (`sql = "...FROM tblA"`; `sql = sql & "..."`) → edge to `tblA` retained.
- [x] 3.2 RED: three-plus fragment accumulation → edges to `tblA` AND `tblB`.
- [x] 3.3 RED: fresh (non-self-ref) reassignment after use resets tracking → only new table's edge. (Already passed pre-fix — reset is the existing overwrite behavior, unchanged; kept as a regression guard.)
- [x] 3.4 RED: case-insensitive self-reference (`Sql = sql & "..."`) → accumulation still applies.
- [x] 3.5 GREEN: implement accumulate/reset in `trackSqlVariableAssignment` per design's `#13 fix shape` (case-insensitive `<varName> &` RHS match; else replace).

**Commit**: `fix(vba): accumulate SQL fragments across self-referential concatenation (#13)` — SDD: vba-graph-connectivity-fixes
**Verification**: `npx vitest run __tests__/extraction-vba.test.ts`

## Phase 4: #12a — Stub Tagging (`vba-extractor.ts`)

- [x] 4.1 RED: qualified paren-form call (`scanCallSites`) → stub node has `metadata.stub===true`; stub edge metadata has `{stub:true, synthesizedBy:'vba-name-resolution', receiverType, member}`.
- [x] 4.2 RED: class-typed qualified statement-form call → same stub-tagging contract (node + edge metadata).
- [x] 4.3 GREEN: add `metadata.stub:true` to stub node creation and `stub:true`/`receiverType`/`member` to edge metadata at both emission sites; `synthId` generation UNCHANGED (added a `resolveReceiverType()` helper — for class-typed local vars it substitutes the RESOLVED class name in place of the raw receiver text before building `qualified`; for everything else, unchanged).
- [x] 4.4 Update existing OLD-behavior `.name` assertions in `__tests__/extraction-vba.test.ts` (~lines 1304, 1321, 1396) to assert the new interim-stub contract (`metadata.stub===true` alongside the raw-name check) — separate commit-worthy step, not conflated with 4.1/4.2. Confirmed: exactly 3 assertions changed (`m_NCOp.Registrar`→`NCOperaciones.Registrar`, `m_Obj.Init`→`SomeClass.Init`, `m_AROp.Eliminar`→`ARAuditoriaOperaciones.Eliminar`), matching the proposal's "~3-4 class-typed `.name` assertions updated (intentional, not regressions)" forecast exactly.

**Commit**: `feat(vba): tag call-stub nodes and edges with metadata.stub (#12a)` — SDD: vba-graph-connectivity-fixes
**Verification**: `npx vitest run __tests__/extraction-vba.test.ts`

## Phase 5: #12b — Resolver Pass + Wiring

- [x] 5.1 Implement `resolveVbaCallStubs(): number` in `src/resolution/index.ts`: `getVbaCallStubs()` → per stub, class-typed exact-QN match (`getNodesByQualifiedNameExact`, decline on 2+) OR `.bas` module-scoped fallback (narrow candidates by containing module identity, decline on 0/2+) → per incoming edge (`getIncomingEdges(stubId,['calls'])`): dedupe via pass-level `Set` + `edgeExists` check → `deleteEdgeById` (collapse) or `repointEdgeTarget` (clear `stub:false`, keep `synthesizedBy`/`receiverType`/`member`) → `deleteNode(stubId)` once all incoming edges handled. Added a `stubIds` exclusion set (all live stub node ids for the pass) so an exact-qualifiedName lookup never mistakes a SIBLING stub (same qualifiedName, different id/line — the F1 duplicate-call-site case) for a second real candidate.
- [x] 5.2 Wire `resolveVbaCallStubs()` at the `resolveChainedCallsViaConformance()` slot in `src/index.ts`'s `indexAll()` (~428) and `sync()` (~552).

**Commit**: `feat(vba): resolve call stubs to real cross-file nodes (#12b)` — SDD: vba-graph-connectivity-fixes

## Phase 6: E2E Scenarios (`extraction-vba-realfixtures.test.ts`, CI-covered)

- [x] 6.1 RED: class-typed stub resolves to real method node (`ACAuditoriaOperaciones.Registrar`) after `indexAll()`; `metadata.stub` absent/false.
- [x] 6.2 RED: `.bas`-qualified stub resolves to real bare-name node (`mdlCursor.MouseCursor`), module-scoped narrowing.
- [x] 6.3 RED: duplicate-collapse (F1) — two call sites → same target → EXACTLY one `(source,target,'calls')` edge row; `getCallers` lists the caller ONCE.
- [x] 6.4 RED: ambiguous/unmatched stub keeps stub metadata (`metadata.stub===true`), no throw. Implemented in an isolated purpose-built temp project (own `CodeGraph` instance) alongside 6.5/6.6 rather than the shared fixtures instance, to avoid a second `indexAll()`/`sync()` corrupting state for 6.1–6.3's shared `cg`. Per the nodes.metadata gap (2.1's deviation note), "keeps stub metadata" is verified on the EDGE's `metadata.stub` (which persists), not the retained stub node's own metadata (which doesn't).
- [x] 6.5 RED: idempotency (F6) — run `indexAll()` twice unchanged → node+edge counts stable.
- [x] 6.6 RED: target-only resync self-heal (F4) — resync only the target `.cls`/`.bas` (caller untouched) → edge still points at valid non-stub real node.
- [x] 6.7 GREEN: confirm all 6.1–6.6 pass against Phase 5 implementation. All 6 passed on the first run — no narrowing/collapse edge cases needed fixing.

**Commit**: `test(vba): e2e coverage for call-stub repoint, collapse, idempotency, self-heal` — SDD: vba-graph-connectivity-fixes
**Verification**: `npx vitest run __tests__/extraction-vba-realfixtures.test.ts __tests__/extraction-vba.test.ts` (Windows CI VBA regression subset also runs `extraction-vba-events.test.ts`, `extraction-vba-enums-consts.test.ts`, `extraction-vba-control-modeling.test.ts` — unaffected by this change, run for full-suite confidence)

## Phase 7: Cleanup

- [x] 7.1 Add `### Fixes` entry to `CHANGELOG.md` under `[Unreleased]` (user-facing: cross-file VBA calls now resolve to their real target instead of dead-ending; multi-fragment SQL concatenation table references are no longer dropped).
- [x] 7.2 Full verification: `npm run build && npm test`. Build green. Full suite: all VBA-related tests green (142/142 in the Windows CI VBA regression subset run standalone); 3 unrelated tests (`mcp-roots.test.ts`) intermittently failed only under full-suite parallel worker load (`EPERM`/"Worker exited unexpectedly") — confirmed pre-existing flakiness, not a regression: re-running `npm test` produced a DIFFERENT failing set each time, and all failing tests passed 100% in isolation. None touch VBA extraction/resolution/DB code.

**Verification**: all green; no non-VBA regression (per project's cross-platform validation convention).
