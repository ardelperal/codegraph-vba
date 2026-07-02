# Proposal: VBA Graph Connectivity Fixes (issues #12, #13)

## Intent

Two confirmed VBA-extractor bugs break the call/SQL graph, so `codegraph_explore` returns dead ends and loses table references. Bundled into one change (user decision): same files, same theme — call/SQL connectivity.

- **#12** (https://github.com/ardelperal/codegraph-vba/issues/12): qualified calls (`m_NCOp.Registrar x`, `modUtils.Foo(...)`) synthesize permanent dead-end stub nodes instead of connecting to the real target. Node ids hash on the caller's file+line, so a stub can NEVER id-collide with a cross-file real node — extractor-only renaming provably cannot fix connectivity.
- **#13** (https://github.com/ardelperal/codegraph-vba/issues/13): `trackSqlVariableAssignment` overwrites instead of accumulating across `sql = sql & "..."`, losing earlier fragments' tables (typically the initial `FROM <table>`).

## Scope

### In Scope
- #12: tag stubs `metadata.stub=true` (mirrors `DoCmd.OpenForm` precedent); use resolved type name for class-typed receivers via `localVarTypeMap`; new post-extraction resolver pass (Approach 3 / B2) at the `resolveChainedCallsViaConformance` lifecycle slot that repoints stub edges to real nodes via exact `matchByQualifiedName` lookup + bare-member `.bas` fallback.
- #13: detect self-referential concat (RHS starts with `<varName> &`, case-insensitive) → append; genuine fresh assignment still resets.
- Failing tests first (strict TDD), placed in the Windows-CI regression subset (see Dependencies).

### Out of Scope
- Unifying Dim/WithEvents/SQL-table stub families onto one pipeline (full B1 `UnresolvedReference` migration) — larger separate investment; strategic fork left for later.
- Multi-hop cross-variable SQL (`sql2 = sql1 & ...`) — stays a documented limitation.

## Capabilities

### New Capabilities
- None

### Modified Capabilities
- `vba-code-extraction`: "Call Sites Emit Edges" — qualified stub edges now repoint to real cross-file targets; class-typed stub name uses resolved type. "SQL String Table References" — accumulate concatenated fragments instead of overwrite.

## Approach

Approach 3 (targeted post-extraction reconciliation, B2): keep eager stub emission (preserves test shape), add a VBA-scoped resolver pass that repoints edges — never touches shared multi-language dispatch. #13 is a localized append/reset in `trackSqlVariableAssignment`.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `src/extraction/vba-extractor.ts` | Modified | Stub tagging + resolved-type name (#12); accumulate SQL (#13) |
| `src/resolution/index.ts` | Modified | New `resolveVbaCallStubs`-style pass |
| `src/db/queries.ts` | Modified | Metadata-filtered stub lookup + edge-target repoint |
| `__tests__/extraction-vba.test.ts` | Modified | ~3-4 class-typed `.name` assertions updated (intentional, not regressions); new unit tests |
| `__tests__/extraction-vba-realfixtures.test.ts` | Modified | New e2e repoint + SQL-accumulation coverage |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Edge repoint orphans edges (no update-in-place primitive) | Med | delete+reinsert with focused query tests |
| Touching heavily-tested blacklist/same-file branch | Low | Leave `scanCallSites` unqualified branch + blacklists untouched |
| Same file region diffed by both fixes | Low | Sequence #13 before #12; follow file's audit-comment convention |

## Rollback Plan

Revert the PR; both fixes are additive/isolated to the named files. No schema or data migration — a re-index restores prior graph state.

## Dependencies

- **CI hard requirement**: `.github/workflows/ci.yml` "Run Windows VBA regression subset" step MUST pass. New/changed load-bearing tests MUST live in that subset: `extraction-vba.test.ts` (unit), `extraction-vba-realfixtures.test.ts` (e2e). Do NOT put load-bearing assertions only in `extraction-vba-preprocess.test.ts` / `extraction-vba-form.test.ts` (excluded).
- Strict TDD: every behavior change needs a failing test first.
- Design/apply must verify existing Dysflow fixtures cover a class-to-fixture-class call, else add a minimal fixture pair.

## Success Criteria

- [ ] #12: qualified call edges resolve to real cross-file nodes (class-typed + `.bas`-module) proven by e2e `indexAll()` test
- [ ] #13: multi-fragment SQL retains all table refs; fresh reassignment resets; case-mismatched self-ref treated as self-referential
- [ ] Windows VBA regression subset green; full vitest suite green
- [ ] Delivery: single PR (target); revisit only if tasks-phase forecast flags 400-line risk
