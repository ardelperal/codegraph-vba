# Design: VBA Graph Connectivity Fixes (issues #12, #13)

## Technical Approach

Approach 3 (targeted post-extraction reconciliation, B2). Keep eager stub emission in `vba-extractor.ts` (preserves the synchronous-`extract()` contract ~all unit tests rely on), tag qualified-call stubs with `metadata.stub`, and add one VBA-scoped resolver pass that repoints stub `calls` edges to real cross-file nodes via DB-level edge UPDATE + duplicate collapse. #13 is a localized append/reset in `trackSqlVariableAssignment`. Zero touch to the shared multi-language dispatch (`matchReference`/`matchMethodCall`).

## Architecture Decisions

### Decision: Pass location — method on `ReferenceResolver`
**Choice**: Add `resolveVbaCallStubs(): number` to `src/resolution/index.ts`, invoked at the `resolveChainedCallsViaConformance()` slot (`src/index.ts` ~428/431 `indexAll`, ~552/555 `sync`).
**Alternatives**: new `vba-call-stub-resolver.ts`.
**Rationale**: Needs `this.queries`, mirrors two sibling second-pass methods, ~50 lines. Cohesion beats a standalone file.

### Decision: Repoint via UPDATE-in-place + explicit duplicate collapse + stub-flag clear (F1, F5)
**Choice**: `repointEdgeTarget(edgeId, newTargetId, newMetadataJson)` → `UPDATE edges SET target=?, metadata=? WHERE id=?`. The resolver drives it per incoming stub edge (`getIncomingEdges(stubId, ['calls'])`) and enforces uniqueness in JS:
1. Compute new metadata in JS: `{...parsed, stub: false}` (strip/flip the stub flag — F5) while keeping `synthesizedBy`/`receiverType`/`member`.
2. Maintain a pass-level `Set<"source\0newTarget\0kind">`. If the tuple was already produced this pass OR a DB row `(source,newTarget,'calls')` already exists → `deleteEdgeById(edge.id)` (drop the loser) instead of repointing.
3. Otherwise repoint (UPDATE target + cleared metadata) and record the tuple.
After all incoming edges of a stub are handled, `deleteNode(stubId)` (stubs are never edge sources).
**Alternatives**: delete+reinsert; blind bulk `UPDATE ... WHERE target=@old` (prior design — produces duplicates).
**Rationale**: `edges` has an AUTOINCREMENT PK and NO unique constraint (schema.sql:45-56), so two stubs at different call-site lines (`synthId` keys on `lineNum`, vba-extractor.ts:1338-1343) targeting one real node would yield two identical `(source,target,'calls')` rows. `traversal.ts` dedups via a `visited` set added AFTER a depth guard (`traversal.ts:246-249`); at the `maxDepth=1` used across `src/mcp/tools.ts` the child is never marked visited before the guard returns, so duplicate same-source edges surface TWICE in `codegraph_node`/explore trails — an agent-visible regression. Collapsing at repoint keeps exactly one row. **Accepted info loss**: the surviving row keeps the first repoint's `line`/`col`; a second identical call at another line loses only its line number. Justified: the graph models "A calls B" (callers/callees dedupe by node), not per-line call sites.

### Decision: Stub lookup follows the JSON-in-JS convention, NOT `json_extract` (F2)
**Choice**: `getVbaCallStubs()` = `SELECT * FROM nodes WHERE kind='function' AND language='vba' AND metadata LIKE '%"stub":true%'`, then in JS after `rowToNode`/`safeJsonParse` keep only rows where `metadata?.stub === true`. The `LIKE` is a cheap pre-filter only; correctness is the JS check.
**Alternatives**: `json_extract(metadata,'$.stub')=1` predicate (prior design).
**Rationale**: `json_extract` has ZERO precedent in `src/`; the established convention is `JSON.stringify` TEXT parsed via `safeJsonParse` after fetch (`queries.ts:1304`, 135/150/169). The supported runtime is Node `>=22.5 <25` (hard exit on 25.x); JSON1 availability was only smoke-tested on Node 25.2.1 (out of range) — unverified for the shipped range, so we do not depend on it. Not a perf concern: stub count is bounded per project (one narrow scan per full index).

### Decision: `.bas` fallback narrows by module identity BEFORE deciding (F3)
**Choice**: For a `.bas`-qualified stub `receiver.member`:
1. `getNodesByName(member)` filtered `kind==='function' && language==='vba' && filePath.endsWith('.bas') && !metadata.stub`.
2. Narrow to candidates whose containing `.bas` file's module identity equals `receiver` (case-insensitive). Module identity = the file's `module` node name (VB_Name attribute, else file basename without ext — vba-extractor.ts:149/292). Resolve `receiver`'s file(s) via `getNodesByName(receiver)` filtered `kind==='module' && language==='vba'`, intersect on `filePath`.
3. Exactly 1 candidate after narrowing → repoint. 0 or 2+ AFTER narrowing → decline (keep stub).
**Alternatives**: match on bare `member` across ALL `.bas` files, decline on 2+ project-wide (prior design).
**Rationale**: The `receiver` text IS the target module's name (`.bas` calls are always `Module.Member`, never an indirect variable). Deciding on the unscoped project-wide collision count made common helper names (`Init`, `Log`, `GetValue`) always decline, defeating the fix. The correct precedent is the narrow-then-broaden pattern at `resolution/index.ts:1290-1307` (prefer the ref's own file, broaden only if empty) — NOT `matchByQualifiedName`'s bare fallback, which takes the first match and never declines (`name-matcher.ts:428-437`). Note `.bas` function `qualifiedName` is bare (no class prefix), so the module name lives on the sibling `module` node, not the function's `qualifiedName`.

### Decision: Class-typed matching
`getNodesByQualifiedNameExact(stub.qualifiedName)` — `.cls` methods carry `${className}.${proc}`, so `NCOperaciones.Registrar` matches exactly. Resolve only when exactly 1 candidate (mirrors `matchByQualifiedName`); 2+ declines.

### Decision: Repointed edge self-heals on target-only incremental resync (F4 — RESOLVED, not deferred)
**Conclusion (traced, definitive)**: A repointed `calls` edge SURVIVES a resync of ONLY the target `.cls`/`.bas` file (caller untouched). No silent data loss. The prior "open question / documented limitation" framing was WRONG.
**Trace**: `storeExtractionResult` (`extraction/index.ts:1788`) snapshots the repointed edge via `getCrossFileIncomingEdgesWithTarget(targetFile)` (`queries.ts:1466` — `tgt.file_path=target ∧ kind!='contains' ∧ src.file_path!=target`) BEFORE `deleteFile` cascades it. It captures the target's `(kind,name)`. After re-inserting the target file's fresh nodes, the reinsert loop (`index.ts:1825-1839`) re-resolves by `${targetKind}\0${targetName}` against the fresh nodes and reinserts the edge (with its snapshot metadata, so the F5-cleared `stub:false` is preserved). The real VBA proc node's `name` is the bare proc name, stable across same-content re-extraction → the key matches → edge repoints to the fresh real node id. The resolver pass then finds no stub (the caller wasn't re-extracted, so no new stub emitted) → no-op, no duplicate. If the caller IS resynced, the extractor re-emits the stub and the pass re-repoints. Both paths converge.

## Data Flow

    extract() ─▶ stub function node {metadata.stub}
                 + calls edge (target=stubId, synthesizedBy, stub:true)
                         │  (persisted per-file)
                         ▼
    indexAll/sync ─▶ resolveVbaCallStubs()
        getVbaCallStubs() (LIKE prefilter → JS stub===true)
          exact QN match | .bas module-scoped fallback
            per incoming edge:
              dup(source,newTarget,'calls')? ─▶ deleteEdgeById  (collapse)
              else ─▶ repointEdgeTarget(edgeId, realId, meta{stub:false})
            all done ─▶ deleteNode(stubId)
          no/ambiguous match ─▶ leave stub (dead-end preserved)

## File Changes

| File | Action | Description |
|------|--------|-------------|
| `src/extraction/vba-extractor.ts` | Modify | Class-typed receiver resolved via `localVarTypeMap`; stub node `metadata:{stub:true}`, edge metadata `{stub:true, synthesizedBy, receiverType, member}`. `synthId` UNCHANGED. `trackSqlVariableAssignment`: accumulate on self-ref concat (#13). Blacklists + unqualified branch untouched. |
| `src/resolution/index.ts` | Modify | New `resolveVbaCallStubs()` with per-edge repoint + duplicate collapse + `.bas` module-scoped narrowing; call at both lifecycle slots. |
| `src/index.ts` | Modify | Invoke pass after `resolveDeferredThisMemberRefs()` in `indexAll` + `sync`. |
| `src/db/queries.ts` | Modify | New `getVbaCallStubs()` (LIKE prefilter + JS `stub===true`), `repointEdgeTarget(edgeId, newTargetId, metadataJson)`, `deleteEdgeById(id)`, `edgeExists(source,target,kind)`. Reuse `getNodesByName`, `getNodesByQualifiedNameExact`, `getIncomingEdges`, `deleteNode`. |
| `__tests__/extraction-vba.test.ts` | Modify | Update ~3-4 class-typed `.name` expectations; #13 SQL accumulation units. |
| `__tests__/extraction-vba-realfixtures.test.ts` | Modify | New e2e: class-typed + `.bas` repoint; duplicate-collapse; idempotency; target-only resync self-heal. |
| `__tests__/fixtures/vba/src/...` | Create | Minimal caller (var typed `ACAuditoriaOperaciones` calling `.Registrar`; two same-target call sites; `.bas`-module paren call) — current fixtures lack a resolving cross-class call. |

## #13 fix shape

In `trackSqlVariableAssignment`, after `newFragment = collectStringLiteralText(...)`: if RHS (`m[2]`.trim) matches `/^<escapedVarName>\s*&/i` AND `sqlVariables.has(varName)` → `set(varName, existing + ' ' + newFragment)`; else `set(varName, newFragment)`. Case-insensitive (map keyed lowercase). Independent of `collectStringLiteralText`'s forward `&`-continuation walk.

## Testing Strategy

| Layer | What | Approach |
|-------|------|----------|
| Unit | resolved-type stub name; SQL accumulate/reset/case-mismatch/multi-hop-silent | `extraction-vba.test.ts` (Windows-CI subset) |
| E2E | class-typed + `.bas` repoint; orphan stub deleted; unresolved/ambiguous stub kept | `extraction-vba-realfixtures.test.ts` via `indexAll()` |
| E2E — duplicate collapse (F1) | two call sites in one Sub → same qualified target, both resolve to one real node | assert EXACTLY one `(source,target,'calls')` edge row survives; assert `getCallers` on the target lists the caller ONCE |
| E2E — idempotency (F6) | run the resolver pass twice with no intervening change | assert node count AND edge count are stable (no duplicates, no re-created stubs) |
| E2E — target-only resync self-heal (F4) | after a successful repoint, resync ONLY the target `.cls`/`.bas` (caller file untouched) | assert the edge still points at a valid non-stub real node afterward |

Strict TDD: failing tests first; load-bearing assertions only in the two CI-subset files.

## Migration / Rollout

No schema/data migration. Re-index restores prior state. Sequence #13 before #12 (shared file region).

## Open Questions

None — all prior open questions resolved (F4 self-heal traced and confirmed above).
