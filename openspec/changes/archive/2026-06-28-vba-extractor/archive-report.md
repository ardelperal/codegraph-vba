# Archive Report — `vba-extractor`

## Status

**`partial`** — mirrors the verify verdict `PASS-WITH-WARNINGS`. No critical issues. The partial status reflects the pre-existing Windows platform test debt (23 failures in non-VBA files, all documented in `CLAUDE.md`) and two SUGGESTION items from verify.

## Summary

The `vba-extractor` change adds VBA / Access language support to codegraph via two regex extractors: `VbaExtractor` for `.bas`/`.cls` files and `VbaFormExtractor` for Dysflow-canonical `.form.txt`/`.report.txt` files. Both are wired into `extractFromSource()` alongside the existing extractor dispatch chain. All 28 spec scenarios pass (75/75 VBA-specific vitest atoms green); both non-negotiable invariants hold (`.form.txt` emits zero `function`/`class`/`module` nodes; literal `Sub` in form UI produces no function nodes). The six implementation commits are properly tagged with `SDD: vba-extractor`, reachable from `feature/vba-extractor`, and no later commits have reverted or overwritten any of the change's files. The change is ready for the user to open a PR against `ardelperal/codegraph:main`.

## Verify Verdict

**PASS-WITH-WARNINGS** (from `openspec/changes/archive/2026-06-28-vba-extractor/verify-report.md`)

- All 28 spec scenarios covered by passing tests
- VBA-specific test counts: 75/75 green
- Pre-existing Windows failures (23 total): git-hooks, MCP daemon/initialize/roots, worktree detection, Drupal/GoFrame routing, prepare-release, exclude-config, upgrade-index-stamp — none in VBA files
- Source-vs-binary: N/A (TypeScript fork, no Access `.accdb`)
- Schema additions: additive, non-breaking

## Delta Specs Archived

Both specs were full specs (no main spec existed), copied to canonical location:

| Canonical path | Action |
|---|---|
| `openspec/specs/vba-code-extraction/spec.md` | **Created** — 11 requirements, 22 scenarios, 186 lines |
| `openspec/specs/vba-form-ui-extraction/spec.md` | **Created** — 4 requirements, 6 scenarios, 68 lines |

### Internal Cross-References

- `vba-form-ui-extraction/spec.md` references `vba-code-extraction` as the sibling spec covering the canonical `.cls` (REQ-FORM-1: *"the sibling `.cls` of the same basename"*).
- The proposal references the two spec domains by name.
- The design references `vba-code-extraction` and `vba-form-ui-extraction` by name.
- No broken links or stale references after the move.

## Implementation Commits

| Commit | Work unit | SDD tasks | Verification | Access sync |
|---|---|---|---|---|
| `76e7454` | T1 wiring + detectVbaFormFile + types/grammars/dispatch | T1.1–T1.10 | `npx vitest run __tests__/detect-vba-form-file.test.ts` 8/8 green | N/A |
| `60146e9` | T2 pre-processing helpers | T2.1–T2.7 | `npx vitest run __tests__/extraction-vba-preprocess.test.ts` 36/36 green | N/A |
| `176a667` | T3 VbaExtractor + 22 spec scenarios | T3.1–T3.19 | `npx vitest run __tests__/extraction-vba.test.ts` 22/22 green | N/A |
| `1ba73f9` | T4 VbaFormExtractor + 7 spec scenarios (6 spec + 1 bonus) | T4.1–T4.7 | `npx vitest run __tests__/extraction-vba-form.test.ts` 7/7 green | N/A |
| `d189b50` | T5 server-instructions + E2E regression | T5.1–T5.3 | `npx vitest run __tests__/extraction-vba-e2e.test.ts` 2/2 green | N/A |
| `e538532` | Implementation-commits table fill-in | T5.4 | `git log --oneline` shows table content matches | N/A |

## Target Branch

`feature/vba-extractor` (current branch, based on `main` at `a4e1bda`).

Per the session preflight (`single-pr-default`), the merge to `ardelperal/codegraph:main` is the user's gate — this archive does not push or open a PR.

## Reachability Verification

All 6 commits confirmed reachable from `feature/vba-extractor`:

| SHA | Subject | Reachability |
|---|---|---|
| `76e7454` | `feat(vba): wire VBA language into types, grammars, and extractFromSource dispatch` | ✅ `git merge-base --is-ancestor 76e7454 feature/vba-extractor` |
| `60146e9` | `feat(vba): add pre-processing pipeline (joinLineContinuations, stripVbaComments, extractStringLiterals)` | ✅ |
| `176a667` | `feat(vba): implement VbaExtractor for .bas/.cls/.frm/.dsr with 22 spec scenarios` | ✅ |
| `1ba73f9` | `feat(vba): implement VbaFormExtractor for .form.txt/.report.txt with 7 spec scenarios` | ✅ |
| `d189b50` | `docs(vba): add server-instructions paragraph for VBA/Access; add E2E regression test` | ✅ |
| `e538532` | `docs(vba-extractor): record implementation commit table` | ✅ |

No later commits on `feature/vba-extractor` (confirmed: only the 6 implementation commits + 1 pre-fork commit from `main`).

## Later-Commit Overwrite Check

`git diff a4e1bda..feature/vba-extractor --stat` shows only the expected VBA files changed:
- `src/extraction/vba-extractor.ts` (+698 lines)
- `src/extraction/vba-form-extractor.ts` (+230 lines)
- `src/extraction/vba-preprocess.ts` (+230 lines)
- `src/extraction/grammars.ts` (+46 lines / -N)
- `src/extraction/tree-sitter.ts` (+13 lines / -N)
- `src/types.ts` (+20 lines / -N)
- `src/mcp/server-instructions.ts` (+20 lines / -N)
- 5 test files (+866 total lines)

No overwrite, no revert. All files are additive or mechanical (grammar wiring, type additions).

## Source-vs-Binary Sync

N/A — TypeScript fork, no Access `.accdb` involved.

## Open Follow-Ups

### From verify-report SUGGESTION items:

1. **E2E test cleanup hardening** (`__tests__/extraction-vba-e2e.test.ts`): The `afterAll` cleanup runs but if the test process is hard-killed between `beforeAll` and `afterAll`, `.codegraph/` is left behind in the main checkout. Recommend wrapping the `indexAll()` call in a `try/finally` that triggers cleanup on hard failure. Low-priority: `beforeAll` self-heals on the next run.

2. **Spec scenario count documentation**: `tasks.md` Phase 3 listed 18 scenarios; `extraction-vba.test.ts` ships 22 tests. The spec grew during sdd-spec and apply kept tests in sync. No action required; informational only.

3. **Per-language node/edge count regression vs upstream v1.1.2**: Not run in this verify slice. A follow-up `ab-new-vs-baseline.sh` run per `scripts/agent-eval/` could quantify the delta. Out of scope for this archive.

### User-owned gates:

- **Merge to `main`**: Per session preflight (`single-pr-default`), the user owns the PR gate. Run `branch-pr` skill or open a PR from `feature/vba-extractor` → `ardelperal/codegraph:main` when ready.
- **Follow-up change: VBA fixtures** (per Engram obs #14702): Add real `.bas`/`.cls`/`.form.txt` fixtures to `__tests__/fixtures/vba/` for future regression coverage. The extractor is functional but lives off synthetic test cases until real-world fixtures are provided.

## Archive Contents

| Artifact | Status |
|---|---|
| `proposal.md` | ✅ In archive |
| `specs/vba-code-extraction/spec.md` | ✅ In archive |
| `specs/vba-form-ui-extraction/spec.md` | ✅ In archive |
| `design.md` | ✅ In archive |
| `tasks.md` | ✅ In archive (all checkboxes reconciled — see below) |
| `verify-report.md` | ✅ In archive |

### Task Completion Reconciliation

`tasks.md` uses `- [ ]` checkboxes for all implementation tasks. All were unchecked at archive time despite the implementation being complete (verified by `apply-progress` observation #14714 and `verify-report` observation #14717). Per the SDD archive policy exception for stale checkboxes with proof from `apply-progress` + `verify-report`:

- `apply-progress` (#14714): 6 commits, 75/75 tests green, 5 phases complete
- `verify-report` (#14717): 28/28 spec scenarios green, both invariants hold, all commits reachable and tagged

The unchecked boxes are stale — the work is done. Reconciliation reason recorded in this archive report.

## Engram Observation IDs (Traceability)

| Artifact | Engram ID | topic_key |
|---|---|---|
| Exploration | #14703 | `sdd/vba-extractor/explore` |
| Proposal | #14706 | `sdd/vba-extractor/proposal` |
| Spec | #14707 | `sdd/vba-extractor/spec` |
| Tasks | #14710 | `sdd/vba-extractor/tasks` |
| Design | #14709 | `sdd/vba-extractor/design` |
| Apply-progress | #14714 | `sdd/vba-extractor/apply-progress` |
| Verify-report | #14717 | `sdd/vba-extractor/verify-report` |
| Archive-report | — (this doc) | `sdd/vba-extractor/archive-report` |

## Archive Metadata

- **Archive date**: 2026-06-28
- **Archived by**: `sdd-archive` (executor)
- **Artifact store**: `hybrid` (openspec filesystem + Engram)
- **Verify verdict**: `PASS-WITH-WARNINGS`
- **Archive status**: `partial` (non-critical: pre-existing Windows test debt + 2 SUGGESTION items)
- **CHANGELOG entry**: added under `## [Unreleased]` — `docs(changelog): note VBA / Access extractor under [Unreleased]`

## SDD Cycle Complete

The `vba-extractor` change has been fully planned, implemented, verified, and archived. The canonical specs now live at `openspec/specs/vba-code-extraction/spec.md` and `openspec/specs/vba-form-ui-extraction/spec.md`. The archived change lives at `openspec/changes/archive/2026-06-28-vba-extractor/`. All 6 implementation commits are reachable and tagged. Ready for the user to open a PR against `ardelperal/codegraph:main`.
