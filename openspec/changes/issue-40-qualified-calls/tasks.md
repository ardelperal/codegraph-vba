# Tasks: Issue 40 Qualified Calls

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | 140-220 |
| 400-line budget risk | Low |
| Chained PRs recommended | No |
| Suggested split | Single PR |
| Delivery strategy | exception-ok |
| Chain strategy | size-exception |

Decision needed before apply: No
Chained PRs recommended: No
Chain strategy: size-exception
400-line budget risk: Low

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | Add qualified-call gate and keep current behavior green | PR 1 | `src/extraction/vba-extractor.ts`; RED tests first |
| 2 | Cover paren/statement qualified-call scenarios and polish | PR 1 | `__tests__/extraction-vba.test.ts`; keep TDD order |

## Phase 1: RED — Failing Tests

- [x] 1.1 Add a failing test in `__tests__/extraction-vba.test.ts` for `nCount.ToString()` being silent when `nCount As Long` is declared.
- [x] 1.2 Add a failing test in `__tests__/extraction-vba.test.ts` for `modUtils.Foo arg` emitting a heuristic calls edge when `modUtils` is undeclared.

## Phase 2: GREEN — Extraction Logic

- [x] 2.1 Add `private shouldProcessQualifiedCall(receiverName: string): boolean` to `src/extraction/vba-extractor.ts` beside `isLocalProjectClassVar`.
- [x] 2.2 Use `shouldProcessQualifiedCall()` in `scanCallSites` before emitting qualified paren-form stubs for `Receiver.Member(...)`.
- [x] 2.3 Use `shouldProcessQualifiedCall()` in `sweepCallsAndSql` for qualified statement-form calls so module names are kept and primitive/external locals stay silent.

## Phase 3: REFACTOR — Contract Cleanup

- [x] 3.1 Update comments in `src/extraction/vba-extractor.ts` so the receiver eligibility rule is described once and matches both scanners.
- [x] 3.2 Re-run the affected `__tests__/extraction-vba.test.ts` cases and remove any duplicate setup created while making RED tests pass.

## Phase 4: Verification / Spec Alignment

- [x] 4.1 Verify the new scenarios in `openspec/changes/issue-40-qualified-calls/specs/vba-code-extraction/spec.md` stay aligned with the implementation behavior.
- [x] 4.2 Confirm the full `extraction-vba.test.ts` suite still passes for existing class-local qualified calls and same-file calls.
