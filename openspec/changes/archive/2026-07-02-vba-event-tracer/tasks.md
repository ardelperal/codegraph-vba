# Tasks: VBA Event Tracer

## Review Workload Forecast

Decision needed before apply: No
Chained PRs recommended: No
Chain strategy: size-exception
400-line budget risk: Medium

## Phase 1: Event Signature Extraction

- [x] 1.1 **RED**: Add a unit test to [__tests__/extraction-vba-roadmap-25-26.test.ts](file:///C:/00repos/codigo/00_codegraph_main/__tests__/extraction-vba-roadmap-25-26.test.ts) asserting that the extracted `event` nodes populate the `signature` column with the trimmed source declaration line.
- [x] 1.2 **GREEN**: Update `sweepEventsTypesAndDeclares` in [src/extraction/vba-extractor.ts](file:///C:/00repos/codigo/00_codegraph_main/src/extraction/vba-extractor.ts) to populate the `signature` field of `event` nodes with the trimmed source line.
- [x] 1.3 **REFACTOR**: Run `npm test` to verify extraction tests pass cleanly.

## Phase 2: Custom Skill Creation

- [x] 2.1 **RED**: Define E2E skill validation cases detailing the exact expected output for normal trace requests, circular reference modules, and unqualified ambiguous queries before implementing the skill.
- [x] 2.2 **GREEN**: Create the custom agent skill at `.agents/skills/vba-event-tracer/SKILL.md` containing dynamic resolution logic, ambiguity check (`EVENT_AMBIGUOUS`), circular reference handling, and warnings formatting.
- [x] 2.3 **REFACTOR**: Verify the skill instructions resolve variable names cleanly and run under 100ms when simulated on sample indexes.

## Phase 3: Integration & Final Verification

- [x] 3.1 **REFACTOR**: Execute full test suite `npm run build && npm test` to confirm zero regressions across all VBA parsing features.
