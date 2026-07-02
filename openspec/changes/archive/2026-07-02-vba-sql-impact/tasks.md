# Tasks: VBA SQL Impact Analysis (vba-sql-impact)

## Review Workload Forecast
```text
Decision needed before apply: No
Chained PRs recommended: No
Chain strategy: size-exception
400-line budget risk: Medium
```

## Implementation Tasks

### Phase 1: Test & Utility Helpers (TDD)
- [x] **Task 1 [RED]:** Create helper tests file `__tests__/vba-sql-impact.test.ts` with failing tests.
  - Write test cases for VBA caller tracing regex matching (`OpenRecordset` and `QueryDefs`).
  - Write test cases for Form/Report layout parsing (`RecordSource` and `RowSource` properties, including ComboBox/ListBox parent control contexts).
  - Write test cases for SQL table/column alias resolution (implicit/explicit FROM/JOIN aliases).
- [x] **Task 2 [GREEN]:** Implement utility helpers in `src/utils/sql-impact-helpers.ts`.
  - Implement regex patterns and parsing methods to make all tests in `__tests__/vba-sql-impact.test.ts` pass.
  - Run `npm run test` or `npx vitest` to verify green status.
- [x] **Task 3 [REFACTOR]:** Refactor helper parsing logic and type exports.
  - Clean up regex definitions, simplify pattern-matching constructs, and structure return types cleanly.
  - Verify zero regressions by running tests.

### Phase 2: Skill Definition & Integration
- [x] **Task 4 [GREEN]:** Create the agent-side custom skill at `.agents/skills/vba-sql-impact/SKILL.md`.
  - Author the custom skill markdown explaining step-by-step how the agent executes dynamic SQL impact analysis.
  - Detail indexer queries, workspace file searches, helper invocations, and output payload formatting.
- [x] **Task 5 [RED]:** Write integration tests for full-flow lineage extraction.
  - In `__tests__/vba-sql-impact.test.ts`, add a test case that creates a mock database connection, registers control nodes and edges, writes mock `.form.txt` and `.sql` query files to a temporary folder, and runs the combined backtrace/impact logic.
- [x] **Task 6 [GREEN]:** Make integration tests pass and finalize verification.
  - Fix any integration/resolution path discrepancies.
  - Run full test suite (`npm test`) and check lint rules to ensure zero regressions across the codebase.
