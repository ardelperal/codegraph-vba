# Tasks: VBA Handler Backtrace Custom Agent Skill

This task list breaks down the implementation of the `vba-handler-backtrace` capability following a strict test-driven development (TDD) approach in a single PR.

## Review Workload Forecast

```text
Decision needed before apply: No
Chained PRs recommended: No
Chain strategy: size-exception
400-line budget risk: Medium
```

## Task List

### 1. Test Harness Setup & RED Phase: Parsing and Concatenation
- [x] **Task 1 (RED):** Create `__tests__/vba-handler-backtrace.test.ts` with failing unit tests for basic signature and SQL query string parsing utilities.
  - Test UDT parameter parser with signature: `Public Sub ProcessOrder(ByRef ctx As OrderContext, ByVal flags As Long)` to extract `ctx: OrderContext`.
  - Test SQL concatenation logic with multi-line input ending in line continuation `_` and string concatenation `&`.
  - Verify tests run and fail using `npm test`.

### 2. GREEN Phase: Parsing and Concatenation Helper Implementation
- [x] **Task 2 (GREEN):** Create `src/utils/backtrace-helpers.ts` and implement helper functions for parameter signature parsing and SQL concatenation extraction.
  - Implement signature parsing regex matching `/(?:ByVal|ByRef)?\s*(\w+)\s+As\s+(\w+)/gi` and filter out primitive types.
  - Implement SQL reconstruction logic that accumulates multiline strings up to a limit of 200 characters.
  - Verify that the parsing tests now pass.

### 3. RED Phase: Traversal, Cycle Detection & Depth Capping
- [x] **Task 3 (RED):** Add tests in `__tests__/vba-handler-backtrace.test.ts` for dynamic graph traversal.
  - Mock or use test SQLite database with form control, call edges, and circular dependencies.
  - Verify traversal resolves complete call paths.
  - Test cycle detection (terminates traversal, flags `cycle_detected: true`).
  - Test depth capping (stops at max depth, adds `MAX_DEPTH_EXCEEDED` warning).
  - Verify tests fail.

### 4. GREEN Phase: Traversal Algorithm Implementation
- [x] **Task 4 (GREEN):** Implement traversal, cycle detection, and depth capping in `src/utils/backtrace-helpers.ts`.
  - Write graph traversal logic querying database paths from control to events to methods.
  - Track `visited` node set for cycle detection.
  - Implement depth tracking and append warnings if depth is exceeded.
  - Verify all unit tests pass with `npm test`.

### 5. REFACTOR Phase & Agent Custom Skill Definition
- [x] **Task 5 (REFACTOR):** Refactor the backtrace helper code for performance, readability, and error handling.
  - Simplify DB queries, ensure file read descriptors are closed, and ensure clean types in `src/utils/backtrace-helpers.ts`.
  - Run `npm test` to ensure no regressions.
- [x] **Task 6 (GREEN/REFACTOR):** Create the custom agent skill file at `.agents/skills/vba-handler-backtrace/SKILL.md`.
  - Document the instructions for the agent to call the traversal and parsing helpers, detailing the query schemas and extraction rules.
  - Provide instructions for formatting the JSON tree output.
