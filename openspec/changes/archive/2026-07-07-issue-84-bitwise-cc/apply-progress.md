# Apply Progress: Bitwise Conditional-Compilation Truthiness (-1/0 VBA Semantics)

## TDD Cycle Evidence

| Phase/Cycle | Target File | Test Command & Result | Status | Code Change Summary |
|---|---|---|---|---|
| Phase 1: RED | `__tests__/extraction-vba-preprocess.test.ts` | `npx vitest run __tests__/extraction-vba-preprocess.test.ts` -> Failed | **RED** | Added unit tests for bitwise operators (And, Or, Not, Xor), operator precedence, comparisons, non-zero truthiness, and syntax recovery. |
| Phase 1: GREEN | `src/extraction/vba-preprocess.ts` | `npx vitest run __tests__/extraction-vba-preprocess.test.ts` -> Passed | **GREEN** | Implemented custom tokenization and a recursive descent expression parser enforcing VBA precedence and 32-bit signed integer coercion. |
| Phase 2: GREEN | `src/extraction/vba-preprocess.ts` | `npx vitest run __tests__/extraction-vba-preprocess.test.ts` -> Passed | **GREEN** | Removed regex-based `normalizeConditionalExpression` and wired `evaluateConditionalExpression`/`evaluateConstRhs` directly to tokenization and parser. |
| Phase 3: GREEN | `__tests__/extraction-vba-preprocess.test.ts` | `npx vitest run __tests__/extraction-vba-preprocess.test.ts` -> Passed (66 tests)<br>`npx vitest run __tests__/extraction-vba` -> Passed (332 tests) | **GREEN** | Added triangulation/edge case tests for 32-bit signed overflow, case insensitivity of `#Const` references, and nested negation. Verified zero regressions across the whole VBA extractor suite. |
