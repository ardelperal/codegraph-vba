# Tasks: Issue 84 Bitwise Conditional-Compilation

## Review Workload Forecast

Decision needed before apply: No
Chained PRs recommended: No
Chain strategy: size-exception
400-line budget risk: Low

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1    | Implement tokenizer, parser, and evaluator in `vba-preprocess.ts` | PR 1 | Phase 1 & 2 changes |
| 2    | Add comprehensive unit tests and run tests | PR 1 | Phase 3 changes |

## Phase 1: Core parser/evaluator

- [x] 1.1 Define the `Token` interface and token types in [vba-preprocess.ts](file:///C:/00repos/codigo/codegraph-vba/src/extraction/vba-preprocess.ts).
- [x] 1.2 Implement the tokenizer function `tokenize` in `src/extraction/vba-preprocess.ts` to scan:
  - Whitespace (skipped)
  - Decimal integers (`\d+`)
  - Parentheses (`(`, `)`)
  - Comparison operators (`<=`, `>=`, `<>`, `<`, `>`, `=`)
  - Unary operators (`-`, `+`)
  - Keywords/Identifiers case-insensitively (`And`, `Or`, `Not`, `Xor`, constants)
- [x] 1.3 Implement identifier resolution within the tokenizer:
  - Check the scoped `#Const` table (`constTable`).
  - Check hardcoded environment constants: `VBA7` -> `-1`, `Win64` -> `-1`, `Win32` -> `-1`, `Win16` -> `0`, `Mac` -> `0`, `True` -> `-1`, `False` -> `0`.
  - Fall back undefined identifiers to `0`.
  - Emit resolved identifiers directly as `NUMBER` tokens containing the resolved integer value.
- [x] 1.4 Implement the recursive descent `Parser` class in `src/extraction/vba-preprocess.ts` enforcing VBA's operator precedence:
  - Level 1: `( ... )`, number literals (Primary expressions)
  - Level 2: Unary `-`, `+`
  - Level 3: `Not`
  - Level 4: `=`, `<>`, `<`, `<=`, `>`, `>=` (Comparisons, returning `-1` for True, `0` for False)
  - Level 5: `And` (Bitwise AND)
  - Level 6: `Xor` (Bitwise XOR)
  - Level 7: `Or` (Bitwise OR)
- [x] 1.5 Enforce explicit 32-bit signed integer coercion (via `| 0`) at each evaluation step (unary negation, logical bitwise operations, comparisons).
- [x] 1.6 Ensure the parser throws an error on syntax errors, mismatched parentheses, or unexpected trailing tokens.

## Phase 2: Wiring

- [x] 2.1 Remove the regex-based `normalizeConditionalExpression` function from [vba-preprocess.ts](file:///C:/00repos/codigo/codegraph-vba/src/extraction/vba-preprocess.ts).
- [x] 2.2 Rewrite `evaluateConditionalExpression` to invoke `tokenize`, parse the expression using the new `Parser`, convert the resulting integer to boolean (non-zero is `true`, zero is `false`), and catch any errors to safely return `false`.
- [x] 2.3 Rewrite `evaluateConstRhs` to invoke `tokenize`, parse the expression using the new `Parser`, convert the resulting integer to string representation, and catch any errors to safely return `null`.

## Phase 3: Testing

- [x] 3.1 Add new unit tests to [extraction-vba-preprocess.test.ts](file:///C:/00repos/codigo/codegraph-vba/__tests__/extraction-vba-preprocess.test.ts) covering bitwise logical operations (e.g. `True And False`, `True Or False`, `Not True`, `1 And 2`, `Not 0`).
- [x] 3.2 Add tests for VBA non-zero truthiness (e.g. `#If 2 Then` keeps the branch, `#If 0 Then` blanks the branch).
- [x] 3.3 Add tests for correct precedence levels (e.g. `3 And 4 Or 1`, comparison operator precedence over bitwise operators).
- [x] 3.4 Add tests for syntax fallbacks/errors (e.g. invalid syntax such as mismatched parentheses or unsupported symbols default to `false`/`null` instead of crashing).
- [x] 3.5 Run the full unit test suite `npm test` to verify everything is green.
