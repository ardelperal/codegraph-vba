# Proposal: Bitwise-Precise Conditional Compilation Truthiness (-1/0 VBA Semantics) (Issue #84)

## Intent
Implement bitwise-precise evaluation for conditional compilation in the VBA preprocess pipeline, addressing the functional gaps of the current implementation which uses standard JavaScript boolean evaluation via `Function`. This ensures correct evaluation of VBA-specific logic, including bitwise operations, `-1`/`0` comparison outcomes, case-insensitive logic, lack of short-circuiting, and support for the `Xor` operator.

## Scope

### In Scope
- Replace the JS `Function`-based evaluation inside `evaluateConditionalExpression` and `evaluateConstRhs` in [vba-preprocess.ts](file:///C:/00repos/codigo/codegraph-vba/src/extraction/vba-preprocess.ts) with a secure, custom expression tokenizer, parser, and evaluator.
- Support the following expression components:
  - Logical/bitwise operators: `And`, `Or`, `Not`, `Xor` (case-insensitive).
  - Comparison operators: `=`, `<>`, `<`, `>`, `<=`, `>=`.
  - Constants: Hardcoded platform and environment flags (`VBA7`, `Win64`, `Win32` to `-1`; `Win16`, `Mac` to `0`), booleans (`True` to `-1`, `False` to `0`), and numeric literals (signed 32-bit integers).
  - Custom file-scoped constants declared with `#Const`.
  - Grouping parentheses.
- Ensure bitwise operations (`~` for `Not`, `&` for `And`, `|` for `Or`, `^` for `Xor`) are correctly performed on 32-bit signed integers without short-circuiting.
- Ensure comparison operations return `-1` for true and `0` for false.
- Ensure truthiness treats any non-zero value as true and `0` as false.
- Keep the existing line-count preservation invariant intact for all preprocess sweeps.

### Out of Scope
- Configurable platform constants via `codegraph.json` (to be addressed by issue 82).

## Capabilities
- **New Capabilities**: None.
- **Modified Capabilities**:
  - `vba-code-extraction` (modified to support bitwise-precise, secure pre-processing of conditional compilation branches).

## Technical Approach
Implement a pure TypeScript tokenizer and recursive descent (or Pratt) parser/evaluator.

### Tokenizer
Splits the conditional expression string into a stream of tokens:
- **NUMBER**: Integer literals (e.g. `123`, `0`).
- **IDENTIFIER**: Environment flags/variables (e.g. `Win64`, `MODO_DEBUG`, `True`, `False`, `And`, `Or`, `Not`, `Xor`).
- **OPERATOR**: `=`, `<>`, `<`, `>`, `<=`, `>=`, `(`, `)`.

### Parser & Precedence
The parser will evaluate expressions following standard VBA operator precedence:
1. **Primary**: Paren expressions `( ... )`, number literals, resolved identifiers.
2. **Unary Numeric**: `-` (negation), `+` (identity).
3. **Bitwise Not**: `Not` (operator mapping to 32-bit bitwise NOT: `~val`).
4. **Comparisons**: `=`, `<>`, `<`, `<=`, `>`, `>=`.
5. **Bitwise And**: `And` (bitwise AND: `a & b`).
6. **Bitwise Xor**: `Xor` (bitwise XOR: `a ^ b`).
7. **Bitwise Or**: `Or` (bitwise OR: `a | b`).

### Identifier Resolution
Lookup values (case-insensitively):
1. `#Const` table map.
2. Hardcoded environment constants:
   - `VBA7` -> `-1`
   - `Win64` -> `-1`
   - `Win32` -> `-1`
   - `Win16` -> `0`
   - `Mac` -> `0`
   - `True` -> `-1`
   - `False` -> `0`
3. Fallback for undefined/unknown identifiers -> `0` (equivalent to `False` in VBA).

### Truthiness & Safety
- **No `Function` or `eval` usage**: Eliminates code execution security risks.
- **No short-circuiting**: Since logical operators in VBA are bitwise, both sides are always evaluated.
- **Truthy conversion**: An expression `#If <expr>` will be evaluated to a number, and if the number is non-zero, the branch is active.

## Affected Areas
- **[src/extraction/vba-preprocess.ts](file:///C:/00repos/codigo/codegraph-vba/src/extraction/vba-preprocess.ts)**:
  - Replace `normalizeConditionalExpression`, `evaluateConditionalExpression`, and `evaluateConstRhs` with the custom tokenizer/parser/evaluator implementation.
- **[__tests__/extraction-vba-preprocess.test.ts](file:///C:/00repos/codigo/codegraph-vba/__tests__/extraction-vba-preprocess.test.ts)**:
  - Extend test suites to cover bitwise operations (`Not`, `And`, `Or`, `Xor`), complex parenthesized expressions, and comparisons returning `-1`/`0`.

## Risks & Mitigation
- **Syntax Errors**: If the preprocessor encounters syntax errors in `#If` expressions (e.g., mismatched parentheses or unrecognized operators/syntax), it will catch the error and fallback to treating the expression as false (blanking the branch). This matches the existing conservative behavior.
- **Whitespace / Normalization quirks**: The tokenizer handles whitespace and operator formatting naturally, avoiding regex-based substitution edge cases.

## Rollback Plan
Revert the edits to `src/extraction/vba-preprocess.ts` and `__tests__/extraction-vba-preprocess.test.ts` using git:
```bash
git checkout -- src/extraction/vba-preprocess.ts __tests__/extraction-vba-preprocess.test.ts
```

## Success Criteria
- Successful execution of all tests in `__tests__/extraction-vba-preprocess.test.ts` via Vitest.
- Addition of comprehensive tests demonstrating correct bitwise arithmetic, correct comparison logic (e.g., `#If (Win64 = -1) And (Mac = 0) Then` evaluates to true), and correct truthiness evaluation (non-zero is true).
