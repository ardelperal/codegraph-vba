## Exploration: feat(vba-preprocess): bitwise-precise conditional-compilation truthiness (-1/0 VBA semantics) (Issue #84)

### Current State
Today, `preprocessConditionalCompilation` evaluates VBA conditional compilation (`#If`, `#ElseIf`) by performing basic regex substitutions before passing the expression to a JavaScript `Function` constructor:
- Replacing `=` with `===` and `<>` with `!==`
- Substituting `And`/`Or`/`Not` with `&&`/`||`/`!`
- Substituting `true`/`false` with `-1`/`0`

This simplified approach fails to accurately represent VBA conditional compilation semantics:
1. **Bitwise Operations**: In VBA, logical operators (`And`, `Or`, `Not`, `Xor`) are bitwise operations on signed integers.
2. **Boolean Outputs**: VBA comparison operators (`=`, `<>`, `<`, `<=`, `>`, `>=`) return `-1` for True and `0` for False.
3. **No Short-Circuiting**: VBA evaluates all operands of bitwise logical operations, unlike JavaScript's short-circuiting `&&` and `||`.
4. **Lack of `Xor`**: `Xor` is not supported in the current implementation.
5. **Comparison Constraints**: The current whitelist only permits `===` and `!==`. It lacks support for other comparison operators like `<`, `>`, `<=`, `>=`.
6. **Security / Safety**: The use of the JS `Function` constructor is an unnecessary security risk and is flagged by static analysis scanners.

### Affected Areas
- [vba-preprocess.ts](file:///C:/00repos/codigo/codegraph-vba/src/extraction/vba-preprocess.ts) — Replace the regex substitutions and `Function`-based evaluation inside `evaluateConditionalExpression` and `evaluateConstRhs` with a custom expression evaluator.
- [extraction-vba-preprocess.test.ts](file:///C:/00repos/codigo/codegraph-vba/__tests__/extraction-vba-preprocess.test.ts) — Add comprehensive test suites for VBA-precise bitwise operations, comparisons returning `-1`/`0`, case-insensitive keyword parsing, and `#Const` table integration.

### Approaches

#### 1. Pure TS Pratt/Recursive Descent Parser & Evaluator (Recommended)
Replace the regex normalization and `Function` constructor execution with a custom tokenizer and recursive descent parser/evaluator in TypeScript.
- **Tokenizer**: Splitting expressions into `NUMBER`, `IDENTIFIER`, and `OPERATOR` tokens.
- **Precedence & Grammar**:
  1. Primary: `NUMBER`, `IDENTIFIER`, parenthesized expressions.
  2. Unary: `-` (negation), `+` (identity).
  3. Comparisons: `=`, `<>`, `<`, `<=`, `>`, `>=` (evaluate to `-1` or `0`).
  4. Bitwise Unary: `Not` (mapped to JS `~`).
  5. Bitwise Binary: `And` (JS `&`), `Xor` (JS `^`), `Or` (JS `|`).
- **Identifier Resolution**: Check against case-insensitive `constTable` or hardcoded VBA constants (`VBA7`, `Win64`, `Win32` -> `-1`; `Win16`, `Mac` -> `0`; `True` -> `-1`; `False` -> `0`).
- **Truthiness**: Evaluates the expression to a single number and treats any non-zero value as `true` (VBA truthiness).
- **Pros**:
  - 100% accurate VBA bitwise and comparison semantics.
  - Zero safety risks (no `Function` or `eval`).
  - Strict validation of expressions built directly into the parser.
  - Simple, robust, and easy to maintain (~120 lines of TypeScript).
- **Cons**:
  - Slightly larger code size in `vba-preprocess.ts` compared to a simple regex.
- **Effort**: Low-Medium.

#### 2. Enhanced Regex + JS Function Helper Context
Augment the current implementation by mapping operators and comparisons using complex regular expressions to JavaScript helper functions (e.g. `vbaAnd`, `vbaOr`, `vbaNot`, `vbaXor`, `vbaEq`, etc.), and run the final string inside the `Function` context where these helpers are declared.
- **Pros**:
  - Keeps the existing `Function` evaluator.
- **Cons**:
  - Extremely complex and fragile regexes are required to handle precedence, unary negation, parentheses, and ternary operators.
  - Retains the unsafe `Function` evaluation.
- **Effort**: High.

### Recommendation
Proceed with **Approach 1** (Pratt/Recursive Descent Parser & Evaluator). It is secure, correct, and far easier to verify than trying to map precedence and types with complex regex patterns.

### Risks
- **Syntax Errors**: If the preprocessor encounters syntax errors in `#If` expressions (e.g., mismatched parentheses), it will throw an error and blank the branch. This is the correct conservative behavior matching the current codebase's fallback.
- **Integer Size**: JavaScript performs bitwise operations on 32-bit signed integers, which matches the behavior of modern VBA7/x64 environments.

### Ready for Proposal
Yes.
