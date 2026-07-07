# Design: Issue 84 Bitwise Conditional-Compilation

## Technical Approach

We will replace the regex-based substitutions and JavaScript `Function`-based evaluation inside `evaluateConditionalExpression` and `evaluateConstRhs` in [vba-preprocess.ts](file:///C:/00repos/codigo/codegraph-vba/src/extraction/vba-preprocess.ts) with a secure, custom expression tokenizer and a recursive descent expression parser and evaluator.

### 1. Tokenizer (Lexer)
The tokenizer scans the input conditional expression character by character:
- **Whitespace**: Skipped.
- **Numbers**: Decimal integers matched by `\d+`.
- **Parentheses**: `(` and `)`.
- **Comparison Operators**: `<=`, `>=`, `<>`, `<`, `>`, `=`.
- **Unary Arithmetic Operators**: `-` (unary minus) and `+` (unary plus) to support signed numbers and negated expressions.
- **Keywords/Identifiers**: `[A-Za-z_][A-Za-z0-9_]*`.
  - Case-insensitive operators: `And`, `Or`, `Not`, `Xor`.
  - Identifiers resolved case-insensitively using:
    1. The scoped `#Const` table (`constTable`).
    2. Hardcoded environment constants:
       - `VBA7` -> `-1`
       - `Win64` -> `-1`
       - `Win32` -> `-1`
       - `Win16` -> `0`
       - `Mac` -> `0`
       - `True` -> `-1`
       - `False` -> `0`
    3. Unknown/undefined identifiers fallback to `0`.
  - Resolved identifiers are emitted directly as `NUMBER` tokens containing the resolved integer value.

### 2. Parser & Evaluator
We use a recursive descent parser that performs on-the-fly expression evaluation. The parser enforces VBA's operator precedence:

| Level | Operators | Description | Associativity |
|---|---|---|---|
| 1 (Highest) | `( ... )`, number literals | Primary expressions | N/A |
| 2 | Unary `-`, `+` | Unary arithmetic signs | Right-to-Left |
| 3 | `Not` | Unary bitwise negation | Right-to-Left |
| 4 | `=`, `<>`, `<`, `<=`, `>`, `>=` | Comparisons | Left-to-Right |
| 5 | `And` | Bitwise AND | Left-to-Right |
| 6 | `Xor` | Bitwise XOR | Left-to-Right |
| 7 (Lowest) | `Or` | Bitwise OR | Left-to-Right |

All bitwise operations are evaluated using JavaScript's bitwise operators (`&`, `|`, `^`, `~`), which naturally coerce operands to 32-bit signed integers. We explicitly force 32-bit signed integer coercion (via `| 0`) at every evaluation step (e.g. unary negation, comparisons, and logical operations).

### 3. Comparison Semantics
Comparison operators return VBA's truthy/falsy values:
- `True` is represented as `-1`.
- `False` is represented as `0`.

### 4. Safety & Fallbacks
- No `eval` or `Function` calls.
- If the tokenizer or parser encounters an unrecognized character or syntax error (e.g., mismatched parentheses or unexpected trailing tokens), it throws an error.
- The entry points `evaluateConditionalExpression` and `evaluateConstRhs` catch any parsing/lexing errors and return safe defaults (`false` and `null` respectively), preventing runtime crashes.

---

## Architecture Decisions

| Decision | Choice | Alternatives considered | Rationale |
|---|---|---|---|
| **Parsing Strategy** | Custom tokenizer + recursive descent parser | Regular expression substitution + JS `Function` execution | Prevents security/sandbox escaping risks, and allows precise control over VBA bitwise precedence and comparison semantics. |
| **AST-less Evaluation** | On-the-fly evaluation inside the parser | Build a AST (Abstract Syntax Tree) then walk it | Since we only need the final evaluated scalar value, evaluating during parsing is faster and reduces code complexity. |
| **Token-Level Resolution** | Look up identifiers during tokenization | Resolve names via regex before parsing | Regex lookup is prone to word boundary / substring bugs. Resolving names in the tokenizer is robust and keeps the parser clean. |
| **Unary Sign Tokenization** | Tokenize `-`/`+` as operators | Tokenize negative numbers (e.g. `-5`) as a single number token | Treating `-` as an operator cleanly supports unary negation in front of variables (`-Win64`) and paren blocks (`-(5)`) without special-casing in the lexer. |

---

## Data Flow

```text
VBA Expression String
  │
  ▼
tokenize()
  ├─ Skip whitespace
  ├─ Match parentheses & comparisons
  ├─ Match & resolve identifiers case-insensitively (constTable -> hardcoded -> 0)
  └─ Produce Token Stream
  │
  ▼
Parser.parseExpression()
  ├─ parseOr() [Level 7]
  │   └─ parseXor() [Level 6]
  │       └─ parseAnd() [Level 5]
  │           └─ parseComparison() [Level 4]
  │               └─ parseNot() [Level 3]
  │                   └─ parseUnary() [Level 2]
  │                       └─ parsePrimary() [Level 1]
  ▼
Final Number Result (32-bit signed integer)
  ├─ For #If / #ElseIf (evaluateConditionalExpression): Convert non-zero to true, zero to false
  └─ For #Const (evaluateConstRhs): Convert to string representation (e.g. "-1", "0")
```

---

## File Changes

| File | Action | Description |
|---|---|---|
| [vba-preprocess.ts](file:///C:/00repos/codigo/codegraph-vba/src/extraction/vba-preprocess.ts) | Modify | Remove `normalizeConditionalExpression`. Add custom `Token` interface, `tokenize` function, and `Parser` class. Replace the implementations of `evaluateConditionalExpression` and `evaluateConstRhs` with tokenizer/parser invocation. |
| [extraction-vba-preprocess.test.ts](file:///C:/00repos/codigo/codegraph-vba/__tests__/extraction-vba-preprocess.test.ts) | Modify | Add comprehensive tests for precedence, bitwise logic (`And`/`Or`/`Not`/`Xor`), comparison return values (`-1`/`0`), unary operators, error recovery, and case-insensitivity. |

---

## Interfaces / Contracts

```ts
interface Token {
  type:
    | 'NUMBER'
    | 'PAREN_OPEN'
    | 'PAREN_CLOSE'
    | 'OP_NOT'
    | 'OP_AND'
    | 'OP_OR'
    | 'OP_XOR'
    | 'COMP_LE'
    | 'COMP_GE'
    | 'COMP_NE'
    | 'COMP_LT'
    | 'COMP_GT'
    | 'COMP_EQ'
    | 'OP_MINUS'
    | 'OP_PLUS'
    | 'EOF';
  numberValue?: number;
}

function tokenize(expr: string, constTable: ReadonlyMap<string, string>): Token[];

class Parser {
  constructor(tokens: Token[]);
  public parseExpression(): number;
  public ensureEOF(): void;
}
```

---

## NodeKind / Edge Mapping

Not applicable. Conditional compilation preprocessing happens before any AST nodes or edges are generated.

---

## Testing Strategy

| Layer | What to Test | Approach |
|---|---|---|
| Unit | Correct operator precedence | Verify that `3 And 4 Or 1` evaluates correctly depending on order (And > Or). |
| Unit | Signed 32-bit integer arithmetic | Test bitwise overflow/negative semantics (e.g. `Not 0` is `-1`, `Not -1` is `0`). |
| Unit | Comparison return values | Verify comparison operations return `-1` and `0` specifically. |
| Unit | Case-insensitivity | Test keywords (`aNd`, `xOr`) and constants (`vBa7`, `WIN64`) with different casings. |
| Unit | Robust error recovery | Ensure syntax errors (`FLAGS InvalidSyntax @@@`) fallback to `false` without crashing the process. |

---

## Migration / Rollout

No migration is required as this is a drop-in replacement for the internal preprocessing evaluator.

---

## Open Questions

None.
