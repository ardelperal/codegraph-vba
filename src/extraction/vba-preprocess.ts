/**
 * VBA pre-processing pipeline.
 *
 * Pure helpers used by `VbaExtractor` before and during regex extraction:
 *
 *   1. joinLineContinuations(src)
 *      VBA uses a trailing ` _` (space + underscore) at the end of a line to
 *      mark line continuation. We collapse these into single logical lines so
 *      the regex sweeps downstream see one statement per regex match.
 *
 *   2. stripVbaComments(src)
 *      VBA has two comment forms:
 *        - apostrophe `'` — anything from the apostrophe to end-of-line
 *        - `Rem ` (case-insensitive keyword, with trailing space) — the legacy
 *          alternative to `'`
 *      Both are STRIPPED only when outside double-quoted strings. We walk
 *      character-by-character so a `'` inside `"..."` is preserved.
 *
 *   3. preprocessConditionalCompilation(src)
 *      Blanks inactive VBA conditional-compilation branches (`#If`, `#ElseIf`,
 *      `#Else`, `#End If`) while preserving line count.
 *
 *   4. extractStringLiterals(src)
 *      Used by SQL-variable tracking to read literal fragments from assignments
 *      such as `m_SQL = "SELECT ..." & ...`. It returns every `"..."` span
 *      with its 1-based line and 0-based column. VBA doubles `"` inside literals
 *      (`""`) as the escape for a single `"` — the escape is consumed, the
 *      literal is returned without the doubling.
 *
 * Each helper is < 50 LOC, pure, and tested in isolation in
 * `__tests__/extraction-vba-preprocess.test.ts`.
 *
 * Source unchanged guarantee: every helper treats `src` as read-only.
 */

// -----------------------------------------------------------------------------
// 1. joinLineContinuations
// -----------------------------------------------------------------------------

/**
 * Collapse VBA line continuations: any line whose last non-newline character is
 * `_` (preceded by exactly one space, per VBA convention) is joined to the
 * next line. A single space replaces the joined boundary.
 *
 * - Input  : multi-line source
 * - Output : source with continuation lines merged into the previous line
 *
 * Does NOT inspect string literals — VBA's continuation convention is line-
 * physical (the underscore must be the last non-space char on the line), so a
 * `"` followed by `_` inside a string is still a line continuation. In real
 * Dysflow-exported code this never happens (continuations are syntactic, not
 * inside string content), so we skip the cost of tracking string state.
 */
export function joinLineContinuations(src: string): string {
  if (!src) return src;
  // VBA convention: " _\n" → join. **Preserve the newline AND the
  // leading space** so the transformed source has the same line count
  // AND the same intra-line whitespace as the original — the sweep
  // patterns downstream are anchored and the space before `_` is
  // meaningful for the joined line.
  //
  // Dangling ` _` at end-of-file (malformed input defensive) has no
  // newline to preserve; replace with a single space.
  return src.replace(/ _(?:(\r?\n)|$)/g, (_m, nl) => ' ' + (nl ?? ''));
}

// -----------------------------------------------------------------------------
// 2. stripVbaComments
// -----------------------------------------------------------------------------

const REM_PREFIX = /^Rem(\s|$)/i;
// Fix 5: handle trailing bare `Rem` at EOL (`\s|$` instead of `\s` alone).
const REM_MIDLINE = /\s+Rem(\s|$)/i;
const OPTION_DIRECTIVE = /^\s*Option\s+\w+/i;

/**
 * Strip VBA comments from source. Walks character-by-character:
 *  - inside a double-quoted string: skip until the closing `"` (handling the
 *    `""` doubled-quote escape so it doesn't terminate the string early)
 *  - outside a string:
 *      * `'` to EOL → drop the rest of the line
 *      * line-start `Rem ` (case-insensitive) → drop the whole line
 *      * mid-line ` Rem ` (case-insensitive, with surrounding whitespace) →
 *        drop the rest of the line
 *      * `Option <word>` directive lines → drop the whole line
 *      * newline → end-of-line marker for the apostrophe rule
 *
 * The trailing whitespace of each stripped line is trimmed so a `code ' comment`
 * becomes `code` (not `code `), keeping downstream regex anchored patterns
 * from drifting on whitespace artifacts.
 *
 * Source is treated as read-only.
 */
export function stripVbaComments(src: string): string {
  if (!src) return src;
  const lines = src.split('\n');
  const out: string[] = [];

  for (const rawLine of lines) {
    // Whole-line `Rem` comment.
    if (REM_PREFIX.test(rawLine)) {
      // Push an empty placeholder so the line count matches the original
      // source — node `startLine` values computed downstream as `i + 1`
      // depend on line-count parity. Same for Option directives below.
      out.push('');
      continue;
    }

    // Option directives — inert compiler hints that emit no symbol.
    if (OPTION_DIRECTIVE.test(rawLine)) {
      out.push('');
      continue;
    }

    let codeOnly = '';
    let inString = false;
    let i = 0;

    while (i < rawLine.length) {
      const ch = rawLine[i];
      const next = rawLine[i + 1];

      if (inString) {
        if (ch === '"' && next === '"') {
          codeOnly += '""';
          i += 2;
          continue;
        }
        if (ch === '"') {
          inString = false;
          codeOnly += '"';
          i++;
          continue;
        }
        codeOnly += ch;
        i++;
        continue;
      }

      if (ch === '"') {
        inString = true;
        codeOnly += '"';
        i++;
        continue;
      }
      if (ch === "'") {
        // Apostrophe starts a comment — drop the rest of the line.
        break;
      }
      codeOnly += ch;
      i++;
    }

    // Trim trailing whitespace left over from a stripped comment. Also
    // strip a mid-line ` Rem ` (legacy Rem syntax allowed anywhere with
    // surrounding whitespace) — the character-by-character walk above only
    // handles `'`; VBA's `Rem` is keyword-driven and may appear mid-line.
    //
    // **String-aware**: the previous implementation ran `REM_MIDLINE`
    // against the whole stripped line and would silently truncate string
    // literals containing " Rem " (e.g. `"SELECT Rem FROM tbl"` became
    // `"SELECT`). Audit W1 (June 2026). The fix splits into alternating
    // code / string segments and only applies the regex to code segments.
    let stripped = codeOnly.replace(/\s+$/, '');
    stripped = stripRemInCodeSegments(stripped);
    out.push(stripped);
  }

  return out.join('\n');
}

// -----------------------------------------------------------------------------
// 3. preprocessConditionalCompilation
// -----------------------------------------------------------------------------

interface ConditionalFrame {
  parentActive: boolean;
  active: boolean;
  branchTaken: boolean;
}

const IF_DIRECTIVE = /^\s*#If\s+(.+?)\s+Then\s*$/i;
const ELSEIF_DIRECTIVE = /^\s*#ElseIf\s+(.+?)\s+Then\s*$/i;
const ELSE_DIRECTIVE = /^\s*#Else\s*$/i;
const ENDIF_DIRECTIVE = /^\s*#(?:End\s*If|EndIf)\s*$/i;
// #Const NAME = <value> — captured name (identifier) and value (anything
// to EOL). The directive line is always blanked; the value is fed through
// the same normalize+evaluate pipeline as a #If expression so it inherits
// the substitution table for free.
const CONST_DIRECTIVE = /^\s*#Const\s+([A-Za-z_][A-Za-z_0-9]*)\s*=\s*(.+?)\s*$/i;

/**
 * Evaluate VBA conditional-compilation directives for the modern Windows
 * Access/VBA target this extractor is designed around:
 *   - VBA7   = true  (VBA7 runtime available)
 *   - Win64  = true  (64-bit VBA host — Access/VBA on x64 Windows)
 *   - Win32  = true  (legacy guard; always true on modern Windows, incl. Win64)
 *   - Win16  = false (legacy 16-bit Windows — not a current target)
 *   - Mac    = false (Mac host)
 *
 * In addition, the file-scoped `#Const NAME = <value>` table is consulted
 * before the hardcoded constants above — `#Const MODO_DEBUG = True` then
 * `#If MODO_DEBUG Then` keeps the branch. The #Const line itself is
 * blanked to preserve line-count parity.
 *
 * Truthiness follows VBA semantics: True = -1, False = 0. We achieve this
 * by substituting the JS boolean literals `true`/`false` with their VBA
 * numeric equivalents `-1`/`0` after the identifier+operator rewrites and
 * before the whitelist check. This is a deliberate simplification — full
 * VBA CC is bitwise on -1/0 (so `True And True = -1`), but the supported
 * expression surface here is truthy comparison / `=` / `<>` / `And` /
 * `Or` / `Not`, for which JS `&&`/`||`/`!` truthy evaluation is
 * equivalent: any non-zero operand is truthy, matching VBA's "non-zero is
 * true" convention. If a future task needs bitwise-precise `-1` semantics
 * (e.g. distinguishing `#If X = 1` from `#If X = -1` on a `True` const),
 * promote this evaluator to full integer arithmetic.
 *
 * Directives and inactive branch lines are replaced with empty strings so
 * downstream extraction keeps source-line parity. Unsupported/unsafe
 * expressions evaluate to false rather than throwing.
 */
export function preprocessConditionalCompilation(src: string, customTargets?: Record<string, boolean>): string {
  if (!src) return src;
  const lines = src.split('\n');
  const out: string[] = [];
  const stack: ConditionalFrame[] = [];
  // Per-call #Const table. Name → post-evaluation numeric string (e.g.
  // "-1" for True, "0" for False, "1" for an integer literal). Stored as a
  // string so it can be substituted verbatim into the normalized
  // expression. A name that fails to evaluate is NOT stored — the lookup
  // simply misses and the conservative fallback (unknown identifier →
  // false) applies, exactly as for an unrecognised hardcoded constant.
  const constTable = new Map<string, string>();

  for (const line of lines) {
    // Parse #Const BEFORE the other directives so the table is up-to-date
    // when a subsequent #If expression references it. The line itself is
    // always blanked (line-count parity invariant).
    const constMatch = CONST_DIRECTIVE.exec(line);
    if (constMatch) {
      const name = (constMatch[1] ?? '').trim();
      const rhs = (constMatch[2] ?? '').trim();
      const value = evaluateConstRhs(rhs, constTable, customTargets);
      if (value !== null) {
        constTable.set(name, value);
      }
      out.push('');
      continue;
    }

    const ifMatch = IF_DIRECTIVE.exec(line);
    if (ifMatch) {
      const parentActive = stack.every((frame) => frame.active);
      const active = parentActive && evaluateConditionalExpression(ifMatch[1] ?? '', constTable, customTargets);
      stack.push({ parentActive, active, branchTaken: active });
      out.push('');
      continue;
    }

    const elseIfMatch = ELSEIF_DIRECTIVE.exec(line);
    if (elseIfMatch) {
      const frame = stack[stack.length - 1];
      if (frame) {
        const active =
          frame.parentActive &&
          !frame.branchTaken &&
          evaluateConditionalExpression(elseIfMatch[1] ?? '', constTable, customTargets);
        frame.active = active;
        if (active) frame.branchTaken = true;
      }
      out.push('');
      continue;
    }

    if (ELSE_DIRECTIVE.test(line)) {
      const frame = stack[stack.length - 1];
      if (frame) {
        const active = frame.parentActive && !frame.branchTaken;
        frame.active = active;
        if (active) frame.branchTaken = true;
      }
      out.push('');
      continue;
    }

    if (ENDIF_DIRECTIVE.test(line)) {
      stack.pop();
      out.push('');
      continue;
    }

    out.push(stack.every((frame) => frame.active) ? line : '');
  }

  return out.join('\n');
}

/**
 * Apply the full substitution pipeline to a conditional-compilation
 * expression: identifier substitutions (#Const table first, then the
 * hardcoded constants), operator rewrites, and the True→-1 / False→0
 * conversion that gives the evaluator correct VBA equality semantics.
 *
 * Returns the normalized string if it passes the whitelist (signed
 * integer literal + operators + parens + whitespace only), or `null` if
 * any substitution left a token that does not match the whitelist. A
 * `null` return signals "do not evaluate" — the caller falls back to
 * false (the conservative behaviour preserved from the original
 * implementation).
 */
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

function tokenize(
  expr: string,
  constTable: ReadonlyMap<string, string>,
  customTargets?: Record<string, boolean>
): Token[] {
  const tokens: Token[] = [];
  let i = 0;

  while (i < expr.length) {
    const ch = expr[i];

    if (ch === undefined) {
      break;
    }

    // Skip whitespace
    if (/\s/.test(ch)) {
      i++;
      continue;
    }

    // Parentheses
    if (ch === '(') {
      tokens.push({ type: 'PAREN_OPEN' });
      i++;
      continue;
    }
    if (ch === ')') {
      tokens.push({ type: 'PAREN_CLOSE' });
      i++;
      continue;
    }

    // Comparison operators (multi-character first)
    if (ch === '<' && expr[i + 1] === '=') {
      tokens.push({ type: 'COMP_LE' });
      i += 2;
      continue;
    }
    if (ch === '>' && expr[i + 1] === '=') {
      tokens.push({ type: 'COMP_GE' });
      i += 2;
      continue;
    }
    if (ch === '<' && expr[i + 1] === '>') {
      tokens.push({ type: 'COMP_NE' });
      i += 2;
      continue;
    }
    if (ch === '<') {
      tokens.push({ type: 'COMP_LT' });
      i++;
      continue;
    }
    if (ch === '>') {
      tokens.push({ type: 'COMP_GT' });
      i++;
      continue;
    }
    if (ch === '=') {
      tokens.push({ type: 'COMP_EQ' });
      i++;
      continue;
    }

    // Unary/Binary signs
    if (ch === '-') {
      tokens.push({ type: 'OP_MINUS' });
      i++;
      continue;
    }
    if (ch === '+') {
      tokens.push({ type: 'OP_PLUS' });
      i++;
      continue;
    }

    // Number literals (decimal integers)
    if (/\d/.test(ch)) {
      const start = i;
      while (i < expr.length && /\d/.test(expr[i] ?? '')) {
        i++;
      }
      const numStr = expr.substring(start, i);
      tokens.push({ type: 'NUMBER', numberValue: parseInt(numStr, 10) | 0 });
      continue;
    }

    // Identifiers and Keywords
    if (/[A-Za-z_]/.test(ch)) {
      const start = i;
      while (i < expr.length && /[A-Za-z0-9_]/.test(expr[i] ?? '')) {
        i++;
      }
      const id = expr.substring(start, i);
      const idUpper = id.toUpperCase();

      // Keywords
      if (idUpper === 'NOT') {
        tokens.push({ type: 'OP_NOT' });
        continue;
      }
      if (idUpper === 'AND') {
        tokens.push({ type: 'OP_AND' });
        continue;
      }
      if (idUpper === 'OR') {
        tokens.push({ type: 'OP_OR' });
        continue;
      }
      if (idUpper === 'XOR') {
        tokens.push({ type: 'OP_XOR' });
        continue;
      }

      // Resolve Identifier case-insensitively
      let resolvedValue: number | null = null;
      for (const [key, val] of constTable.entries()) {
        if (key.toUpperCase() === idUpper) {
          resolvedValue = parseInt(val, 10) | 0;
          break;
        }
      }

      // Configured Preprocessor Targets (case-insensitively)
      if (resolvedValue === null && customTargets) {
        for (const [key, val] of Object.entries(customTargets)) {
          if (key.toUpperCase() === idUpper) {
            resolvedValue = val ? -1 : 0;
            break;
          }
        }
      }

      // Hardcoded environment constants
      if (resolvedValue === null) {
        switch (idUpper) {
          case 'VBA7':
          case 'WIN64':
          case 'WIN32':
          case 'TRUE':
            resolvedValue = -1;
            break;
          case 'WIN16':
          case 'MAC':
          case 'FALSE':
            resolvedValue = 0;
            break;
        }
      }

      // Fallback
      if (resolvedValue === null) {
        resolvedValue = 0;
      }

      tokens.push({ type: 'NUMBER', numberValue: resolvedValue });
      continue;
    }

    // Any other character is a lexing error
    throw new Error(`Unexpected character in expression: ${ch}`);
  }

  tokens.push({ type: 'EOF' });
  return tokens;
}

class Parser {
  private tokens: Token[];
  private current = 0;

  constructor(tokens: Token[]) {
    this.tokens = tokens;
  }

  private peek(): Token {
    return this.tokens[this.current] || { type: 'EOF' };
  }

  private advance(): Token {
    const t = this.peek();
    if (t.type !== 'EOF') {
      this.current++;
    }
    return t;
  }

  private match(type: Token['type']): boolean {
    if (this.peek().type === type) {
      this.advance();
      return true;
    }
    return false;
  }

  public parseExpression(): number {
    return this.parseOr();
  }

  public ensureEOF(): void {
    if (this.peek().type !== 'EOF') {
      throw new Error(`Unexpected trailing token: ${this.peek().type}`);
    }
  }

  // Level 7: Or (lowest precedence)
  private parseOr(): number {
    let expr = this.parseXor();
    while (this.match('OP_OR')) {
      const right = this.parseXor();
      expr = (expr | right) | 0;
    }
    return expr;
  }

  // Level 6: Xor
  private parseXor(): number {
    let expr = this.parseAnd();
    while (this.match('OP_XOR')) {
      const right = this.parseAnd();
      expr = (expr ^ right) | 0;
    }
    return expr;
  }

  // Level 5: And
  private parseAnd(): number {
    let expr = this.parseComparison();
    while (this.match('OP_AND')) {
      const right = this.parseComparison();
      expr = (expr & right) | 0;
    }
    return expr;
  }

  // Level 4: Comparisons (=, <>, <, <=, >, >=)
  private parseComparison(): number {
    let expr = this.parseAdditive();
    while (true) {
      const op = this.peek().type;
      if (
        op === 'COMP_EQ' ||
        op === 'COMP_NE' ||
        op === 'COMP_LT' ||
        op === 'COMP_LE' ||
        op === 'COMP_GT' ||
        op === 'COMP_GE'
      ) {
        this.advance();
        const right = this.parseAdditive();
        let cmpResult = false;
        switch (op) {
          case 'COMP_EQ':
            cmpResult = expr === right;
            break;
          case 'COMP_NE':
            cmpResult = expr !== right;
            break;
          case 'COMP_LT':
            cmpResult = expr < right;
            break;
          case 'COMP_LE':
            cmpResult = expr <= right;
            break;
          case 'COMP_GT':
            cmpResult = expr > right;
            break;
          case 'COMP_GE':
            cmpResult = expr >= right;
            break;
        }
        expr = (cmpResult ? -1 : 0) | 0;
      } else {
        break;
      }
    }
    return expr;
  }

  // Level 3.5: Binary +, - (arithmetic binds tighter than comparison in VBA).
  // Wrapped with `| 0` so 32-bit signed overflow matches VBA integer math
  // (e.g. 2147483647 + 1 === -2147483648).
  private parseAdditive(): number {
    let expr = this.parseNot();
    while (true) {
      const op = this.peek().type;
      if (op === 'OP_PLUS') {
        this.advance();
        expr = (expr + this.parseNot()) | 0;
      } else if (op === 'OP_MINUS') {
        this.advance();
        expr = (expr - this.parseNot()) | 0;
      } else {
        break;
      }
    }
    return expr;
  }

  // Level 3: Not
  private parseNot(): number {
    if (this.match('OP_NOT')) {
      const expr = this.parseNot();
      return (~expr) | 0;
    }
    return this.parseUnary();
  }

  // Level 2: Unary -, +
  private parseUnary(): number {
    if (this.match('OP_MINUS')) {
      const expr = this.parseUnary();
      return (-expr) | 0;
    }
    if (this.match('OP_PLUS')) {
      const expr = this.parseUnary();
      return expr | 0;
    }
    return this.parsePrimary();
  }

  // Level 1: Primary (Paren, Number)
  private parsePrimary(): number {
    const t = this.peek();
    if (this.match('NUMBER')) {
      return (t.numberValue ?? 0) | 0;
    }
    if (this.match('PAREN_OPEN')) {
      const expr = this.parseExpression();
      if (!this.match('PAREN_CLOSE')) {
        throw new Error('Mismatched parenthesis: expected )');
      }
      return expr;
    }
    throw new Error(`Unexpected token: ${t.type}`);
  }
}

function evaluateConditionalExpression(
  expr: string,
  constTable: ReadonlyMap<string, string> = new Map(),
  customTargets?: Record<string, boolean>,
): boolean {
  try {
    let exprClean = expr.trim();
    if (exprClean.toLowerCase().endsWith('then')) {
      exprClean = exprClean.slice(0, -4).trim();
    }
    const tokens = tokenize(exprClean, constTable, customTargets);
    const parser = new Parser(tokens);
    const result = parser.parseExpression();
    parser.ensureEOF();
    return result !== 0;
  } catch {
    return false;
  }
}

function evaluateConstRhs(
  rhs: string,
  constTable: ReadonlyMap<string, string>,
  customTargets?: Record<string, boolean>,
): string | null {
  try {
    const tokens = tokenize(rhs.trim(), constTable, customTargets);
    const parser = new Parser(tokens);
    const result = parser.parseExpression();
    parser.ensureEOF();
    return String(result);
  } catch {
    return null;
  }
}

/**
 * Split `line` into alternating code / string-literal segments and apply
 * `REM_MIDLINE` only to code segments. String segments are returned
 * verbatim — a string like `"SELECT Rem FROM tbl"` is NOT truncated.
 */
function stripRemInCodeSegments(line: string): string {
  const segments: Array<{ text: string; isString: boolean }> = [];
  let buf = '';
  let inString = false;
  let i = 0;
  while (i < line.length) {
    const ch = line[i];
    const next = line[i + 1];
    if (inString) {
      buf += ch;
      if (ch === '"' && next === '"') {
        buf += next;
        i += 2;
        continue;
      }
      if (ch === '"') {
        inString = false;
        segments.push({ text: buf, isString: true });
        buf = '';
        i++;
        continue;
      }
      i++;
      continue;
    }
    if (ch === '"') {
      if (buf) {
        segments.push({ text: buf, isString: false });
        buf = '';
      }
      inString = true;
      buf += ch;
      i++;
      continue;
    }
    buf += ch;
    i++;
  }
  if (buf) segments.push({ text: buf, isString: inString });
  // Fix 5: once a `Rem` comment marker is found in a code segment, discard
  // that segment from the Rem position onward AND all subsequent segments
  // (which may be string literals from the comment body — they must not
  // produce false SQL table references).
  let remFound = false;
  for (const s of segments) {
    if (remFound) {
      s.text = '';
      continue;
    }
    if (!s.isString) {
      const m = REM_MIDLINE.exec(s.text);
      if (m) {
        s.text = s.text.slice(0, m.index).replace(/\s+$/, '');
        remFound = true;
      }
    }
  }
  return segments.map((s) => s.text).join('');
}

// -----------------------------------------------------------------------------
// 4. extractStringLiterals
// -----------------------------------------------------------------------------

export interface StringLiteralSpan {
  /** The literal content (no surrounding quotes; `""` collapsed to `"`). */
  text: string;
  /** 1-based line number where the opening quote sits. */
  line: number;
  /** 0-based column where the opening quote sits. */
  column: number;
}

/**
 * Walk source and record every double-quoted string literal with its 1-based
 * line and 0-based column. VBA's `""` escape is collapsed to a single `"` in
 * the recorded `text`.
 *
 * The source is NOT mutated.
 */
export function extractStringLiterals(src: string): StringLiteralSpan[] {
  if (!src) return [];
  const out: StringLiteralSpan[] = [];
  let line = 1;
  let col = 0;
  let i = 0;

  while (i < src.length) {
    const ch = src[i];
    if (ch === '"') {
      const startLine = line;
      const startCol = col;
      let text = '';
      i++;
      col++;
      while (i < src.length) {
        const c = src[i];
        if (c === '"' && src[i + 1] === '"') {
          // Escape: keep one `"` in the recorded text.
          text += '"';
          i += 2;
          col += 2;
          continue;
        }
        if (c === '"') {
          // Closing quote.
          i++;
          col++;
          break;
        }
        if (c === '\n') {
          // Unterminated string at EOL — stop walking it but record what we have.
          // (VBA actually does support multi-line string literals via embedded
          // vbCrLf, but Dysflow never emits them; defensive stop is fine.)
          break;
        }
        text += c;
        i++;
        col++;
      }
      out.push({ text, line: startLine, column: startCol });
      continue;
    }
    if (ch === '\n') {
      line++;
      col = 0;
      i++;
      continue;
    }
    i++;
    col++;
  }

  return out;
}
