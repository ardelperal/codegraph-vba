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

/**
 * Evaluate VBA conditional-compilation directives for the modern Windows
 * Access/VBA target this extractor is designed around:
 *   - VBA7 = true
 *   - Win64 = true
 *   - Mac = false
 *
 * Directives and inactive branch lines are replaced with empty strings so
 * downstream extraction keeps source-line parity. Unsupported/unsafe
 * expressions evaluate to false rather than throwing.
 */
export function preprocessConditionalCompilation(src: string): string {
  if (!src) return src;
  const lines = src.split('\n');
  const out: string[] = [];
  const stack: ConditionalFrame[] = [];

  for (const line of lines) {
    const ifMatch = IF_DIRECTIVE.exec(line);
    if (ifMatch) {
      const parentActive = stack.every((frame) => frame.active);
      const active = parentActive && evaluateConditionalExpression(ifMatch[1] ?? '');
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
          evaluateConditionalExpression(elseIfMatch[1] ?? '');
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

function evaluateConditionalExpression(expr: string): boolean {
  let normalized = expr.trim();
  normalized = normalized.replace(/\bThen\s*$/i, '');
  normalized = normalized.replace(/<>/g, '!==');
  normalized = normalized.replace(/(?<![<>=])=(?![=])/g, '===');
  normalized = normalized.replace(/\bVBA7\b/gi, 'true');
  normalized = normalized.replace(/\bWin64\b/gi, 'true');
  normalized = normalized.replace(/\bMac\b/gi, 'false');
  normalized = normalized.replace(/\bAnd\b/gi, '&&');
  normalized = normalized.replace(/\bOr\b/gi, '||');
  normalized = normalized.replace(/\bNot\b/gi, '!');
  normalized = normalized.trim();

  if (!/^(?:true|false|\d+|&&|\|\||!|===|!==|\(|\)|\s)+$/.test(normalized)) {
    return false;
  }

  try {
    return Boolean(Function(`"use strict"; return (${normalized});`)());
  } catch {
    return false;
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
