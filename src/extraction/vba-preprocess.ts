/**
 * VBA pre-processing pipeline.
 *
 * Three pure helpers run in order at the top of `VbaExtractor.extract()`:
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
 *   3. extractStringLiterals(src)
 *      Walks the source once and returns every `"..."` span with its 1-based
 *      line and 0-based column. VBA doubles `"` inside literals (`""`) as the
 *      escape for a single `"` — the escape is consumed, the literal is
 *      returned without the doubling.
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

const REM_PREFIX = /^Rem\s/i;
const REM_MIDLINE = /\s+Rem\s/i;
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
    let stripped = codeOnly.replace(/\s+$/, '');
    const remMid = REM_MIDLINE.exec(stripped);
    if (remMid) {
      stripped = stripped.slice(0, remMid.index).replace(/\s+$/, '');
    }
    out.push(stripped);
  }

  return out.join('\n');
}

// -----------------------------------------------------------------------------
// 3. extractStringLiterals
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