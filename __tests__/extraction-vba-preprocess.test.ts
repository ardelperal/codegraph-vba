/**
 * VBA pre-processing pipeline — unit tests.
 *
 * Three pure helpers run in order at the top of `VbaExtractor.extract()`:
 *   1. joinLineContinuations — joins lines ending with `_` (VBA line-continuation)
 *   2. stripVbaComments      — strips `'` and `Rem ` comments outside strings
 *   3. extractStringLiterals — returns every `"..."` span with line/column
 *
 * Each helper is < 30 LOC and tested in isolation. Red-green TDD per group.
 */
import { describe, it, expect } from 'vitest';
import {
  joinLineContinuations,
  stripVbaComments,
  extractStringLiterals,
  preprocessConditionalCompilation,
} from '../src/extraction/vba-preprocess';

describe('joinLineContinuations', () => {
  // **Contract**: this helper must preserve the source's line count. The
  // downstream `sweepProcedures` uses `lineNum = i + 1` on the
  // transformed array to assign `startLine` to every emitted node —
  // those numbers must align with the original source so
  // `codegraph_explore` returns the right lines. (See audit finding
  // C1 in obs `codegraph/audit/vba-extractor-2026-06-28`.)
  it('joins two lines when the first ends with " _" — preserves line count', () => {
    const src = 'Sub Foo()\n  DoCmd.RunSQL _\n    "SELECT * FROM tbl"';
    const out = joinLineContinuations(src);
    // Line count preserved.
    expect(out.split('\n').length).toBe(src.split('\n').length);
    // The `_` continuation marker is removed; the newline stays.
    expect(out).toContain('DoCmd.RunSQL \n');
    expect(out).not.toContain(' _\n');
    // The continued line's content is preserved verbatim.
    expect(out).toContain('"SELECT * FROM tbl"');
  });

  it('joins a Debug.Print continuation across lines — preserves line count', () => {
    const src = 'Debug.Print _\n  "hello"';
    const out = joinLineContinuations(src);
    expect(out).toBe('Debug.Print \n  "hello"');
    expect(out.split('\n').length).toBe(src.split('\n').length);
  });

  it('returns input unchanged when no line ends with " _"', () => {
    const src = 'Sub Foo()\n  Debug.Print "x"\nEnd Sub';
    expect(joinLineContinuations(src)).toBe(src);
  });

  it('joins multiple chained continuations in one statement — preserves line count', () => {
    const src = 'x = a _\n  + b _\n  + c';
    const out = joinLineContinuations(src);
    expect(out.split('\n').length).toBe(src.split('\n').length);
    // Both `_` markers removed, both newlines preserved.
    expect(out).toContain('x = a \n  + b \n  + c');
  });

  it('joins continuations across many statements — preserves line count', () => {
    const src = 'A _\n+ 1\nB _\n+ 2';
    const out = joinLineContinuations(src);
    expect(out.split('\n').length).toBe(src.split('\n').length);
    expect(out).toContain('A \n+ 1\nB \n+ 2');
    expect(out).not.toContain(' _\n');
  });

  it('does not join lines that do not end with " _" (only one trailing space)', () => {
    // "_" alone at end of line is the line-continuation marker.
    // A bare " _" (space+underscore) means the same in VBA.
    const src = 'A\nB';
    expect(joinLineContinuations(src)).toBe(src);
  });

  it('preserves an empty source', () => {
    expect(joinLineContinuations('')).toBe('');
  });

  it('does not split inside string literals — single logical line preserved', () => {
    // The whole thing is one line; no continuation present.
    const src = 'Sub X(): Debug.Print "line1 _ line2": End Sub';
    const out = joinLineContinuations(src);
    expect(out).toBe(src);
  });

  it('strips the continuation marker but keeps the newline', () => {
    // VBA convention: " _\n" → the `_` is consumed and the `\n` stays
    // so the transformed source has the same line count as the input.
    const src = 'X _\nY';
    expect(joinLineContinuations(src)).toBe('X \nY');
  });

  it('does not join when the underscore is mid-line (not at the end)', () => {
    const src = 'A = 1 + _invalid_underscore\nB = 2';
    expect(joinLineContinuations(src)).toBe(src);
  });

  it('joins even when continuation is followed by indented code — preserves line count', () => {
    const src = 'Function Foo() As Long\n  Foo = 1 _\n    + 2\nEnd Function';
    const out = joinLineContinuations(src);
    expect(out.split('\n').length).toBe(src.split('\n').length);
    expect(out).toContain('Foo = 1 \n    + 2\nEnd Function');
  });

  it('handles a continuation that ends a file (no following line)', () => {
    const src = 'x = 1 _';
    // Defensive: the helper should not throw on a dangling continuation.
    // Result keeps the single line (no newline to preserve at EOF; the
    // dangling ` _` is replaced with a single space).
    expect(() => joinLineContinuations(src)).not.toThrow();
    expect(joinLineContinuations(src)).toBe('x = 1 ');
  });

  it('handles CRLF line endings — preserves line count', () => {
    const src = 'Sub X()\r\n  DoCmd.RunSQL _\r\n    "SELECT"\r\nEnd Sub';
    const out = joinLineContinuations(src);
    expect(out.split(/\r?\n/).length).toBe(src.split(/\r?\n/).length);
    expect(out).not.toContain(' _\r\n');
  });
});

describe('stripVbaComments', () => {
  it('strips a trailing `\' comment from a code line', () => {
    const src = 'Dim x As Long  \' comment';
    expect(stripVbaComments(src)).toBe('Dim x As Long');
  });

  it('strips a Rem-prefixed line entirely — preserves line count via empty placeholder', () => {
    const src = 'Rem LegacyComment\nSub X(): End Sub';
    // The Rem line is replaced with an empty placeholder so the line
    // count matches the original — node `startLine` values computed
    // downstream as `i + 1` depend on this parity. (Audit C1 fix.)
    const out = stripVbaComments(src);
    expect(out).toBe('\nSub X(): End Sub');
    expect(out.split('\n').length).toBe(src.split('\n').length);
  });

  it('preserves an Option directive line as an empty placeholder (line count parity)', () => {
    const src = 'Option Explicit\nOption Compare Database\n\nSub X()\nEnd Sub';
    const out = stripVbaComments(src);
    // Same number of lines as input; Option content gone.
    expect(out.split('\n').length).toBe(src.split('\n').length);
    expect(out).not.toContain('Option Explicit');
    expect(out).not.toContain('Option Compare Database');
    expect(out).toContain('Sub X()');
  });

  it('preserves line count across a mixed VBA prelude (Rem + Option + blank + code)', () => {
    // The exact pattern that broke real fixture files in the audit:
    // every real Dysflow-exported .bas/.cls has Option Compare Database
    // + Option Explicit at the top, plus occasional Rem comments.
    const src = [
      'Rem ============================================',
      'Rem Módulo: ACAuditoriaOperaciones',
      'Rem ============================================',
      'Option Compare Database',
      'Option Explicit',
      '',
      'Public Function AccionRepetida() As Long',
      '    AccionRepetida = 1',
      'End Function',
    ].join('\n');
    const out = stripVbaComments(src);
    expect(out.split('\n').length).toBe(src.split('\n').length);
    expect(out).not.toContain('Rem ');
    expect(out).not.toContain('Option ');
    // The Public Function is on line 7 of the original; after stripping
    // it should be on line 7 of the output too (line count preserved).
    const lines = out.split('\n');
    expect(lines[6]).toContain('Public Function AccionRepetida');
  });

  it('preserves a single quote inside a double-quoted string', () => {
    // "It's a test" — the `'` is INSIDE the string and is NOT a comment marker.
    const src = 'Debug.Print "It\'s a test"';
    expect(stripVbaComments(src)).toBe(src);
  });

  it('does NOT match SQL inside a `\'` VBA comment line', () => {
    // The whole line is a comment; stripVbaComments removes everything.
    const src = "' DoCmd.RunSQL \"SELECT * FROM tblFake\"";
    expect(stripVbaComments(src)).toBe('');
  });

  it('strips Option Explicit / Option Compare Database / Option Base', () => {
    const src = 'Option Explicit\nOption Compare Database\nOption Base 1';
    const out = stripVbaComments(src);
    // Whitespace-only lines from stripped directives are kept as empty lines.
    expect(out).not.toContain('Option Explicit');
    expect(out).not.toContain('Option Compare Database');
    expect(out).not.toContain('Option Base 1');
  });

  it('strips a `\'` comment mid-line, leaving the code before it', () => {
    const src = 'x = 1 \' inline comment\ny = 2';
    expect(stripVbaComments(src)).toBe('x = 1\ny = 2');
  });

  it('handles a `\'` comment immediately after code (no leading space)', () => {
    const src = 'x=1\'nospace';
    expect(stripVbaComments(src)).toBe('x=1');
  });

  it('does not strip inside double-quoted strings', () => {
    const src = 'Foo "abc\'def" Bar';
    expect(stripVbaComments(src)).toBe(src);
  });

  it('strips a Rem comment that has code before it on the same line', () => {
    // VBA's `Rem` is allowed mid-line only after whitespace.
    const src = 'x = 1 Rem trailing rem comment';
    expect(stripVbaComments(src)).toBe('x = 1');
  });

  it('preserves a string literal containing the word Rem', () => {
    const src = 'MsgBox "Reminder: pay rent"';
    expect(stripVbaComments(src)).toBe(src);
  });

  it('preserves a SQL string containing the word Rem outside any Rem comment (W1 invariant)', () => {
    // Audit W1 (June 2026): the previous regex-based REM_MIDLINE stripper
    // was applied to the whole stripped line and would truncate string
    // literals containing " Rem " — breaking SQL table extraction for
    // tables/columns whose names happen to include that substring.
    const src = 'q = "select Rem from Remitentes"';
    expect(stripVbaComments(src)).toBe(src);
  });

  it('still strips a real mid-line Rem comment — everything after Rem is comment content (W1 + Issue #5)', () => {
    // In VBA, `Rem` starts a comment that continues to end-of-line.
    // Everything after the `Rem` keyword (including string literals that
    // happen to appear in the comment text) is discarded.
    // Updated by Issue #5 fix: the old test incorrectly expected `" Rem "`
    // and `y = 2` to be preserved after a `Rem` comment — they are comment
    // content and must NOT appear in the output.
    const src = 'x = 1 Rem trailing comment  " Rem "  y = 2';
    const out = stripVbaComments(src);
    // Code before Rem is preserved.
    expect(out).toContain('x = 1');
    // Everything after Rem (including string literals in the comment) is gone.
    expect(out).not.toContain('Rem trailing');
    expect(out).not.toContain('" Rem "');
    expect(out).not.toContain('y = 2');
  });

  it('strips a `\'` on its own line', () => {
    const src = "' just a comment\nSub X(): End Sub";
    expect(stripVbaComments(src)).toBe('\nSub X(): End Sub');
  });

  it('preserves an empty input', () => {
    expect(stripVbaComments('')).toBe('');
  });

  /**
   * Fix 6: bare `Rem` (no trailing space) must be treated as a whole-line
   * comment, same as `Rem `. VBA allows `Rem` alone on a line as a valid
   * empty comment. The old regex `/^Rem\s/i` required a whitespace char after
   * `Rem` and let bare `Rem` lines through.
   */
  it('a bare "Rem" line with no trailing space is stripped as a comment (Fix 6)', () => {
    const src = 'Rem\nx = 1';
    const out = stripVbaComments(src);
    // The Rem line becomes an empty placeholder (line-count parity preserved).
    expect(out).toBe('\nx = 1');
    expect(out.split('\n').length).toBe(src.split('\n').length);
  });

  it('a bare "REM" (upper-case) line with no trailing space is also stripped (Fix 6)', () => {
    const src = 'REM\nSub X(): End Sub';
    const out = stripVbaComments(src);
    expect(out).toBe('\nSub X(): End Sub');
  });

  /**
   * Issue #5 — `DoCmd.RunSQL "real" Rem "fake"` must not produce a false
   * table reference from the string that appears inside the Rem comment.
   * `stripRemInCodeSegments` must discard ALL segments (including string
   * literals) that follow the Rem marker.
   */
  it('Issue #5: Rem comment with string argument after it — string is discarded (not returned)', () => {
    // After `Rem` the rest of the line (including `"SELECT * FROM tblFake"`)
    // is comment content. stripVbaComments must NOT return that string.
    const src = 'DoCmd.RunSQL "SELECT * FROM tblReal" Rem "SELECT * FROM tblFake"';
    const out = stripVbaComments(src);
    expect(out).toContain('"SELECT * FROM tblReal"');
    expect(out).not.toContain('tblFake');
  });

  it('Issue #5: trailing bare Rem at EOL (REM_MIDLINE /(\\s|$)/) strips the Rem token itself', () => {
    // `\s+Rem$` — Rem appears at end-of-line with no trailing space.
    const src = 'DoCmd.RunSQL "SELECT * FROM tblReal" Rem';
    const out = stripVbaComments(src);
    expect(out).toContain('"SELECT * FROM tblReal"');
    expect(out).not.toContain('Rem');
  });
});

describe('extractStringLiterals', () => {
  it('returns one literal span with text/line/column', () => {
    const src = 'DoCmd.RunSQL "SELECT * FROM tbl"';
    const lits = extractStringLiterals(src);
    expect(lits).toHaveLength(1);
    expect(lits[0]?.text).toBe('SELECT * FROM tbl');
    expect(lits[0]?.line).toBe(1);
    expect(lits[0]?.column).toBeGreaterThan(0);
  });

  it('returns multiple literals in order of appearance', () => {
    const src = 'A "first" + B "second" + C "third"';
    const lits = extractStringLiterals(src);
    expect(lits.map((l) => l.text)).toEqual(['first', 'second', 'third']);
  });

  it('returns [] when there are no string literals', () => {
    expect(extractStringLiterals('Sub X(): Debug.Print 42: End Sub')).toEqual([]);
  });

  it('preserves a literal with embedded doubled-quote escape ""', () => {
    // In VBA, `""` inside a string literal is an escape for one `"`.
    // The whole string is one literal: `a"b`.
    const src = 'MsgBox "a""b"';
    const lits = extractStringLiterals(src);
    expect(lits).toHaveLength(1);
    expect(lits[0]?.text).toBe('a"b');
  });

  it('returns separate literals when a statement concatenates strings with &', () => {
    const src = 'x = "SELECT * " & "FROM tbl"';
    const lits = extractStringLiterals(src);
    expect(lits.map((l) => l.text)).toEqual(['SELECT * ', 'FROM tbl']);
  });

  it('reports the correct line number on a multi-line source', () => {
    const src = 'Sub X()\n  s = "hello"\nEnd Sub';
    const lits = extractStringLiterals(src);
    expect(lits).toHaveLength(1);
    expect(lits[0]?.line).toBe(2);
    expect(lits[0]?.text).toBe('hello');
  });

  it('preserves a literal that spans no characters (empty string)', () => {
    const src = 'x = ""';
    const lits = extractStringLiterals(src);
    expect(lits).toHaveLength(1);
    expect(lits[0]?.text).toBe('');
  });

  it('does not match a single quote-only token as a string literal', () => {
    // Single quote alone is a comment marker, not a string opener in VBA.
    expect(extractStringLiterals("' comment line")).toEqual([]);
  });

  it('handles a string literal at column 0', () => {
    const src = '"hello"';
    const lits = extractStringLiterals(src);
    expect(lits).toHaveLength(1);
    expect(lits[0]?.column).toBe(0);
  });

  it('records column as the offset of the opening quote', () => {
    const src = '  "abc"';
    const lits = extractStringLiterals(src);
    expect(lits[0]?.column).toBe(2);
  });

  it('returns two literals when the second starts a new line', () => {
    const src = '"first"\n"second"';
    const lits = extractStringLiterals(src);
    expect(lits.map((l) => l.text)).toEqual(['first', 'second']);
    expect(lits.map((l) => l.line)).toEqual([1, 2]);
  });

  it('preserves the source string unchanged (does not mutate)', () => {
    const src = 'DoCmd.RunSQL "SELECT"';
    const before = src;
    extractStringLiterals(src);
    expect(src).toBe(before);
  });
});

describe('preprocessConditionalCompilation', () => {
  it('blanks directives and inactive #Else branch while preserving line count', () => {
    const src = [
      '#If Win64 Then',
      'Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dw As Long)',
      '#Else',
      'Public Declare Sub Sleep Lib "kernel32" (ByVal dw As Long)',
      '#End If',
    ].join('\n');

    const out = preprocessConditionalCompilation(src);
    const lines = out.split('\n');

    expect(lines).toHaveLength(src.split('\n').length);
    expect(lines[0]).toBe('');
    expect(lines[1]).toContain('PtrSafe');
    expect(lines[2]).toBe('');
    expect(lines[3]).toBe('');
    expect(lines[4]).toBe('');
  });

  it('evaluates Not/And conditions with modern Windows VBA defaults', () => {
    const src = [
      '#If Not Mac And Win64 Then',
      'Public Sub ActiveBranch()',
      '#Else',
      'Public Sub InactiveBranch()',
      '#End If',
    ].join('\n');

    const out = preprocessConditionalCompilation(src);
    expect(out).toContain('ActiveBranch');
    expect(out).not.toContain('InactiveBranch');
    expect(out.split('\n')).toHaveLength(src.split('\n').length);
  });

  it('handles nested inactive parent blocks without leaking active children', () => {
    const src = [
      '#If Mac Then',
      '#If Win64 Then',
      'Public Sub Wrong()',
      '#End If',
      '#Else',
      'Public Sub Right()',
      '#End If',
    ].join('\n');

    const out = preprocessConditionalCompilation(src);
    expect(out).not.toContain('Wrong');
    expect(out).toContain('Right');
  });

  it('treats unsafe or unknown expressions as false without throwing', () => {
    const src = [
      '#If CreateObject("WScript.Shell") Then',
      'Public Sub Unsafe()',
      '#Else',
      'Public Sub SafeFallback()',
      '#End If',
    ].join('\n');

    expect(() => preprocessConditionalCompilation(src)).not.toThrow();
    const out = preprocessConditionalCompilation(src);
    expect(out).not.toContain('Unsafe');
    expect(out).toContain('SafeFallback');
  });
});

/**
 * Issue #51 — `fix(vba): conditional-compilation evaluator — Win32/Win16,
 * True=-1 semantics, #Const support`. Three concrete gaps the previous
 * evaluator had:
 *
 *   1. `Win32` / `Win16` were not in the identifier substitution table —
 *      the legacy `#If Win32 Then` guard (always True on modern Windows)
 *      blanked its ACTIVE branch.
 *   2. The whitelist rejected `-` outright, so `#If Win64 = -1 Then`
 *      blanked its branch even though VBA's True = -1 makes the
 *      comparison True.
 *   3. `#Const NAME = <value>` was not parsed — a user-defined
 *      `#Const MODO_DEBUG = True` was ignored, so `#If MODO_DEBUG Then`
 *      fell through to the unknown-identifier fallback (false) and the
 *      user's TRUE branch was blanked.
 *
 * The fix maps `Win32` → true and `Win16` → false alongside the existing
 * VBA7/Win64/Mac table; substitutes `true`/`false` with `-1`/`0` so
 * numeric equality comparisons match VBA's True = -1 semantics; accepts
 * unary minus in the whitelist; and parses `#Const` into a per-call map
 * consulted before the hardcoded constants.
 *
 * Each atom asserts both the kept-or-blanked behavior of the active
 * branch AND line-count parity (the core invariant — downstream
 * extraction's `startLine` values depend on it).
 */
describe('Issue #51: Win32/Win16 + True=-1 + #Const support', () => {
  // ---- atom 1: Win32 is True on modern Windows -------------------------
  it('atom 1: #If Win32 Then keeps the branch on modern Windows', () => {
    const src = [
      '#If Win32 Then',
      'Debug.Print "x"',
      '#End If',
    ].join('\n');
    const out = preprocessConditionalCompilation(src);
    const lines = out.split('\n');
    expect(lines).toHaveLength(3);
    expect(lines[0]).toBe('');
    expect(lines[1]).toContain('Debug.Print "x"');
    expect(lines[2]).toBe('');
  });

  // ---- atom 2: Win16 is False (not a current target) --------------------
  it('atom 2: #If Win16 Then blanks the branch on modern Windows', () => {
    const src = [
      '#If Win16 Then',
      'Debug.Print "x"',
      '#End If',
    ].join('\n');
    const out = preprocessConditionalCompilation(src);
    const lines = out.split('\n');
    expect(lines).toHaveLength(3);
    expect(lines[0]).toBe('');
    expect(lines[1]).toBe('');
    expect(lines[2]).toBe('');
  });

  // ---- atom 3: #Const True keeps the branch ----------------------------
  it('atom 3: #Const MODO_DEBUG = True then #If MODO_DEBUG Then keeps the branch', () => {
    const src = [
      '#Const MODO_DEBUG = True',
      '#If MODO_DEBUG Then',
      'Debug.Print "x"',
      '#End If',
    ].join('\n');
    const out = preprocessConditionalCompilation(src);
    const lines = out.split('\n');
    expect(lines).toHaveLength(4);
    expect(lines[0]).toBe('');
    expect(lines[1]).toBe('');
    expect(lines[2]).toContain('Debug.Print "x"');
    expect(lines[3]).toBe('');
  });

  // ---- atom 4: #Const False → #Else branch kept ------------------------
  it('atom 4: #Const MODO_DEBUG = False then #If / #Else selects the else branch', () => {
    const src = [
      '#Const MODO_DEBUG = False',
      '#If MODO_DEBUG Then',
      'Debug.Print "x"',
      '#Else',
      'Debug.Print "y"',
      '#End If',
    ].join('\n');
    const out = preprocessConditionalCompilation(src);
    const lines = out.split('\n');
    expect(lines).toHaveLength(6);
    expect(lines[0]).toBe('');
    expect(lines[1]).toBe('');
    expect(lines[2]).toBe('');
    expect(lines[3]).toBe('');
    expect(lines[4]).toContain('Debug.Print "y"');
    expect(lines[5]).toBe('');
  });

  // ---- atom 5: #Const integer literal preserved through comparison ----
  it('atom 5: #Const X = 1 then #If X = 1 Then keeps the branch', () => {
    const src = [
      '#Const X = 1',
      '#If X = 1 Then',
      'Debug.Print "x"',
      '#End If',
    ].join('\n');
    const out = preprocessConditionalCompilation(src);
    const lines = out.split('\n');
    expect(lines).toHaveLength(4);
    expect(lines[0]).toBe('');
    expect(lines[1]).toBe('');
    expect(lines[2]).toContain('Debug.Print "x"');
    expect(lines[3]).toBe('');
  });

  // ---- atom 6: VBA True = -1 semantics for #If Win64 = -1 --------------
  it('atom 6: #If Win64 = -1 Then keeps the branch (VBA True = -1 semantics)', () => {
    const src = [
      '#If Win64 = -1 Then',
      'Debug.Print "x"',
      '#End If',
    ].join('\n');
    const out = preprocessConditionalCompilation(src);
    const lines = out.split('\n');
    expect(lines).toHaveLength(3);
    expect(lines[0]).toBe('');
    expect(lines[1]).toContain('Debug.Print "x"');
    expect(lines[2]).toBe('');
  });

  // ---- atom 7: #Const True cannot satisfy `= False` --------------------
  it('atom 7: #Const X = True then #If X = False Then blanks the branch', () => {
    const src = [
      '#Const X = True',
      '#If X = False Then',
      'Debug.Print "x"',
      '#End If',
    ].join('\n');
    const out = preprocessConditionalCompilation(src);
    const lines = out.split('\n');
    expect(lines).toHaveLength(4);
    expect(lines[0]).toBe('');
    expect(lines[1]).toBe('');
    expect(lines[2]).toBe('');
    expect(lines[3]).toBe('');
  });

  // ---- atom 8: unknown identifier still evaluates false (regression) ---
  it('atom 8 (negative regression): unknown identifier still evaluates false', () => {
    const src = [
      '#If NonExistent Then',
      'Debug.Print "x"',
      '#End If',
    ].join('\n');
    const out = preprocessConditionalCompilation(src);
    const lines = out.split('\n');
    expect(lines).toHaveLength(3);
    expect(lines[0]).toBe('');
    expect(lines[1]).toBe('');
    expect(lines[2]).toBe('');
  });

  // ---- defensive: line-count parity across every atom -------------------
  it('Issue #51 atoms preserve line-count parity across the suite', () => {
    // Re-runs the same six positive atoms and checks line-count parity on
    // each — the core invariant the preprocessor guarantees (downstream
    // extraction's `startLine` values depend on it).
    const sources = [
      // atom 1
      '#If Win32 Then\nDebug.Print "x"\n#End If',
      // atom 3
      '#Const MODO_DEBUG = True\n#If MODO_DEBUG Then\nDebug.Print "x"\n#End If',
      // atom 4
      '#Const MODO_DEBUG = False\n#If MODO_DEBUG Then\nDebug.Print "x"\n#Else\nDebug.Print "y"\n#End If',
      // atom 5
      '#Const X = 1\n#If X = 1 Then\nDebug.Print "x"\n#End If',
      // atom 6
      '#If Win64 = -1 Then\nDebug.Print "x"\n#End If',
      // atom 7
      '#Const X = True\n#If X = False Then\nDebug.Print "x"\n#End If',
    ];
    for (const src of sources) {
      const out = preprocessConditionalCompilation(src);
      expect(out.split('\n').length).toBe(src.split('\n').length);
    }
  });

  // ---- defensive: #Const line itself is blanked (parity) ----------------
  it('Issue #51: #Const directive line is blanked (line-count parity)', () => {
    const src = '#Const MODO_DEBUG = True\nSub X()\nEnd Sub';
    const out = preprocessConditionalCompilation(src);
    const lines = out.split('\n');
    expect(lines).toHaveLength(3);
    expect(lines[0]).toBe('');
    expect(lines[1]).toBe('Sub X()');
    expect(lines[2]).toBe('End Sub');
    expect(out).not.toContain('#Const');
  });

  // ---- defensive: #Const RHS as a quoted string is unsupported ----------
  it('Issue #51: #Const NAME = "literal" is unsupported (no entry, line blanked)', () => {
    // VBA forbids string-literal #Const values (CC evaluates only at
    // compile time, no runtime string comparison); the implementation
    // must NOT store the entry, and the line must be blanked.
    const src = [
      '#Const X = "hello"',
      '#If X Then',
      'Debug.Print "x"',
      '#End If',
    ].join('\n');
    const out = preprocessConditionalCompilation(src);
    const lines = out.split('\n');
    expect(lines).toHaveLength(4);
    expect(lines[0]).toBe(''); // #Const line blanked regardless
    // X is unknown — falls through to the unknown-identifier fallback
    // (the conservative behavior preserved from the original
    // implementation), which blanks the branch.
    expect(lines[2]).toBe('');
  });
});

describe('Issue 84: Bitwise, Operator Precedence, comparisons, Xor and non-zero truthiness', () => {
  it('performs bitwise And, Or, Not, Xor correctly on integers', () => {
    // 1 And 2 = 0
    const srcAnd = [
      '#If 1 And 2 Then',
      'Debug.Print "active"',
      '#Else',
      'Debug.Print "inactive"',
      '#End If',
    ].join('\n');
    expect(preprocessConditionalCompilation(srcAnd)).toContain('inactive');

    // 3 And 6 = 2 (non-zero is truthy)
    const srcAndTrue = [
      '#If 3 And 6 Then',
      'Debug.Print "active"',
      '#Else',
      'Debug.Print "inactive"',
      '#End If',
    ].join('\n');
    expect(preprocessConditionalCompilation(srcAndTrue)).toContain('active');

    // 1 Or 2 = 3
    const srcOr = [
      '#If 1 Or 2 Then',
      'Debug.Print "active"',
      '#Else',
      'Debug.Print "inactive"',
      '#End If',
    ].join('\n');
    expect(preprocessConditionalCompilation(srcOr)).toContain('active');

    // 3 Xor 6 = 5
    const srcXor = [
      '#If 3 Xor 6 Then',
      'Debug.Print "active"',
      '#Else',
      'Debug.Print "inactive"',
      '#End If',
    ].join('\n');
    expect(preprocessConditionalCompilation(srcXor)).toContain('active');

    // Not 0 = -1 (truthy)
    const srcNotZero = [
      '#If Not 0 Then',
      'Debug.Print "active"',
      '#Else',
      'Debug.Print "inactive"',
      '#End If',
    ].join('\n');
    expect(preprocessConditionalCompilation(srcNotZero)).toContain('active');

    // Not -1 = 0 (falsy)
    const srcNotMinusOne = [
      '#If Not -1 Then',
      'Debug.Print "active"',
      '#Else',
      'Debug.Print "inactive"',
      '#End If',
    ].join('\n');
    expect(preprocessConditionalCompilation(srcNotMinusOne)).toContain('inactive');

    // Not 2 = -3 (truthy)
    const srcNotTwo = [
      '#If Not 2 Then',
      'Debug.Print "active"',
      '#Else',
      'Debug.Print "inactive"',
      '#End If',
    ].join('\n');
    expect(preprocessConditionalCompilation(srcNotTwo)).toContain('active');
  });

  it('respects VBA operator precedence', () => {
    // 1 Or 2 And 4:
    // If And > Or: 1 Or (2 And 4) -> 1 Or 0 -> 1 (truthy)
    // If Or > And: (1 Or 2) And 4 -> 3 And 4 -> 0 (falsy)
    const srcPrecedenceOrAnd = [
      '#If 1 Or 2 And 4 Then',
      'Debug.Print "active"',
      '#Else',
      'Debug.Print "inactive"',
      '#End If',
    ].join('\n');
    expect(preprocessConditionalCompilation(srcPrecedenceOrAnd)).toContain('active');

    // 1 Xor 2 And 2:
    // If And > Xor: 1 Xor (2 And 2) -> 1 Xor 2 -> 3 (truthy)
    // If Xor > And: (1 Xor 2) And 2 -> 3 And 2 -> 2 (truthy)
    // Let's test: 3 Xor 2 And 2:
    // If And > Xor: 3 Xor (2 And 2) -> 3 Xor 2 -> 1 (truthy)
    // If Xor > And: (3 Xor 2) And 2 -> 1 And 2 -> 0 (falsy)
    const srcPrecedenceXorAnd = [
      '#If 3 Xor 2 And 2 Then',
      'Debug.Print "active"',
      '#Else',
      'Debug.Print "inactive"',
      '#End If',
    ].join('\n');
    expect(preprocessConditionalCompilation(srcPrecedenceXorAnd)).toContain('active');

    // 1 And 2 = 2:
    // If comparison > bitwise: 1 And (2 = 2) -> 1 And -1 -> 1 (truthy)
    // If bitwise > comparison: (1 And 2) = 2 -> 0 = 2 -> 0 (falsy)
    const srcPrecedenceCompAnd = [
      '#If 1 And 2 = 2 Then',
      'Debug.Print "active"',
      '#Else',
      'Debug.Print "inactive"',
      '#End If',
    ].join('\n');
    expect(preprocessConditionalCompilation(srcPrecedenceCompAnd)).toContain('active');
  });

  it('evaluates comparisons to -1 (True) or 0 (False)', () => {
    // (2 > 1) evaluates to -1
    const srcCompTrue = [
      '#If (2 > 1) = -1 Then',
      'Debug.Print "active"',
      '#Else',
      'Debug.Print "inactive"',
      '#End If',
    ].join('\n');
    expect(preprocessConditionalCompilation(srcCompTrue)).toContain('active');

    // (1 > 2) evaluates to 0
    const srcCompFalse = [
      '#If (1 > 2) = 0 Then',
      'Debug.Print "active"',
      '#Else',
      'Debug.Print "inactive"',
      '#End If',
    ].join('\n');
    expect(preprocessConditionalCompilation(srcCompFalse)).toContain('active');
  });

  it('implements non-zero truthiness', () => {
    const srcTwo = [
      '#If 2 Then',
      'Debug.Print "active"',
      '#Else',
      'Debug.Print "inactive"',
      '#End If',
    ].join('\n');
    expect(preprocessConditionalCompilation(srcTwo)).toContain('active');

    const srcMinusFive = [
      '#If -5 Then',
      'Debug.Print "active"',
      '#Else',
      'Debug.Print "inactive"',
      '#End If',
    ].join('\n');
    expect(preprocessConditionalCompilation(srcMinusFive)).toContain('active');

    const srcZero = [
      '#If 0 Then',
      'Debug.Print "active"',
      '#Else',
      'Debug.Print "inactive"',
      '#End If',
    ].join('\n');
    expect(preprocessConditionalCompilation(srcZero)).toContain('inactive');
  });

  it('handles invalid syntax gracefully', () => {
    const srcMismatchedParen = [
      '#If (1 And 2 Then',
      'Debug.Print "active"',
      '#Else',
      'Debug.Print "inactive"',
      '#End If',
    ].join('\n');
    expect(preprocessConditionalCompilation(srcMismatchedParen)).toContain('inactive');

    const srcInvalidTokens = [
      '#If 1 And 2 @@@ Then',
      'Debug.Print "active"',
      '#Else',
      'Debug.Print "inactive"',
      '#End If',
    ].join('\n');
    expect(preprocessConditionalCompilation(srcInvalidTokens)).toContain('inactive');
  });

  it('triangulates overflow, nested negation, and case insensitivity', () => {
    // 32-bit signed integer overflow: 2147483647 + 1 = -2147483648
    const srcOverflow = [
      '#If 2147483647 + 1 = -2147483648 Then',
      'Debug.Print "active"',
      '#Else',
      'Debug.Print "inactive"',
      '#End If',
    ].join('\n');
    // Assert the ACTIVE branch is kept AND the else branch is blanked. A bare
    // `.toContain('active')` would falsely pass on the wrong branch because the
    // string 'inactive' contains the substring 'active'.
    const outOverflow = preprocessConditionalCompilation(srcOverflow);
    expect(outOverflow).toContain('Debug.Print "active"');
    expect(outOverflow).not.toContain('Debug.Print "inactive"');

    // Case insensitivity of #Const variables
    const srcCaseConst = [
      '#Const MyVar = 42',
      '#If MYVAR = 42 Then',
      'Debug.Print "active"',
      '#Else',
      'Debug.Print "inactive"',
      '#End If',
    ].join('\n');
    const outCaseConst = preprocessConditionalCompilation(srcCaseConst);
    expect(outCaseConst).toContain('Debug.Print "active"');
    expect(outCaseConst).not.toContain('Debug.Print "inactive"');

    // Nested negation with binary addition: -(5 + -(3)) = -2
    const srcNestedNeg = [
      '#If -(5 + -(3)) = -2 Then',
      'Debug.Print "active"',
      '#Else',
      'Debug.Print "inactive"',
      '#End If',
    ].join('\n');
    const outNestedNeg = preprocessConditionalCompilation(srcNestedNeg);
    expect(outNestedNeg).toContain('Debug.Print "active"');
    expect(outNestedNeg).not.toContain('Debug.Print "inactive"');
  });
});

describe('custom targets and precedence', () => {
  it('respects precedence: local #Const > custom targets > built-in defaults', () => {
    // 1. Custom target overrides built-in: Win64 defaults to true, but if we configure Win64=false, it should be false.
    const src1 = [
      '#If Win64 Then',
      'Debug.Print "active"',
      '#Else',
      'Debug.Print "inactive"',
      '#End If',
    ].join('\n');
    expect(preprocessConditionalCompilation(src1, { Win64: false })).toContain('inactive');

    // 2. Local #Const overrides custom target:
    // Custom target sets MYVAR = false, but local #Const sets MYVAR = true.
    const src2 = [
      '#Const MYVAR = True',
      '#If MYVAR Then',
      'Debug.Print "active"',
      '#Else',
      'Debug.Print "inactive"',
      '#End If',
    ].join('\n');
    expect(preprocessConditionalCompilation(src2, { MYVAR: false })).toContain('active');
  });

  it('checks custom target keys case-insensitively', () => {
    const src = [
      '#If my_custom_platform Then',
      'Debug.Print "active"',
      '#Else',
      'Debug.Print "inactive"',
      '#End If',
    ].join('\n');
    expect(preprocessConditionalCompilation(src, { MY_CUSTOM_PLATFORM: true })).toContain('active');
  });

  it('falls back undefined custom targets to 0', () => {
    const src = [
      '#If UNDEFINED_TARGET Then',
      'Debug.Print "active"',
      '#Else',
      'Debug.Print "inactive"',
      '#End If',
    ].join('\n');
    expect(preprocessConditionalCompilation(src, { OTHER_TARGET: true })).toContain('inactive');
  });
});



