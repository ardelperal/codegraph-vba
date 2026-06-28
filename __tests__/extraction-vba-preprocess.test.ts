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
} from '../src/extraction/vba-preprocess';

describe('joinLineContinuations', () => {
  it('joins two lines when the first ends with " _"', () => {
    const src = 'Sub Foo()\n  DoCmd.RunSQL _\n    "SELECT * FROM tbl"';
    const out = joinLineContinuations(src);
    expect(out).toContain('DoCmd.RunSQL     "SELECT * FROM tbl"');
    // The original line break must be gone.
    expect(out).not.toContain('DoCmd.RunSQL _\n');
  });

  it('joins a Debug.Print continuation across lines', () => {
    const src = 'Debug.Print _\n  "hello"';
    const out = joinLineContinuations(src);
    expect(out).toBe('Debug.Print   "hello"');
  });

  it('returns input unchanged when no line ends with " _"', () => {
    const src = 'Sub Foo()\n  Debug.Print "x"\nEnd Sub';
    expect(joinLineContinuations(src)).toBe(src);
  });

  it('joins multiple chained continuations in one statement', () => {
    const src = 'x = a _\n  + b _\n  + c';
    const out = joinLineContinuations(src);
    expect(out).toContain('x = a   + b   + c');
  });

  it('joins continuations across many statements', () => {
    const src = 'A _\n+ 1\nB _\n+ 2';
    const out = joinLineContinuations(src);
    expect(out).toContain('A + 1');
    expect(out).toContain('B + 2');
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

  it('strips the trailing CRLF before the underscore + a single space', () => {
    // VBA convention: " _\n" → join to next line with one space inserted.
    const src = 'X _\nY';
    expect(joinLineContinuations(src)).toBe('X Y');
  });

  it('does not join when the underscore is mid-line (not at the end)', () => {
    const src = 'A = 1 + _invalid_underscore\nB = 2';
    expect(joinLineContinuations(src)).toBe(src);
  });

  it('joins even when continuation is followed by indented code', () => {
    const src = 'Function Foo() As Long\n  Foo = 1 _\n    + 2\nEnd Function';
    const out = joinLineContinuations(src);
    expect(out).toContain('Foo = 1     + 2');
  });

  it('handles a continuation that ends a file (no following line)', () => {
    const src = 'x = 1 _';
    // Defensive: the helper should not throw on a dangling continuation.
    // Result is `x = 1` (the trailing ` _` is replaced with a single space,
    // which downstream regex sweeps tolerate; downstream stripVbaComments
    // and the extractor regex sweep don't anchor on trailing whitespace).
    expect(() => joinLineContinuations(src)).not.toThrow();
    expect(joinLineContinuations(src)).toBe('x = 1 ');
  });
});

describe('stripVbaComments', () => {
  it('strips a trailing `\' comment from a code line', () => {
    const src = 'Dim x As Long  \' comment';
    expect(stripVbaComments(src)).toBe('Dim x As Long');
  });

  it('strips a Rem-prefixed line entirely', () => {
    const src = 'Rem LegacyComment\nSub X(): End Sub';
    expect(stripVbaComments(src)).toBe('Sub X(): End Sub');
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

  it('strips a `\'` on its own line', () => {
    const src = "' just a comment\nSub X(): End Sub";
    expect(stripVbaComments(src)).toBe('\nSub X(): End Sub');
  });

  it('preserves an empty input', () => {
    expect(stripVbaComments('')).toBe('');
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