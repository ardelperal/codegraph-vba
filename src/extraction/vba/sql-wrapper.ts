/**
 * SQL-in-strings sweep (REQ-CODE-8). Scans SQL wrapper calls
 * (`DoCmd.RunSQL "…"`, `*db.OpenRecordset "…"`, `getdb().Execute "…"`) and
 * variable-form executions, tracks `sql = sql & "…"` accumulation, and emits
 * `references` edges (via `ctx.emitReference`) to the table names found in the
 * FROM/JOIN/INTO/UPDATE clauses.
 *
 * Issue #203: the `FROM/JOIN/INTO/UPDATE <table>` regex is the canonical
 * source from `src/extraction/sql-table-scan.ts` — every table-name
 * capture path in the project imports from there so a reserved word
 * (`WHERE`, `ORDER`, `SET`, …) can never be emitted as a table
 * reference, and a non-literal operand dropped by `&`-concatenation
 * (`"DELETE FROM " & tabla & " WHERE x"`) is replaced with a `?`
 * sentinel that the regex can never match. The shared module also
 * emits the read/write access direction so this file no longer
 * re-implements `classifySqlAccess`.
 */
import { escapeRegExpLiteral } from './text-utils';
import { VbaExtractorContext } from './context';
import { scanSqlTables } from '../sql-table-scan';

/** SQL wrapper helpers — order matters because `db.Execute` is a suffix of others. */
const SQL_WRAPPERS: ReadonlyArray<{ name: string; re: RegExp }> = [
  { name: 'DoCmd.RunSQL', re: /\bDoCmd\.RunSQL\s+"((?:[^"]|"")*)"/giu },
  { name: '*db.OpenRecordset', re: /\b(?:\p{L}[\p{L}\p{N}_]*)?db\b(?:\(\))?\.OpenRecordset\s+"((?:[^"]|"")*)"/giu },
  { name: '*db.Execute', re: /\b(?:\p{L}[\p{L}\p{N}_]*)?db\b(?:\(\))?\.Execute\s+"((?:[^"]|"")*)"/giu },
];

/** SQL assigned to a local variable, e.g. `m_SQL = "SELECT ..." & ...`. */
const SQL_VAR_ASSIGN_RE =
  /^\s*(\p{L}[\p{L}\p{N}_]*)\s*=\s*(.*)$/iu;

/** SQL wrapper called with a variable, e.g. `getdb().Execute m_SQL`. */
const SQL_VAR_EXEC_RE =
  /\b(?:\p{L}[\p{L}\p{N}_]*)?db\b(?:\(\))?\.(?:OpenRecordset|Execute)\s*\(?\s*(\p{L}[\p{L}\p{N}_]*)\s*\)?/giu;

/**
 * Issue #42: `DoCmd.RunSQL <identifier>` (variable form) — the dominant
 * Access idiom for executing a dynamically-built SQL string. Today only
 * the literal form `DoCmd.RunSQL "DELETE FROM X"` is tracked via the
 * `SQL_WRAPPERS` regex; the variable form silently dropped table impact for
 * every procedure that builds SQL in a string and runs it through
 * `DoCmd.RunSQL`.
 *
 * This regex is the DoCmd.RunSQL analogue of `SQL_VAR_EXEC_RE` above and
 * is iterated by `scanSqlInLine`. When a match is found, the captured
 * identifier is resolved against `sqlVariables` (populated by
 * `trackSqlVariableAssignment` with `&`-accumulate semantics — Issue #13)
 * and the resulting SQL string drives `scanSqlTables`.
 *
 * The optional `(?:\(\))?` + `\s*\(?` shape lets the regex match both
 * the parenthesised form `DoCmd.RunSQL(strSQL)` and the no-paren form
 * `DoCmd.RunSQL strSQL` that the existing SQL_WRAPPERS literal regex
 * does not cover. The captured identifier is the only thing we need —
 * we DO NOT try to parse what the variable points at; that's the
 * existing `sqlVariables` map's job.
 */
const SQL_VAR_DOCMD_RUNSQL_RE =
  /\bDoCmd\.RunSQL\s*\(?\s*(\p{L}[\p{L}\p{N}_]*)\s*\)?/giu;

/**
 * Regex matching the chained `& "..."` literals that may follow a
 * wrapper's first literal on the same physical line. Captures the
 * literal CONTENT (group 1); the surrounding `&` and quotes are
 * structural, not data. VBA allows whitespace around `&` and around
 * the inner quotes — handled with `\s*`. The `((?:[^"]|"")*)` body
 * mirrors the wrapper regex so a `""` inside a chained literal still
 * decodes to a single `"`.
 *
 * Cross-physical-line concat via `_` continuation is OUT OF SCOPE for
 * v1 (deferred; see commit message).
 */
/**
 * Given the text that follows a SQL wrapper's first literal on the same
 * physical line, return one fragment per `&`-concatenated operand in
 * source order. Each fragment is either the CONTENT of a `"..."` literal
 * or a `?` sentinel for any non-literal operand (a variable, function
 * call, expression). The `?` sentinel can never match the
 * `scanSqlTables` identifier class, so `"DELETE FROM " & tabla & " WHERE x"`
 * becomes `"DELETE FROM   ?   WHERE x"` — the reserved-word reject list
 * drops the `WHERE` capture and no `vba-sql-table` edge is emitted.
 *
 * Operates per-physical-line only — VBA `_` line continuation across
 * physical lines is handled separately by `collectConcatFragments` for
 * the variable-assignment path (see below).
 */
function collectSqlWrapperChain(rest: string): string[] {
  return collectConcatFragments(rest);
}

/**
 * Issue #13: `sql = sql & "..."` (self-referential concatenation) must
 * ACCUMULATE the new fragment onto whatever was already tracked for
 * `varName`, not overwrite it. Overwriting silently dropped earlier
 * fragments' tables — typically the initial `FROM <table>` in
 * `sql = "SELECT * FROM tblA"` followed by `sql = sql & " WHERE x=1"`.
 *
 * Detection: the RHS (`m[2]`, trimmed) starts with `<varName> &`,
 * case-insensitively — matching VBA's case-insensitive identifiers (`Sql`
 * and `sql` are the same variable). A genuine fresh assignment (RHS does
 * NOT start with the self-reference) still RESETS tracking — that
 * behavior is unchanged.
 */
export function trackSqlVariableAssignment(
  lines: string[],
  lineIndex: number,
  sqlVariables: Map<string, string>,
): void {
  const line = lines[lineIndex] ?? '';
  const m = SQL_VAR_ASSIGN_RE.exec(line);
  if (!m) return;
  const rawVarName = m[1] ?? '';
  const varName = rawVarName.toLowerCase();
  const rhs = (m[2] ?? '').trim();
  const newFragment = collectStringLiteralText(lines, lineIndex);
  if (!newFragment) return;

  const selfRefRe = new RegExp(`^${escapeRegExpLiteral(rawVarName)}\\s*&`, 'i');
  const existing = sqlVariables.get(varName);
  if (existing !== undefined && selfRefRe.test(rhs)) {
    sqlVariables.set(varName, `${existing} ${newFragment}`);
  } else {
    sqlVariables.set(varName, newFragment);
  }
}

/**
 * Walk the lines starting at `startIndex` and collect every
 * `&`-concatenated fragment — `"..."` literal CONTENT or a `?`
 * sentinel for non-literal operands. Multi-line concat (via `_`
 * continuation) is handled by carrying on while the current physical
 * line ends with `&`.
 *
 * Replaces the legacy implementation that used `extractStringLiterals`
 * to extract every `"..."` and silently dropped every non-literal
 * operand between them — the source of Issue #203's silently-wrong
 * `vba-sql-table` captures.
 */
function collectStringLiteralText(lines: string[], startIndex: number): string {
  const fragments: string[] = [];
  for (let i = startIndex; i < lines.length; i++) {
    const line = lines[i] ?? '';
    fragments.push(...collectConcatFragments(line));
    if (!line.trimEnd().endsWith('&')) break;
  }
  return fragments.join(' ');
}

/**
 * Walk a single source string and emit one fragment per
 * `&`-concatenated operand.
 *
 *   - A `"..."` literal → its content (with `""` doubled-quote escapes
 *     collapsed to a single `"`).
 *   - A non-literal operand (any non-whitespace token — variable,
 *     function call, expression, the `m_SQL = ` prefix on an
 *     assignment) that sits BETWEEN two operands in the chain → emit
 *     a `?` sentinel so the gap can never match an identifier in
 *     `scanSqlTables`.
 *   - Leading / trailing / inter-fragment whitespace is skipped (the
 *     `[...].join(' ')` in the callers handles the gap).
 *
 * Implementation: a small state machine that walks the string,
 * alternating between two modes:
 *
 *   1. IN-LITERAL — collect until matching `"`, honour `""` escapes.
 *   2. IN-GAP — scan until next `&` (concat operator) or next `"` (start
 *      of next literal). If the gap contained any non-whitespace
 *      characters AND we previously emitted a literal, emit a `?`
 *      sentinel — that's the "we don't know this table" signal the
 *      regex needs.
 *
 * Why mode 2 stops at `"`: a `"` inside a non-literal operand is the
 * start of the NEXT literal in the chain; treating it as part of the
 * gap would silently consume a real literal and re-introduce the
 * Issue #203 bug.
 */
function collectConcatFragments(src: string): string[] {
  const out: string[] = [];
  let i = 0;
  let lastEmittedLiteral = false;
  while (i < src.length) {
    const ch = src[i] ?? '';
    if (ch === '"') {
      // Mode 1 — literal.
      let text = '';
      i++;
      while (i < src.length) {
        const c = src[i] ?? '';
        if (c === '"' && src[i + 1] === '"') {
          text += '"';
          i += 2;
          continue;
        }
        if (c === '"' || c === '\n') {
          i++;
          break;
        }
        text += c;
        i++;
      }
      out.push(text);
      lastEmittedLiteral = true;
      continue;
    }
    if (/\s/.test(ch)) {
      i++;
      continue;
    }
    // Mode 2 — gap (non-literal operand). Walk until the next `"` or
    // `&`. If we saw non-whitespace content AND a literal came before
    // us in this chain, emit the `?` sentinel so the gap can never
    // bridge two SQL keywords.
    let sawContent = false;
    while (i < src.length) {
      const c = src[i] ?? '';
      if (c === '"' || c === '&') break;
      if (!/\s/.test(c)) sawContent = true;
      i++;
    }
    if (src[i] === '&') {
      i++;
    }
    if (sawContent && lastEmittedLiteral) {
      out.push('?');
    }
    lastEmittedLiteral = false;
  }
  return out;
}

export function scanSqlInLine(
  ctx: VbaExtractorContext,
  line: string,
  lineNum: number,
  dedupe: Set<string>,
  sqlVariables: Map<string, string>,
): void {
  for (const { re } of SQL_WRAPPERS) {
    // Each wrapper regex is stateful (has /g); reset before use.
    const localRe = new RegExp(re.source, re.flags);
    let m: RegExpExecArray | null;
    while ((m = localRe.exec(line)) !== null) {
      const firstLiteral = m[1] ?? '';
      // After the wrapper regex consumes up to and including the closing
      // `"` of the first literal, walk the rest of the line for any
      // `& "..."` chains and concatenate every literal's content. Joining
      // with a space (mirrors `collectStringLiteralText`) keeps adjacent
      // `FROM tblA` & `FROM tblB` separated so `scanSqlTables` finds both.
      const rest = line.slice(m.index + m[0].length);
      const chain = collectSqlWrapperChain(rest);
      const joined = [firstLiteral, ...chain].join(' ');
      emitSqlTableReferences(ctx, joined, lineNum, dedupe);
    }
  }

  const localRe = new RegExp(SQL_VAR_EXEC_RE.source, SQL_VAR_EXEC_RE.flags);
  let vm: RegExpExecArray | null;
  while ((vm = localRe.exec(line)) !== null) {
    const varName = (vm[1] ?? '').toLowerCase();
    const sqlString = sqlVariables.get(varName);
    if (!sqlString) continue;
    emitSqlTableReferences(ctx, sqlString, lineNum, dedupe);
  }

  // Issue #42: `DoCmd.RunSQL <identifier>` (variable form). Mirrors the
  // SQL_VAR_EXEC_RE path above but for the Access-style `DoCmd.RunSQL`
  // idiom — the dominant pattern in real-world VBA modules. Resolve the
  // captured identifier against `sqlVariables` (populated by
  // `trackSqlVariableAssignment` with `&`-accumulate semantics, Issue
  // #13) and feed the resolved SQL string into `scanSqlTables`.
  // Unresolved identifiers (no row in the map) are silently skipped —
  // same graceful-no-op contract as SQL_VAR_EXEC_RE.
  const docmdLocalRe = new RegExp(
    SQL_VAR_DOCMD_RUNSQL_RE.source,
    SQL_VAR_DOCMD_RUNSQL_RE.flags,
  );
  let dm: RegExpExecArray | null;
  while ((dm = docmdLocalRe.exec(line)) !== null) {
    const varName = (dm[1] ?? '').toLowerCase();
    const sqlString = sqlVariables.get(varName);
    if (!sqlString) continue;
    emitSqlTableReferences(ctx, sqlString, lineNum, dedupe);
  }
}

function emitSqlTableReferences(
  ctx: VbaExtractorContext,
  sqlString: string,
  lineNum: number,
  dedupe: Set<string>,
): void {
  // Issue #203: delegate to the shared scanner. It owns the
  // `FROM/JOIN/INTO/UPDATE <table>` regex, the reserved-word reject
  // list, the `?`-sentinel fallback (see `vba/sql-wrapper.ts:144-178`)
  // and the read/write access direction.
  for (const row of scanSqlTables(sqlString)) {
    const key = `${lineNum}:${row.table}`;
    if (dedupe.has(key)) continue;
    dedupe.add(key);
    ctx.emitReference(row.table, lineNum, 0, 'vba-sql-table', row.access);
  }
}
