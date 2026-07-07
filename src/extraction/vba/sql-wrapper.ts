/**
 * SQL-in-strings sweep (REQ-CODE-8). Scans SQL wrapper calls
 * (`DoCmd.RunSQL "…"`, `*db.OpenRecordset "…"`, `getdb().Execute "…"`) and
 * variable-form executions, tracks `sql = sql & "…"` accumulation, and emits
 * `references` edges (via `ctx.emitReference`) to the table names found in the
 * FROM/JOIN/INTO/UPDATE clauses.
 */
import { extractStringLiterals } from '../vba-preprocess';
import { escapeRegExpLiteral } from './text-utils';
import { VbaExtractorContext } from './context';

/** SQL wrapper helpers — order matters because `db.Execute` is a suffix of others. */
const SQL_WRAPPERS: ReadonlyArray<{ name: string; re: RegExp }> = [
  { name: 'DoCmd.RunSQL', re: /\bDoCmd\.RunSQL\s+"((?:[^"]|"")*)"/g },
  { name: '*db.OpenRecordset', re: /\b(?:\p{L}[\p{L}\p{N}_]*)?db\b(?:\(\))?\.OpenRecordset\s+"((?:[^"]|"")*)"/giu },
  { name: '*db.Execute', re: /\b(?:\p{L}[\p{L}\p{N}_]*)?db\b(?:\(\))?\.Execute\s+"((?:[^"]|"")*)"/giu },
  // Fix 4 (Issue #4): inline-literal forms `getdb().Execute "..."` and
  // `getdb().OpenRecordset "..."` — the variable form is covered by
  // SQL_VAR_EXEC_RE but the direct-literal form was missing.
  { name: 'getdb().Execute', re: /\bgetdb\(\)\.Execute\s+"((?:[^"]|"")*)"/g },
  { name: 'getdb().OpenRecordset', re: /\bgetdb\(\)\.OpenRecordset\s+"((?:[^"]|"")*)"/g },
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
 * and the resulting SQL string drives `emitSqlTableReferences`.
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
 * SQL table-name regex scoped to the clauses that introduce a table
 * reference: `FROM <t>`, `JOIN <t>`, `INTO <t>`, `UPDATE <t>`. Adding
 * `JOIN` lets the scanner pick up tables from joined fragments that
 * arrive via `&`-concatenated wrapper literals (e.g.
 * `db.Execute "FROM A" & " JOIN B"`); without it the second literal's
 * table was silently dropped even though the wrapper regex now matches
 * the chain.
 *
 * The captured table name is an optional bracketed/unbracketed schema
 * prefix followed by a `.`, then a bracketed-or-bare identifier — so
 * `FROM dbo.tblCustomers` and `FROM [My Schema].[My Table]` come
 * through as one composite reference. Without the prefix the regex
 * still matches a single identifier byte-identical to the old shape.
 * Brackets in the captured composite are stripped by
 * `emitSqlTableReferences` (`replace(/[\[\]]/g, '')`), so the public
 * node name is the unwrapped form `dbo.tblCustomers` /
 * `My Schema.My Table` — matching how plain `[Order Details]` is also
 * unwrapped to `Order Details`. The identifier class
 * `\[[^\]]+\]|\p{L}[\p{L}\p{N}_]*` (same as the saved-queries
 * `TABLE_RE` in `sql-query-extractor.ts`) ensures bracketed names
 * with spaces — `[Order Details]`, `[My Schema]`, `[My Table]` —
 * are captured whole.
 */
const SQL_TABLE_RE =
  /\b(?:FROM|JOIN|INTO|UPDATE)\s+((?:(?:\[[^\]]+\]|\p{L}[\p{L}\p{N}_]*)\.)?(?:\[[^\]]+\]|\p{L}[\p{L}\p{N}_]*))/giu;

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
const SQL_WRAPPER_CHAIN_RE = /&\s*"((?:[^"]|"")*)"/g;

/**
 * Given the text that follows a SQL wrapper's first literal on the same
 * physical line, return the contents of every `& "..."` chained literal
 * in source order. Operates per-physical-line only — VBA `_` line
 * continuation across physical lines is handled separately by
 * `collectStringLiteralText` for the variable-assignment path.
 */
function collectSqlWrapperChain(rest: string): string[] {
  const out: string[] = [];
  const re = new RegExp(SQL_WRAPPER_CHAIN_RE.source, SQL_WRAPPER_CHAIN_RE.flags);
  let m: RegExpExecArray | null;
  while ((m = re.exec(rest)) !== null) {
    out.push(m[1] ?? '');
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
      // `FROM tblA` & `FROM tblB` separated so `SQL_TABLE_RE` finds both.
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
  // #13) and feed the resolved SQL string into `emitSqlTableReferences`.
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

/**
 * #13 fix: `sql = sql & "..."` (self-referential concatenation) must
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

function collectStringLiteralText(lines: string[], startIndex: number): string {
  const fragments: string[] = [];
  for (let i = startIndex; i < lines.length; i++) {
    const line = lines[i] ?? '';
    for (const lit of extractStringLiterals(line)) {
      fragments.push(lit.text);
    }
    if (!line.trimEnd().endsWith('&')) break;
  }
  return fragments.join(' ');
}

function emitSqlTableReferences(
  ctx: VbaExtractorContext,
  sqlString: string,
  lineNum: number,
  dedupe: Set<string>,
): void {
  // Scan the SQL string for FROM/INTO/UPDATE <table>.
  // Preserve the source regex's `/u` flag (Unicode property classes)
  // — hardcoding `'gi'` here would silently break non-ASCII identifiers.
  const tableRe = new RegExp(SQL_TABLE_RE.source, SQL_TABLE_RE.flags);
  let tm: RegExpExecArray | null;
  while ((tm = tableRe.exec(sqlString)) !== null) {
    const table = (tm[1] ?? '').replace(/[\[\]]/g, '');
    if (!table) continue;
    const key = `${lineNum}:${table}`;
    if (dedupe.has(key)) continue;
    dedupe.add(key);
    ctx.emitReference(table, lineNum, 0, 'vba-sql-table');
  }
}
