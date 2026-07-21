/**
 * # SQL table-name scanner — leaf module shared by every VBA + SQL
 * table-extraction path (Issue #203).
 *
 * Before this module existed, the same `FROM/JOIN/INTO/UPDATE` capture
 * regex lived in three places:
 *
 *   - `src/extraction/sql-query-extractor.ts:57`     (saved queries)
 *   - `src/extraction/vba/sql-wrapper.ts:80-81`      (in-code SQL sweep)
 *   - `src/extraction/vba-form-extractor.ts:617-618` (RecordSource/RowSource)
 *
 * The three copies had already started to diverge (`vba-form-extractor`
 * grew a `SQL_PREFIX_RE` the other two never adopted), and every
 * caller was vulnerable to the same silently-wrong captures the issue
 * lists:
 *
 *   ```vba
 *   getdb().Execute "DELETE FROM " & tabla & " WHERE activo = 1"
 *   ```
 *
 * Today, `collectSqlWrapperChain` silently drops non-literal operands
 * (variables, function calls) and joins surviving literal fragments
 * with a space. The dropped operand leaves a whitespace gap that
 * `\s+` happily crosses, so `DELETE FROM ` + ` ` + ` WHERE activo = 1`
 * becomes `DELETE FROM   WHERE activo = 1` — and `SQL_TABLE_RE`
 * captures `WHERE` as a table reference.
 *
 * ## What this module guarantees
 *
 * 1. The shared regex (with optional schema prefix + bracketed
 *    identifiers, identical bytes to the old shape) lives in ONE
 *    place.
 * 2. `SQL_RESERVED_TABLE_TOKENS` rejects every SQL reserved word that
 *    can legitimately appear immediately after a
 *    `FROM`/`JOIN`/`INTO`/`UPDATE` keyword in a real statement —
 *    `WHERE`, `ORDER`, `GROUP`, `HAVING`, `SET`, `VALUES`, `SELECT`,
 *    `INNER`, `LEFT`, `RIGHT`, `OUTER`, `FULL`, `CROSS`, `JOIN`,
 *    `ON`, `UNION`, `AS`, `DISTINCT`, `TOP`, `IN`, `EXISTS`. The list
 *    is exhaustive for the SQL grammar subset codegraph models.
 * 3. `scanSqlTables` reads a SINGLE joined SQL string. It does NOT
 *    know about VBA concatenation — the JOIN-WITH-SPACE shape
 *    `collectSqlWrapperChain` produces is its input contract.
 *    Operands dropped by `collectSqlWrapperChain` are now replaced
 *    with a `?` sentinel (see `vba/sql-wrapper.ts`); the reserved-
 *    word reject list catches any `?` keyword bridge the concat could
 *    still create.
 * 4. Each row carries the SQL `clause` (`FROM`/`JOIN`/`INTO`/`UPDATE`)
 *    so callers can classify `access: 'read' | 'write'` without
 *    re-running the SQL classifier.
 *
 * ## Defense in depth
 *
 * The reserved-word check uses the **unwrapped first identifier
 * component** so schema-qualified inputs like `FROM WHERE.ID` are also
 * rejected (the unwrapped form is `WHERE.ID`, first component is
 * `WHERE`). The check is case-insensitive (SQL keywords are
 * case-insensitive by spec).
 *
 * Returning `[]` for a reserved-word capture is the "silent beats
 * wrong" doctrine the project documents in `CLAUDE.md`: emitting a
 * confident wrong edge ("WHERE is a table that gets written to")
 * pollutes downstream queries far more than emitting no edge at all.
 */
export interface SqlTableScanRow {
  /** The unwrapped table name (brackets + surrounding whitespace stripped). */
  table: string;
  /** The SQL clause that introduced the reference (`FROM`/`JOIN`/`INTO`/`UPDATE`). */
  clause: 'FROM' | 'JOIN' | 'INTO' | 'UPDATE';
  /** Whether this row reads or mutates the table — derived from the SQL verb. */
  access: 'read' | 'write';
}

/**
 * Canonical reject list — every SQL reserved word that can appear
 * immediately after `FROM`/`JOIN`/`INTO`/`UPDATE` in a real statement.
 * The list is exported so consumers can introspect it; it is also the
 * authoritative source the scanner consults internally.
 *
 * The list intentionally omits `FROM`/`JOIN`/`INTO`/`UPDATE` itself
 * (those are the capturing keywords) and `ALL` / `ANY` (they are
 * legitimate table-ish tokens in some dialects but rarely appear as
 * the first token after a FROM-style keyword in a well-formed
 * statement).
 */
export const SQL_RESERVED_TABLE_TOKENS: ReadonlySet<string> = new Set([
  'WHERE',
  'ORDER',
  'GROUP',
  'HAVING',
  'INNER',
  'LEFT',
  'RIGHT',
  'OUTER',
  'FULL',
  'CROSS',
  'JOIN',
  'ON',
  'SET',
  'VALUES',
  'SELECT',
  'UNION',
  'AS',
  'DISTINCT',
  'TOP',
  'IN',
  'EXISTS',
]);

/**
 * Shared `FROM` / `JOIN` / `INTO` / `UPDATE <table>` regex. Same
 * shape as the three duplicates this module consolidates: group 1 is
 * the clause keyword (`FROM`/`JOIN`/`INTO`/`UPDATE`), group 2 is the
 * composite identifier — an optional bracketed/unbracketed schema
 * prefix followed by `.`, then a bracketed-or-bare identifier.
 *
 *   `FROM dbo.tblCustomers`         → m[2] = `dbo.tblCustomers`
 *   `FROM [My Schema].[My Table]`   → m[2] = `[My Schema].[My Table]`
 *   `FROM tblCustomers`             → m[2] = `tblCustomers`
 *   `FROM [Order Details]`          → m[2] = `[Order Details]`
 *
 * Brackets in `m[2]` are stripped by `scanSqlTables`, so the public
 * table name is the unwrapped form (`dbo.tblCustomers` /
 * `My Schema.My Table`) — matching how plain `[Order Details]` is
 * also unwrapped to `Order Details`. `\p{L}` covers accented
 * identifiers common in localized schemas.
 *
 * Keeping the single composite capture (rather than splitting
 * schema/table into separate groups) preserves byte-identity with the
 * three duplicates this module replaces, so every existing
 * regression test continues to apply without rewrites.
 */
const TABLE_RE =
  /\b(FROM|JOIN|INTO|UPDATE)\s+((?:(?:\[[^\]]+\]|\p{L}[\p{L}\p{N}_]*)\.)?(?:\[[^\]]+\]|\p{L}[\p{L}\p{N}_]*))/giu;

/**
 * Scan a SQL string for `FROM` / `JOIN` / `INTO` / `UPDATE` table
 * references. Returns one row per match; the caller is responsible
 * for cross-row deduplication if it wants one node per table.
 *
 * Captures that resolve to a SQL reserved word (Issue #203 — e.g.
 * `FROM   WHERE x=1` → `WHERE`) are DROPPED. This encodes
 * "we don't know this table" rather than guessing — matching the
 * "silent beats wrong" / "partial coverage is worse than none"
 * doctrine the project documents in `CLAUDE.md`.
 *
 * Schema-qualified inputs (`FROM WHERE.ID` → `WHERE.ID`) are also
 * dropped: the check inspects the FIRST unwrapped identifier
 * component (everything up to the first `.`), case-insensitive.
 *
 * Empty / whitespace-only / non-DML input returns `[]`.
 */
export function scanSqlTables(sql: string): SqlTableScanRow[] {
  if (!sql) return [];
  const out: SqlTableScanRow[] = [];
  const re = new RegExp(TABLE_RE.source, TABLE_RE.flags);
  let m: RegExpExecArray | null;
  while ((m = re.exec(sql)) !== null) {
    const clauseRaw = (m[1] ?? '').toUpperCase();
    if (
      clauseRaw !== 'FROM' &&
      clauseRaw !== 'JOIN' &&
      clauseRaw !== 'INTO' &&
      clauseRaw !== 'UPDATE'
    ) {
      continue;
    }
    const clause = clauseRaw as SqlTableScanRow['clause'];

    // m[2] is the composite (optional schema `.` table) with any
    // combination of brackets. Strip ALL brackets and surrounding
    // whitespace — the public node name is the unwrapped form.
    const table = (m[2] ?? '').replace(/[\[\]]/g, '').trim();
    if (!table) continue;

    // Reserved-word rejection — the FIRST identifier component is the
    // canonical SQL keyword to test (handles `WHERE.ID` too).
    const firstId = table.split('.')[0] ?? '';
    if (SQL_RESERVED_TABLE_TOKENS.has(firstId.toUpperCase())) continue;

    out.push({ table, clause, access: classifyAccess(sql, clause) });
  }
  return out;
}

/**
 * `access` classifier — ported verbatim from
 * `src/extraction/vba/sql-wrapper.ts:91-96`. Lifted into this leaf
 * module so all three call sites produce identical tagging.
 *
 * The mutating targets are writes: `INSERT INTO <t>`, `UPDATE <t>`,
 * and the `FROM <t>` of a `DELETE` (Access's `DELETE FROM x` makes
 * that FROM the delete target). Every other `FROM`/`JOIN` source
 * table — including the source of an `INSERT ... SELECT` — is a read.
 */
function classifyAccess(sqlString: string, clause: SqlTableScanRow['clause']): 'read' | 'write' {
  if (clause === 'INTO' || clause === 'UPDATE') return 'write';
  if (clause === 'FROM' && /^\s*DELETE\b/i.test(sqlString)) return 'write';
  return 'read';
}