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
 * ## Issue #256 — DDL verbs and the Access `IN` clause
 *
 * The clause set also covers `CREATE TABLE`, `ALTER TABLE` and
 * `DROP TABLE`; all three target a table they mutate, so they carry
 * `access: 'write'`. The reserved-word reject list applies to them
 * unchanged.
 *
 * Separately, `scanSqlExternalBackends` captures the Access-specific
 * `IN "<path>"` clause, which points a query at ANOTHER database file.
 * That operand is a file, not a table, so it gets its own scanner and
 * its own node shape (`buildExternalBackendNode`) rather than being
 * folded into the table rows.
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
import { Node } from '../types';
import { generateNodeId } from './tree-sitter-helpers';

export interface SqlTableScanRow {
  /** The unwrapped table name (brackets + surrounding whitespace stripped). */
  table: string;
  /**
   * The SQL clause that introduced the reference. The four DML clauses
   * (`FROM`/`JOIN`/`INTO`/`UPDATE`) were the original set; Issue #256
   * added the three DDL verbs, whose targets are always writes.
   */
  clause: SqlTableClause;
  /** Whether this row reads or mutates the table — derived from the SQL verb. */
  access: 'read' | 'write';
}

/**
 * Every clause keyword the scanner captures a table name after. The DDL
 * three are two-word keywords normalized to a single interior space
 * (`create   table` → `CREATE TABLE`).
 */
export type SqlTableClause =
  | 'FROM'
  | 'JOIN'
  | 'INTO'
  | 'UPDATE'
  | 'CREATE TABLE'
  | 'ALTER TABLE'
  | 'DROP TABLE';

/**
 * The set of clause keywords, in the same order as the regex alternation.
 * Used to validate a capture's clause before it is narrowed to
 * {@link SqlTableClause}.
 */
const SQL_TABLE_CLAUSES: ReadonlySet<string> = new Set<SqlTableClause>([
  'FROM',
  'JOIN',
  'INTO',
  'UPDATE',
  'CREATE TABLE',
  'ALTER TABLE',
  'DROP TABLE',
]);

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
  /\b(FROM|JOIN|INTO|UPDATE|CREATE\s+TABLE|ALTER\s+TABLE|DROP\s+TABLE)\s+((?:(?:\[[^\]]+\]|\p{L}[\p{L}\p{N}_]*)\.)?(?:\[[^\]]+\]|\p{L}[\p{L}\p{N}_]*))/giu;

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
    // Normalize the clause keyword: uppercase, and collapse the interior
    // whitespace of the two-word DDL verbs (`create   table` →
    // `CREATE TABLE`) so the emitted `clause` is a stable enum value.
    const clauseRaw = (m[1] ?? '').toUpperCase().replace(/\s+/g, ' ');
    if (!SQL_TABLE_CLAUSES.has(clauseRaw)) continue;
    const clause = clauseRaw as SqlTableClause;

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
 *
 * Issue #256: the three DDL verbs are unconditional writes. `CREATE
 * TABLE` and `DROP TABLE` change whether the table exists at all, and
 * `ALTER TABLE` rewrites its shape — every one of them is a mutation of
 * the named table, so "who writes to table X" must report them.
 */
function classifyAccess(sqlString: string, clause: SqlTableClause): 'read' | 'write' {
  if (clause === 'INTO' || clause === 'UPDATE') return 'write';
  if (clause === 'CREATE TABLE' || clause === 'ALTER TABLE' || clause === 'DROP TABLE') {
    return 'write';
  }
  if (clause === 'FROM' && /^\s*DELETE\b/i.test(sqlString)) return 'write';
  return 'read';
}

// ---------------------------------------------------------------------------
// Access `IN "<path>"` — the external-backend clause (Issue #256)
// ---------------------------------------------------------------------------

/**
 * Issue #256: Jet/ACE lets a single query read or write a table that
 * lives in ANOTHER database file:
 *
 *   ```sql
 *   SELECT * FROM TbRemota IN "C:\datos\otra.accdb"
 *   ```
 *
 * The `IN` operand is a FILE, not a table, so it needs its own capture
 * and its own representation. Without it a cross-backend access is
 * indistinguishable from a local one, and the second database file is
 * invisible to the graph.
 *
 * Only a QUOTED operand matches. The value-list operator (`WHERE Id IN
 * (1,2)`, `WHERE N IN ("a","b")`) always opens with `(`, so requiring a
 * quote immediately after the keyword separates the two meanings without
 * a parser. `IN ""` — the empty path of the Access ODBC/dBASE
 * connect-string form (`IN "" [ODBC;DSN=x]`) — yields an empty operand
 * and is dropped: that shape names a DSN, not a file we can key on.
 *
 * The first alternative handles the VBA doubled-quote shape. A SQL
 * string built in VBA source (`"… IN ""C:\x.accdb"""`) reaches the
 * scanner with its inner quotes still doubled, because the wrapper
 * regex in `vba/sql-wrapper.ts` captures the first literal verbatim.
 * Matching `""…""` before `"…"` decodes that shape without a
 * whole-string rewrite that could disturb the table scan.
 */
const EXTERNAL_BACKEND_RE = /\bIN\s+(?:""([^"]+)""|"([^"]*)"|'([^']*)')/giu;

/**
 * Issue #256: the `synthesizedBy` tag stamped on every `references`
 * edge that points at an external backend file. Exported so consumers
 * and tests share one spelling.
 */
export const EXTERNAL_BACKEND_SYNTHESIZED_BY = 'vba-external-backend';

/**
 * Issue #256: normalize an external-database path so the SAME file named
 * from two different queries converges on ONE node.
 *
 * Access backends are Windows paths: the separator may be `\` or `/`,
 * the case is not significant, and a hand-written path often carries a
 * doubled separator or a trailing one. Normalizing to lowercase
 * forward-slash form makes `C:\Datos\Otra.accdb` and
 * `c:/datos/otra.accdb` the same key.
 *
 * A leading `//` is preserved so a UNC share (`\\SERVIDOR\Datos\x.accdb`)
 * stays distinguishable from an absolute local path.
 *
 * Returns `''` for blank input; callers drop empty paths.
 */
export function normalizeBackendPath(raw: string): string {
  const trimmed = raw.trim();
  if (!trimmed) return '';
  const slashed = trimmed.replace(/\\/g, '/');
  const uncPrefix = slashed.startsWith('//') ? '//' : '';
  const body = slashed
    .slice(uncPrefix.length)
    .replace(/\/{2,}/g, '/')
    .replace(/\/+$/, '');
  if (!body) return '';
  return (uncPrefix + body).toLowerCase();
}

/**
 * Issue #256: scan a SQL string for Access `IN "<path>"` external-backend
 * clauses. Returns the NORMALIZED paths in first-appearance order, with
 * duplicates collapsed (two spellings of the same file are one entry).
 *
 * Empty / unquoted / value-list `IN` operands yield nothing — the same
 * "silent beats wrong" doctrine the table scanner follows: an operand we
 * cannot key on produces no node rather than a guessed one.
 */
export function scanSqlExternalBackends(sql: string): string[] {
  if (!sql) return [];
  const out: string[] = [];
  const seen = new Set<string>();
  const re = new RegExp(EXTERNAL_BACKEND_RE.source, EXTERNAL_BACKEND_RE.flags);
  let m: RegExpExecArray | null;
  while ((m = re.exec(sql)) !== null) {
    const raw = m[1] ?? m[2] ?? m[3] ?? '';
    const normalized = normalizeBackendPath(raw);
    if (!normalized) continue;
    if (seen.has(normalized)) continue;
    seen.add(normalized);
    out.push(normalized);
  }
  return out;
}

/**
 * Issue #256: build the graph node for an external backend file.
 *
 * `file` is the honest kind — it IS a file, just not one this index
 * parsed — so no new `NodeKind` is needed, and `metadata.external`
 * keeps it distinguishable from indexed files in every query.
 *
 * Two properties matter and are both deliberate:
 *
 *  1. **Convergence.** The node id is keyed on the normalized path via a
 *     synthetic `synthetic:external-db/<path>` file path (the same trick
 *     `vba/tempvars.ts` uses for TempVars keys), so the same backend named
 *     from two different queries — or from a saved query AND from in-code
 *     SQL — collapses to ONE node instead of one per referencing file.
 *  2. **Byte-identity.** Everything on the node derives from the
 *     normalized path alone: no line, no column, no per-caller language.
 *     Every emitter therefore produces an identical node, so whichever
 *     one is written last cannot silently rewrite the others' fields.
 *     Per-site position lives on the `references` EDGE, which is where
 *     call-site information belongs.
 *
 * The `language` is pinned to `vba` because the node models an Access
 * database file, whichever extractor happened to discover it.
 */
export function buildExternalBackendNode(normalizedPath: string): Node {
  const filePath = `synthetic:external-db/${normalizedPath}`;
  return {
    id: generateNodeId(filePath, 'file', normalizedPath, 0),
    kind: 'file',
    name: normalizedPath.split('/').filter(Boolean).pop() ?? normalizedPath,
    qualifiedName: normalizedPath,
    filePath,
    language: 'vba',
    startLine: 1,
    endLine: 1,
    startColumn: 0,
    endColumn: normalizedPath.length,
    metadata: { external: true, backendPath: normalizedPath },
    updatedAt: Date.now(),
  };
}
