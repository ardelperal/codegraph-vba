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
 *
 * Issue #244: which receivers count as a database handle is no longer baked
 * into the scanning regexes. `DEFAULT_SQL_WRAPPERS` lists the built-in ones
 * and `codegraph.json` → `vba.sqlWrappers` extends that list with a project's
 * own accessors. Config entries are identifier fragments or `receiver.method`
 * pairs — never raw regex, which in this per-line hot path would be a
 * catastrophic-backtracking foot-gun.
 */
import { escapeRegExpLiteral } from './text-utils';
import { VbaExtractorContext, ProcInfo } from './context';
import { scanSqlTables, scanSqlExternalBackends } from '../sql-table-scan';

/** SQL assigned to a local variable, e.g. `m_SQL = "SELECT ..." & ...`. */
const SQL_VAR_ASSIGN_RE =
  /^\s*(\p{L}[\p{L}\p{N}_]*)\s*=\s*(.*)$/iu;

/**
 * `metadata.synthesizedBy` for a table reached through a SQL wrapper call
 * (`DoCmd.RunSQL`, `db.Execute`, `db.OpenRecordset`, …).
 */
const SQL_TABLE_SYNTHESIZER = 'vba-sql-table';

/**
 * Issue #252 — `metadata.synthesizedBy` for a table reached through SQL
 * assigned to a binding property at RUNTIME. Deliberately distinct from the
 * `.form.txt` sweep's static `vba-row-source` so an audit can tell the two
 * provenances apart: one is what the designer stored, the other is what the
 * code actually binds.
 */
export const RUNTIME_BINDING_SYNTHESIZER = 'vba-row-source-dynamic';

/**
 * Issue #252 — one receiver-chain segment on the left of a binding
 * assignment: a plain identifier or a `[bracketed name]`, optionally
 * followed by a single call/index parenthesis (`Controls("cbo")`,
 * `Item(0)`). Parentheses are matched non-nested on purpose — a
 * bounded, backtracking-safe shape for a per-line hot path.
 */
const BINDING_SEGMENT =
  String.raw`(?:\[[^\]\r\n]+\]|[\p{L}_][\p{L}\p{N}_]*)(?:\s*\([^()\r\n]*\))?`;

/**
 * Issue #252 — `<receiver>.(RowSource|RecordSource|ControlSource|Filter|OrderBy) = <value>`.
 *
 * Group 1 is the property, group 2 the raw right-hand side.
 *
 * Shape notes, because each piece earns its place:
 *
 *   - Anchored at `^\s*`, and the receiver is built out of identifier
 *     segments joined by `.` or `!` — never a permissive `.*`. That is what
 *     rejects a COMPARISON (`If Me.Filter = "…" Then`): `If` is an
 *     identifier segment but no `.` follows it, and the pattern cannot
 *     restart mid-line. A `'` comment line is rejected by the same anchor.
 *   - The receiver chain is OPTIONAL so the bare `With` form
 *     (`With Me.cbo` … `.RowSource = "…"`) resolves; that spelling is
 *     pervasive in real Access code.
 *   - `(?![\p{L}\p{N}_])` after the property name keeps `Me.FilterOn = True`
 *     and `Me.OrderByOn = True` out — they are booleans, not SQL.
 *   - Every quantifier is bounded by a mandatory separator (`.`/`!`/`(`), so
 *     the pattern cannot backtrack catastrophically on a long line.
 */
const BINDING_ASSIGN_RE = new RegExp(
  String.raw`^\s*(?:${BINDING_SEGMENT}(?:\s*[.!]\s*${BINDING_SEGMENT})*)?\s*[.!]\s*(RowSource|RecordSource|ControlSource|Filter|OrderBy)(?![\p{L}\p{N}_])\s*=\s*(.+)$`,
  'iu',
);

/** Issue #252 — a right-hand side that is exactly one bare identifier. */
const BINDING_VAR_RE = /^\p{L}[\p{L}\p{N}_]*$/u;

/**
 * Issue #244 — the two execution-site regexes.
 *
 * Both capture the RECEIVER (group 1) and the METHOD (group 2) generically
 * and leave the "is this a database handle?" decision to
 * {@link receiverMatches}, driven by the wrapper list. The pre-#244 shape
 * hard-coded the receiver as "any identifier ENDING in `db`", which silently
 * dropped every per-backend accessor a real multi-database Access project
 * uses (`getdbHPS()`, `getdbExpedientes()`, `dbToUse`, `m_dbLanzadera`, and
 * every DAO `QueryDef` receiver).
 *
 * Shape notes, in the order they appear:
 *
 *   - `(?:\(\))?` — the accessor-call form `getdb().Execute`.
 *   - `(?=[\s("])` / `(?=[\s(])` — the method name must END at whitespace, an
 *     opening paren, or (literal form) the opening quote. Without it the
 *     greedy method group can give characters back and read `db.Executed` as
 *     method `Execute` + argument `d`.
 *   - `\s*\(?\s*` — the parenthesised call form. Its absence in the literal
 *     regex was the second defect the issue reports:
 *     `getdb().OpenRecordset("SELECT * FROM TbX")` fell through BOTH the
 *     literal path (which demanded whitespace before the quote) and the
 *     variable path (a literal is not an identifier), so the statement
 *     produced no table reference at all.
 *
 * Every quantifier is bounded and the alternatives inside the literal body
 * (`[^"]` vs `""`) are disjoint, so neither regex can backtrack
 * catastrophically — which is also why config entries are identifier
 * fragments and never raw regex.
 */
const SQL_WRAPPER_LITERAL_RE =
  /\b(\p{L}[\p{L}\p{N}_]*)(?:\(\))?\.(\p{L}[\p{L}\p{N}_]*)(?=[\s("])\s*\(?\s*"((?:[^"]|"")*)"/giu;

/**
 * Variable form — `getdb().Execute strSQL`, `DoCmd.RunSQL(strSQL)`,
 * `qdf.Execute sql`. Group 3 is the identifier, resolved against the
 * `sqlVariables` map that `trackSqlVariableAssignment` fills with
 * `&`-accumulate semantics (Issue #13). We deliberately do NOT try to parse
 * what the variable points at here; unresolved identifiers are skipped.
 */
const SQL_WRAPPER_VAR_RE =
  /\b(\p{L}[\p{L}\p{N}_]*)(?:\(\))?\.(\p{L}[\p{L}\p{N}_]*)(?=[\s(])\s*\(?\s*(\p{L}[\p{L}\p{N}_]*)\s*\)?/giu;

/**
 * The methods a BARE identifier entry (`"getdb"`) implies. An explicit
 * `receiver.method` entry (`"cnn.Execute"`) names its own method and ignores
 * this set.
 */
const DEFAULT_WRAPPER_METHODS: ReadonlySet<string> = new Set([
  'openrecordset',
  'execute',
]);

/**
 * Issue #244 — the wrapper list used when `codegraph.json` →
 * `vba.sqlWrappers` is absent. A STRICT SUPERSET of the pre-#244 behaviour:
 *
 *   - `db` covers every receiver the old `…db\b` regex reached (see
 *     {@link receiverMatches} for the suffix arm that preserves it) plus the
 *     `db`-prefixed locals it missed (`dbUse`, `dbToUse`, `m_dbLanzadera`).
 *   - `getdb` covers the per-backend accessor family (`getdbHPS()`,
 *     `getdbExpedientes()`, `getdbLanzadera()`) the issue measured at 26-36%
 *     of all execution sites.
 *   - `CurrentDb` / `DBEngine` are the Access built-ins.
 *   - `qd` / `qdf` are the conventional DAO `QueryDef` receivers.
 *   - `DoCmd.RunSQL` is the Access statement form (both the literal and the
 *     variable spelling, previously two dedicated regexes).
 *   - `Connection.Execute` / `Recordset.Open` are the ADO pair.
 *
 * Project-specific accessors that do not fit these names are added through
 * `vba.sqlWrappers`, which EXTENDS this list rather than replacing it.
 */
export const DEFAULT_SQL_WRAPPERS: readonly string[] = [
  'db',
  'getdb',
  'CurrentDb',
  'DBEngine',
  'qd',
  'qdf',
  'DoCmd.RunSQL',
  'Connection.Execute',
  'Recordset.Open',
];

/** One parsed wrapper entry. Both fields are lowercase — VBA is case-insensitive. */
interface SqlWrapperMatcher {
  /** Receiver name fragment, e.g. `getdb` from `"getdb"` or `cnn` from `"cnn.Execute"`. */
  receiver: string;
  /** Method name, or `null` for a bare entry (which implies {@link DEFAULT_WRAPPER_METHODS}). */
  method: string | null;
}

/**
 * The per-extractor compiled wrapper state. Built ONCE per `VbaExtractor`
 * (see `vba-extractor.ts`) and cached on `VbaExtractorContext` — the two
 * RegExps are stateful (`/g`) and `scanSqlInLine` runs on every line of every
 * file, so re-compiling them per line (which is what the pre-#244 code did)
 * is pure waste on the hottest path in VBA extraction.
 */
export interface CompiledSqlWrappers {
  /** Literal form — `<receiver>[()].<method> "…"`. Stateful; reset before use. */
  readonly literalRe: RegExp;
  /** Variable form — `<receiver>[()].<method> <identifier>`. Stateful; reset before use. */
  readonly varRe: RegExp;
  /** Parsed entries, defaults first, project entries appended. */
  readonly matchers: readonly SqlWrapperMatcher[];
}

/**
 * Parse the plain-string wrapper entries into matchers and pair them with a
 * fresh pair of scanning RegExps.
 *
 * `configured` entries are APPENDED to {@link DEFAULT_SQL_WRAPPERS}, never
 * substituted for them: a project that names its own accessor must not lose
 * `CurrentDb` in the trade. Entries are accepted in exactly two forms — a
 * bare identifier fragment (`"getdb"`) or a `receiver.method` pair
 * (`"cnn.Execute"`). Anything else is dropped here; `project-config.ts`
 * already warned about it at load time.
 */
export function compileSqlWrappers(
  configured?: readonly string[],
): CompiledSqlWrappers {
  const matchers: SqlWrapperMatcher[] = [];
  const seen = new Set<string>();
  for (const raw of [...DEFAULT_SQL_WRAPPERS, ...(configured ?? [])]) {
    const entry = (raw ?? '').trim().toLowerCase();
    if (!entry || seen.has(entry)) continue;
    const dot = entry.indexOf('.');
    if (dot < 0) {
      if (!/^\p{L}[\p{L}\p{N}_]*$/u.test(entry)) continue;
      seen.add(entry);
      matchers.push({ receiver: entry, method: null });
      continue;
    }
    const receiver = entry.slice(0, dot);
    const method = entry.slice(dot + 1);
    if (
      !/^\p{L}[\p{L}\p{N}_]*$/u.test(receiver) ||
      !/^\p{L}[\p{L}\p{N}_]*$/u.test(method)
    ) {
      continue;
    }
    seen.add(entry);
    matchers.push({ receiver, method });
  }
  return {
    literalRe: new RegExp(SQL_WRAPPER_LITERAL_RE.source, SQL_WRAPPER_LITERAL_RE.flags),
    varRe: new RegExp(SQL_WRAPPER_VAR_RE.source, SQL_WRAPPER_VAR_RE.flags),
    matchers,
  };
}

/**
 * The default wrappers, compiled once for the whole module. Used by any
 * `VbaExtractorContext` built without options (tests, out-of-repo callers) so
 * "no config" behaves exactly like "config absent" and still never compiles a
 * RegExp per line. Safe to share: `scanSqlInLine` resets `lastIndex` before
 * every scan and never re-enters itself.
 */
let defaultCompiled: CompiledSqlWrappers | null = null;
function defaultSqlWrappers(): CompiledSqlWrappers {
  if (defaultCompiled === null) defaultCompiled = compileSqlWrappers();
  return defaultCompiled;
}

/**
 * Does `receiver` name a database handle under `pattern`?
 *
 * A leading `m_` / `p_` scope prefix is stripped first — `m_dbLanzadera` is
 * the same handle as `dbLanzadera`, and VBA codebases use both spellings for
 * the same variable.
 *
 * Then TWO arms, both comparing case-insensitively:
 *
 *   - PREFIX — the arm this issue adds, and the one every per-backend
 *     accessor needs: `getdbHPS`, `getdbExpedientes`, `dbUse`, `dbToUse`.
 *     The pattern must END AT A SEGMENT BOUNDARY: what follows it is either
 *     nothing, or a character that is neither a lowercase letter nor `_`
 *     (a capital or a digit).
 *     Without that, `db` would also claim `dbg`, `dbase` and `db_test` —
 *     names that merely START with the same two letters. Every execution
 *     receiver in the three measured corpora (`getdbNC`, `getdbAGEDO`,
 *     `dbUse`, `p_db`, `qdf`, …) clears the boundary.
 *   - SUFFIX — `MiBaseDatosdb` for pattern `db`. This is what the pre-#244
 *     regex `\b(?:\p{L}[\p{L}\p{N}_]*)?db\b` matched, kept so the defaults
 *     stay a strict superset of the old behaviour instead of trading one
 *     blind spot for another.
 *
 * The method check in {@link matchesWrapper} is the second gate: `dbg.Print`
 * fails both the boundary rule AND the method rule.
 */
function receiverMatches(receiver: string, pattern: string): boolean {
  const lower = receiver.toLowerCase();
  const bare = lower.startsWith('m_') || lower.startsWith('p_') ? lower.slice(2) : lower;
  if (bare.endsWith(pattern)) return true;
  if (!bare.startsWith(pattern)) return false;
  if (bare.length === pattern.length) return true;
  // Segment boundary: read the ORIGINAL (un-lowercased) character so
  // `dbUse` splits at the capital and `dbase` does not split at all.
  const next = receiver.slice(receiver.length - bare.length).charAt(pattern.length);
  return next !== '_' && next === next.toUpperCase();
}

/** Is `<receiver>.<method>` an execution site under any configured wrapper? */
function matchesWrapper(
  wrappers: CompiledSqlWrappers,
  receiver: string,
  method: string,
): boolean {
  const methodLower = method.toLowerCase();
  for (const matcher of wrappers.matchers) {
    const methodOk =
      matcher.method === null
        ? DEFAULT_WRAPPER_METHODS.has(methodLower)
        : matcher.method === methodLower;
    if (!methodOk) continue;
    if (receiverMatches(receiver, matcher.receiver)) return true;
  }
  return false;
}

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

/**
 * Issue #253: `QueryDefs("nombreConsulta")` names a saved query directly,
 * with no wrapper method involved, so it gets its own pattern rather than a
 * row in the wrapper table. String-literal argument only — a variable holds a
 * name nothing static can know.
 */
const QUERY_DEFS_RE = /\bQueryDefs\s*\(\s*"([^"]*)"\s*\)/giu;

/**
 * Issue #253: does this literal read as a SQL statement rather than the name
 * of a saved query?
 *
 * The gate is deliberately WIDER than the four DML verbs the issue names. A
 * literal is rejected as SQL if it contains any of these keywords anywhere,
 * because the cost of the two mistakes is not symmetric: treating SQL as a
 * query name emits a reference to a name no `queries/*.sql` will ever carry
 * (a permanent `failed` reference and a confusing one), while treating an
 * oddly-named query as SQL emits nothing at all. Silent beats wrong.
 */
const SQL_KEYWORD_RE =
  /\b(SELECT|INSERT|UPDATE|DELETE|FROM|WHERE|JOIN|INTO|VALUES|SET|CREATE|ALTER|DROP|UNION|GROUP|ORDER|HAVING|TRANSFORM|PARAMETERS|EXEC|EXECUTE|DISTINCT|TABLE)\b/iu;

/**
 * Issue #253: an Access object name — letters, digits, underscores and
 * interior spaces, nothing else. No punctuation, no operators, no wildcards.
 *
 * The length cap is not cosmetic: it is the last line of defence against a
 * long keyword-free string (a message, a `Value List` payload) being read as
 * a query name. Access itself caps object names at 64 characters.
 */
const ACCESS_OBJECT_NAME_RE = /^[\p{L}_][\p{L}\p{N}_ ]{0,63}$/u;

/** True when a wrapper's literal argument names a saved query, not a statement. */
function looksLikeSavedQueryName(text: string): boolean {
  const trimmed = text.trim();
  if (!trimmed) return false;
  if (SQL_KEYWORD_RE.test(trimmed)) return false;
  return ACCESS_OBJECT_NAME_RE.test(trimmed);
}

/**
 * Issue #253: emit ONE `dao-query` unresolved reference for a saved-query
 * name, from the CALLING PROCEDURE (not the module).
 *
 * The procedure is the source on purpose: `SqlQueryExtractor` already links
 * `query -> table`, so a `procedure -> query` reference is the hop that makes
 * `procedure -> query -> table` traversable end to end. An edge hung off the
 * module would leave that flow broken at its first hop.
 *
 * No node is created. The resolver binds this to the REAL `query` node built
 * from `queries/<Name>.sql`, and declines when there is none — a name that
 * matches neither a query nor a table stays an actionable `failed` reference
 * rather than becoming a fabricated table placeholder.
 */
function emitSavedQueryReference(
  ctx: VbaExtractorContext,
  queryName: string,
  lineNum: number,
  column: number,
  caller: ProcInfo,
  dedupe: Set<string>,
): void {
  const trimmed = queryName.trim();
  if (!trimmed) return;
  // The `query-name:` prefix keeps this de-dup bucket disjoint from the table
  // and external-backend buckets that share the same set.
  const key = `${lineNum}:query-name:${trimmed.toLocaleLowerCase('en-US')}`;
  if (dedupe.has(key)) return;
  dedupe.add(key);
  ctx.unresolvedReferences.push({
    fromNodeId: ctx.findOrCreateFunctionNodeId(caller),
    referenceName: trimmed,
    referenceKind: 'dao-query',
    line: lineNum,
    column,
    filePath: ctx.filePath,
    language: 'vba',
    metadata: { synthesizedBy: 'vba-query-name' },
  });
}

export function scanSqlInLine(
  ctx: VbaExtractorContext,
  line: string,
  lineNum: number,
  dedupe: Set<string>,
  sqlVariables: Map<string, string>,
  caller?: ProcInfo,
): void {
  // Issue #244: one compiled wrapper set per extractor, cached on the
  // context. A context built without options (tests, out-of-repo callers)
  // falls back to the module-level defaults — never to a per-line compile.
  const wrappers = ctx.sqlWrappers ?? defaultSqlWrappers();

  // Literal form — `getdb().Execute "DELETE FROM T"`,
  // `CurrentDb.OpenRecordset("SELECT * FROM T")`, `DoCmd.RunSQL "…"`.
  const literalRe = wrappers.literalRe;
  literalRe.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = literalRe.exec(line)) !== null) {
    if (!matchesWrapper(wrappers, m[1] ?? '', m[2] ?? '')) continue;
    const firstLiteral = m[3] ?? '';
    // After the wrapper regex consumes up to and including the closing
    // `"` of the first literal, walk the rest of the line for any
    // `& "..."` chains and concatenate every literal's content. Joining
    // with a space (mirrors `collectStringLiteralText`) keeps adjacent
    // `FROM tblA` & `FROM tblB` separated so `scanSqlTables` finds both.
    const rest = line.slice(m.index + m[0].length);
    const chain = collectSqlWrapperChain(rest);
    const joined = [firstLiteral, ...chain].join(' ');
    // Issue #253: a wrapper handed a literal that is NOT a statement but IS a
    // bare Access object name is naming a saved query. Verb detection wins:
    // anything that reads as SQL takes the table path below, unchanged.
    if (caller && chain.length === 0 && looksLikeSavedQueryName(joined)) {
      emitSavedQueryReference(ctx, joined, lineNum, m.index, caller, dedupe);
      continue;
    }
    emitSqlTableReferences(ctx, joined, lineNum, dedupe);
  }

  // Variable form — `getdb().Execute strSQL`, `qdf.Execute sql`,
  // `DoCmd.RunSQL(strSQL)` (Issue #42). The captured identifier is resolved
  // against `sqlVariables` (populated by `trackSqlVariableAssignment` with
  // `&`-accumulate semantics, Issue #13); an identifier with no row in the
  // map is silently skipped — we never guess what a variable holds.
  const varRe = wrappers.varRe;
  varRe.lastIndex = 0;
  let vm: RegExpExecArray | null;
  while ((vm = varRe.exec(line)) !== null) {
    if (!matchesWrapper(wrappers, vm[1] ?? '', vm[2] ?? '')) continue;
    const varName = (vm[3] ?? '').toLowerCase();
    const sqlString = sqlVariables.get(varName);
    if (!sqlString) continue;
    // A variable that accumulated a bare name rather than a statement names a
    // saved query, exactly as the literal form does.
    if (caller && looksLikeSavedQueryName(sqlString)) {
      emitSavedQueryReference(ctx, sqlString, lineNum, vm.index, caller, dedupe);
      continue;
    }
    emitSqlTableReferences(ctx, sqlString, lineNum, dedupe);
  }

  // Issue #253: `db.QueryDefs("nombreConsulta")` names a saved query with no
  // wrapper method in sight, so it is scanned independently of the wrapper
  // table. Literal argument only.
  if (caller) {
    QUERY_DEFS_RE.lastIndex = 0;
    let qm: RegExpExecArray | null;
    while ((qm = QUERY_DEFS_RE.exec(line)) !== null) {
      const name = qm[1] ?? '';
      if (!looksLikeSavedQueryName(name)) continue;
      emitSavedQueryReference(ctx, name, lineNum, qm.index, caller, dedupe);
    }
  }

  // Issue #252 — SQL bound to a form/control property at runtime.
  scanRuntimeBindingAssignment(ctx, line, lineNum, dedupe, sqlVariables);
}

/**
 * Issue #252 — SQL assigned to a binding property at runtime.
 *
 * This is the half of the data-binding problem the `.form.txt` sweep cannot
 * see. The measured corpus binds in code, not in the designer: 161
 * `.RowSource =` assignments across `.cls`/`.bas`, against ZERO
 * `RecordSource` and ZERO `ControlSource` occurrences in `00_EXPEDIENTES`'s
 * `.form.txt` files. Everything those statements touch was invisible.
 *
 * The right-hand side goes through the SAME pipeline the wrapper path uses —
 * {@link collectConcatFragments} then `scanSqlTables` — deliberately, and not
 * a parallel one. That is what makes the `?` sentinel for non-literal
 * operands and the shared reserved-word reject list apply here for free:
 * `.RowSource = "SELECT " & campo & " FROM " & tabla` becomes
 * `SELECT  ?  FROM   ?` and yields no table at all, rather than the confident
 * wrong `FROM` / `WHERE` captures Issue #203 documents.
 *
 * A value that is NOT SQL — a bare saved-query name, a `Value List` payload,
 * a plain field name on `ControlSource` — produces nothing here on its own,
 * because `scanSqlTables` finds no clause keyword in it. Connecting a bare
 * query name to its saved query is T11's job, not this one; guessing would
 * be exactly the "confident wrong edge" the project forbids.
 */
function scanRuntimeBindingAssignment(
  ctx: VbaExtractorContext,
  line: string,
  lineNum: number,
  dedupe: Set<string>,
  sqlVariables: Map<string, string>,
): void {
  const m = BINDING_ASSIGN_RE.exec(line);
  if (!m) return;
  const rhs = (m[2] ?? '').trim();
  if (!rhs) return;

  // Literal form — one or more `"..."` operands, possibly `&`-chained.
  const fragments = collectConcatFragments(rhs);
  if (fragments.length > 0) {
    emitSqlTableReferences(
      ctx,
      fragments.join(' '),
      lineNum,
      dedupe,
      RUNTIME_BINDING_SYNTHESIZER,
    );
    return;
  }

  // Variable form — `.RowSource = strSQL`. Resolved against the SAME
  // `sqlVariables` map the wrapper path uses, so the `&`-accumulate
  // semantics of Issue #13 and the procedure scoping of Issue #204 apply
  // unchanged. An identifier with no row in the map is skipped: we never
  // guess what a variable holds.
  if (!BINDING_VAR_RE.test(rhs)) return;
  const sqlString = sqlVariables.get(rhs.toLowerCase());
  if (!sqlString) return;
  emitSqlTableReferences(
    ctx,
    sqlString,
    lineNum,
    dedupe,
    RUNTIME_BINDING_SYNTHESIZER,
  );
}

function emitSqlTableReferences(
  ctx: VbaExtractorContext,
  sqlString: string,
  lineNum: number,
  dedupe: Set<string>,
  synthesizedBy: string = SQL_TABLE_SYNTHESIZER,
): void {
  // Issue #203: delegate to the shared scanner. It owns the
  // `FROM/JOIN/INTO/UPDATE <table>` regex, the reserved-word reject
  // list, the `?`-sentinel fallback (see `vba/sql-wrapper.ts:144-178`)
  // and the read/write access direction.
  //
  // Issue #252: the de-dup key is namespaced by `synthesizedBy` so the
  // runtime-binding path and the wrapper path keep disjoint buckets. A
  // single line can legitimately be both (`Set rs = db.OpenRecordset(sql)`
  // has no binding, but a `With` body can interleave the two shapes), and
  // one path must never silence the other's edge.
  for (const row of scanSqlTables(sqlString)) {
    const key = `${synthesizedBy}:${lineNum}:${row.table}`;
    if (dedupe.has(key)) continue;
    dedupe.add(key);
    ctx.emitReference(row.table, lineNum, 0, synthesizedBy, row.access);
  }

  // Issue #256: the Access `IN "<path>"` clause points the statement at
  // another database file. It is a cross-backend edge, so it gets its
  // own `file`-kind target rather than being mistaken for a table.
  // The `external-db:` prefix keeps this de-dup bucket disjoint from the
  // table bucket above (a table name can never contain a colon).
  for (const backendPath of scanSqlExternalBackends(sqlString)) {
    const key = `${lineNum}:external-db:${backendPath}`;
    if (dedupe.has(key)) continue;
    dedupe.add(key);
    ctx.emitExternalBackendReference(backendPath, lineNum, 0);
  }
}
