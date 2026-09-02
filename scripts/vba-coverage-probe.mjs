/**
 * VBA coverage probe — the committed measurement tool behind every
 * acceptance criterion in `docs/vba-node-discovery-plan.md`.
 *
 * Every task in the VBA node-discovery roadmap states its acceptance as a
 * number ("stub nodes drop by >= 3,800", "table references rise"). This
 * script is the only supported way to produce those numbers, so a PR can
 * prove it landed a win instead of asserting one.
 *
 * Usage (from the repo root):
 *   npm run probe:vba -- <root> [<root> ...]
 *   npx tsx scripts/vba-coverage-probe.mjs <root> [<root> ...]
 *   npx tsx scripts/vba-coverage-probe.mjs --json <root> [<root> ...]
 *
 * It imports `VbaExtractor`, `VbaFormExtractor` and `SqlQueryExtractor`
 * DIRECTLY. No index, no SQLite, no CLI, no `codegraph.json`. It walks the
 * given roots, dispatches each file by extension, aggregates, and prints a
 * human-readable report (or JSON with `--json`).
 *
 * ## The stub discriminator — read this before changing anything
 *
 * A synthesized call target and a declared procedure are BOTH
 * `kind: 'function'` and BOTH carry `visibility: 'public'`. The ONLY
 * discriminator is `metadata.stub === true`, set in
 * `src/extraction/vba/calls.ts`.
 *
 * Without it the reference corpus reads 9,972 procedures instead of 3,840
 * — a 2.6x inflation that silently corrupts every later measurement.
 *
 * ## Reported fields
 *
 * Adding fields later is fine; REMOVING one breaks comparability with
 * earlier runs. The contract is:
 *
 *   files                     total files dispatched, by extension
 *   nodesByKind               { [kind]: count }
 *   declaredProcedures        function nodes WITHOUT metadata.stub
 *   stubProcedures            function nodes WITH metadata.stub
 *   stubTargetsTop            top 30 stub node names by count
 *   edgesByKind               { [kind]: count }
 *   edgesBySynthesizer        { [metadata.synthesizedBy]: count }
 *   unresolvedByKind          { [referenceKind]: count }
 *   sqlTablesReferenced       distinct table names reached via VBA-embedded SQL
 *                             (`vba-sql-table` + `vba-row-source-dynamic`)
 *   formsReferenced           distinct form names reached via opens-form / form refs
 *   proceduresWithNoOutgoing  declared procedures with zero edges and zero unresolved refs
 *   errorHandling             the error-handling census — see below
 *   errors                    extraction errors by code
 *
 * `files` additionally reports `total` (every regular file walked, including
 * ones no extractor claims) and `skipped` so the walked-vs-dispatched split
 * is visible rather than inferred.
 *
 * ## The `errorHandling` block — task E1 of `docs/vba-error-handling-plan.md`
 *
 * Every number in §2 of that plan came from a throwaway scan. This block is
 * the committed replacement, so E2–E5 measure against something a reviewer
 * can re-run.
 *
 * It is the ONE part of the probe that does NOT read extractor output: the
 * extractor emits no error-handling metadata yet (that is task E2). It reads
 * the raw module source instead, procedure by procedure, using the procedure
 * START lines the extractor already reports so `proceduresTotal` can never
 * drift from `declaredProcedures`.
 *
 * Two classifier decisions are load-bearing, and both exist because the
 * throwaway scan got them wrong:
 *
 * 1. **The error channel is a configurable list, not a regex literal.** The
 *    original scan hardcoded a narrow pattern, missed every `p_Error = …`
 *    write, and reported 889 channel writes plus a 1,899-strong "other"
 *    bucket that a 12-of-12 audit showed was channel writes all along.
 *    `DEFAULT_ERROR_CHANNEL_NAMES` holds the names; `--error-channel` and the
 *    `errorChannelNames` option override them; the matcher is built from the
 *    list at run time and covers the bare (`m_Error = …`), `Me.`-qualified
 *    (`Me.Error = …`) and object-qualified (`obj.Error = …`) write forms.
 * 2. **A body that both writes the channel and displays is `mixed`.** The
 *    original scan tested one signal after the other and let evaluation order
 *    silently pick a winner. Here all three signals (channel / display /
 *    reraise) are collected independently, and `mixed` is what more than one
 *    of them means. `handlerSignals` additionally reports the raw,
 *    non-exclusive per-signal totals, so the exclusive classification and the
 *    "how many handlers display at all" question are both answerable without
 *    re-deriving one from the other.
 *
 * Reported fields:
 *
 *   proceduresTotal                declared procedures in `.bas`/`.cls`-style files
 *   proceduresByProtection         { handler, resumeNext, none }
 *   unprotectedOverFiveLoc         `none` + more than 5 statement lines
 *   unprotectedTouchingIo          the above, narrowed to bodies touching I/O
 *   handlersByBehavior             { channel, display, reraise, mixed, unknown }
 *   handlerSignals                 non-exclusive per-signal handler counts
 *   handlerLabelNames              { [label]: `On Error GoTo <label>` sites }
 *   resumeNextOpenScopes           `Resume Next` scopes never closed by `GoTo 0`
 *   danglingGotoTargets            [{ file, procedure, label }]
 *   proceduresWithMultipleHandlers procedures with 2+ `On Error GoTo <label>`
 *   labels                         { defined, handlerTargets, controlFlow }
 *   statements                     raw construct counts (the §2.2 table)
 *   config                         the lists this run classified with
 */
import * as fs from 'fs';
import * as path from 'path';
import { pathToFileURL } from 'url';

/** `.form.txt` / `.report.txt` — Dysflow SaveAsText UI exports. */
const FORM_FILE_RE = /\.(form|report)\.txt$/i;
/** `.bas` / `.cls` / `.frm` / `.dsr` — VBA code modules. */
const CODE_FILE_RE = /\.(bas|cls|frm|dsr)$/i;
/** `queries/<Name>.sql` — Dysflow-exported saved Access queries. */
const SQL_FILE_RE = /\.sql$/i;

/** How many stub target names `stubTargetsTop` reports. */
const STUB_TARGETS_TOP_N = 30;

/**
 * The `metadata.synthesizedBy` tag `src/extraction/vba/sql-wrapper.ts` puts
 * on every `references` edge that reaches a table named inside VBA-embedded
 * SQL executed through a wrapper call. See `SQL_TABLE_SYNTHESIZERS` below for
 * the full set `sqlTablesReferenced` counts.
 */
const SQL_TABLE_SYNTHESIZER = 'vba-sql-table';
/**
 * Issue #252 — the tag `src/extraction/vba/sql-wrapper.ts` puts on a table
 * reached through SQL assigned to a binding property at runtime
 * (`.RowSource = "SELECT …"`). It is a different PROVENANCE from
 * `vba-sql-table`, but the same REACH: a table the code touches. Both feed
 * `sqlTablesReferenced` so the metric keeps meaning "distinct tables this
 * corpus reaches from VBA-embedded SQL".
 */
const RUNTIME_BINDING_SYNTHESIZER = 'vba-row-source-dynamic';
/** Every synthesizer whose target is a table named inside VBA-embedded SQL. */
const SQL_TABLE_SYNTHESIZERS = new Set([
  SQL_TABLE_SYNTHESIZER,
  RUNTIME_BINDING_SYNTHESIZER,
]);
/**
 * The `metadata.synthesizedBy` tag `src/extraction/vba-form-extractor.ts`
 * puts on the `.form.txt` -> sibling `.cls` unresolved reference. Together
 * with `opens-form` edges this is what `formsReferenced` counts.
 */
const FORM_BINDING_SYNTHESIZER = 'vba-form-binding';

/**
 * Two-segment extension for `.form.txt` / `.report.txt`, plain
 * `path.extname()` otherwise. `path.extname('X.form.txt')` is `.txt`, which
 * would collapse forms and reports into one meaningless bucket.
 */
export function fileExtensionKey(filePath) {
  const base = path.basename(filePath).toLowerCase();
  const twoSegment = base.match(/(\.(?:form|report)\.txt)$/);
  if (twoSegment) return twoSegment[1];
  const ext = path.extname(base);
  return ext === '' ? '(none)' : ext;
}

/** Which of the three extractors claims this file, or `null`. */
export function dispatchFor(filePath) {
  if (FORM_FILE_RE.test(filePath)) return 'form';
  if (CODE_FILE_RE.test(filePath)) return 'code';
  if (SQL_FILE_RE.test(filePath)) return 'sql';
  return null;
}

/**
 * Walk `roots` and return every regular file, sorted, deduplicated. Sorting
 * makes the aggregate order-independent numbers reproducible run to run and
 * keeps `stubTargetsTop` tie-breaking stable.
 */
export function collectFiles(roots) {
  const seen = new Set();
  for (const root of roots) {
    const stat = fs.statSync(root);
    if (stat.isFile()) {
      seen.add(path.resolve(root));
      continue;
    }
    const stack = [path.resolve(root)];
    while (stack.length > 0) {
      const dir = stack.pop();
      for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
        const p = path.join(dir, entry.name);
        if (entry.isDirectory()) stack.push(p);
        else if (entry.isFile()) seen.add(p);
      }
    }
  }
  return [...seen].sort();
}

/**
 * Resolve the three extractor classes. `src/` is the source of truth — a
 * stale `dist/` would silently change the measurement — so it wins whenever
 * it is present. `dist/` is the fallback for an installed package, where
 * `src/` does not ship.
 */
export async function loadExtractors(repoRoot) {
  const srcDir = path.join(repoRoot, 'src', 'extraction');
  const distDir = path.join(repoRoot, 'dist', 'extraction');
  const useSrc = fs.existsSync(path.join(srcDir, 'vba-extractor.ts'));
  const base = useSrc ? srcDir : distDir;
  const ext = useSrc ? '.ts' : '.js';
  const load = async (name) =>
    import(pathToFileURL(path.join(base, name + ext)).href);
  const [code, form, sql] = await Promise.all([
    load('vba-extractor'),
    load('vba-form-extractor'),
    load('sql-query-extractor'),
  ]);
  return {
    VbaExtractor: code.VbaExtractor,
    VbaFormExtractor: form.VbaFormExtractor,
    SqlQueryExtractor: sql.SqlQueryExtractor,
  };
}

function bump(map, key) {
  map.set(key, (map.get(key) ?? 0) + 1);
}

/** Map -> plain object, sorted by count desc then key asc. Deterministic. */
function sortedObject(map) {
  const out = {};
  for (const [k, v] of [...map.entries()].sort(
    (a, b) => b[1] - a[1] || (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0),
  )) {
    out[k] = v;
  }
  return out;
}

/* ───────────────────────── error handling (task E1) ───────────────────── */

/**
 * The error-propagation channel this corpus actually uses: a module-level or
 * object field the failing procedure writes and the caller reads. Names only
 * — never substrings — so `ErrorCount` cannot match `Error`.
 *
 * These are the DEFAULTS, overridable per run (`errorChannelNames` /
 * `--error-channel`). Keeping them a list rather than a baked-in alternation
 * is the entire point of E1: the previous approximate 74% figure came from a
 * hardcoded pattern that silently missed `p_Error`.
 *
 * The same four names are what `docs/vba-error-handling-plan.md` task E4
 * proposes as the `vba.errorChannel` config default, so the probe and the
 * extractor will classify the same set.
 */
export const DEFAULT_ERROR_CHANNEL_NAMES = Object.freeze([
  'm_Error',
  'p_Error',
  'g_Error',
  'Error',
]);

/**
 * Calls that make an error visible to a human. `MsgBox` is the form-code
 * shape; `Debug.Print` is the developer-only shape. §2.3 of the plan counts
 * both as "display", so the default list does too.
 */
export const DEFAULT_DISPLAY_CALLS = Object.freeze(['MsgBox', 'Debug.Print']);

/**
 * The I/O surface that turns an unprotected procedure from "noise" into "a
 * finding": database access, an Access macro action, the filesystem, or a
 * conversion that can raise a type-mismatch. Deliberately a heuristic list —
 * it narrows §3.1's 813 unprotected procedures down to the short actionable
 * one a human can review, and it is configurable for exactly that reason.
 *
 * A trailing `.` means "this receiver, any member" (`DoCmd.` matches
 * `DoCmd.OpenForm` but not a local variable called `DoCmd`).
 */
export const DEFAULT_IO_MARKERS = Object.freeze([
  // database
  'OpenRecordset', 'Execute', 'CurrentDb', 'DBEngine',
  // Access macro actions
  'DoCmd.',
  // late binding / automation
  'CreateObject', 'GetObject',
  // filesystem
  'FreeFile', 'Kill', 'FileCopy', 'MkDir', 'RmDir', 'SetAttr', 'FileSystemObject',
  // conversions (type mismatch is the classic unhandled raise here)
  'CLng', 'CInt', 'CDbl', 'CSng', 'CCur', 'CDate', 'CBool', 'CByte', 'CStr',
  'CVar', 'CDec', 'Val',
]);

/** Escape a literal so it can be interpolated into a `RegExp`. */
function escapeRe(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * `End Sub` / `End Function` / `End Property`. A local copy of
 * `PROCEDURE_END_RE` from `src/extraction/vba/constants.ts`: the probe is a
 * standalone ESM script that must run against `src/` OR a shipped `dist/`,
 * so it cannot statically import a `.ts` module. Keep the two in sync — the
 * `(?:^|:\s*)` prefix is what tolerates the colon-separated single-line
 * procedure form (#208).
 */
const PROBE_PROCEDURE_END_RE = /(?:^|:\s*)End\s+(?:Sub|Function|Property)\b/i;

/** `On Error GoTo <label|0|-1>` and `On Error Resume Next`, in source order. */
const ON_ERROR_RE =
  /\bOn\s+Error\s+(?:GoTo\s+(-?\d+|\p{L}[\p{L}\p{N}_]*)|(Resume\s+Next))/giu;

/** Every `GoTo` form, including the `On Error GoTo` ones. */
const GOTO_RE = /\bGoTo\s+(-?\d+|\p{L}[\p{L}\p{N}_]*)/giu;

/** `Resume`, `Resume Next`, `Resume <label>`. */
const RESUME_RE = /\bResume\b(?:\s+(Next|\p{L}[\p{L}\p{N}_]*))?/giu;

/**
 * A line label definition. VBA allows a statement to follow on the same line
 * (`errores: MsgBox "x"`), so the trailing-code form is accepted and the
 * remainder is treated as the first line of the handler body.
 *
 * `(?!=)` keeps the named-argument form (`Foo bar:=1`) out, and the keyword
 * guard below keeps `Case 1:` / `Do:` / `Else:` out.
 */
const LINE_LABEL_RE = /^\s*(\p{L}[\p{L}\p{N}_]*)\s*:(?!=)/u;

/** Statement keywords that can head a colon-terminated line but name no label. */
const NOT_A_LABEL = new Set([
  'case', 'do', 'else', 'end', 'exit', 'for', 'if', 'loop', 'next', 'select',
  'then', 'wend', 'while', 'with', 'rem', 'call', 'set', 'let', 'dim',
]);

const ERR_MEMBER_RES = {
  errRaise: /\bErr\s*\.\s*Raise\b/gi,
  errRaiseSentinel: /\bErr\s*\.\s*Raise\s+1000\b/gi,
  errNumber: /\bErr\s*\.\s*Number\b/gi,
  errDescription: /\bErr\s*\.\s*Description\b/gi,
  errClear: /\bErr\s*\.\s*Clear\b/gi,
  errSource: /\bErr\s*\.\s*Source\b/gi,
};

function countMatches(text, re) {
  re.lastIndex = 0;
  let n = 0;
  while (re.exec(text) !== null) n += 1;
  return n;
}

/**
 * Replace every string-literal character with a space and drop a trailing
 * comment. Column positions survive 1-for-1, and nothing inside a literal can
 * be parsed as VBA — which is the whole reason `s = "On Error GoTo errores"`
 * classifies as nothing at all (the discipline #209 established for the
 * extractor's own sweeps).
 */
export function maskVbaLine(line) {
  let masked = '';
  let i = 0;
  while (i < line.length) {
    const ch = line[i];
    if (ch !== '"') {
      masked += ch;
      i += 1;
      continue;
    }
    masked += ' '; // opening quote
    i += 1;
    while (i < line.length) {
      if (line[i] === '"' && line[i + 1] === '"') {
        masked += '  '; // doubled-quote escape
        i += 2;
      } else if (line[i] === '"') {
        masked += ' '; // closing quote
        i += 1;
        break;
      } else {
        masked += ' ';
        i += 1;
      }
    }
  }
  const apostrophe = masked.indexOf("'");
  if (apostrophe >= 0) masked = masked.slice(0, apostrophe);
  const rem = /(?:^|:)\s*Rem\b/i.exec(masked);
  if (rem) masked = masked.slice(0, rem.index);
  return masked;
}

/**
 * Fold `foo _\n bar` into one logical line, blanking the absorbed lines so
 * array indices stay aligned with physical line numbers. A handler split
 * across a continuation is otherwise read as two unrelated fragments — §7.3
 * of the plan calls this out explicitly.
 */
export function joinContinuations(lines) {
  const out = lines.slice();
  for (let i = out.length - 2; i >= 0; i--) {
    if (/(^|\s)_\s*$/.test(out[i])) {
      out[i] = out[i].replace(/(^|\s)_\s*$/, '$1') + out[i + 1];
      out[i + 1] = '';
    }
  }
  return out;
}

/**
 * Build the channel-WRITE matcher from a list of names.
 *
 * A write is the channel name in the assignment-TARGET position of a
 * statement — bare (`p_Error = …`), `Me.`-qualified (`Me.Error = …`) or
 * object-qualified (`obj.Error = …`). Anchoring at the start of the statement
 * is what separates a write from a read: `If m_Error <> "" Then` never
 * matches, and neither does `x = m_Error`.
 */
export function buildChannelWriteMatcher(names) {
  const patterns = names.map(
    (name) =>
      new RegExp(
        `^\\s*(?:Set\\s+)?(?:(?:Me|\\p{L}[\\p{L}\\p{N}_]*)\\s*\\.\\s*)?${escapeRe(name)}\\s*=(?!=)`,
        'iu',
      ),
  );
  return (statement) => patterns.some((re) => re.test(statement));
}

function buildCallMatcher(calls) {
  const patterns = calls.map(
    (call) =>
      new RegExp(
        `\\b${call.split('.').map(escapeRe).join('\\s*\\.\\s*')}\\b`,
        'i',
      ),
  );
  return (text) => patterns.some((re) => re.test(text));
}

function buildMarkerMatcher(markers) {
  const patterns = markers.map((marker) => {
    if (marker.endsWith('.')) {
      return new RegExp(`\\b${escapeRe(marker.slice(0, -1))}\\s*\\.`, 'i');
    }
    return new RegExp(`\\b${escapeRe(marker)}\\b`, 'i');
  });
  return (text) => patterns.some((re) => re.test(text));
}

/**
 * Split a masked line into statements and strip any leading `If … Then` /
 * `ElseIf … Then` / `Else` guard, so the house shape
 * `If Err.Number <> 1000 Then p_Error = "…"` presents `p_Error = "…"` to the
 * channel matcher as a statement in its own right.
 */
function statementsOf(maskedLine) {
  return maskedLine.split(':').map((part) => {
    const guard = /^\s*(?:(?:If|ElseIf)\b.*?\bThen\b|Else\b)/i.exec(part);
    return guard ? part.slice(guard[0].length) : part;
  });
}

function labelDefinedOn(maskedLine) {
  const m = LINE_LABEL_RE.exec(maskedLine);
  if (!m) return null;
  const name = m[1];
  if (NOT_A_LABEL.has(name.toLowerCase())) return null;
  return { name, rest: maskedLine.slice(m[0].length) };
}

function resolveErrorConfig(options = {}) {
  return {
    errorChannelNames: options.errorChannelNames ?? [...DEFAULT_ERROR_CHANNEL_NAMES],
    displayCalls: options.displayCalls ?? [...DEFAULT_DISPLAY_CALLS],
    ioMarkers: options.ioMarkers ?? [...DEFAULT_IO_MARKERS],
  };
}

/**
 * Classify every procedure in one module's source.
 *
 * @param {string} filePath  reported verbatim in `danglingGotoTargets`
 * @param {string} source    the raw module text
 * @param {Array<{ startLine: number, name: string }>} procedures
 *   Procedure declaration lines, taken from the extractor's own non-stub
 *   `function` nodes. Deriving the starts from the extractor rather than from
 *   a second `PROC_RE` is what keeps `proceduresTotal` and
 *   `declaredProcedures` from drifting apart.
 * @param {object} [config]  `{ errorChannelNames, displayCalls, ioMarkers }`
 * @returns {Array<object>} one record per procedure
 */
export function analyzeVbaErrorHandling(filePath, source, procedures, config = {}) {
  const cfg = resolveErrorConfig(config);
  const writesChannel = buildChannelWriteMatcher(cfg.errorChannelNames);
  const displays = buildCallMatcher(cfg.displayCalls);
  const touchesIo = buildMarkerMatcher(cfg.ioMarkers);

  const lines = joinContinuations(source.split('\n').map(maskVbaLine));

  // Two procedures can share one physical line only in the colon-separated
  // form; dedupe so a region is never scanned twice.
  const byStart = new Map();
  for (const proc of procedures) {
    if (!byStart.has(proc.startLine)) byStart.set(proc.startLine, proc);
  }
  const starts = [...byStart.keys()].sort((a, b) => a - b);

  const records = [];
  for (let s = 0; s < starts.length; s++) {
    const startLine = starts[s];
    const nextStart = starts[s + 1] ?? lines.length + 1;
    const searchCap = Math.min(lines.length, nextStart - 1);
    let endLine = searchCap;
    for (let ln = startLine; ln <= searchCap; ln++) {
      if (PROBE_PROCEDURE_END_RE.test(lines[ln - 1] ?? '')) {
        endLine = ln;
        break;
      }
    }

    const statements = {
      onErrorGoToLabel: 0,
      onErrorResumeNext: 0,
      onErrorGoToZero: 0,
      onErrorGoToMinusOne: 0,
      gotoStatements: 0,
      resumeNext: 0,
      resumeLabel: 0,
      resumeBare: 0,
      errRaise: 0,
      errRaiseSentinel: 0,
      errNumber: 0,
      errDescription: 0,
      errClear: 0,
      errSource: 0,
    };

    const handlerTargets = [];      // `On Error GoTo <label>` sites, in order
    const gotoTargets = new Set();  // every non-numeric `GoTo` target
    const definedLabels = [];       // { name, line, rest }
    let openResumeNextScopes = 0;
    let statementLines = 0;
    let bodyText = '';

    for (let ln = startLine; ln <= endLine; ln++) {
      const line = lines[ln - 1] ?? '';
      bodyText += line + '\n';

      ON_ERROR_RE.lastIndex = 0;
      let m;
      while ((m = ON_ERROR_RE.exec(line)) !== null) {
        if (m[2]) {
          statements.onErrorResumeNext += 1;
          openResumeNextScopes += 1;
          continue;
        }
        const target = m[1];
        if (target === '0') {
          statements.onErrorGoToZero += 1;
          openResumeNextScopes = 0;
        } else if (target === '-1') {
          // Valid VBA (clears the current error) and absent from this corpus.
          // Treated as a reset, exactly like `GoTo 0`, per the plan's E2 rules.
          statements.onErrorGoToMinusOne += 1;
          openResumeNextScopes = 0;
        } else {
          statements.onErrorGoToLabel += 1;
          handlerTargets.push(target);
        }
      }

      GOTO_RE.lastIndex = 0;
      while ((m = GOTO_RE.exec(line)) !== null) {
        statements.gotoStatements += 1;
        const target = m[1];
        if (!/^-?\d+$/.test(target)) gotoTargets.add(target);
      }

      RESUME_RE.lastIndex = 0;
      while ((m = RESUME_RE.exec(line)) !== null) {
        const tail = m[1];
        if (!tail) statements.resumeBare += 1;
        else if (/^next$/i.test(tail)) statements.resumeNext += 1;
        else statements.resumeLabel += 1;
      }

      for (const [key, re] of Object.entries(ERR_MEMBER_RES)) {
        statements[key] += countMatches(line, re);
      }

      const label = labelDefinedOn(line);
      if (label) definedLabels.push({ name: label.name, line: ln, rest: label.rest });

      // "Statement lines" for §3.1's size filter: body lines only, excluding
      // the declaration, the `End`, blanks, comments (already stripped by the
      // mask) and bare label definitions.
      if (ln > startLine && ln < endLine) {
        if (line.trim() !== '' && !(label && label.rest.trim() === '')) {
          statementLines += 1;
        }
      }
    }

    const targetSet = new Set(handlerTargets.map((t) => t.toLowerCase()));
    const definedSet = new Set(definedLabels.map((l) => l.name.toLowerCase()));

    const protection =
      handlerTargets.length > 0
        ? 'handler'
        : statements.onErrorResumeNext > 0
          ? 'resumeNext'
          : 'none';

    // A label nobody targets is control flow, not a handler (guardrail 1 in
    // §6 of the plan — 136 such labels in the corpus).
    const handlerLabelDefs = definedLabels.filter((l) =>
      targetSet.has(l.name.toLowerCase()),
    );

    let behavior = null;
    const signals = { channel: false, display: false, reraise: false };
    if (protection === 'handler') {
      const first = handlerLabelDefs[0] ?? null;
      if (first) {
        // The handler body runs from the label to the procedure's end; on the
        // label's own line only the text AFTER the colon belongs to it.
        const handlerLines = [first.rest];
        for (let ln = first.line + 1; ln <= endLine; ln++) {
          handlerLines.push(lines[ln - 1] ?? '');
        }
        const handlerText = handlerLines.join('\n');
        signals.channel = handlerLines.some((l) =>
          statementsOf(l).some((st) => writesChannel(st)),
        );
        signals.display = displays(handlerText);
        signals.reraise = /\bErr\s*\.\s*Raise\b/i.test(handlerText);
      }
      // `mixed` is what MORE THAN ONE signal means. The signals are collected
      // independently first precisely so evaluation order cannot decide.
      const hit = Object.entries(signals)
        .filter(([, on]) => on)
        .map(([name]) => name);
      behavior = hit.length === 0 ? 'unknown' : hit.length === 1 ? hit[0] : 'mixed';
    }

    const dangling = [...gotoTargets]
      .filter((t) => !definedSet.has(t.toLowerCase()))
      .sort();

    records.push({
      file: filePath,
      procedure: byStart.get(startLine).name,
      startLine,
      endLine,
      protection,
      behavior,
      signals,
      handlerTargets,
      handlerCount: handlerTargets.length,
      definedLabels: definedLabels.map((l) => l.name),
      handlerLabelDefinitions: handlerLabelDefs.map((l) => l.name),
      danglingTargets: dangling,
      openResumeNextScopes,
      statementLines,
      touchesIo: touchesIo(bodyText),
      statements,
    });
  }
  return records;
}

/** Fold per-procedure records into the reported `errorHandling` block. */
export function summarizeErrorHandling(records, config = {}) {
  const cfg = resolveErrorConfig(config);
  const byProtection = { handler: 0, resumeNext: 0, none: 0 };
  const byBehavior = { channel: 0, display: 0, reraise: 0, mixed: 0, unknown: 0 };
  const handlerSignals = { channel: 0, display: 0, reraise: 0 };
  const handlerLabelNames = new Map();
  const danglingGotoTargets = [];
  const statements = {
    onErrorGoToLabel: 0,
    onErrorResumeNext: 0,
    onErrorGoToZero: 0,
    onErrorGoToMinusOne: 0,
    gotoStatements: 0,
    resumeNext: 0,
    resumeLabel: 0,
    resumeBare: 0,
    errRaise: 0,
    errRaiseSentinel: 0,
    errNumber: 0,
    errDescription: 0,
    errClear: 0,
    errSource: 0,
  };

  let unprotectedOverFiveLoc = 0;
  let unprotectedTouchingIo = 0;
  let resumeNextOpenScopes = 0;
  let proceduresWithMultipleHandlers = 0;
  let labelsDefined = 0;
  let labelsHandlerTargets = 0;

  for (const rec of records) {
    byProtection[rec.protection] += 1;
    if (rec.behavior) byBehavior[rec.behavior] += 1;
    if (rec.protection === 'handler') {
      for (const key of ['channel', 'display', 'reraise']) {
        if (rec.signals[key]) handlerSignals[key] += 1;
      }
    }
    for (const label of rec.handlerTargets) {
      handlerLabelNames.set(label, (handlerLabelNames.get(label) ?? 0) + 1);
    }
    for (const label of rec.danglingTargets) {
      danglingGotoTargets.push({
        file: rec.file,
        procedure: rec.procedure,
        label,
      });
    }
    // §3.1's narrowing is cumulative: 813 unprotected -> 310 over five lines
    // -> 185 of those touching I/O. `unprotectedTouchingIo` is a SUBSET of
    // `unprotectedOverFiveLoc`, not an independent count.
    if (rec.protection === 'none' && rec.statementLines > 5) {
      unprotectedOverFiveLoc += 1;
      if (rec.touchesIo) unprotectedTouchingIo += 1;
    }
    resumeNextOpenScopes += rec.openResumeNextScopes;
    if (rec.handlerCount > 1) proceduresWithMultipleHandlers += 1;
    labelsDefined += rec.definedLabels.length;
    labelsHandlerTargets += rec.handlerLabelDefinitions.length;
    for (const key of Object.keys(statements)) statements[key] += rec.statements[key];
  }

  return {
    proceduresTotal: records.length,
    proceduresByProtection: byProtection,
    unprotectedOverFiveLoc,
    unprotectedTouchingIo,
    handlersByBehavior: byBehavior,
    handlerSignals,
    handlerLabelNames: sortedObject(handlerLabelNames),
    resumeNextOpenScopes,
    danglingGotoTargets,
    proceduresWithMultipleHandlers,
    labels: {
      defined: labelsDefined,
      handlerTargets: labelsHandlerTargets,
      controlFlow: labelsDefined - labelsHandlerTargets,
    },
    statements,
    config: cfg,
  };
}

/**
 * Run the probe over `roots`.
 *
 * @param {string[]} roots
 * @param {{ extractors?: object, repoRoot?: string, errorChannelNames?: string[],
 *           displayCalls?: string[], ioMarkers?: string[] }} [options]
 *   `extractors` injects the three classes directly (the test does this so
 *   it exercises the same aggregation the CLI does without a second module
 *   resolution path). Otherwise they are loaded from `repoRoot`.
 */
export async function runProbe(roots, options = {}) {
  const extractors =
    options.extractors ??
    (await loadExtractors(options.repoRoot ?? process.cwd()));
  const { VbaExtractor, VbaFormExtractor, SqlQueryExtractor } = extractors;

  const filesByExt = new Map();
  const skippedByExt = new Map();
  const nodesByKind = new Map();
  const edgesByKind = new Map();
  const edgesBySynthesizer = new Map();
  const unresolvedByKind = new Map();
  const stubTargets = new Map();
  const errorsByCode = new Map();
  const sqlTables = new Set();
  const forms = new Set();
  const errorConfig = resolveErrorConfig(options);
  const errorRecords = [];

  let dispatched = 0;
  let skipped = 0;
  let declaredProcedures = 0;
  let stubProcedures = 0;
  let proceduresWithNoOutgoing = 0;

  for (const absPath of collectFiles(roots)) {
    const extKey = fileExtensionKey(absPath);
    const dispatch = dispatchFor(absPath);
    if (dispatch === null) {
      skipped += 1;
      bump(skippedByExt, extKey);
      continue;
    }
    dispatched += 1;
    bump(filesByExt, extKey);

    let result;
    let source = '';
    try {
      source = fs.readFileSync(absPath, 'utf8');
      if (dispatch === 'form') {
        result = new VbaFormExtractor(absPath, source).extract();
      } else if (dispatch === 'sql') {
        result = new SqlQueryExtractor(absPath, source).extract();
      } else {
        result = new VbaExtractor(absPath, source).extract();
      }
    } catch {
      bump(errorsByCode, 'probe_extractor_threw');
      continue;
    }

    // Node id -> name, so edges (which carry ids) can be attributed to the
    // name a human reads in the roadmap tables.
    const nodeNameById = new Map();
    for (const node of result.nodes) nodeNameById.set(node.id, node.name);

    for (const node of result.nodes) {
      bump(nodesByKind, node.kind);
      if (node.kind !== 'function') continue;
      // THE stub discriminator. See the header comment.
      if (node.metadata?.stub === true) {
        stubProcedures += 1;
        bump(stubTargets, node.name);
      } else {
        declaredProcedures += 1;
      }
    }

    const outgoing = new Set();
    for (const edge of result.edges) {
      bump(edgesByKind, edge.kind);
      bump(edgesBySynthesizer, edge.metadata?.synthesizedBy ?? '(none)');
      if (edge.source) outgoing.add(edge.source);
      if (SQL_TABLE_SYNTHESIZERS.has(edge.metadata?.synthesizedBy)) {
        const name = nodeNameById.get(edge.target);
        if (name) sqlTables.add(name);
      }
      if (edge.kind === 'opens-form') {
        const name =
          edge.metadata?.targetFormName ?? nodeNameById.get(edge.target);
        if (name) forms.add(name);
      }
    }

    for (const ref of result.unresolvedReferences) {
      bump(unresolvedByKind, ref.referenceKind);
      if (ref.fromNodeId) outgoing.add(ref.fromNodeId);
      if (ref.metadata?.synthesizedBy === FORM_BINDING_SYNTHESIZER) {
        forms.add(ref.referenceName);
      }
    }

    for (const node of result.nodes) {
      if (node.kind !== 'function') continue;
      if (node.metadata?.stub === true) continue;
      if (!outgoing.has(node.id)) proceduresWithNoOutgoing += 1;
    }

    for (const error of result.errors) {
      bump(errorsByCode, error.code ?? '(none)');
    }

    // Task E1. Only code modules — a `.form.txt` carries no procedures, and a
    // `.sql` file carries no VBA at all. The declaration lines come from the
    // extractor's own non-stub `function` nodes so `proceduresTotal` cannot
    // drift from `declaredProcedures`.
    if (dispatch === 'code') {
      const procedures = result.nodes
        .filter((n) => n.kind === 'function' && n.metadata?.stub !== true)
        .map((n) => ({
          startLine: n.startLine,
          name: n.qualifiedName ?? n.name,
        }));
      errorRecords.push(
        ...analyzeVbaErrorHandling(absPath, source, procedures, errorConfig),
      );
    }
  }

  const stubTargetsTop = [...stubTargets.entries()]
    .sort((a, b) => b[1] - a[1] || (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0))
    .slice(0, STUB_TARGETS_TOP_N)
    .map(([name, count]) => ({ name, count }));

  return {
    roots: roots.map((r) => path.resolve(r)),
    files: {
      total: dispatched + skipped,
      dispatched,
      skipped,
      byExtension: sortedObject(filesByExt),
      skippedByExtension: sortedObject(skippedByExt),
    },
    nodesByKind: sortedObject(nodesByKind),
    declaredProcedures,
    stubProcedures,
    stubTargetsTop,
    edgesByKind: sortedObject(edgesByKind),
    edgesBySynthesizer: sortedObject(edgesBySynthesizer),
    unresolvedByKind: sortedObject(unresolvedByKind),
    sqlTablesReferenced: { count: sqlTables.size, names: [...sqlTables].sort() },
    formsReferenced: { count: forms.size, names: [...forms].sort() },
    proceduresWithNoOutgoing,
    errorHandling: summarizeErrorHandling(errorRecords, errorConfig),
    errors: sortedObject(errorsByCode),
  };
}

function table(title, obj, keyHeader) {
  const lines = [`### ${title}`, '', `| ${keyHeader} | count |`, '|---|--:|'];
  const entries = Object.entries(obj);
  if (entries.length === 0) lines.push('| _(none)_ | 0 |');
  for (const [k, v] of entries) lines.push(`| \`${k}\` | ${v} |`);
  lines.push('');
  return lines.join('\n');
}

/** Render the report as markdown — the shape pasted into PR bodies. */
export function formatReport(report) {
  const out = ['# VBA coverage probe', ''];
  out.push('Roots:');
  for (const root of report.roots) out.push(`  - \`${root}\``);
  out.push('');

  out.push('### Files', '', '| extension | files |', '|---|--:|');
  for (const [k, v] of Object.entries(report.files.byExtension)) {
    out.push(`| \`${k}\` | ${v} |`);
  }
  for (const [k, v] of Object.entries(report.files.skippedByExtension)) {
    out.push(`| \`${k}\` _(no extractor)_ | ${v} |`);
  }
  out.push(`| **dispatched** | **${report.files.dispatched}** |`);
  out.push(`| **total walked** | **${report.files.total}** |`, '');

  out.push(table('Nodes by kind', report.nodesByKind, 'kind'));

  const fnTotal = report.declaredProcedures + report.stubProcedures;
  out.push(
    '### Procedures',
    '',
    '| | count |',
    '|---|--:|',
    `| declared procedures (no \`metadata.stub\`) | ${report.declaredProcedures} |`,
    `| stub function nodes (\`metadata.stub === true\`) | ${report.stubProcedures} |`,
    `| \`function\` nodes total | ${fnTotal} |`,
    `| declared procedures with no outgoing | ${report.proceduresWithNoOutgoing} |`,
    '',
  );

  out.push(
    `### Top ${STUB_TARGETS_TOP_N} stub targets`,
    '',
    '| stub name | count |',
    '|---|--:|',
  );
  if (report.stubTargetsTop.length === 0) out.push('| _(none)_ | 0 |');
  for (const { name, count } of report.stubTargetsTop) {
    out.push(`| \`${name}\` | ${count} |`);
  }
  out.push('');

  out.push(table('Edges by kind', report.edgesByKind, 'kind'));
  out.push(
    table('Edges by synthesizer', report.edgesBySynthesizer, 'synthesizedBy'),
  );
  out.push(
    table('Unresolved references by kind', report.unresolvedByKind, 'referenceKind'),
  );

  out.push(
    '### Reach',
    '',
    '| | count |',
    '|---|--:|',
    `| distinct SQL tables (\`${[...SQL_TABLE_SYNTHESIZERS].join('` + `')}\`) | ${report.sqlTablesReferenced.count} |`,
    `| distinct forms (\`opens-form\` / form refs) | ${report.formsReferenced.count} |`,
    '',
  );

  out.push(formatErrorHandling(report.errorHandling));
  out.push(table('Extraction errors by code', report.errors, 'code'));
  return out.join('\n');
}

/** The `errorHandling` block, rendered as the tables §2 of the plan uses. */
export function formatErrorHandling(eh) {
  const pct = (n) =>
    eh.proceduresTotal === 0
      ? '—'
      : `${((n / eh.proceduresTotal) * 100).toFixed(1)}%`;
  const out = [];

  out.push(
    '### Error handling — protection',
    '',
    '| | count | share |',
    '|---|--:|--:|',
    `| procedures | ${eh.proceduresTotal} | 100% |`,
    `| …with \`On Error GoTo <label>\` | ${eh.proceduresByProtection.handler} | ${pct(eh.proceduresByProtection.handler)} |`,
    `| …with only \`On Error Resume Next\` | ${eh.proceduresByProtection.resumeNext} | ${pct(eh.proceduresByProtection.resumeNext)} |`,
    `| …with no protection | ${eh.proceduresByProtection.none} | ${pct(eh.proceduresByProtection.none)} |`,
    `| …unprotected and > 5 statement lines | ${eh.unprotectedOverFiveLoc} | ${pct(eh.unprotectedOverFiveLoc)} |`,
    `| …of those, touching DB/DoCmd/files/conversions | ${eh.unprotectedTouchingIo} | ${pct(eh.unprotectedTouchingIo)} |`,
    '',
  );

  out.push(
    '### Error handling — handler behaviour',
    '',
    'Exclusive classification. `mixed` means more than one signal fired.',
    '',
    '| behaviour | count |',
    '|---|--:|',
    `| \`channel\` | ${eh.handlersByBehavior.channel} |`,
    `| \`display\` | ${eh.handlersByBehavior.display} |`,
    `| \`reraise\` | ${eh.handlersByBehavior.reraise} |`,
    `| \`mixed\` | ${eh.handlersByBehavior.mixed} |`,
    `| \`unknown\` | ${eh.handlersByBehavior.unknown} |`,
    '',
    'Raw per-signal totals (non-exclusive — a `mixed` body is counted twice).',
    '',
    '| signal | handlers |',
    '|---|--:|',
    `| writes the error channel | ${eh.handlerSignals.channel} |`,
    `| displays | ${eh.handlerSignals.display} |`,
    `| re-raises | ${eh.handlerSignals.reraise} |`,
    '',
  );

  out.push(table('Error handling — handler label names', eh.handlerLabelNames, 'label'));

  out.push(
    '### Error handling — labels and scopes',
    '',
    '| | count |',
    '|---|--:|',
    `| line labels defined | ${eh.labels.defined} |`,
    `| …that are handler targets | ${eh.labels.handlerTargets} |`,
    `| …that are pure control flow | ${eh.labels.controlFlow} |`,
    `| \`Resume Next\` scopes never closed | ${eh.resumeNextOpenScopes} |`,
    `| procedures with more than one handler | ${eh.proceduresWithMultipleHandlers} |`,
    `| dangling \`GoTo\` targets | ${eh.danglingGotoTargets.length} |`,
    '',
  );

  if (eh.danglingGotoTargets.length > 0) {
    out.push('| file | procedure | label |', '|---|---|---|');
    for (const d of eh.danglingGotoTargets) {
      out.push(`| \`${d.file}\` | \`${d.procedure}\` | \`${d.label}\` |`);
    }
    out.push('');
  }

  out.push(table('Error handling — statements', eh.statements, 'construct'));

  out.push(
    '### Error handling — classifier configuration',
    '',
    `- error channel: ${eh.config.errorChannelNames.map((n) => `\`${n}\``).join(', ')}`,
    `- display calls: ${eh.config.displayCalls.map((n) => `\`${n}\``).join(', ')}`,
    `- I/O markers: ${eh.config.ioMarkers.length} patterns`,
    '',
  );

  return out.join('\n');
}

/** `--error-channel=m_Error,p_Error` -> `['m_Error', 'p_Error']`. */
function parseListFlag(argv, name) {
  const prefix = `--${name}=`;
  const flag = argv.find((a) => a.startsWith(prefix));
  if (!flag) return undefined;
  const values = flag
    .slice(prefix.length)
    .split(',')
    .map((v) => v.trim())
    .filter((v) => v !== '');
  return values.length > 0 ? values : undefined;
}

async function main(argv) {
  const asJson = argv.includes('--json');
  const roots = argv.filter((a) => !a.startsWith('--'));
  if (roots.length === 0) {
    console.error(
      'usage: npx tsx scripts/vba-coverage-probe.mjs [--json] ' +
        '[--error-channel=<name,...>] [--display-calls=<name,...>] <root> [<root> ...]',
    );
    process.exit(2);
  }
  for (const root of roots) {
    if (!fs.existsSync(root)) {
      console.error(`root does not exist: ${root}`);
      process.exit(2);
    }
  }
  const report = await runProbe(roots, {
    repoRoot: process.cwd(),
    errorChannelNames: parseListFlag(argv, 'error-channel'),
    displayCalls: parseListFlag(argv, 'display-calls'),
    ioMarkers: parseListFlag(argv, 'io-markers'),
  });
  console.log(asJson ? JSON.stringify(report, null, 2) : formatReport(report));
}

// Only run the CLI when invoked as a script, never when imported by a test.
const invokedPath = process.argv[1] ? pathToFileURL(process.argv[1]).href : '';
if (invokedPath === import.meta.url) {
  main(process.argv.slice(2)).catch((err) => {
    console.error(err);
    process.exit(1);
  });
}
