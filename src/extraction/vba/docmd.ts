/**
 * `DoCmd.<Verb>` modelling (hueco 6 / issues #48, #52, #254).
 *
 * Three shapes live here:
 *   1. `OpenForm` / `OpenReport` — a synthetic form-layout/report-layout stub
 *      plus an `opens-form` / `opens-report` edge.
 *   2. `OpenQuery` — an `UnresolvedReference` the resolver binds to the real
 *      `query` node emitted for `queries/<Name>.sql`.
 *   3. Issue #254: every OTHER `DoCmd` verb that names an Access object in a
 *      positional argument (`RunMacro`, `TransferSpreadsheet`, `OutputTo`, …).
 *      These emit an `UnresolvedReference` and NEVER a node — see
 *      `DOCMD_OBJECT_DISPATCH`.
 *
 * `DoCmd` is in `RUNTIME_RECEIVER_BLACKLIST`, so these methods are skipped by
 * the generic call-site scan and handled here with their own emission path.
 */
import { generateNodeId } from '../tree-sitter-helpers';
import { unwrapVbaStringLiteral } from './text-utils';
import { VbaExtractorContext, ProcInfo } from './context';

/**
 * `DoCmd.OpenForm "<FormName>"` modelling regex — B4 (hueco 6). Captures the
 * first argument (group 1) as either a `"literal"` or a bare identifier; the
 * trailing positional args (`acNormal`, `acFormEdit`, …) are NOT captured.
 */
const OPEN_FORM_ARG_RE =
  /\bDoCmd\.OpenForm\s+("(?:(?:[^"]|"")*)"|\p{L}[\p{L}\p{N}_]*)/gu;

/**
 * Issue #48: `DoCmd.OpenReport "<ReportName>"` modelling regex — sibling
 * of `OPEN_FORM_ARG_RE`. Same literal-or-bare-id argument capture and same
 * trailing positional-args drop.
 */
const OPEN_REPORT_ARG_RE =
  /\bDoCmd\.OpenReport\s+("(?:(?:[^"]|"")*)"|\p{L}[\p{L}\p{N}_]*)/gu;

/**
 * Issue #48: `DoCmd.OpenQuery "<QueryName>"` modelling regex. Emits an
 * `UnresolvedReference` (NOT a stub + edge) so the resolver binds to the
 * REAL `query` node `SqlQueryExtractor` emits for `queries/<Name>.sql`.
 */
const OPEN_QUERY_ARG_RE =
  /\bDoCmd\.OpenQuery\s+("(?:(?:[^"]|"")*)"|\p{L}[\p{L}\p{N}_]*)/gu;

type DoCmdOpenDispatch = {
  method: 'OpenForm' | 'OpenReport';
  re: RegExp;
  edgeKind: 'opens-form' | 'opens-report';
  stubKind: 'form-layout' | 'report-layout';
  syntheticPrefix: 'synthetic:opensFormStub' | 'synthetic:opensReportStub';
  syntheticExtension: '.form.txt' | '.report.txt';
  moduleNamePrefix: 'Form_' | 'Report_';
  cacheKey: 'OpenForm' | 'OpenReport';
  metadataTargetKey: 'targetFormName' | 'targetReportName';
  synthesizedBy: 'vba-opens-form' | 'vba-opens-report';
};

/**
 * Issue #48 dispatch table — shared literal-or-Const argument resolution
 * for `DoCmd.OpenForm` and `DoCmd.OpenReport`. Each entry carries everything
 * `scanDoCmdOpenCalls` + `emitOpensStubEdge` need to share the pipeline
 * between methods while keeping the per-method names distinct.
 *
 * OpenQuery is intentionally NOT in this dispatch — it emits an
 * `UnresolvedReference` (not a stub + edge). See `scanDoCmdOpenQuery`.
 */
const DOCMD_OPEN_DISPATCH: ReadonlyArray<DoCmdOpenDispatch> = [
  {
    method: 'OpenForm',
    re: OPEN_FORM_ARG_RE,
    edgeKind: 'opens-form',
    stubKind: 'form-layout',
    syntheticPrefix: 'synthetic:opensFormStub',
    syntheticExtension: '.form.txt',
    moduleNamePrefix: 'Form_',
    cacheKey: 'OpenForm',
    metadataTargetKey: 'targetFormName',
    synthesizedBy: 'vba-opens-form',
  },
  {
    method: 'OpenReport',
    re: OPEN_REPORT_ARG_RE,
    edgeKind: 'opens-report',
    stubKind: 'report-layout',
    syntheticPrefix: 'synthetic:opensReportStub',
    syntheticExtension: '.report.txt',
    moduleNamePrefix: 'Report_',
    cacheKey: 'OpenReport',
    metadataTargetKey: 'targetReportName',
    synthesizedBy: 'vba-opens-report',
  },
];

/**
 * B4 (hueco 6) extended by Issue #48: scan one line of VBA source for
 * `DoCmd.OpenX "Target"` calls where X ∈ {Form, Report}. For each match emit
 * a cached stub node (form-layout / report-layout) and an
 * `opens-form` / `opens-report` heuristic edge from the calling Sub.
 */
export function scanDoCmdOpenCalls(
  ctx: VbaExtractorContext,
  line: string,
  maskedLine: string,
  caller: ProcInfo,
  lineNum: number,
): void {
  for (const dispatch of DOCMD_OPEN_DISPATCH) {
    // Each regex has /g so we MUST reset `lastIndex` before use; cloning
    // the regex is the simplest way to avoid leaking state across lines
    // AND across dispatch iterations.
    const localRe = new RegExp(dispatch.re.source, dispatch.re.flags);
    let m: RegExpExecArray | null;
    while ((m = localRe.exec(line)) !== null) {
      if (maskedLine.slice(m.index, m.index + 5).toLowerCase() !== 'docmd') continue;
      const rawArg = (m[1] ?? '').trim();
      // Issue #52: const lookup is now per-proc-bucket with module
      // fallback (see `resolveLocalConst`). Two procs declaring the
      // same Const name with different values no longer collide —
      // each call site uses the value visible at its own scope.
      const targetName = rawArg.startsWith('"')
        ? unwrapVbaStringLiteral(rawArg)
        : (ctx.resolveLocalConst(rawArg) ?? rawArg);
      if (!targetName) continue;
      emitOpensStubEdge(ctx, dispatch, caller, targetName, lineNum, m.index);
    }
  }
}

/**
 * B4 (hueco 6) extended by Issue #48: emit a stub `form-layout` /
 * `report-layout` node for `targetName` (cached per dispatch entry so
 * duplicates collapse and OpenForm/OpenReport de-dup buckets stay
 * disjoint) and a single `opens-form` / `opens-report` heuristic edge
 * from `caller` to that stub.
 */
function emitOpensStubEdge(
  ctx: VbaExtractorContext,
  dispatch: DoCmdOpenDispatch,
  caller: ProcInfo,
  targetName: string,
  lineNum: number,
  column: number,
): void {
  const key = `${dispatch.cacheKey}:${targetName.toLowerCase()}`;
  let stubId = ctx.opensStubIdsByKey.get(key);
  if (!stubId) {
    // Synthetic file path keeps the stub's id deterministic AND
    // disambiguates it from any real `.form.txt` / `.report.txt`
    // indexed later. The directory prefix is intentionally not a real
    // filesystem path — it just namespaces the id space. The file
    // extension DOES mirror the real form/report file extension so a
    // reader of the synthetic path can tell the stub's intent at a glance.
    const syntheticFilePath = `${dispatch.syntheticPrefix}/${targetName}${dispatch.syntheticExtension}`;
    stubId = generateNodeId(
      syntheticFilePath,
      dispatch.stubKind,
      targetName,
      0,
    );
    ctx.opensStubIdsByKey.set(key, stubId);
    ctx.nodes.push({
      id: stubId,
      kind: dispatch.stubKind,
      name: targetName,
      // Convention: form module names in Access are `Form_<Name>` and
      // report module names are `Report_<Name>`. We follow the same
      // convention in the synthetic stub's qualifiedName so cross-file
      // lookups can find it consistently.
      qualifiedName: `${dispatch.moduleNamePrefix}${targetName}`,
      filePath: syntheticFilePath,
      language: 'vba',
      startLine: lineNum,
      endLine: lineNum,
      startColumn: 0,
      endColumn: 0,
      metadata: { stub: true },
      updatedAt: Date.now(),
    });
  }
  ctx.edges.push({
    source: ctx.findOrCreateFunctionNodeId(caller),
    target: stubId,
    kind: dispatch.edgeKind,
    provenance: 'heuristic',
    metadata: {
      synthesizedBy: dispatch.synthesizedBy,
      [dispatch.metadataTargetKey]: targetName,
    },
    line: lineNum,
    column,
  });
}

/**
 * Issue #48: scan one line of VBA source for `DoCmd.OpenQuery "X"` calls.
 * Each match emits ONE `UnresolvedReference` (NOT a stub + edge) so the
 * resolver binds to the REAL `query` node that `SqlQueryExtractor`
 * produces for `queries/<Name>.sql`, tagged `synthesizedBy: 'vba-opens-query'`.
 */
export function scanDoCmdOpenQuery(
  ctx: VbaExtractorContext,
  line: string,
  maskedLine: string,
  caller: ProcInfo,
  lineNum: number,
): void {
  const localRe = new RegExp(OPEN_QUERY_ARG_RE.source, OPEN_QUERY_ARG_RE.flags);
  let m: RegExpExecArray | null;
  while ((m = localRe.exec(line)) !== null) {
    if (maskedLine.slice(m.index, m.index + 5).toLowerCase() !== 'docmd') continue;
    const rawArg = (m[1] ?? '').trim();
    // Issue #52: same per-proc-with-module-fallback lookup as
    // `scanDoCmdOpenCalls`.
    const targetName = rawArg.startsWith('"')
      ? unwrapVbaStringLiteral(rawArg)
      : (ctx.resolveLocalConst(rawArg) ?? rawArg);
    if (!targetName) continue;
    ctx.unresolvedReferences.push({
      fromNodeId: ctx.findOrCreateFunctionNodeId(caller),
      referenceName: targetName,
      referenceKind: 'dao-query',
      line: lineNum,
      column: m.index,
      filePath: ctx.filePath,
      language: 'vba',
      metadata: { synthesizedBy: 'vba-opens-query' },
    });
  }
}

// ---------------------------------------------------------------------------
// Issue #254 (task T12) — the remaining `DoCmd` verbs that name an object.
//
// Everything below is ONE table-driven dispatch: a verb, the zero-based
// position of the argument that names the object, and the provenance tag the
// emitted reference carries. Adding a verb is adding a row — there is
// deliberately no per-verb regex.
//
// Why a reference and never a node: Dysflow exports `classes/`, `forms/`,
// `modules/` and `queries/` to the source tree and nothing else. A macro or a
// table named by one of these verbs has no indexed artifact behind it, so a
// node for it would be a stub nothing can ever back — the fabricated-target
// shape the runtime-allowlist work is removing elsewhere. An
// `UnresolvedReference` costs nothing, stays visible as an actionable `failed`
// reference, and resolves by itself the day the artifact IS exported.
//
// Two verbs from the issue list are deliberately NOT here:
//   - `RunSQL` names a SQL statement, not an object. Its literal and
//     variable forms are already modelled by the SQL-wrapper sweep
//     (`vba/sql-wrapper.ts`), which emits `vba-sql-table` references for the
//     tables the statement touches. A second emission would double-count.
//   - `RunCommand` takes an intrinsic `acCmd*` command constant. There is no
//     project artifact to point at, so there is nothing to reference.
// ---------------------------------------------------------------------------

type DoCmdObjectDispatch = {
  /** The `DoCmd` method name, matched case-insensitively. */
  method: string;
  /**
   * Zero-based position, in the call's argument list, of the argument that
   * names an EXISTING Access object. Where a verb takes both a source and a
   * destination name (`CopyObject`, `Rename`), this is the source — the
   * destination does not exist yet, so it cannot be referenced.
   */
  argIndex: number;
  /** Per-verb provenance tag stamped on the emitted reference. */
  synthesizedBy: string;
};

/**
 * Issue #254 dispatch table — verb, argument position, provenance tag.
 *
 * Argument positions follow the documented Access `DoCmd` signatures:
 *   RunMacro            MacroName, RepeatCount, RepeatExpression
 *   OpenTable           TableName, View, DataMode
 *   ApplyFilter         FilterName, WhereCondition, ControlName
 *   CopyObject          DestinationDatabase, NewName, SourceObjectType, SourceObjectName
 *   DeleteObject        ObjectType, ObjectName
 *   Rename              NewName, ObjectType, OldName
 *   SelectObject        ObjectType, ObjectName, InDatabaseWindow
 *   BrowseTo            ObjectType, ObjectName, PathToSubformControl, …
 *   OutputTo            ObjectType, ObjectName, OutputFormat, OutputFile, …
 *   SendObject          ObjectType, ObjectName, OutputFormat, To, …
 *   TransferSpreadsheet TransferType, SpreadsheetType, TableName, FileName, …
 *   TransferText        TransferType, SpecificationName, TableName, FileName, …
 *   TransferDatabase    TransferType, DatabaseType, DatabaseName, ObjectType, Source, Destination, …
 */
const DOCMD_OBJECT_DISPATCH: ReadonlyArray<DoCmdObjectDispatch> = [
  { method: 'RunMacro', argIndex: 0, synthesizedBy: 'vba-runs-macro' },
  { method: 'OpenTable', argIndex: 0, synthesizedBy: 'vba-opens-table' },
  { method: 'ApplyFilter', argIndex: 0, synthesizedBy: 'vba-applies-filter' },
  { method: 'CopyObject', argIndex: 3, synthesizedBy: 'vba-copies-object' },
  { method: 'DeleteObject', argIndex: 1, synthesizedBy: 'vba-deletes-object' },
  { method: 'Rename', argIndex: 2, synthesizedBy: 'vba-renames-object' },
  { method: 'SelectObject', argIndex: 1, synthesizedBy: 'vba-selects-object' },
  { method: 'BrowseTo', argIndex: 1, synthesizedBy: 'vba-browses-to' },
  { method: 'OutputTo', argIndex: 1, synthesizedBy: 'vba-outputs-to' },
  { method: 'SendObject', argIndex: 1, synthesizedBy: 'vba-sends-object' },
  {
    method: 'TransferSpreadsheet',
    argIndex: 2,
    synthesizedBy: 'vba-transfers-spreadsheet',
  },
  { method: 'TransferText', argIndex: 2, synthesizedBy: 'vba-transfers-text' },
  {
    method: 'TransferDatabase',
    argIndex: 4,
    synthesizedBy: 'vba-transfers-database',
  },
];

/**
 * One regex per dispatch row, built once at module load. It matches only the
 * `DoCmd.<Verb>` head plus the optional opening paren — the argument list is
 * split by `splitDoCmdArguments`, because a positional lookup through a list
 * that mixes literals, intrinsic constants and omitted arguments is not
 * something a regex should be asked to do.
 */
const DOCMD_OBJECT_HEAD_RES: ReadonlyArray<RegExp> = DOCMD_OBJECT_DISPATCH.map(
  (d) => new RegExp(`\\bDoCmd\\.${d.method}\\b[ \\t]*\\(?[ \\t]*`, 'giu'),
);

/**
 * Split a `DoCmd` argument list into its top-level arguments, in source order.
 *
 * `text` is the remainder of the line AFTER the `DoCmd.<Verb>` head. Commas
 * inside a string literal or inside a nested paren/bracket group do not split;
 * an omitted argument (`DoCmd.CopyObject , "X", …`) yields an empty string so
 * the positions of the arguments that ARE present stay correct.
 *
 * The scan stops at the first top-level `)` (closing the parenthesised call
 * form), `:` (VBA statement separator) or a comment quote — comments are
 * already stripped upstream, so that last one is defence in depth.
 */
function splitDoCmdArguments(text: string): string[] {
  const args: string[] = [];
  let current = '';
  let depth = 0;
  let inString = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i] ?? '';
    if (inString) {
      current += ch;
      if (ch === '"') {
        if (text[i + 1] === '"') {
          current += '"';
          i++;
        } else {
          inString = false;
        }
      }
      continue;
    }
    if (ch === '"') {
      inString = true;
      current += ch;
      continue;
    }
    if (ch === '(' || ch === '[') {
      depth++;
      current += ch;
      continue;
    }
    if (ch === ')' || ch === ']') {
      if (depth === 0) break;
      depth--;
      current += ch;
      continue;
    }
    if (depth === 0 && (ch === "'" || ch === ':')) break;
    if (depth === 0 && ch === ',') {
      args.push(current.trim());
      current = '';
      continue;
    }
    current += ch;
  }
  args.push(current.trim());
  return args;
}

/**
 * Resolve one `DoCmd` object argument to the name it denotes, or `null` when
 * it denotes nothing we can know statically.
 *
 * A `"literal"` is taken verbatim. A bare identifier is resolved through the
 * per-proc-then-module `Const` lookup (Issue #52) — a `Const` holding a string
 * literal IS a static name. Anything else (an undeclared variable, a
 * concatenation, a function call) resolves to `null` and the call site stays
 * SILENT. This is deliberately stricter than the `OpenForm`/`OpenReport`
 * bare-identifier fallback: a guessed target on a verb nobody in the corpus
 * uses would be noise nobody can act on.
 */
function resolveDoCmdObjectArgument(
  ctx: VbaExtractorContext,
  rawArg: string,
): string | null {
  const arg = rawArg.trim();
  if (!arg) return null;
  if (arg.startsWith('"')) return unwrapVbaStringLiteral(arg) || null;
  if (!/^\p{L}[\p{L}\p{N}_]*$/u.test(arg)) return null;
  return ctx.resolveLocalConst(arg) ?? null;
}

/**
 * Issue #254: scan one line of VBA source for every `DoCmd` verb in
 * `DOCMD_OBJECT_DISPATCH`. Each match whose object argument resolves to a
 * static name emits ONE `UnresolvedReference` — no node, ever.
 *
 * `maskedLine` has string CONTENT blanked out, so the `docmd` prefix check
 * rejects a verb name that only appears inside a string literal. The object
 * name itself lives inside a literal, which is why the split runs on the
 * original `line`.
 */
export function scanDoCmdObjectCalls(
  ctx: VbaExtractorContext,
  line: string,
  maskedLine: string,
  caller: ProcInfo,
  lineNum: number,
): void {
  for (let d = 0; d < DOCMD_OBJECT_DISPATCH.length; d++) {
    const dispatch = DOCMD_OBJECT_DISPATCH[d]!;
    // The shared regexes carry /g, so clone before use to keep `lastIndex`
    // from leaking across lines and across dispatch rows.
    const head = DOCMD_OBJECT_HEAD_RES[d]!;
    const localRe = new RegExp(head.source, head.flags);
    let m: RegExpExecArray | null;
    while ((m = localRe.exec(line)) !== null) {
      if (maskedLine.slice(m.index, m.index + 5).toLowerCase() !== 'docmd') continue;
      const args = splitDoCmdArguments(line.slice(m.index + m[0].length));
      const targetName = resolveDoCmdObjectArgument(
        ctx,
        args[dispatch.argIndex] ?? '',
      );
      if (!targetName) continue;
      ctx.unresolvedReferences.push({
        fromNodeId: ctx.findOrCreateFunctionNodeId(caller),
        referenceName: targetName,
        // `references` is an EdgeKind literal, so if the named artifact ever
        // IS indexed the reference becomes a plain `references` edge with no
        // extractor change.
        referenceKind: 'references',
        line: lineNum,
        column: m.index,
        filePath: ctx.filePath,
        language: 'vba',
        metadata: {
          synthesizedBy: dispatch.synthesizedBy,
          docmdVerb: dispatch.method,
          targetName,
        },
      });
    }
  }
}
