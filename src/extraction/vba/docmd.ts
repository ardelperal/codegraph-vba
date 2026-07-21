/**
 * `DoCmd.OpenForm` / `DoCmd.OpenReport` / `DoCmd.OpenQuery` modelling
 * (hueco 6 / issues #48, #52). OpenForm/OpenReport emit a synthetic
 * form-layout/report-layout stub + `opens-form`/`opens-report` edge; OpenQuery
 * emits an `UnresolvedReference` the resolver binds to the real `query` node.
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
