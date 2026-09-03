/**
 * Issue #263 (task E6 of `docs/vba-error-handling-plan.md`): turn the line
 * labels the error-policy classifier has ALREADY found into `label` nodes,
 * plus the edges that bind them to their procedure.
 *
 * This module is EMISSION ONLY, in the same sense `parameters.ts` is. Every
 * fact it publishes was computed by `errors.ts` while the procedure body was
 * open — label definitions, `On Error GoTo` targets, the handler region, the
 * dangling-target resolution, and the handler's behaviour. Nothing here
 * re-reads a source line, and nothing here re-classifies. If a change to this
 * file needs a regex, the fact is missing upstream and belongs there instead:
 * two scanners disagreeing about what a label is, is exactly the failure #259
 * spent a whole file's worth of documentation avoiding.
 *
 * ## Why a node at all — and why it is not a container
 *
 * §4 of the plan rejected this design for #259, and §4.3 named the condition
 * that reopens it: a consumer that needs to address a handler *as a thing*
 * rather than ask a procedure a yes/no question. That is what this is. The
 * label node makes a handler searchable (`kind:label`), linkable by a stable
 * id, and queryable for the defects — dangling targets, duplicate handlers,
 * pure control-flow labels — that `errorPolicy`'s scalar fields flatten away.
 *
 * What it deliberately is NOT is a container. **Calls inside a handler stay
 * attributed to the enclosing procedure.** Re-parenting them onto the label
 * would change `callers` / `callees` for the 3,774 procedures in the
 * reference corpus that have a handler, and silently break every consumer
 * query that assumes a call belongs to the procedure containing it. #260's
 * `metadata.inErrorHandler` already answers "did this call come from the
 * error path"; the label is *addressable*, not a parent.
 *
 * ## The collision trap
 *
 * VBA scopes a line label to its procedure, and this corpus writes the same
 * label everywhere: `errores` is defined 3,735 times. `qualifiedName` is
 * therefore always `<ModuleOrClass>.<Procedure>.<label>` — the same shape
 * `parameters.ts` (#257) and the module-variable sweep (#251) chose, and for
 * the same reason. Drop the procedure segment and every handler in a project
 * becomes one symbol.
 *
 * ## Never fabricate a target
 *
 * `On Error GoTo <label>` where `<label>` is never defined in that procedure
 * is a handler that can never run — the defect `errorPolicy.danglingTarget`
 * reports. It emits an `UnresolvedReference` and **no node**. A graph that
 * invents its own targets cannot be used to find this defect, which is the
 * only reason to look for it.
 */
import { generateNodeId } from '../tree-sitter-helpers';
import type {
  VbaErrorPolicy,
  VbaErrorPolicyState,
  VbaExtractorContext,
  VbaLabelSite,
} from './context';

/** `handles-error` edges and their unresolved twin carry this provenance. */
export const ERROR_HANDLER_SYNTHESIZED_BY = 'vba-error-handler';

/** A plain `GoTo <label>` jump — a generic `references` edge. */
export const GOTO_SYNTHESIZED_BY = 'vba-goto';

/**
 * A `GoTo` (plain or `On Error`) naming a label the procedure never defines.
 * One shared tag for both, because from a consumer's point of view they are
 * the same defect: a jump with no landing site.
 */
export const GOTO_UNRESOLVED_SYNTHESIZED_BY = 'vba-goto-unresolved';

/**
 * Emit this procedure's `label` nodes and the three edge shapes that reach
 * them, then hand control back to `closeErrorPolicy`.
 *
 * `policy` is the object that was just folded from `state`; it is read, never
 * recomputed. `endLine` is the procedure's terminating `End` line.
 *
 * Ordering matters and is the caller's responsibility: this runs BEFORE
 * #260's `markErrorHandlerRegion`, so a `GoTo` or a second `On Error GoTo`
 * written inside the handler region is stamped `metadata.inErrorHandler` by
 * that single stamping point like every other edge, instead of quietly
 * becoming the one emitter that opted out of it.
 */
export function emitLabelNodes(
  ctx: VbaExtractorContext,
  state: VbaErrorPolicyState,
  policy: VbaErrorPolicy,
  endLine: number,
): void {
  const procedureNode = ctx.functionNodeByStartLine.get(state.startLine);
  // No procedure node means the pre-walk never saw a declaration for this
  // body. There is nothing to attach a label to, and a parentless label node
  // is worse than no node.
  if (!procedureNode) return;
  if (
    state.labelDefs.length === 0 &&
    state.onErrorSites.length === 0 &&
    state.gotoSites.length === 0
  ) {
    return;
  }

  const procedureName = procedureNode.name;
  const qualifiedPrefix = ctx.moduleName
    ? `${ctx.moduleName}.${procedureName}`
    : procedureName;

  // The ONE label whose region `errorPolicy` resolved — `openedTarget`, the
  // earliest definition an `On Error GoTo` names. Only it can copy #260's
  // derived `behavior` and region lines, because only it is what #260
  // derived them FOR. A procedure that swaps to a second handler label has a
  // second region nobody has classified, and inventing one here would be the
  // re-classification this module exists to avoid.
  const regionOwner = policy.handlerStartLine === null ? null : state.openedTarget;

  const nodeIdByKey = new Map<string, string>();

  for (const def of state.labelDefs) {
    const isHandler = state.targets.has(def.key);
    const nodeId = generateNodeId(ctx.filePath, 'label', def.name, def.line);
    // `definedLabels` already kept only the first definition per key, so a
    // repeat can only arrive from malformed VBA. Keep the first.
    if (nodeIdByKey.has(def.key)) continue;
    nodeIdByKey.set(def.key, nodeId);

    const metadata: Record<string, unknown> = { isHandler };
    if (isHandler && def.key === regionOwner) {
      metadata.handlerBehavior = policy.behavior;
      metadata.regionStartLine = policy.handlerStartLine;
      metadata.regionEndLine = policy.handlerEndLine;
    }

    ctx.nodes.push({
      id: nodeId,
      kind: 'label',
      name: def.name,
      qualifiedName: `${qualifiedPrefix}.${def.name}`,
      filePath: ctx.filePath,
      language: 'vba',
      startLine: def.line,
      // A handler's body runs from its label to the procedure's end, so the
      // node spans it. A control-flow label is a jump target and nothing
      // more — it owns exactly its own line.
      endLine: isHandler ? endLine : def.line,
      startColumn: def.column,
      endColumn: def.column + def.name.length,
      metadata,
      updatedAt: Date.now(),
    });

    ctx.edges.push({
      source: procedureNode.id,
      target: nodeId,
      kind: 'contains',
    });
  }

  for (const site of state.onErrorSites) {
    emitJump(
      ctx,
      procedureNode.id,
      nodeIdByKey,
      site,
      'handles-error',
      ERROR_HANDLER_SYNTHESIZED_BY,
    );
  }

  for (const site of state.gotoSites) {
    emitJump(
      ctx,
      procedureNode.id,
      nodeIdByKey,
      site,
      'references',
      GOTO_SYNTHESIZED_BY,
    );
  }
}

/**
 * One jump: an edge onto the label's node when the procedure defines it, an
 * `UnresolvedReference` when it does not.
 *
 * The unresolved row keeps `referenceKind: 'references'` for BOTH jump kinds.
 * `handles-error` would be wrong there: the row records a target that does
 * not exist, so nothing routes errors to it, and a resolver that later found
 * a same-named symbol elsewhere in the project would materialise a
 * cross-procedure error edge VBA's own scoping rules forbid.
 */
function emitJump(
  ctx: VbaExtractorContext,
  procedureNodeId: string,
  nodeIdByKey: ReadonlyMap<string, string>,
  site: VbaLabelSite,
  edgeKind: 'handles-error' | 'references',
  synthesizedBy: string,
): void {
  const targetId = nodeIdByKey.get(site.key);
  if (targetId) {
    ctx.edges.push({
      source: procedureNodeId,
      target: targetId,
      kind: edgeKind,
      provenance: 'heuristic',
      metadata: { synthesizedBy },
      line: site.line,
      column: site.column,
    });
    return;
  }

  ctx.unresolvedReferences.push({
    fromNodeId: procedureNodeId,
    referenceName: site.name,
    referenceKind: 'references',
    line: site.line,
    column: site.column,
    filePath: ctx.filePath,
    language: 'vba',
    metadata: { synthesizedBy: GOTO_UNRESOLVED_SYNTHESIZED_BY },
  });
}
