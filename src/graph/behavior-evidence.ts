/**
 * Behavior evidence — the bounded, typed answer to "what does this Access
 * control actually do?", assembled from facts that are already indexed.
 *
 * Issue #299. Every other consumer surface returns prose keyed on a fuzzy
 * symbol name; a machine consumer (Dysflow's behavior map) needs a typed
 * array with deterministic call paths and an explicit account of what could
 * NOT be answered. That is why this is its own operation rather than another
 * mode of `codegraph_node` / `codegraph_impact`: those return agent-facing
 * text and would have to grow a second, contradictory output shape.
 *
 * It is READ-ONLY. Nothing here parses source, opens Access, writes to the
 * index or triggers a re-index.
 *
 * ## The graph contract it reads
 *
 * - An `event-handler` edge is stored HANDLER -> control/layout, for the
 *   `<Control>_<Event>` code-behind convention AND for `=Expression()`
 *   wiring. Reaching a handler from its control means following it backwards.
 * - A call is a `calls` edge, or a `references` edge landing on a procedure:
 *   VBA's statement-form Sub call (`SaveRecord` on a line of its own) is
 *   ambiguous at parse time, so it stays `unqualified-ident` (#265) and
 *   resolves to `references`, not `calls`.
 * - Data access is a `references` edge onto a table/query node, carrying
 *   `metadata.access` ('read' | 'write') when the extractor could tell. Those
 *   edges are sourced from the MODULE node by design (see
 *   `VbaExtractionContext.emitReference`), so they are attributed to a
 *   procedure by the edge's LINE falling inside that procedure's range — a
 *   fact the index already stores, not a name guess.
 *
 * ## What it deliberately does not do
 *
 * No effect is inferred. An empty `effects` or `tables` list means the index
 * holds no such fact for that path — never that the code has no runtime
 * effect. Dynamic dispatch, late binding and anything the extractor declined
 * are reported in `context.unresolved` instead of being silently dropped.
 */
import * as path from 'path';
import { QueryBuilder } from '../db/queries';
import { Edge, Node } from '../types';

/**
 * The consumer compatibility payload.
 *
 * Field types are fixed by the consumer that reads them (Dysflow's
 * `CodeGraphBehaviorEvidence`) and must not be widened here.
 */
export interface CodeGraphBehaviorEvidence {
  /** Name of the procedure the event is bound to. */
  handler: string;
  /**
   * One root-to-leaf execution path, handler first, callees in the order the
   * graph stores them (by node id, so the output is stable across runs).
   * Distinct branches are separate entries — they are never concatenated into
   * a single sequence the runtime would not take. A path that re-enters a
   * procedure already on it ends there, repeating that name once.
   */
  callPath: string[];
  /** Tables reached along this path, directly or through a saved query. */
  tables?: string[];
  /** See {@link BEHAVIOR_EFFECT_VOCABULARY}. */
  effects?: string[];
}

/**
 * The closed vocabulary of `effects` strings. Anything outside this list is a
 * bug, not a new fact.
 *
 * - `read:<name>` / `write:<name>` — a data reference the extractor typed.
 * - `data-access:<name>` — a data reference whose direction is unknown.
 *   NOT a read and NOT a write; it means the index cannot say.
 * - `opens-form:<Name>` / `opens-report:<Name>` — an Access object opened.
 * - `raises-event:<Name>` — an event raised from the path.
 */
export const BEHAVIOR_EFFECT_VOCABULARY = [
  'read:<name>',
  'write:<name>',
  'data-access:<name>',
  'opens-form:<Name>',
  'opens-report:<Name>',
  'raises-event:<Name>',
] as const;

/** How a handler is bound to the thing it answers for. */
export type BehaviorBindingScope =
  /** `<Control>_<Event>` code-behind on a control. */
  | 'control'
  /** The form's or report's own lifecycle event. */
  | 'form'
  /** An `=Expression()` property on the control or layout. */
  | 'expression'
  /** The request named a procedure directly; no binding was involved. */
  | 'direct';

export interface BehaviorEvidenceRequest {
  /** Stable node id. The unambiguous form — always prefer it. */
  nodeId?: string;
  /** Control, handler or layout name. Requires `layout` when not unique. */
  name?: string;
  /**
   * Layout context for `name`: a form/report name or its layout file name.
   * Without it an ambiguous name is REFUSED, never silently narrowed to the
   * first match.
   */
  layout?: string;
  /** Call-path depth budget. Default 5, clamped to 1..20. */
  maxCallDepth?: number;
  /** Maximum evidence entries. Default 50, clamped to 1..500.
   * Traversal also visits at most (maxResults + 1) * maxCallDepth path steps. */
  maxResults?: number;
}

export interface BehaviorEvidenceTarget {
  id: string;
  name: string;
  kind: string;
  filePath: string;
  /** Owning form/report, from the layout's `contains` edge. */
  layout: string | null;
}

export interface BehaviorHandlerContext {
  handler: string;
  handlerId: string;
  /** Access event name (`Click`, `Load`, …), or null when not bound. */
  event: string | null;
  scope: BehaviorBindingScope;
  provenance: string;
  /** `file:line` of the wiring site. */
  location: string;
  /** `metadata.synthesizedBy` of the binding edge, when present. */
  wiredBy: string | null;
}

export interface BehaviorDataEvidence {
  /** Procedure the reference was attributed to. */
  procedure: string;
  name: string;
  /** Node kind of the referenced object (`class`, `table`, `query`, …). */
  targetKind: string;
  access: 'read' | 'write' | 'unknown';
  /** `metadata.synthesizedBy` of the reference edge. */
  via: string | null;
  /** Set when the table was reached through a saved query. */
  throughQuery?: string;
  /**
   * `edge-source` — the reference edge starts at the procedure itself.
   * `source-line` — it starts at the module node and was attributed to this
   * procedure because the edge's line falls inside its range.
   */
  attributedBy: 'edge-source' | 'source-line';
  location: string;
}

export interface BehaviorAmbiguity {
  name: string;
  matches: Array<{ id: string; kind: string; filePath: string; layout: string | null }>;
}

export interface BehaviorUnresolvedReference {
  /** Procedure the unresolved reference sits in, when it could be attributed. */
  procedure: string | null;
  referenceName: string;
  referenceKind: string;
  location: string;
}

export interface BehaviorEvidenceResult {
  /** What the request resolved to, or null when it resolved to nothing. */
  target: BehaviorEvidenceTarget | null;
  /** The consumer compatibility payload. */
  evidence: CodeGraphBehaviorEvidence[];
  context: {
    handlers: BehaviorHandlerContext[];
    data: BehaviorDataEvidence[];
    unresolved: BehaviorUnresolvedReference[];
    /** Populated only when a name matched more than one node. */
    ambiguous: BehaviorAmbiguity[];
    truncated: {
      /** A path hit `maxCallDepth` and was cut short. */
      callDepth: boolean;
      /** Entries were dropped, or the traversal work budget left paths unexplored. */
      results: boolean;
      /** A path re-entered a procedure already on it. */
      cycle: boolean;
    };
    /** See {@link BEHAVIOR_EVIDENCE_NOTES}. */
    notes: string[];
  };
}

/**
 * The closed vocabulary of `context.notes`.
 *
 * - `MISSING_TARGET_SELECTOR` — neither `nodeId` nor `name` was given.
 * - `TARGET_NOT_FOUND` — nothing in the index matches. A lookup miss, not
 *   proof the control has no behavior.
 * - `AMBIGUOUS_TARGET` — the name matches several nodes; see
 *   `context.ambiguous`. No evidence is returned, and nothing is guessed.
 * - `NO_HANDLER_BOUND` — the target exists but no handler is wired to it in
 *   the index. Access macros and `[Event Procedure]` entries with no
 *   code-behind land here.
 * - `STATIC_SOURCE_EVIDENCE` — always present. The answer comes from indexed
 *   exported source: it does not prove the code ran, and it says nothing
 *   about whether the `.accdb` binary matches the export.
 */
export const BEHAVIOR_EVIDENCE_NOTES = [
  'MISSING_TARGET_SELECTOR',
  'TARGET_NOT_FOUND',
  'AMBIGUOUS_TARGET',
  'NO_HANDLER_BOUND',
  'STATIC_SOURCE_EVIDENCE',
] as const;

const UI_NODE_KINDS = new Set([
  'form-instance-control',
  'form-layout',
  'report-layout',
]);

const LAYOUT_NODE_KINDS = new Set(['form-layout', 'report-layout']);

const PROCEDURE_KINDS = new Set(['function', 'method']);

/**
 * Node kinds a data reference can land on. An Access table has no kind of its
 * own — it is a placeholder `class` node, the same shape a VBA type gets —
 * which is why {@link isDataReference} exists.
 */
const DATA_TARGET_KINDS = new Set<string>(['class', 'query']);

/**
 * The `metadata.synthesizedBy` channels that mean "this names a table or a
 * saved query".
 *
 * An Access table is modelled as a placeholder `class` node, so the target's
 * kind alone cannot tell a table apart from an ordinary type. Without this
 * allowlist a `Dim rs As DAO.Recordset` — a TYPE reference emitted by
 * `vba-name-resolution` — would be reported as a table the handler touches,
 * which is exactly the kind of guessed evidence this contract forbids.
 */
const DATA_REFERENCE_CHANNELS = new Set([
  'vba-sql-table',
  'vba-record-source',
  'vba-row-source',
  'vba-row-source-dynamic',
  'vba-control-source',
  'vba-source-object',
  'vba-query-name',
  'vba-opens-query',
  'vba-opens-table',
  'sql-query-table',
]);

/**
 * Is this `references` edge a data access, as opposed to a type, name or
 * binding reference that happens to land on the same node shape?
 */
function isDataReference(edge: Edge, target: Node): boolean {
  if (target.kind === 'query') return true;
  const access = edge.metadata?.access;
  if (access === 'read' || access === 'write') return true;
  const via = metadataString(edge, 'synthesizedBy');
  return via !== null && DATA_REFERENCE_CHANNELS.has(via);
}

const DEFAULT_MAX_CALL_DEPTH = 5;
const DEFAULT_MAX_RESULTS = 50;

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}

function locationOf(filePath: string | undefined, line: number | undefined): string {
  if (!filePath) return 'unknown';
  return line ? `${filePath}:${line}` : filePath;
}

function metadataString(edge: Edge, key: string): string | null {
  const value = edge.metadata?.[key];
  return typeof value === 'string' ? value : null;
}

/**
 * Assembles behavior evidence for one control, layout or handler.
 *
 * Pure read path over {@link QueryBuilder}: no new SQL, no second traversal
 * implementation, no writes.
 */
export function buildBehaviorEvidence(
  queries: QueryBuilder,
  request: BehaviorEvidenceRequest,
): BehaviorEvidenceResult {
  const maxCallDepth = clamp(
    Math.trunc(request.maxCallDepth ?? DEFAULT_MAX_CALL_DEPTH),
    1,
    20,
  );
  const maxResults = clamp(
    Math.trunc(request.maxResults ?? DEFAULT_MAX_RESULTS),
    1,
    500,
  );

  const result: BehaviorEvidenceResult = {
    target: null,
    evidence: [],
    context: {
      handlers: [],
      data: [],
      unresolved: [],
      ambiguous: [],
      truncated: { callDepth: false, results: false, cycle: false },
      notes: ['STATIC_SOURCE_EVIDENCE'],
    },
  };

  const owningLayout = (node: Node): string | null => {
    if (LAYOUT_NODE_KINDS.has(node.kind)) return node.name;
    for (const edge of queries.getIncomingEdges(node.id, ['contains'])) {
      const owner = queries.getNodeById(edge.source);
      if (owner && LAYOUT_NODE_KINDS.has(owner.kind)) return owner.name;
    }
    const base = path.basename(node.filePath ?? '');
    const stripped = base.replace(/\.(form|report)\.txt$/i, '');
    return stripped === base ? null : stripped;
  };

  // ---- 1. Resolve the target ------------------------------------------
  let target: Node | null = null;

  if (request.nodeId) {
    target = queries.getNodeById(request.nodeId);
    if (!target) result.context.notes.push('TARGET_NOT_FOUND');
  } else if (request.name) {
    const wanted = request.name.toLowerCase();
    const candidates = queries
      .getNodesByLowerName(wanted)
      .filter((node) => UI_NODE_KINDS.has(node.kind) || PROCEDURE_KINDS.has(node.kind));
    const scoped = request.layout
      ? candidates.filter((node) => {
          const layoutName = owningLayout(node);
          const base = path.basename(node.filePath ?? '');
          const wantedLayout = request.layout!.toLowerCase();
          return (
            layoutName?.toLowerCase() === wantedLayout ||
            base.toLowerCase() === wantedLayout ||
            base.toLowerCase().startsWith(`${wantedLayout}.`)
          );
        })
      : candidates;

    if (scoped.length === 0) {
      result.context.notes.push('TARGET_NOT_FOUND');
    } else if (scoped.length > 1) {
      result.context.notes.push('AMBIGUOUS_TARGET');
      result.context.ambiguous.push({
        name: request.name,
        matches: scoped
          .map((node) => ({
            id: node.id,
            kind: node.kind,
            filePath: node.filePath,
            layout: owningLayout(node),
          }))
          .sort((a, b) => a.id.localeCompare(b.id)),
      });
      return result;
    } else {
      target = scoped[0]!;
    }
  } else {
    result.context.notes.push('MISSING_TARGET_SELECTOR');
  }

  if (!target) return result;

  result.target = {
    id: target.id,
    name: target.name,
    kind: target.kind,
    filePath: target.filePath,
    layout: owningLayout(target),
  };

  // ---- 2. Collect the handlers bound to it ----------------------------
  interface ResolvedHandler {
    node: Node;
    context: BehaviorHandlerContext;
  }
  const handlers: ResolvedHandler[] = [];

  if (UI_NODE_KINDS.has(target.kind)) {
    const bindings = queries
      .getIncomingEdges(target.id, ['event-handler'])
      .sort((a, b) => a.source.localeCompare(b.source));
    for (const edge of bindings) {
      const handlerNode = queries.getNodeById(edge.source);
      if (!handlerNode) continue;
      const wiredBy = metadataString(edge, 'synthesizedBy');
      const scope: BehaviorBindingScope =
        metadataString(edge, 'scope') === 'form'
          ? 'form'
          : wiredBy === 'vba-expression-handler'
            ? 'expression'
            : 'control';
      handlers.push({
        node: handlerNode,
        context: {
          handler: handlerNode.name,
          handlerId: handlerNode.id,
          event: metadataString(edge, 'eventName'),
          scope,
          provenance: edge.provenance ?? 'unknown',
          location: locationOf(handlerNode.filePath, handlerNode.startLine),
          wiredBy,
        },
      });
    }
  } else if (PROCEDURE_KINDS.has(target.kind)) {
    // The request named the procedure itself. Report the binding it carries,
    // if any, instead of pretending it was reached through a control.
    const binding = queries
      .getOutgoingEdges(target.id, ['event-handler'])
      .sort((a, b) => a.target.localeCompare(b.target))[0];
    const wiredBy = binding ? metadataString(binding, 'synthesizedBy') : null;
    handlers.push({
      node: target,
      context: {
        handler: target.name,
        handlerId: target.id,
        event: binding ? metadataString(binding, 'eventName') : null,
        scope: binding
          ? metadataString(binding, 'scope') === 'form'
            ? 'form'
            : wiredBy === 'vba-expression-handler'
              ? 'expression'
              : 'control'
          : 'direct',
        provenance: binding?.provenance ?? 'unknown',
        location: locationOf(target.filePath, target.startLine),
        wiredBy,
      },
    });
  }

  if (handlers.length === 0) {
    result.context.notes.push('NO_HANDLER_BOUND');
    return result;
  }
  result.context.handlers = handlers.map((h) => h.context);

  // ---- 3. Walk each handler's call paths ------------------------------
  const calleesOf = (nodeId: string): Node[] => {
    const out: Node[] = [];
    const seen = new Set<string>();
    const edges = queries
      .getOutgoingEdges(nodeId, ['calls', 'references'])
      .sort((a, b) => a.target.localeCompare(b.target));
    for (const edge of edges) {
      if (seen.has(edge.target)) continue;
      const node = queries.getNodeById(edge.target);
      if (!node || !PROCEDURE_KINDS.has(node.kind)) continue;
      seen.add(edge.target);
      out.push(node);
    }
    return out;
  };

  // Stream paths: output deduplication must never permit exponential work.
  // One extra result's depth allows exact-fit requests to finish honestly.
  let remainingSteps = (maxResults + 1) * maxCallDepth;
  function* walk(chain: Node[]): Generator<Node[]> {
    if (remainingSteps === 0) {
      result.context.truncated.results = true;
      return;
    }
    remainingSteps--;
    const head = chain[chain.length - 1]!;
    if (chain.length >= maxCallDepth) {
      if (calleesOf(head.id).length > 0) result.context.truncated.callDepth = true;
      yield chain;
      return;
    }
    const callees = calleesOf(head.id).filter((node) => {
      if (chain.some((step) => step.id === node.id)) {
        result.context.truncated.cycle = true;
        return false;
      }
      return true;
    });
    if (callees.length === 0) {
      yield chain;
      return;
    }
    for (const callee of callees) {
      yield* walk([...chain, callee]);
      if (result.context.truncated.results) return;
    }
  }
  function* paths(): Generator<Node[]> {
    for (const handler of handlers) {
      yield* walk([handler.node]);
      if (result.context.truncated.results) return;
    }
  }

  // ---- 4. Attribute data references and effects to each procedure -----
  const dataCache = new Map<string, BehaviorDataEvidence[]>();

  const dataFor = (procedure: Node): BehaviorDataEvidence[] => {
    const cached = dataCache.get(procedure.id);
    if (cached) return cached;

    const collected: BehaviorDataEvidence[] = [];
    const record = (
      edge: Edge,
      node: Node,
      attributedBy: BehaviorDataEvidence['attributedBy'],
    ): void => {
      const access = edge.metadata?.access;
      const direction: BehaviorDataEvidence['access'] =
        access === 'read' || access === 'write' ? access : 'unknown';
      collected.push({
        procedure: procedure.name,
        name: node.name,
        targetKind: node.kind,
        access: direction,
        via: metadataString(edge, 'synthesizedBy'),
        attributedBy,
        location: locationOf(procedure.filePath, edge.line),
      });
      // A saved query is a stored fact about the tables it reads; follow it
      // one hop so a procedure that only names the query still reports them.
      if (node.kind === 'query') {
        for (const tableEdge of queries
          .getOutgoingEdges(node.id, ['references'])
          .sort((a, b) => a.target.localeCompare(b.target))) {
          const table = queries.getNodeById(tableEdge.target);
          if (!table || !DATA_TARGET_KINDS.has(table.kind) || table.kind === 'query') continue;
          collected.push({
            procedure: procedure.name,
            name: table.name,
            targetKind: table.kind,
            access: direction,
            via: metadataString(tableEdge, 'synthesizedBy'),
            throughQuery: node.name,
            attributedBy,
            location: locationOf(node.filePath, tableEdge.line),
          });
        }
      }
    };

    for (const edge of queries
      .getOutgoingEdges(procedure.id, ['references'])
      .sort((a, b) => a.target.localeCompare(b.target))) {
      const node = queries.getNodeById(edge.target);
      if (!node || !DATA_TARGET_KINDS.has(node.kind)) continue;
      if (!isDataReference(edge, node)) continue;
      record(edge, node, 'edge-source');
    }

    // Module-scoped data references, attributed by source line. See the file
    // header: `emitReference` always sources these from the module node.
    for (const container of queries.getIncomingEdges(procedure.id, ['contains'])) {
      const owner = queries.getNodeById(container.source);
      if (!owner || (owner.kind !== 'module' && owner.kind !== 'class')) continue;
      for (const edge of queries
        .getOutgoingEdges(owner.id, ['references'])
        .sort((a, b) => a.target.localeCompare(b.target))) {
        if (
          edge.line === undefined ||
          edge.line < procedure.startLine ||
          edge.line > procedure.endLine
        ) {
          continue;
        }
        const node = queries.getNodeById(edge.target);
        if (!node || !DATA_TARGET_KINDS.has(node.kind)) continue;
        if (!isDataReference(edge, node)) continue;
        record(edge, node, 'source-line');
      }
    }

    dataCache.set(procedure.id, collected);
    return collected;
  };

  const effectCache = new Map<string, string[]>();
  const objectEffectsFor = (procedure: Node): string[] => {
    const cached = effectCache.get(procedure.id);
    if (cached) return cached;
    const effects: string[] = [];
    for (const edge of queries
      .getOutgoingEdges(procedure.id, ['opens-form', 'opens-report', 'raises-event'])
      .sort((a, b) => a.target.localeCompare(b.target))) {
      const node = queries.getNodeById(edge.target);
      if (!node) continue;
      effects.push(`${edge.kind}:${node.name}`);
    }
    effectCache.set(procedure.id, effects);
    return effects;
  };

  // ---- 5. Build the compatibility payload -----------------------------
  const seenEntries = new Set<string>();
  const seenData = new Set<string>();

  const reported = new Map<string, Node>();
  for (const chain of paths()) {
    const chainData: BehaviorDataEvidence[] = [];
    const tables = new Set<string>();
    const effects = new Set<string>();

    for (const step of chain) {
      for (const item of dataFor(step)) {
        chainData.push(item);
        if (item.targetKind !== 'query') tables.add(item.name);
        effects.add(
          item.access === 'unknown'
            ? `data-access:${item.name}`
            : `${item.access}:${item.name}`,
        );
      }
      for (const effect of objectEffectsFor(step)) effects.add(effect);
    }

    const entry: CodeGraphBehaviorEvidence = {
      handler: chain[0]!.name,
      callPath: chain.map((step) => step.name),
      tables: [...tables].sort(),
      effects: [...effects].sort(),
    };
    const key = JSON.stringify(entry);
    if (seenEntries.has(key)) continue;
    seenEntries.add(key);
    if (result.evidence.length >= maxResults) {
      result.context.truncated.results = true;
      break;
    }
    result.evidence.push(entry);
    for (const step of chain) reported.set(step.id, step);
    for (const item of chainData) {
      const dataKey = `${item.procedure}|${item.name}|${item.access}|${item.attributedBy}|${item.location}`;
      if (!seenData.has(dataKey)) {
        seenData.add(dataKey);
        result.context.data.push(item);
      }
    }
  }

  // ---- 6. Unresolved references inside the reported procedures --------

  const byFile = new Map<string, Node[]>();
  for (const node of reported.values()) {
    if (!node.filePath) continue;
    const list = byFile.get(node.filePath) ?? [];
    list.push(node);
    byFile.set(node.filePath, list);
  }
  for (const [filePath, procedures] of [...byFile.entries()].sort((a, b) =>
    a[0].localeCompare(b[0]),
  )) {
    for (const ref of queries.getUnresolvedReferencesForFile(filePath)) {
      const owner = procedures.find(
        (proc) => ref.line >= proc.startLine && ref.line <= proc.endLine,
      );
      if (!owner) continue;
      result.context.unresolved.push({
        procedure: owner.name,
        referenceName: ref.referenceName,
        referenceKind: ref.referenceKind,
        location: locationOf(filePath, ref.line),
      });
    }
  }

  return result;
}
