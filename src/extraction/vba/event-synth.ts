/**
 * VBA event-handler synthesis pass (issue #150).
 *
 * Problem: `vba-event-tracer` (skill shipped in #33) traces an event end-to-
 * end at QUERY TIME by walking the graph — declaration site → `raises-event`
 * edges → `subscribes-event` edges → handler Subs named by the WithEvents
 * variable convention (`<varName>_<EventName>`). Every `codegraph_explore`
 * call repeats this work because the GRAPH doesn't pre-compute the
 * connection. The graph has the data; what's missing is the connection.
 *
 * Solution: a two-phase post-extraction pass that materializes one
 * `event-handler` edge (reusing the existing kind) from every
 * `RaiseEvent <EventName>` site to its `m_<var>_<EventName>` handler Sub,
 * when resolvable by the WithEvents naming convention. Edge provenance:
 * `heuristic`, `metadata.synthesizedBy: 'vba-event-handler'`,
 * `metadata.eventName: '<EventName>'`, `metadata.variableName: '<var>'`.
 *
 * Phase 1 — collect the pairs (per `vba-withevents` `references` edge):
 *   For every `WithEvents m_X As ClassName` in a `.cls` (the WithEvents
 *   sweep emits a `references` edge with `synthesizedBy: 'vba-withevents'`
 *   and `metadata.variableName: 'm_X'` after the #150 change to
 *   `VbaExtractorContext.emitReference`):
 *     1. Resolve `ClassName` to the real class node by name. The
 *        synthetic class stub created by the extractor (and CASCADE-
 *        deleted by `resolveVbaReferenceStubs`) is irrelevant — we use
 *        the repointed `references` edge's target id, which is the real
 *        class node after the resolver runs.
 *     2. For every `event` node in the real class's file: locate the
 *        handler Sub in the SUBSCRIBER file (the edge's source — the
 *        subscriber's module/class node) named `m_<var>_<EventName>`.
 *     3. If found, register a pending pair keyed by `eventNodeId` →
 *        { handlerNode, variableName }.
 *
 * Phase 2 — connect at RaiseEvent sites (whole project):
 *   For every `raises-event` edge in the DB:
 *     1. Locate the event node (edge.target) and read `eventName` from
 *        the edge metadata.
 *     2. Look up the handler Sub from the pairs registered for that
 *        event node.
 *     3. Emit an `event-handler` edge from the raiser Sub to the handler
 *        Sub with the metadata described above.
 *
 * Failure modes (silent, no error — per the issue spec):
 *   - WithEvents variable not found in any subscriber
 *     (no `references` edge with `synthesizedBy: 'vba-withevents'`).
 *   - Handler Sub `m_<var>_<EventName>` not declared.
 *   - Event declared in a class that has no `WithEvents` binding.
 *   - The `references` edge's target no longer points at a real class
 *     node (resolver couldn't resolve it — e.g. name collision with a
 *     different kind).
 *
 * Idempotent: re-running on a fully-resolved project is a no-op because
 * `edgeExists` short-circuits any duplicate `(source,target,
 * 'event-handler')` tuple. The unique index on
 * `(source, target, kind, IFNULL(line,-1), IFNULL(col,-1))` also
 * backstops this with `INSERT OR IGNORE` semantics in `insertEdge`.
 *
 * Wired into `src/index.ts` next to `resolveVbaReferenceStubs` (it does
 * NOT need to run BEFORE that pass — the `variableName` lives on the
 * `references` edge metadata, which the resolver preserves across its
 * repoint).
 */
import { Edge, Node, NodeKind } from '../../types';
import { QueryBuilder } from '../../db/queries';

/**
 * Pending pair for Phase 1: a `WithEvents m_X As ClassName` resolved to
 * a specific event + a specific handler Sub in the subscriber file.
 * Keyed by `eventNode.id` so Phase 2 can look it up in O(1) per
 * `raises-event` edge.
 */
interface EventHandlerPair {
  handlerNode: Node;
  variableName: string;
}

export function synthesizeVbaEventHandlerEdges(queries: QueryBuilder): number {
  // --- Phase 1: collect pairs ---------------------------------------------
  // Walk every `vba-withevents` `references` edge. The SOURCE is the
  // subscriber's module/class node (its `filePath` scopes the handler
  // Sub lookup); the TARGET is the REAL event-declaring class node
  // (after `resolveVbaReferenceStubs` repointed it from the synthetic
  // class stub). The metadata carries the WithEvents variable name
  // (`m_X`) — issue #150 change to `VbaExtractorContext.emitReference`.
  const bindings = queryWithEventsBindings(queries);
  if (bindings.length === 0) return 0;

  // key = eventNodeId (string), value = the pair. We keep the FIRST
  // binding per event node — multiple `WithEvents` vars on the same
  // class binding to the same event from the same file is rare; if it
  // ever happens the second binding's variable name is silently
  // shadowed by the first, which is fine because we only emit one
  // `event-handler` edge per raiser anyway.
  const pairsByEventId = new Map<string, EventHandlerPair>();

  for (const binding of bindings) {
    if (!binding.variableName) continue;
    // Find every event node in the real event-declaring class's file.
    // iterateNodesByKind streams the events without materializing the
    // whole table (cheap on a few hundred-event project, O(1) memory
    // for million-event projects).
    for (const eventNode of queries.iterateNodesByKind('event' as NodeKind)) {
      if (eventNode.language !== 'vba') continue;
      if (eventNode.filePath !== binding.eventClassFilePath) continue;
      // Handler Sub name follows the WithEvents naming convention:
      // `m_<variableName>_<eventName>` (e.g. `m_Notifier_DataChanged`).
      const handlerName = `${binding.variableName}_${eventNode.name}`;
      const handlerSub = findHandlerSubInFile(
        queries,
        binding.subscriberFilePath,
        handlerName,
      );
      if (!handlerSub) continue;
      if (pairsByEventId.has(eventNode.id)) continue;
      pairsByEventId.set(eventNode.id, {
        handlerNode: handlerSub,
        variableName: binding.variableName,
      });
    }
  }

  if (pairsByEventId.size === 0) return 0;

  // --- Phase 2: connect at RaiseEvent sites --------------------------------
  const raiseRows = queries.getRaisesEventEdges();
  const seenTuples = new Set<string>();
  const newEdges: Edge[] = [];
  for (const re of raiseRows) {
    const pair = pairsByEventId.get(re.target);
    if (!pair) continue;
    const tupleKey = `${re.source}\0${pair.handlerNode.id}`;
    if (seenTuples.has(tupleKey)) continue;
    if (queries.edgeExists(re.source, pair.handlerNode.id, 'event-handler')) {
      seenTuples.add(tupleKey);
      continue;
    }
    seenTuples.add(tupleKey);
    const meta = parseMetadata(re.metadata);
    const eventName =
      typeof meta?.eventName === 'string' ? (meta.eventName as string) : '';
    newEdges.push({
      source: re.source,
      target: pair.handlerNode.id,
      kind: 'event-handler',
      provenance: 'heuristic',
      metadata: {
        synthesizedBy: 'vba-event-handler',
        eventName,
        variableName: pair.variableName,
      },
      line: re.line ?? undefined,
      column: re.col ?? undefined,
    });
  }

  if (newEdges.length > 0) {
    queries.insertEdges(newEdges);
  }
  return newEdges.length;
}

/**
 * Discover every WithEvents binding in the project by walking the
 * `vba-withevents` `references` edges. Each binding carries:
 *   - `subscriberFilePath`  → where the handler Sub must live
 *   - `eventClassFilePath`  → where the event node lives
 *   - `variableName`        → the `m_X` part of `m_X_<EventName>`
 *
 * Uses one SQL query (the new `getReferencesBySynthesizedBy` helper)
 * + per-edge `getNodeById` lookups so the cost is O(bindings)
 * regardless of total node count. Cache-friendly: the QueryBuilder's
 * LRU node cache absorbs repeated `getNodeById` calls.
 */
function queryWithEventsBindings(
  queries: QueryBuilder,
): Array<{
  subscriberFilePath: string;
  eventClassFilePath: string;
  variableName: string;
}> {
  const rows = queries.getReferencesBySynthesizedBy('vba-withevents');
  const out: Array<{
    subscriberFilePath: string;
    eventClassFilePath: string;
    variableName: string;
  }> = [];

  for (const row of rows) {
    const meta = parseMetadata(row.metadata);
    if (meta?.synthesizedBy !== 'vba-withevents') continue;
    const variableName =
      typeof meta?.variableName === 'string' ? (meta.variableName as string) : '';
    if (!variableName) continue;

    // Subscriber file: the source is the subscriber's class node.
    const subscriberNode = queries.getNodeById(row.source);
    if (!subscriberNode || !subscriberNode.filePath) continue;

    // Event-declaring class file: the target is the real class node
    // (after `resolveVbaReferenceStubs` repointed it from the synthetic
    // class stub). We accept `class` / `module` / `interface` / `enum`
    // / `struct` / `type_alias` — whatever kind the resolver landed on.
    const eventClassNode = queries.getNodeById(row.target);
    if (!eventClassNode || !eventClassNode.filePath) continue;
    if (eventClassNode.language !== 'vba') continue;
    if (eventClassNode.filePath === subscriberNode.filePath) continue;

    out.push({
      subscriberFilePath: subscriberNode.filePath,
      eventClassFilePath: eventClassNode.filePath,
      variableName,
    });
  }
  return out;
}

/**
 * Find a function (Sub) node in `filePath` whose name matches
 * `handlerName` (case-insensitive, since VBA identifiers are case-
 * insensitive). Uses a narrow SQL query (the QueryBuilder doesn't
 * expose a public name+file+kind lookup, so we go through
 * `getNodesByFile` and filter — filePath-keyed lookups are fast on
 * the `idx_nodes_file_path` index).
 */
function findHandlerSubInFile(
  queries: QueryBuilder,
  filePath: string,
  handlerName: string,
): Node | null {
  const lower = handlerName.toLowerCase();
  // getNodesByFile returns nodes ordered by start_line; the order
  // doesn't matter for the existence check, only the (name, filePath,
  // kind='function') match.
  const candidates = queries.getNodesByFile(filePath);
  for (const node of candidates) {
    if (node.kind !== 'function') continue;
    if (node.language !== 'vba') continue;
    if (node.name.toLowerCase() !== lower) continue;
    return node;
  }
  return null;
}

/**
 * Safe-parse the metadata JSON column. The schema stores it as TEXT;
 * an absent or malformed cell returns `null` (treated as "no metadata"
 * by callers). Mirrors the same shape used elsewhere in the codebase.
 */
function parseMetadata(
  raw: string | null | undefined,
): Record<string, unknown> | null {
  if (!raw) return null;
  try {
    const parsed = JSON.parse(raw);
    if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
      return parsed as Record<string, unknown>;
    }
  } catch {
    /* fall through */
  }
  return null;
}
