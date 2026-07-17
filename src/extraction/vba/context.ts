/**
 * VbaExtractorContext — the shared state + shared helpers threaded through
 * every VBA extraction pass.
 *
 * It owns the node/edge/reference accumulators, the per-file lookup maps
 * (`localProcs`, `functionNodeByName`, `localVarTypeMap`, `localConstants`,
 * `localEvents`), the lazily-created module/class node, and the de-dup caches.
 * The per-concern sweeps under `src/extraction/vba/*` are free functions that
 * take a `VbaExtractorContext` and read/write this shared state; the handful
 * of helpers that more than one sweep needs (`emitReference`,
 * `pushContainsFromModule`, `findOrCreateFunctionNodeId`, the qualified-call
 * gates, and the per-scope Const lookup) live here as methods.
 */
import {
  Node,
  Edge,
  ExtractionError,
  UnresolvedReference,
} from '../../types';
import { generateNodeId } from '../tree-sitter-helpers';
import { PRIMITIVE_TYPES } from './constants';

export interface ProcInfo {
  name: string;
  qualifiedName: string;
  kind: 'sub' | 'function' | 'property';
  visibility: 'public' | 'private' | 'protected' | 'internal';
  startLine: number;
}

/**
 * Issue #83: the per-concern classifier shape. The single walker in
 * `vba-extractor.ts` calls `classifyLine(line, index, ctx)` once per
 * pre-split line, in a stable order across all six concerns. No internal
 * `.split('\n')` — the split happens once in the walker.
 */
export interface VbaClassifier {
  /** Stable, human-readable name for logs/metrics. */
  readonly name: string;
  /**
   * Number of top-level symbols this classifier emitted. Read by the
   * orchestrator to compute `hasAnySymbols` (skip module-node creation
   * for files with NO symbols per REQ-CODE-10).
   */
  count: number;
  /** Classify one pre-split line. May mutate `ctx` and bump `count`. */
  classifyLine(line: string, index: number, ctx: VbaExtractorContext): void;
  /**
   * Optional end-of-file hook. Used by the calls-and-sql classifier to
   * flush `procEndLines` → function node `endLine`.
   */
  finalize?(ctx: VbaExtractorContext): void;
}

export class VbaExtractorContext {
  public filePath: string;
  public nodes: Node[] = [];
  public edges: Edge[] = [];
  public errors: ExtractionError[] = [];
  public unresolvedReferences: UnresolvedReference[] = [];
  public moduleOrClassNode: Node | null = null;

  /**
   * Class-name prefix for `qualifiedName` composition.
   *
   * B3 (hueco 5): for `.cls` files every function node's `qualifiedName`
   * is composed as `${className}.${procName}` (e.g. `Form_TestForm.Form_Load`)
   * so cross-class callers can disambiguate which form each `Form_Load`
   * belongs to. For `.bas` files this is `null` — module-scoped procs keep
   * their bare-name `qualifiedName`. Resolved once per `extract()` from
   * `isCls` + `vbName`; `sweepProcedures` reads it.
   */
  public classNamePrefix: string | null = null;

  /**
   * Map of procedure name (within this file) → list of ProcInfo. The list
   * (rather than a single value) is needed because VBA allows multiple
   * Property accessors with the same name: `Property Get Foo`, `Property
   * Let Foo`, and `Property Set Foo` all share the name `Foo`.
   */
  public localProcs = new Map<string, ProcInfo[]>();

  /**
   * Issue #83: flat list of every `ProcInfo` produced by the procedures
   * classifier, in declaration order. Mirrors the array the legacy
   * `sweepProcedures` returned; the orchestrator reads it after the single
   * walker pass to (a) decide `hasAnySymbols`, and (b) check for a
   * `New` initializer to set `hasClassInitializer` on the class node.
   */
  public procedures: ProcInfo[] = [];

  /**
   * Cache: procedure name → first matching function node emitted for that
   * name (audit S2 — O(1) same-file call resolution).
   */
  public functionNodeByName = new Map<string, Node>();

  /**
   * Cache: startLine → function node emitted for that line. When Property
   * Get/Let/Set share a name, we find the SPECIFIC accessor's node by its
   * declaration line (Fix 1).
   */
  public functionNodeByStartLine = new Map<number, Node>();

  /** Edges whose source needs to be set to the module/class id once it exists. */
  public pendingModuleOrClassSource: Edge[] = [];

  /** Cache so we don't re-emit the same proc function node per call site. */
  public procNodeIdCache = new Map<string, string>();
  public callDedupe = new Set<string>();
  public synthFunctionNodeIds = new Set<string>();

  /**
   * B4 (hueco 6) extended by Issue #48: cache of stub node ids we've already
   * emitted for a given (method, target name) pair in this file. Keyed by
   * `${cacheKey}:${lowerName}` so the OpenForm and OpenReport de-dup buckets
   * stay disjoint.
   */
  public opensStubIdsByKey = new Map<string, string>();

  public synthClassNodeIds = new Set<string>();

  /**
   * Issue #50: TempVars placeholder-node de-dup cache. Keys are the
   * deterministic node ids produced by `emitTempVarReference` — those ids
   * intentionally use a synthetic `synthetic:tempvar/<key>` path so the SAME
   * key referenced from Form_A.cls AND Form_B.cls collapses to ONE node.
   */
  public synthTempVarNodeIds = new Set<string>();

  /**
   * Maps `variableName.toLowerCase()` → declared type info.
   * Built by `sweepDimsAndWithEvents`; consulted by the unified qualified-call
   * gate (`shouldProcessQualifiedCall`) so declared project-class locals emit
   * edges, declared primitive/external locals stay silent, and undeclared
   * receivers remain module-name candidates for the resolver (Fix 2 / Issue #2).
   */
  public localVarTypeMap = new Map<string, {
    outer: string;
    qualified: boolean;
    withEvents?: boolean;
    variableName?: string;
    assignedWithSet?: boolean;
  }>();

  /**
   * Issue #52: Const resolution buckets, scoped per procedure. Key is
   * `'module'` for module-level Consts, or the procedure's `startLine`
   * (stringified) for proc-local Consts. Each bucket maps the lowercase
   * constant name to its simple-literal value (used by `DoCmd.OpenForm` /
   * `OpenReport` / `OpenQuery` argument resolution via `resolveLocalConst`).
   */
  public localConstants: Map<'module' | string, Map<string, string>> = new Map();

  /**
   * Issue #52: the current Const-lookup scope. `'module'` when no procedure
   * is open, otherwise the top-of-stack proc's `startLine` as a string. Both
   * `sweepEnumsAndConsts` and `sweepCallsAndSql` keep this in sync with their
   * per-line stack walk.
   */
  public currentProcKey: 'module' | string = 'module';

  /**
   * Issue #52: per-extraction proc-stack shared between `sweepEnumsAndConsts`
   * and `sweepCallsAndSql`. Each sweep clears it at the start so the file's
   * mid-proc structural state never leaks across sweeps. Holds the `startLine`
   * of every procedure whose `End` marker the sweep has not yet seen.
   */
  public procStack: number[] = [];

  /**
   * Issue #153: per-extraction state for the events/types/declares
   * classifier's `Type <Name> ... End Type` block tracking. The
   * `currentTypeBlock` is set when a `type-start` rule fires and
   * cleared when the matching `type-end` rule fires. Lives on `ctx`
   * (instead of the factory's closure) so the declarative RULES
   * table's `emit` functions can read/write it without taking a
   * closure-captured mutable reference. `null` means "outside any
   * type block right now". Reset implicitly per-extraction (a fresh
   * `VbaExtractorContext` is constructed every `extract()` call).
   */
  public vbaDeclTypeBlock: { id: string; name: string } | null = null;

  /**
   * Issue #153: per-extraction state for the enum/consts classifier's
   * `Enum <Name> ... End Enum` block tracking. Same shape and
   * rationale as `vbaDeclTypeBlock`: lives on `ctx` so the
   * declarative RULES table's `emit` functions can read/write it
   * without taking a closure-captured mutable reference.
   * `null` means "outside any enum block right now".
   */
  public vbaEnumBlock: { id: string; name: string } | null = null;

  /**
   * Issue #153: per-extraction state for the calls/SQL classifier's
   * `With <receiver> ... End With` block tracking. The with-receiver
   * stack mirrors the textual nesting of `With` blocks inside a
   * procedure body — the top of the stack is the active receiver
   * for `.Member` and `.Member(args)` statement-form calls. Reset
   * implicitly per-extraction (a fresh `VbaExtractorContext` is
   * constructed every `extract()` call). Replaces the per-factory
   * closure variable that lived in `sweepCallsAndSql` before the
   * Issue #153 RULES-table refactor.
   */
  public vbaWithStack: string[] = [];

  /** Local event name (lowercase) → event node for `RaiseEvent` edge emission. */
  public localEvents = new Map<string, Node>();

  /**
   * Issue #152: per-event counter for `RaiseEvent <EventName>` sites. Keyed
   * by event node id (so two different events that share a name across
   * files don't collide — the gate is per-file). The count is incremented
   * in `scanRaiseEvents` BEFORE the edge is pushed, so the count equals
   * the number of raise sites regardless of the gate decision downstream.
   * The orchestrator's fanout-gate pass reads this map to decide which
   * event nodes to flag `metadata.highFanout: true` and which `raises-event`
   * edges to drop.
   */
  public raiseEventCounts = new Map<string, number>();

  /**
   * Same-file function/property return types, keyed by lowercase proc name →
   * declared return type. ONLY non-primitive (project-class) return types are
   * stored. Populated by `sweepProcedures` (which runs before the call sweep);
   * consumed by the call sweep's `Set x = Factory(...)` handling to type the
   * assigned local var so a later `x.Method` qualified call resolves to the
   * factory's class instead of a dead-end `x.Method` stub. Cross-file
   * factories are not covered here (no return type is visible at extraction
   * time) — that stays the resolver's frontier.
   */
  public functionReturnTypes = new Map<string, string>();

  /**
   * Issue #156: per-stage timings, lazily allocated ONLY when
   * `CODEGRAPH_VBA_TIMING=1|2` is set. The default path (env var unset)
   * keeps this `null` so `recordStage` is a single null-check and we
   * pay zero Map-allocation cost in the hot path. Stage names are
   * dot-separated: `preprocess.stripVbaComments`,
   * `preprocess.cc.lexer+parser`, `walk.procedures`, `walk.main`,
   * `classifier.<name>`. Counters are in milliseconds (float).
   */
  public timings: Map<string, number> | null = null;

  /**
   * Issue #156: how many times each classifier's `classifyLine` was
   * invoked. Only allocated together with `timings` so the default path
   * pays no cost.
   */
  public classifierInvokeCounts: Map<string, number> | null = null;

  constructor(filePath: string) {
    this.filePath = filePath;
  }

  /**
   * Issue #156: ensure the timings Maps exist. Called once per
   * `extract()` by the orchestrator right after it reads the env var.
   * No-op when already allocated.
   */
  public ensureTimings(): void {
    if (this.timings === null) {
      this.timings = new Map<string, number>();
      this.classifierInvokeCounts = new Map<string, number>();
    }
  }

  /**
   * Issue #156: add `ms` to the stage named `name`. No-op when
   * `timings` is null (default path).
   */
  public recordStage(name: string, ms: number): void {
    if (this.timings === null) return;
    this.timings.set(name, (this.timings.get(name) ?? 0) + ms);
  }

  /**
   * Emit a `contains` edge from the (lazily-created) module/class node to
   * `targetId`. Mirrors the pending-source pattern the procedure and
   * implements sweeps use: the module node doesn't exist yet, so the edge's
   * source is rewritten once `extract()` creates it.
   */
  public pushContainsFromModule(targetId: string): void {
    if (this.moduleOrClassNode) {
      this.edges.push({
        source: this.moduleOrClassNode.id,
        target: targetId,
        kind: 'contains',
      });
      return;
    }
    const edge: Edge = { source: '', target: targetId, kind: 'contains' };
    this.edges.push(edge);
    this.pendingModuleOrClassSource.push(edge);
  }

  /**
   * Fix 2 (Issue #2): return true iff `receiverName` is a file-local variable
   * (appears in `localVarTypeMap`) whose declared type is a SIMPLE (non-
   * qualified, non-primitive) identifier — a candidate project-defined class.
   * Qualified types (e.g. `DAO.Recordset`) and primitives (`String`, `Long`)
   * return false so runtime/DAO calls are suppressed. Brackets are stripped
   * from the lookup key defensively (Issue #54).
   */
  public isLocalProjectClassVar(receiverName: string): boolean {
    // Issue #54: strip a single leading `[` and/or trailing `]` if present.
    const key = receiverName.replace(/^\[|\]$/g, '').toLowerCase();
    const entry = this.localVarTypeMap.get(key);
    if (!entry) return false; // not declared in this file → silent
    if (entry.qualified) return false; // DAO.Recordset etc. → silent
    if (PRIMITIVE_TYPES.has(entry.outer.toLowerCase())) return false;
    return true;
  }

  /**
   * Qualified-call eligibility is intentionally shared by paren-form
   * (`Receiver.Member(...)`) and statement-form (`Receiver.Member args`) scans:
   * project-class locals are processed after type resolution, declared
   * primitive/external locals are silent, and undeclared receivers remain
   * candidate module names.
   */
  public shouldProcessQualifiedCall(receiverName: string): boolean {
    if (this.isLocalProjectClassVar(receiverName)) return true;
    return !this.localVarTypeMap.has(receiverName.toLowerCase());
  }

  /**
   * #12a: resolve the "receiver type" used to build a qualified call-stub's
   * name/qualifiedName. When `receiverName` is a file-local variable typed
   * as a candidate project class (`isLocalProjectClassVar`), returns the
   * RESOLVED CLASS NAME from `localVarTypeMap` (e.g. `m_NCOp As NCOperaciones`
   * → `'NCOperaciones'`) so the stub matches the real `.cls` method's
   * `${className}.${proc}` shape. Otherwise returns `receiverName` unchanged
   * (the `.bas`-qualified module call case).
   */
  public resolveReceiverType(receiverName: string): string {
    if (this.isLocalProjectClassVar(receiverName)) {
      const key = receiverName.replace(/^\[|\]$/g, '').toLowerCase();
      const entry = this.localVarTypeMap.get(key);
      if (entry) return entry.outer;
    }
    return receiverName;
  }

  public findOrCreateFunctionNodeId(proc: ProcInfo): string {
    // Fix 1: key the cache by `name:startLine` so Property Get/Let/Set
    // accessors with the same name each resolve to their own node.
    const cacheKey = `${proc.name}:${proc.startLine}`;
    const cached = this.procNodeIdCache.get(cacheKey);
    if (cached) return cached;
    // Prefer the startLine-indexed lookup (O(1), correct for all three
    // Property accessors). Fall back to the name-indexed cache for procs
    // whose startLine wasn't recorded (should not happen in practice).
    const fn =
      this.functionNodeByStartLine.get(proc.startLine) ??
      this.findFunctionNodeByName(proc.name);
    if (fn) {
      this.procNodeIdCache.set(cacheKey, fn.id);
      return fn.id;
    }
    // Fallback: synthesize a node id matching generateNodeId's input shape.
    const id = generateNodeId(this.filePath, 'function', proc.name, proc.startLine);
    this.procNodeIdCache.set(cacheKey, id);
    return id;
  }

  public findFunctionNodeByName(name: string): Node | undefined {
    // O(1) via the cache populated as function nodes are emitted in
    // `sweepProcedures` (audit S2 — replaced an O(n) linear scan per call site).
    return this.functionNodeByName.get(name);
  }

  /**
   * Issue #152: apply the per-file `RaiseEvent` fanout gate. For every
   * event whose `raiseEventCounts` value exceeds `maxFanout`:
   *   1. Stamp the event node's `metadata.highFanout = true` and
   *      `metadata.raiseCount = <count>` so consumers (codegraph_explore,
   *      vba-event-tracer) can recognize a high-fanout event.
   *   2. Drop every `raises-event` edge targeting that event id from
   *      `edges`. The event node itself stays — declaration site, handler
   *      linkage via `subscribes-event`, and metadata are all preserved.
   *
   * Events with `<= maxFanout` raise sites are untouched (no flag, edges
   * stay). When `maxFanout` is undefined, the gate is disabled and the
   * method is a no-op. Mirrors the upstream `EVENT_FANOUT_CAP` discipline
   * in `src/resolution/callback-synthesizer.ts:43` (this side suppresses
   * the producer side, the upstream side suppresses the synthesized
   * cross-file dispatcher/handler edges).
   *
   * Returns the number of `raises-event` edges dropped (zero when the
   * gate is disabled or no event exceeds the cap — useful for tests).
   */
  public applyRaiseFanoutGate(maxFanout: number | undefined): number {
    if (maxFanout === undefined) return 0;
    if (maxFanout < 0) return 0;

    // 1. Build the set of event node ids that exceed the cap.
    const overThreshold = new Set<string>();
    for (const [eventId, count] of this.raiseEventCounts) {
      if (count > maxFanout) overThreshold.add(eventId);
    }
    if (overThreshold.size === 0) return 0;

    // 2. Stamp the metadata on the event nodes.
    for (const eventId of overThreshold) {
      const eventNode = this.nodes.find(
        (n) => n.id === eventId && n.kind === 'event',
      );
      if (!eventNode) continue;
      const md = (eventNode.metadata ?? {}) as Record<string, unknown>;
      md.highFanout = true;
      md.raiseCount = this.raiseEventCounts.get(eventId) ?? 0;
      eventNode.metadata = md;
    }

    // 3. Drop every `raises-event` edge targeting an over-threshold event.
    // Re-build the edges array so the dropped edges vanish without
    // disturbing the other edges' order. `edges.length` is the only
    // property kept; this matches how every other sweep mutates `edges`
    // (push-only, never splice).
    const before = this.edges.length;
    const kept: Edge[] = [];
    for (const e of this.edges) {
      if (e.kind === 'raises-event' && overThreshold.has(e.target)) continue;
      kept.push(e);
    }
    this.edges.length = 0;
    for (const e of kept) this.edges.push(e);
    return before - kept.length;
  }

  /**
   * Emit a `references` edge from the file's module/class node to a synthetic
   * node named `targetName`. Used by Dim, WithEvents, Set-New, and SQL sweeps.
   * Fix 5: the synthetic node id is keyed on (filePath, kind, name) WITHOUT
   * lineNum so the same type/table referenced on N lines produces ONE node.
   *
   * `access` (optional): when a caller can classify the reference as a data
   * read or write — SQL table references derive it from the statement verb —
   * it is stamped onto `edge.metadata.access` so consumers can answer "who
   * WRITES table X" vs "who READS table X". Mirrors the read/write tagging the
   * TempVars sweep already emits (`emitTempVarReference`). Omitted for
   * structural references (Dim/WithEvents/Set-New) where the direction is not
   * meaningful.
   *
   * `extras` (optional): additional metadata fields to merge into the
   * emitted edge's `metadata` object. Used by the WithEvents sweep to
   * stamp `variableName` (issue #150) so the post-extraction
   * event-handler synthesis pass can find the `m_<var>_<event>`
   * handler Sub even after `resolveVbaReferenceStubs` CASCADE-deletes
   * the companion `subscribes-event` edge along with the synthetic
   * class stub. Fields here are merged AFTER the standard
   * `{ synthesizedBy [, access] }` shape, so callers can override
   * (rarely useful) but not accidentally drop the stamp.
   */
  public emitReference(
    targetName: string,
    lineNum: number,
    column: number,
    synthesizedBy: string,
    access?: 'read' | 'write',
    extras?: Record<string, unknown>,
  ): void {
    if (!targetName) return;
    const targetId = generateNodeId(
      this.filePath,
      'class', // placeholder kind; cross-file resolution will re-type at lookup
      targetName,
      0,        // stable — line-independent
    );
    if (this.synthClassNodeIds.has(targetId)) {
      // Node already emitted for this name — only add the edge below.
    } else {
      this.synthClassNodeIds.add(targetId);
      this.nodes.push({
        id: targetId,
        kind: 'class',
        name: targetName,
        qualifiedName: targetName,
        filePath: this.filePath,
        language: 'vba',
        startLine: lineNum,
        endLine: lineNum,
        startColumn: column,
        endColumn: column + targetName.length,
        updatedAt: Date.now(),
      });
    }
    const metadata: Record<string, unknown> = access
      ? { synthesizedBy, access, ...(extras ?? {}) }
      : { synthesizedBy, ...(extras ?? {}) };
    const edge: Edge = {
      source: '', // rewritten after module node exists
      target: targetId,
      kind: 'references',
      provenance: 'heuristic',
      metadata,
      line: lineNum,
      column,
    };
    this.edges.push(edge);
    this.pendingModuleOrClassSource.push(edge);
  }

  /**
   * Issue #52: shared lookup helper for `scanDoCmdOpenCalls` and
   * `scanDoCmdOpenQuery`. The current scope is the procedure whose
   * `startLine` is on top of `procStack` (or `'module'` when empty).
   * Per-proc bucket first; module bucket is the fallback.
   */
  public resolveLocalConst(name: string): string | undefined {
    const lower = name.toLowerCase();
    const procBucket = this.localConstants.get(this.currentProcKey);
    if (procBucket) {
      const v = procBucket.get(lower);
      if (v !== undefined) return v;
    }
    const moduleBucket = this.localConstants.get('module');
    return moduleBucket?.get(lower);
  }

  /**
   * Issue #52: shared writer. `scopeKey` is `'module'` or the procedure's
   * startLine-as-string. Creates the bucket lazily. Returns the bucket the
   * value was written to (mostly useful for tests).
   */
  public setLocalConstInScope(
    scopeKey: 'module' | string,
    name: string,
    value: string,
  ): Map<string, string> {
    let bucket = this.localConstants.get(scopeKey);
    if (!bucket) {
      bucket = new Map<string, string>();
      this.localConstants.set(scopeKey, bucket);
    }
    bucket.set(name.toLowerCase(), value);
    return bucket;
  }
}
