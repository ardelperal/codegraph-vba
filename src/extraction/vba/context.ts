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
 * Issue #151: a qualified call site that the first walk could not resolve
 * because the receiver's declared type was not yet known — the receiver
 * was either (a) a `Dim`-declared variable on a LATER line in the same
 * procedure, or (b) a `Set x = New <Type>` / `Set x = <Factory>(...)`
 * assigned on a LATER line. The call is held until the procedure ends so
 * the full `localVarTypeMap` for that procedure scope is available, then
 * re-resolved and emitted (or dropped silently if it still cannot bind).
 */
export interface DeferredQualifiedCall {
  /** The enclosing procedure's `ProcInfo` (startLine is the map key). */
  caller: ProcInfo;
  /** Original receiver text from the call site (e.g. `x`). */
  receiver: string;
  /** Member identifier on the receiver (e.g. `Foo`). */
  member: string;
  /** Source line number (1-based, matches `Node.startLine` shape). */
  lineNum: number;
  /** Match column (0-based; 0 for statement-form, real index for paren). */
  column: number;
  /** Distinguishes the paren-form from the statement-form call shape. */
  kind: 'paren' | 'statement';
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
  /**
   * Issue #151: optional per-procedure end hook. Called once per procedure
   * that this classifier saw open, when the walk hits the matching
   * `End Sub` / `End Function` / `End Property` line. The default
   * `VbaWalker` in `vba-extractor.ts` invokes this BEFORE the line
   * itself is classified (so the proc-stack pop inside `classifyLine`
   * happens after the hook runs). This gives a classifier a chance to
   * drain per-procedure state (e.g. the deferred qualified-call list)
   * while the procedure is still on the stack and `localVarTypeMap` is
   * fully populated for that procedure scope.
   */
  onProcedureEnd?(procStartLine: number, ctx: VbaExtractorContext): void;
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
   *
   * Issue #151: each entry now carries an OPTIONAL `procScope` (the procedure
   * `startLine` as a string) for proc-local declarations. The call sweep
   * deletes those entries at the matching `End Sub` / `End Function` /
   * `End Property` so a `Dim x As MyClassA` inside `Sub First()` does NOT
   * leak into `Sub Second()`'s view of `x`. Module-level declarations
   * (`Dim x As Y` at file top level) keep `procScope === 'module'` and
   * persist across procedures, mirroring VBA's actual scoping rules.
   */
  public localVarTypeMap = new Map<string, {
    outer: string;
    qualified: boolean;
    withEvents?: boolean;
    variableName?: string;
    assignedWithSet?: boolean;
    procScope?: 'module' | string;
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

  /** Local event name (lowercase) → event node for `RaiseEvent` edge emission. */
  public localEvents = new Map<string, Node>();

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
   * Issue #151: per-procedure list of qualified calls the first walk could
   * not resolve because the receiver's declared type was unknown at the
   * call line (forward `Dim` / `Set` / `Set = Factory()` reference). Keyed
   * by the procedure's `startLine` (string) so the call sweep can pop and
   * reprocess the list exactly once at the matching `End Sub` /
   * `End Function` / `End Property` — no cross-procedure leaking.
   *
   * The deferred list is consumed by the calls-and-sql classifier's
   * `onProcedureEnd` hook, which sees the procedure's `procStartLine`
   * (numeric) as the key. The string form is used here because the rest
   * of the per-procedure scoping (`localConstants`) already uses string
   * keys, keeping the persistence shape consistent.
   */
  public deferredQualifiedCalls: Map<string, DeferredQualifiedCall[]> = new Map();

  constructor(filePath: string) {
    this.filePath = filePath;
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
   */
  public emitReference(
    targetName: string,
    lineNum: number,
    column: number,
    synthesizedBy: string,
    access?: 'read' | 'write',
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
    const edge: Edge = {
      source: '', // rewritten after module node exists
      target: targetId,
      kind: 'references',
      provenance: 'heuristic',
      metadata: access ? { synthesizedBy, access } : { synthesizedBy },
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

  /**
   * Issue #151: push a deferred qualified call onto the per-procedure
   * list. `scopeKey` is the procedure's `startLine` as a string
   * (matching `currentProcKey` shape — never the literal `'module'`
   * string, because module-level code has no procedure to defer until).
   * The bucket is created lazily and returned.
   */
  public pushDeferredQualifiedCall(
    scopeKey: string,
    entry: DeferredQualifiedCall,
  ): DeferredQualifiedCall[] {
    let bucket = this.deferredQualifiedCalls.get(scopeKey);
    if (!bucket) {
      bucket = [];
      this.deferredQualifiedCalls.set(scopeKey, bucket);
    }
    bucket.push(entry);
    return bucket;
  }

  /**
   * Issue #151: drain the deferred list for one procedure, returning
   * the calls in registration order. The bucket is deleted from the
   * map so a re-walk of the same procedure scope never replays them.
   * Returns `[]` when the procedure has no deferred calls.
   */
  public drainDeferredQualifiedCalls(
    scopeKey: string,
  ): DeferredQualifiedCall[] {
    const bucket = this.deferredQualifiedCalls.get(scopeKey);
    if (!bucket) return [];
    this.deferredQualifiedCalls.delete(scopeKey);
    return bucket;
  }

  /**
   * Issue #151: clear the procedure-scoped entries from
   * `localVarTypeMap` at `End Sub` / `End Function` / `End Property`.
   * A `Dim x As MyClassA` inside `Sub First()` does NOT leak into
   * `Sub Second()`'s call-site resolution — VBA's actual scoping
   * rule, and the explicit acceptance criterion from the issue
   * ("no leak into the next procedure"). Module-level entries
   * (`procScope === 'module'`) are kept. Returns the number of
   * entries cleared (useful for tests + future metrics).
   */
  public clearLocalVarsInScope(scopeKey: string): number {
    let cleared = 0;
    for (const [key, entry] of this.localVarTypeMap) {
      if (entry.procScope === scopeKey) {
        this.localVarTypeMap.delete(key);
        cleared++;
      }
    }
    return cleared;
  }
}
