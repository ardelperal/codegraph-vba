/**
 * VbaExtractor — regex extractor for `.bas` / `.cls` / `.frm` / `.dsr` source.
 *
 * Emits the symbol + edge shape expected by `codegraph_explore`:
 *  - one `module` node per `.bas` / one `class` node per `.cls` / one `file`
 *    stub for `.frm` / `.dsr` legacy
 *  - one `function` node per `Sub` / `Function` / `Property Get|Let|Set`
 *    declaration, with `node.visibility` set to the canonical lowercase enum
 *    (`'public'` / `'private'`; `Friend`, `Static`, and missing visibility fold
 *    to `'public'`).
 *  - `contains` edges from the module/class node to each function node.
 *  - `calls` edges from a procedure to a same-file procedure (`Sub Outer()` →
 *    `Inner()`); qualified cross-module calls carry `provenance: 'heuristic'`
 *    and `metadata.synthesizedBy: 'vba-name-resolution'`.
 *  - `implements` edges for `Implements IFoo` declarations.
 *  - `references` edges for `Dim x As Foo.Bar` (qualified Dim → outer type)
 *    tagged `synthesizedBy: 'vba-name-resolution'`.
 *  - `references` edges for `WithEvents m_X As Form_Foo` tagged
 *    `synthesizedBy: 'vba-withevents'`.
 *  - `references` edges for SQL table names found inside string literals passed
 *    directly to SQL wrappers, or assigned to variables before
 *    `getdb().OpenRecordset(...)` / `getdb().Execute ...`, tagged
 *    `synthesizedBy: 'vba-sql-table'`.
 *
 * Rejects `.form.txt` / `.report.txt` (REQ-CODE-9 — the form UI extractor
 * owns those). `Option ...` directives are inert — `stripVbaComments()`
 * drops them.
 *
 * Issue #83: this class is now a THIN ORCHESTRATOR that owns a SINGLE
 * per-line walk. Preprocessed source is split into a `readonly string[]`
 * ONCE, then a `VbaWalker` (a small array of `VbaClassifier`s, one per
 * concern) dispatches each line to every classifier in stable order.
 * Previously, each of the 5 per-concern sweeps did its own
 * `src.split('\n')` + full O(n) traversal — 6+ full traversals of every
 * file. Acceptance criterion #1: source is now split at most ONCE for
 * extraction; the 3 preprocessing stages are all `.replace()`-based
 * (no split).
 */
import * as path from 'path';
import { Node, ExtractionResult } from '../types';
import { generateNodeId } from './tree-sitter-helpers';
import {
  joinLineContinuations,
  preprocessConditionalCompilation,
  stripVbaComments,
} from './vba-preprocess';
import { VbaExtractorContext, VbaClassifier } from './vba/context';
import type { VbaExtractionRule } from './vba/rules';
import { RULES as PROCEDURES_RULES, createProceduresClassifier } from './vba/procedures';
import {
  RULES as EVENTS_TYPES_DECLARES_RULES,
  createEventsTypesDeclaresClassifier,
} from './vba/declarations';
import { RULES as DIMS_RULES, createDimsClassifier } from './vba/dims';
import { RULES as ENUMS_CONSTS_RULES, createEnumsConstsClassifier } from './vba/enums-consts';
import { RULES as IMPLEMENTS_RULES, createImplementsClassifier } from './vba/implements';
import {
  RULES as CALL_SWEEP_RULES,
  createCallsAndSqlClassifier,
} from './vba/call-sweep';
import {
  emitPerFileTimings,
  flushAggregate,
  readTimingMode,
  recordAggregateTimings,
  withClassifier,
  withStage,
  type TimingMode,
} from './vba-timing';

/**
 * Issue #152: per-file fanout cap for `RaiseEvent <EventName>` edges. An
 * event raised from more than this many sites in a single file is
 * flagged `metadata.highFanout: true` and ALL `raises-event` edges to it
 * are dropped — the gate suppresses graph noise from events with generic
 * names (`Change`, `Click`, `AfterUpdate`, …) that would otherwise
 * produce hundreds of edges per file. Mirrors the upstream
 * `EVENT_FANOUT_CAP` discipline in
 * `src/resolution/callback-synthesizer.ts:43`. Configurable via
 * `codegraph.json` → `vba.maxRaiseFanout`. Pass `Number.POSITIVE_INFINITY`
 * (or any value `>=` the largest expected count) to disable the gate.
 */
export const DEFAULT_MAX_RAISE_FANOUT = 50;

/**
 * Issue #153: the orchestrator's aggregated view of every VBA
 * classifier's declarative `RULES` table. Each entry is the
 * `readonly VbaExtractionRule[]` constant exported by the
 * corresponding classifier file. Exposed at module scope so:
 *
 *   1. Tests can pin the table shape and assert that the orchestrator
 *      still references each table (a deleted import would
 *      accidentally break a whole concern and the test would catch
 *      the missing reference).
 *   2. Tooling can introspect every rule the orchestrator
 *      dispatches — useful for `codegraph stats` style UIs that
 *      show "what kinds of VBA symbols does codegraph know about".
 *   3. The `validateVbaRuleTables` invariant runs on this aggregate
 *      so a refactor that empties one concern's `RULES` array
 *      (e.g. by deleting the `defineRule` calls during a partial
 *      migration) fails loudly at module load instead of silently
 *      dropping a whole concern.
 *
 * The keys are the FILE BASENAMES of the classifier modules (no
 * `.ts` extension, kebab-case) — the same naming convention the
 * per-classifier export comments use. Stable identifier for
 * `validateVbaRuleTables` consumers.
 */
export const VBA_RULE_TABLES: Readonly<Record<string, readonly VbaExtractionRule[]>> =
  {
    implements: IMPLEMENTS_RULES,
    procedures: PROCEDURES_RULES,
    declarations: EVENTS_TYPES_DECLARES_RULES,
    dims: DIMS_RULES,
    'enums-consts': ENUMS_CONSTS_RULES,
    'call-sweep': CALL_SWEEP_RULES,
  };

/**
 * Issue #153: the canonical list of concern keys `VBA_RULE_TABLES`
 * must contain. Order matches the per-line dispatch order in
 * `extract()`: events/types/declares → implements → dims → enums/consts
 * → calls/sql. The procedures sweep is the only one that runs BEFORE
 * the main walk (so `ctx.localProcs` is populated for the call sweep
 * to consult) and is therefore referenced in `VBA_RULE_TABLES` but NOT
 * in this list — it is intentionally validated as a separate concern
 * because the pre-walk contract is different.
 */
const REQUIRED_DISPATCH_TABLES: readonly string[] = [
  'declarations',
  'implements',
  'dims',
  'enums-consts',
  'call-sweep',
];

/**
 * Issue #153: validate that every per-concern rule table the
 * orchestrator dispatches is non-empty. Called once at module load
 * (see the IIFE at the bottom of this file) so an accidentally-
 * emptied `RULES` array fails loudly when the extractor is first
 * imported, NOT at the first `extract()` call.
 *
 * The result is `{ ok, empty }`:
 *   - `ok: true`  — every required table has at least one rule.
 *   - `ok: false` — at least one table is empty; `empty` lists the
 *                   concern keys that need attention.
 *
 * The function accepts an optional `tables` argument so the unit
 * test in `__tests__/extraction-vba-rule-table.test.ts` can verify
 * the validator's contract independently of the live data.
 */
export function validateVbaRuleTables(
  tables: Readonly<Record<string, readonly VbaExtractionRule[]>> = VBA_RULE_TABLES,
): { ok: boolean; empty: string[] } {
  const empty: string[] = [];
  for (const key of REQUIRED_DISPATCH_TABLES) {
    const table = tables[key];
    if (!table || table.length === 0) empty.push(key);
  }
  return { ok: empty.length === 0, empty };
}

// Run the validator once at module load. The IIFE THROWS on a non-ok
// result so an accidentally-emptied `RULES` array (e.g. from a partial
// refactor that drops the `defineRule` calls) aborts the import
// instead of silently losing an entire concern at the first
// `extract()` call. A user running `codegraph index` without first
// running `npm test` would otherwise see no symptom at all — the
// downstream `extract()` would just emit fewer symbols. The unit
// test in `__tests__/extraction-vba-rule-table.test.ts` pins the
// `validateVbaRuleTables` return-value contract; the IIFE itself is
// pinned by `__tests__/extraction-vba-extractor-import-invariant.test.ts`
// (issue #164).
(function runRuleTableInvariant(): void {
  const result = validateVbaRuleTables();
  if (!result.ok) {
    throw new Error(
      `[vba-extractor] VBA_RULE_TABLES is missing rules for: ${result.empty.join(', ')}. ` +
        `Each concern's RULES array must be non-empty.`,
    );
  }
})();

export class VbaExtractor {
  private filePath: string;
  private source: string;
  private ctx: VbaExtractorContext;
  private vbaTargets?: Record<string, boolean>;
  /**
   * Issue #152: per-file fanout cap for `RaiseEvent <EventName>` edges.
   * When a single event is raised more than this many times in one file,
   * the event node is flagged `metadata.highFanout: true` and ALL
   * `raises-event` edges to it are dropped. Pass `undefined` to disable
   * the gate (the legacy, uncapped behaviour). The 50 default lives in
   * the orchestrator: callers that don't pass a value get the default.
   */
  private maxRaiseFanout: number | undefined;

  constructor(
    filePath: string,
    source: string,
    vbaTargets?: Record<string, boolean>,
    maxRaiseFanout: number = DEFAULT_MAX_RAISE_FANOUT,
  ) {
    this.filePath = filePath;
    this.source = source;
    this.vbaTargets = vbaTargets;
    this.maxRaiseFanout = maxRaiseFanout;
    this.ctx = new VbaExtractorContext(filePath);
  }

  extract(): ExtractionResult {
    const startTime = Date.now();

    // Issue #156: lazy-enable the timing instrumentation ONLY when the
    // env var is set. The Maps stay `null` otherwise so the default path
    // (every real user's `codegraph index` run) pays no Map-allocation
    // cost and no `recordStage` call does any real work.
    const timingMode = readTimingMode();
    if (timingMode !== 'off') {
      this.ctx.ensureTimings();
    }

    try {
      // REQ-CODE-9: hand form/report files off to VbaFormExtractor. If
      // somehow routed here, emit just a file node and return.
      if (this.isFormOrReportFile()) {
        this.ctx.nodes.push(this.createFileNode());
        const result = this.result(startTime);
        this.maybeEmitTimings(result, timingMode);
        return result;
      }

      // Pre-process pipeline (per design): strip comments first to prevent
      // comments from blocking or causing false line continuations, then join
      // continuations, and blank inactive conditional-compilation branches.
      // Line count is preserved by all stages.
      //
      // Issue #156: each stage is wrapped in a timing block that records
      // into the context's `timings` Map. `preprocessConditionalCompilation`
      // itself contains an internal lexer+parser pair; we sum that into a
      // single `preprocess.cc.lexer+parser` stage via a thin wrapper that
      // re-uses the same `withStage` helper.
      const cleanComments = withStage(
        this.ctx,
        'preprocess.stripVbaComments',
        () => stripVbaComments(this.source),
      );
      const joined = withStage(
        this.ctx,
        'preprocess.joinLineContinuations',
        () => joinLineContinuations(cleanComments),
      );
      const uncommented = withStage(
        this.ctx,
        'preprocess.conditionalCompilation',
        () =>
          preprocessConditionalCompilation(
            joined,
            this.vbaTargets,
            this.ctx.timings,
          ),
      );

      // Create the file node.
      this.ctx.nodes.push(this.createFileNode());

      // Resolve the file's "kind" — .cls → class, else module (including .bas
      // and the legacy .frm/.dsr stubs).
      const isCls = this.filePath.toLowerCase().endsWith('.cls');

      // Detect VB_Name attribute (first non-empty line).
      const vbName = this.detectVbName(this.source);

      // B3 (hueco 5): compute the class-name prefix used to compose every
      // `function` node's `qualifiedName` in this file. For `.cls` files
      // the prefix is the resolved VB_Name (or the file basename when
      // VB_Name is absent), producing qualifiedNames like
      // `Form_TestForm.Form_Load`. For `.bas` (and `.frm`/`.dsr`) files
      // the prefix is null, preserving the bare-name qualifiedName that
      // every existing test asserts.
      const fallbackName = path
        .basename(this.filePath)
        .replace(/\.[^.]+$/, '');
      const className = isCls ? (vbName ?? fallbackName) : null;
      this.ctx.classNamePrefix = className;

      // Issue #83: split-then-walk. The preprocessed source is split into
      // lines ONCE into a shared `readonly string[]`. The procedures
      // classifier runs FIRST on the full file so `ctx.localProcs`,
      // `ctx.functionNodeByName`, `ctx.functionNodeByStartLine`, and
      // `ctx.functionReturnTypes` are fully populated before the call/SQL
      // classifier needs to resolve a same-file call target (the legacy
      // `sweepProcedures` ran the entire file first too). The remaining
      // five classifiers then share a single per-line walk, dispatching
      // each line in stable order. The split count drops from 6+ (one per
      // sweep) to 1; the walk count is 2 over the same `lines` array.
      // `hasAnySymbols` is derived from each classifier's `count`.
      const lines = uncommented.split('\n');
      const proceduresCls = createProceduresClassifier();
      const classifiers: VbaClassifier[] = [
        createEventsTypesDeclaresClassifier(),
        createImplementsClassifier(),
        createDimsClassifier(),
        createEnumsConstsClassifier(),
        createCallsAndSqlClassifier(lines),
      ];

      // Pre-walk: procedures only — populates the same-file function index
      // (legacy behaviour: `sweepProcedures` ran first so the call sweep
      // could look up call targets by name).
      withStage(this.ctx, 'walk.procedures', () => {
        for (let i = 0; i < lines.length; i++) {
          withClassifier(this.ctx, proceduresCls, lines[i] ?? '', i);
        }
      });

      // Main walk: every other classifier sees every line in order.
      withStage(this.ctx, 'walk.main', () => {
        for (let i = 0; i < lines.length; i++) {
          const line = lines[i] ?? '';
          for (const c of classifiers) {
            withClassifier(this.ctx, c, line, i);
          }
        }
      });

      // Finalize (calls-sql flushes proc-end line updates to function nodes).
      for (const c of classifiers) c.finalize?.(this.ctx);

      // Issue #152: per-file fanout gate for `RaiseEvent <EventName>`
      // edges. Events raised from more than `maxRaiseFanout` sites get
      // the `metadata.highFanout` flag and ALL of their `raises-event`
      // edges dropped. The gate is a no-op when `maxRaiseFanout` is
      // undefined (the legacy, uncapped behaviour) and runs BEFORE the
      // module/class node is created so `pendingModuleOrClassSource`
      // re-attribution never sees the dropped edges.
      if (this.maxRaiseFanout !== undefined) {
        this.ctx.applyRaiseFanoutGate(this.maxRaiseFanout);
      }

      const hasAnySymbols =
        proceduresCls.count > 0 || classifiers.some((c) => c.count > 0);
      const procs = this.ctx.procedures;

      // Create the module/class node ONLY when the file has actual symbols
      // (REQ-CODE-10 — a file with ONLY Option directives emits zero symbol
      // nodes. A file with Enum/Const but no procedures DOES emit a module
      // node, since those are real symbols — see REQ-CODE-12/13).
      if (hasAnySymbols) {
        this.ctx.moduleOrClassNode = this.createModuleOrClassNode(isCls, vbName);
        this.ctx.nodes.push(this.ctx.moduleOrClassNode);

        // Re-attribute contains/implements/reference edges whose source was
        // held in `pendingModuleOrClassSource` — see below. Since we now
        // know the module id, redirect those edges to it.
        for (const edge of this.ctx.pendingModuleOrClassSource) {
          edge.source = this.ctx.moduleOrClassNode.id;
        }
        this.ctx.pendingModuleOrClassSource.length = 0;

        // Sub New marker on the class node (REQ-CODE-3).
        if (isCls) {
          const hasNew = procs.some((p) => p.name === 'New');
          if (hasNew) {
            const md = (this.ctx.moduleOrClassNode.metadata ?? {}) as Record<string, unknown>;
            md.hasClassInitializer = true;
            md.initializerName = 'New';
            this.ctx.moduleOrClassNode.metadata = md;
          }
        }
      }
    } catch (error) {
      this.ctx.errors.push({
        message: `VBA extraction error: ${error instanceof Error ? error.message : String(error)}`,
        filePath: this.filePath,
        severity: 'error',
        code: 'parse_error',
      });
    }

    const result = this.result(startTime);
    this.maybeEmitTimings(result, timingMode);
    return result;
  }

  // ---------------------------------------------------------------------------
  // Helpers
  // ---------------------------------------------------------------------------

  /**
   * Issue #156: emit the per-file timing block (mode `1` or `2`) and
   * — in mode `2` — append the running aggregate. No-op when
   * `mode === 'off'`. In aggregate mode we flush after every extract()
   * call so the line is visible per-file even without a graceful
   * shutdown (crash-on-file-3 still shows files 1-3 in the aggregate).
   */
  private maybeEmitTimings(_result: ExtractionResult, mode: TimingMode): void {
    if (mode === 'off') return;
    if (mode === 'per-file' || mode === 'aggregate') {
      emitPerFileTimings(this.ctx, this.filePath);
    }
    if (mode === 'aggregate') {
      recordAggregateTimings(this.ctx, this.filePath);
      flushAggregate();
    }
  }

  private result(startTime: number): ExtractionResult {
    return {
      nodes: this.ctx.nodes,
      edges: this.ctx.edges,
      unresolvedReferences: this.ctx.unresolvedReferences,
      errors: this.ctx.errors,
      durationMs: Date.now() - startTime,
    };
  }

  private isFormOrReportFile(): boolean {
    const lower = this.filePath.toLowerCase();
    return lower.endsWith('.form.txt') || lower.endsWith('.report.txt');
  }

  private createFileNode(): Node {
    const lines = this.source.split('\n');
    return {
      id: generateNodeId(this.filePath, 'file', this.filePath, 1),
      kind: 'file',
      name: path.basename(this.filePath),
      qualifiedName: this.filePath,
      filePath: this.filePath,
      language: 'vba',
      startLine: 1,
      endLine: lines.length,
      startColumn: 0,
      endColumn: lines[lines.length - 1]?.length ?? 0,
      updatedAt: Date.now(),
    };
  }

  private detectVbName(src: string): string | null {
    // Walk past the Access class metadata header block (VERSION / BEGIN /
    // MultiUse / END / Attribute …) so VB_Name on a later line is found.
    // The previous implementation returned null at the first non-Attribute
    // line, which meant real Access .cls files — which start with
    // `VERSION 1.0 CLASS / BEGIN / MultiUse = ... / END` before
    // `Attribute VB_Name = "..."` — always fell through to the basename
    // fallback. Audit W2 (June 2026).
    for (const line of src.split(/\r?\n/)) {
      const trimmed = line.trim();
      if (!trimmed) continue;
      const m = /^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"/i.exec(trimmed);
      if (m) return m[1] ?? null;
      // Skip known header keywords so we keep scanning.
      if (
        /^\s*VERSION\b/i.test(trimmed) ||
        /^\s*BEGIN\b/i.test(trimmed) ||
        /^\s*END\b/i.test(trimmed) ||
        /^\s*(?:MultiUse|Persistable|DataBindingBehavior|DataSourceBehavior)\s*=/i.test(trimmed) ||
        /^\s*Attribute\s+/i.test(trimmed)
      ) {
        continue;
      }
      // Any other non-empty line that isn't a known header — stop.
      return null;
    }
    return null;
  }

  private createModuleOrClassNode(isCls: boolean, vbName: string | null): Node {
    const fallbackName = path.basename(this.filePath).replace(/\.[^.]+$/, '');
    const name = vbName ?? fallbackName;
    return {
      id: generateNodeId(this.filePath, isCls ? 'class' : 'module', name, 1),
      kind: isCls ? 'class' : 'module',
      name,
      qualifiedName: name,
      filePath: this.filePath,
      language: 'vba',
      startLine: 1,
      endLine: this.source.split('\n').length,
      startColumn: 0,
      endColumn: 0,
      metadata: {},
      updatedAt: Date.now(),
    };
  }
}
