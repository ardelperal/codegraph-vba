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
import { createProceduresClassifier } from './vba/procedures';
import { createEventsTypesDeclaresClassifier } from './vba/declarations';
import { createDimsClassifier } from './vba/dims';
import { createEnumsConstsClassifier } from './vba/enums-consts';
import { createImplementsClassifier } from './vba/implements';
import { createCallsAndSqlClassifier } from './vba/call-sweep';

export class VbaExtractor {
  private filePath: string;
  private source: string;
  private ctx: VbaExtractorContext;
  private vbaTargets?: Record<string, boolean>;

  constructor(filePath: string, source: string, vbaTargets?: Record<string, boolean>) {
    this.filePath = filePath;
    this.source = source;
    this.vbaTargets = vbaTargets;
    this.ctx = new VbaExtractorContext(filePath);
  }

  extract(): ExtractionResult {
    const startTime = Date.now();

    try {
      // REQ-CODE-9: hand form/report files off to VbaFormExtractor. If
      // somehow routed here, emit just a file node and return.
      if (this.isFormOrReportFile()) {
        this.ctx.nodes.push(this.createFileNode());
        return this.result(startTime);
      }

      // Pre-process pipeline (per design): strip comments first to prevent
      // comments from blocking or causing false line continuations, then join
      // continuations, and blank inactive conditional-compilation branches.
      // Line count is preserved by all stages.
      const cleanComments = stripVbaComments(this.source);
      const joined = joinLineContinuations(cleanComments);
      const uncommented = preprocessConditionalCompilation(joined, this.vbaTargets);

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
      for (let i = 0; i < lines.length; i++) {
        proceduresCls.classifyLine(lines[i] ?? '', i, this.ctx);
      }

      // Main walk: every other classifier sees every line in order.
      // Issue #151: when a line matches `End Sub` / `End Function` /
      // `End Property`, fire each classifier's optional `onProcedureEnd`
      // hook BEFORE `classifyLine` runs on the same line. The hook sees
      // the closing procedure's `startLine` as the key so the deferred
      // qualified-call list (and any other per-procedure state) can be
      // drained while the procedure is still on the call sweep's stack
      // — `localVarTypeMap` is at its full procedure scope at that
      // point. The proc-stack pop inside `classifyLine` happens right
      // after, restoring the parent scope.
      const procedureEndRe = /^\s*End\s+(?:Sub|Function|Property)\b/i;
      for (let i = 0; i < lines.length; i++) {
        const line = lines[i] ?? '';
        if (procedureEndRe.test(line)) {
          // Find the most recent `Sub|Function|Property` opener above
          // this line. The procedures classifier is a single
          // pass over the same `lines` array, so the most recent
          // opening line for the closing `End` is the startLine of
          // the procedure that owns this `End`. We re-derive it here
          // (rather than threading a second stack through the
          // orchestrator) because the procedures classifier's local
          // state is private.
          let openLine = -1;
          for (let j = i - 1; j >= 0; j--) {
            const prev = lines[j] ?? '';
            if (/^\s*End\s+(?:Sub|Function|Property)\b/i.test(prev)) break;
            if (/^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Sub|Function|Property(?:\s+(?:Get|Let|Set))?)\s+\p{L}[\p{L}\p{N}_]*/iu.test(prev)) {
              openLine = j + 1; // 1-based
              break;
            }
          }
          if (openLine > 0) {
            for (const c of classifiers) {
              c.onProcedureEnd?.(openLine, this.ctx);
            }
          }
        }
        for (const c of classifiers) {
          c.classifyLine(line, i, this.ctx);
        }
      }

      // Finalize (calls-sql flushes proc-end line updates to function nodes).
      for (const c of classifiers) c.finalize?.(this.ctx);

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

    return this.result(startTime);
  }

  // ---------------------------------------------------------------------------
  // Helpers
  // ---------------------------------------------------------------------------

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
