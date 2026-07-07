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
 * This class is a THIN ORCHESTRATOR: it owns the input, drives the file-level
 * node creation, and sequences the per-concern sweeps. All shared extraction
 * state (node/edge accumulators, the `functionNodeByName` / `localVarTypeMap`
 * maps, `moduleOrClassNode`) and the sweeps themselves live on
 * `VbaExtractorContext` under `src/extraction/vba/`.
 */
import * as path from 'path';
import { Node, ExtractionResult } from '../types';
import { generateNodeId } from './tree-sitter-helpers';
import {
  joinLineContinuations,
  preprocessConditionalCompilation,
  stripVbaComments,
} from './vba-preprocess';
import { VbaExtractorContext } from './vba/context';
import { sweepProcedures } from './vba/procedures';
import { sweepEventsTypesAndDeclares } from './vba/declarations';
import { sweepDimsAndWithEvents } from './vba/dims';
import { sweepEnumsAndConsts } from './vba/enums-consts';
import { sweepImplements } from './vba/implements';
import { sweepCallsAndSql } from './vba/call-sweep';

export class VbaExtractor {
  private filePath: string;
  private source: string;
  private ctx: VbaExtractorContext;

  constructor(filePath: string, source: string) {
    this.filePath = filePath;
    this.source = source;
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
      const uncommented = preprocessConditionalCompilation(joined);

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

      // Module/class node — created lazily so a file containing only
      // Option directives (REQ-CODE-10: "emits nothing") doesn't carry a
      // module node with no symbols attached.
      let hasAnySymbols = false;

      // Sweep: procedures → function nodes + contains edges.
      const procs = sweepProcedures(this.ctx, uncommented);
      if (procs.length > 0) hasAnySymbols = true;

      // Sweep: first-class Event / Type / Declare declarations (roadmap #26).
      // This runs before call-site scanning so `RaiseEvent Foo` and calls to a
      // `Declare` can resolve to real local nodes rather than fall through.
      const declarationCount = sweepEventsTypesAndDeclares(this.ctx, uncommented);
      if (declarationCount > 0) hasAnySymbols = true;

      // Sweep: Implements (REQ-CODE-5).
      const implCount = sweepImplements(this.ctx, uncommented);
      if (implCount > 0) hasAnySymbols = true;

      // Sweep: qualified Dim (REQ-CODE-6) and WithEvents (REQ-CODE-7).
      const dimCount = sweepDimsAndWithEvents(this.ctx, uncommented);
      if (dimCount > 0) hasAnySymbols = true;

      // Sweep: Enum/Const declarations (REQ-CODE-12, REQ-CODE-13). Dysflow
      // exports the full module text, so a constants module carries its
      // `Const` lines and `Enum ... End Enum` blocks verbatim — these are the
      // project's domain dictionary and must be in the graph.
      const enumConstCount = sweepEnumsAndConsts(this.ctx, uncommented);
      if (enumConstCount > 0) hasAnySymbols = true;

      // Sweep: call sites (REQ-CODE-4) and SQL-in-strings (REQ-CODE-8).
      // Both walk the same per-line view; combining them in one pass is
      // simpler and keeps line tracking consistent.
      sweepCallsAndSql(this.ctx, uncommented);

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
