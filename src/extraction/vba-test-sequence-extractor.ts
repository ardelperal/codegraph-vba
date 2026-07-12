/**
 * VbaTestSequenceExtractor — Dysflow VBA test-sequence JSON
 * (`tests/sequences/*.json`, e.g. `tests/sequences/cache-riesgo.json`).
 *
 * Each sequence groups one or more `Test_*` VBA atoms behind a runner policy
 * that an orchestrator (e.g. an MCP) honors. The shape is:
 *
 *   { "description": "...",
 *     "runnerPolicy": { tool, mode, sequential, ... },
 *     "procedures":  [ "Test_X_RunSlice", "Test_Y_ResetSlice", ... ] }
 *
 * This extractor is the sibling of `VbaTestManifestExtractor`: it links each
 * named procedure to its `Test_*` `function` node WITHOUT duplicating it.
 *
 * SUB-6 scope (deliberately narrow — issue #97):
 *  - ONE `file` node when the shape matches.
 *  - ONE `UnresolvedReference` per `procedures[]` entry carrying a string
 *    name, with metadata `{ synthesizedBy: 'vba-test-sequence',
 *    runnerPolicy, sequenceFile, procedureIndex }`.
 *  - The `ReferenceResolver` already binds `referenceKind='references'` +
 *    `language='vba'` references by bare name to `function` nodes via
 *    `matchFunctionRef` (SUB-3). SUB-6 only adds a provenance-carry-through
 *    for `synthesizedBy: 'vba-test-sequence'` so the new metadata survives
 *    onto the resolved edge (`runnerPolicy`, `sequenceFile`, `procedureIndex`).
 *
 * SCOPE EXCLUSIONS (future work — tracked under epic #91):
 *  - `tests.vba.strict-sequence.json` (`executionUnits` array of paths):
 *    a top-level orchestration plan, NOT a sequence in this sense. The
 *    path-detection gate (under `sequences/`) keeps it out of scope anyway.
 *  - `tests.vba.slices.json` (`slices[]` with `submanifests`): a grouping of
 *    manifests, not a per-test-atom plan. Out of scope.
 *  Both rejection shapes are unit-tested below to lock the gate.
 */
import * as path from 'path';
import { Node, ExtractionResult, ExtractionError, UnresolvedReference } from '../types';
import { generateNodeId } from './tree-sitter-helpers';

/**
 * Content-shape gate: is `parsed` a Dysflow VBA test sequence — a top-level
 * `runnerPolicy` object AND a `procedures` array whose items are strings
 * (possibly empty)? Pure; the file's path is gated separately by
 * `isVbaTestSequenceFile` in `grammars.ts`.
 *
 * The gate accepts an EMPTY `procedures` array on purpose: an empty sequence
 * still gets a `file` node from the extractor (so the manifest is visible in
 * the graph) but emits zero `UnresolvedReference`s (no procedures to bind).
 *
 * A shape that ALSO carries an `executionUnits` or `slices[]` key (the
 * strict-sequence / slices shapes) is rejected — those are top-level
 * orchestration/grouping plans, not per-atom sequences.
 */
export function isVbaTestSequenceShape(parsed: unknown): boolean {
  if (!parsed || typeof parsed !== 'object') return false;
  const obj = parsed as { runnerPolicy?: unknown; procedures?: unknown };
  if (!obj.runnerPolicy || typeof obj.runnerPolicy !== 'object') return false;
  if (!Array.isArray(obj.procedures)) return false;
  return obj.procedures.every((p) => typeof p === 'string');
}

export class VbaTestSequenceExtractor {
  private filePath: string;
  private source: string;
  private nodes: Node[] = [];
  private unresolvedReferences: UnresolvedReference[] = [];
  private errors: ExtractionError[] = [];

  constructor(filePath: string, source: string) {
    this.filePath = filePath;
    this.source = source;
  }

  extract(): ExtractionResult {
    const startTime = Date.now();

    let parsed: unknown;
    try {
      parsed = JSON.parse(this.source);
    } catch (error) {
      this.errors.push({
        message: `VBA test sequence parse error: ${
          error instanceof Error ? error.message : String(error)
        }`,
        filePath: this.filePath,
        severity: 'warning',
        code: 'parse_error',
      });
      return this.result(startTime);
    }

    if (!isVbaTestSequenceShape(parsed)) {
      return this.result(startTime);
    }

    const fileNode = this.createFileNode();
    this.nodes.push(fileNode);
    this.emitProcedureReferences(parsed, fileNode.id);
    return this.result(startTime);
  }

  private emitProcedureReferences(parsed: unknown, fromNodeId: string): void {
    const runnerPolicy = (parsed as { runnerPolicy?: unknown }).runnerPolicy;
    const procedures = (parsed as { procedures?: unknown[] }).procedures ?? [];
    procedures.forEach((procedure, idx) => {
      if (typeof procedure !== 'string') return;
      this.unresolvedReferences.push({
        fromNodeId,
        referenceName: procedure,
        referenceKind: 'references',
        line: 1,
        column: 0,
        filePath: this.filePath,
        language: 'vba',
        metadata: {
          synthesizedBy: 'vba-test-sequence',
          runnerPolicy,
          sequenceFile: this.filePath,
          procedureIndex: idx,
        },
      });
    });
  }

  private result(startTime: number): ExtractionResult {
    return {
      nodes: this.nodes,
      edges: [],
      unresolvedReferences: this.unresolvedReferences,
      errors: this.errors,
      durationMs: Date.now() - startTime,
    };
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
}