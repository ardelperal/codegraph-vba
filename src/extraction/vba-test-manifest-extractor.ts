/**
 * VbaTestManifestExtractor — Dysflow VBA test-manifest JSON
 * (`tests(.<slice>)*.json`, e.g. `tests/tests.vba.smoke.json`).
 *
 * Dysflow registers each VBA test atom in a JSON manifest:
 *
 *   { "tests": [ { "procedure": "Test_X_RunAll", "name": "...", "tags": [...] } ] }
 *
 * `procedure` names an existing `Test_*` VBA `function` node; `name`/`tags` are
 * optional. This extractor is the sibling of `SqlQueryExtractor` /
 * `VbaFormExtractor`: it links each registered atom to its manifest metadata
 * WITHOUT duplicating the procedure node.
 *
 * SUB-1 (this slice) covers detection + the file-node skeleton only:
 *  - a guarded `JSON.parse` (a malformed manifest must never crash the index);
 *  - a content-shape gate — the parsed JSON must have a top-level `tests` array
 *    whose items carry a string `procedure` — so a `tests.*.json` that is not a
 *    VBA manifest (e.g. a `sequences` config) produces nothing;
 *  - one `file` node when the shape matches.
 *
 * The per-entry `vba-test-manifest` `UnresolvedReference`s (SUB-2) and their
 * resolution to `Test_*` function nodes (SUB-3) are added in later slices.
 */
import * as path from 'path';
import { Node, ExtractionResult, ExtractionError, UnresolvedReference } from '../types';
import { generateNodeId } from './tree-sitter-helpers';

/** One `tests[]` entry that carries a string `procedure`. */
interface ManifestTestEntry {
  procedure: string;
  name?: unknown;
  tags?: unknown;
}

/**
 * Content-shape gate: is `parsed` a VBA test manifest — a top-level `tests`
 * array with at least one item carrying a string `procedure`? Pure; the file's
 * basename is gated separately by `isVbaTestManifestFile` in `grammars.ts`.
 */
function isVbaTestManifestShape(parsed: unknown): boolean {
  if (!parsed || typeof parsed !== 'object') return false;
  const tests = (parsed as { tests?: unknown }).tests;
  if (!Array.isArray(tests)) return false;
  return tests.some(
    (entry) =>
      !!entry &&
      typeof entry === 'object' &&
      typeof (entry as { procedure?: unknown }).procedure === 'string',
  );
}

export class VbaTestManifestExtractor {
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
      // Malformed manifest — low-severity so it never breaks the index.
      this.errors.push({
        message: `VBA test manifest parse error: ${
          error instanceof Error ? error.message : String(error)
        }`,
        filePath: this.filePath,
        severity: 'warning',
        code: 'parse_error',
      });
      return this.result(startTime);
    }

    // Not a VBA manifest (valid JSON, wrong shape) → emit nothing, no error.
    if (!isVbaTestManifestShape(parsed)) {
      return this.result(startTime);
    }

    const fileNode = this.createFileNode();
    this.nodes.push(fileNode);
    this.emitTestReferences(parsed, fileNode.id);
    return this.result(startTime);
  }

  /**
   * One `UnresolvedReference` per `tests[]` entry carrying a string `procedure`,
   * pointing at that procedure name. The `ReferenceResolver` (SUB-3) binds each
   * to the existing `Test_*` `function` node — no duplicate node is emitted here.
   * `name` defaults to the procedure name; `tags` defaults to `[]`.
   */
  private emitTestReferences(parsed: unknown, fromNodeId: string): void {
    const tests = (parsed as { tests?: unknown[] }).tests ?? [];
    for (const raw of tests) {
      if (!raw || typeof raw !== 'object') continue;
      const entry = raw as ManifestTestEntry;
      if (typeof entry.procedure !== 'string') continue;

      const testName = typeof entry.name === 'string' ? entry.name : entry.procedure;
      const tags = Array.isArray(entry.tags) ? entry.tags : [];

      this.unresolvedReferences.push({
        fromNodeId,
        referenceName: entry.procedure,
        referenceKind: 'references',
        line: 1,
        column: 0,
        filePath: this.filePath,
        language: 'vba',
        metadata: {
          synthesizedBy: 'vba-test-manifest',
          testName,
          tags,
          manifestFile: this.filePath,
        },
      });
    }
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
