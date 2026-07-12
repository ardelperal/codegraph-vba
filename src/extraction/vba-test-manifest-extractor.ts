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
import { Node, ExtractionResult, ExtractionError } from '../types';
import { generateNodeId } from './tree-sitter-helpers';

/**
 * Content-shape gate: is `parsed` a VBA test manifest — a top-level `tests`
 * array with at least one item carrying a string `procedure`? Pure; the file's
 * basename is gated separately by `isVbaTestManifestFile` in `grammars.ts`.
 */
export function isVbaTestManifestShape(parsed: unknown): boolean {
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

    this.nodes.push(this.createFileNode());
    return this.result(startTime);
  }

  private result(startTime: number): ExtractionResult {
    return {
      nodes: this.nodes,
      edges: [],
      unresolvedReferences: [],
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
