/**
 * VbaExtractor — stub created during T1 wiring.
 *
 * Full implementation lands in T3 (Phase 3 of `tasks.md`). This stub lets
 * `src/extraction/tree-sitter.ts` import a real symbol so the T1 wiring is
 * compilable end-to-end. The stub returns an empty `ExtractionResult` so
 * `extractFromSource('foo.bas', ...)` produces a no-op until T3 lands.
 */
import { ExtractionResult } from '../types';

export class VbaExtractor {
  constructor(_filePath: string, _source: string) {
    void _filePath;
    void _source;
  }

  extract(): ExtractionResult {
    return {
      nodes: [],
      edges: [],
      unresolvedReferences: [],
      errors: [],
      durationMs: 0,
    };
  }
}