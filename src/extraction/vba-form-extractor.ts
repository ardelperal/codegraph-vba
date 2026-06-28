/**
 * VbaFormExtractor — stub created during T1 wiring.
 *
 * Full implementation lands in T4 (Phase 4 of `tasks.md`). Same rationale as
 * `vba-extractor.ts`: this stub lets `tree-sitter.ts` import a real symbol
 * so the T1 wiring is compilable end-to-end. Returns an empty result until T4.
 */
import { ExtractionResult } from '../types';

export class VbaFormExtractor {
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