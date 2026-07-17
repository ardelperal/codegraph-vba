/**
 * Dysflow-export framework resolver (issue #154).
 *
 * The 3 Dysflow-specific extractors ÔÇö form/report SaveAsText UI, test
 * manifests, and test sequences ÔÇö used to be hard-coded into
 * `extractFromSource` as a `detectedLanguage === 'vba' && <isFoo> (filePath)`
 * ladder. They are now lifted into this `FrameworkResolver`-shaped module
 * so they can be:
 *   1. Reached through the standard framework registry (the same path every
 *      other framework ÔÇö Express, React, Spring ÔÇö flows through).
 *   2. Opted out at the project level via `codegraph.json`'s
 *      `vba.dysflowExport: false` ÔÇö useful for projects that carry legacy
 *      `.form.txt`/`.report.txt` files (or test manifests from a different
 *      system) and want them tracked as just a `file` node instead of
 *      being expanded into the graph.
 *
 * Architecture: the underlying `VbaFormExtractor` / `VbaTestManifestExtractor`
 * / `VbaTestSequenceExtractor` classes are the single source of truth for
 * the per-file extraction logic. This module:
 *   - decides WHICH sub-extractor applies to a file (path shape);
 *   - delegates the actual `extract()` to that sub-extractor;
 *   - exposes the `FrameworkResolver` shape (`detect` / `extract` /
 *     `resolve` / `claimsReference`) so the rest of the codebase can treat
 *     "Dysflow export" as a regular framework on top of the base VBA
 *     language.
 *
 * Behavior contract: with `dysflowExport: true` (the default), the emitted
 * nodes/edges/references are byte-identical to the pre-refactor paths in
 * `tree-sitter.ts`. With `dysflowExport: false`, the framework's `detect()`
 * returns `false` so it's never in the per-project detected list ÔÇö meaning
 * its `extract()` is never called and the form/report/manifest/sequence
 * files fall through to a `file`-only node via the language-specific
 * dispatch in `tree-sitter.ts`.
 */
import { FrameworkResolver, UnresolvedRef, ResolvedRef, ResolutionContext, FrameworkExtractionResult } from '../../resolution/types';
import { VbaFormExtractor } from '../vba-form-extractor';
import { VbaTestManifestExtractor } from '../vba-test-manifest-extractor';
import { VbaTestSequenceExtractor } from '../vba-test-sequence-extractor';
import { detectVbaFormFile, isVbaTestManifestFile, isVbaTestSequenceFile } from '../grammars';
import { loadDysflowExportConfig } from '../../project-config';

/**
 * Pick the per-file Dysflow sub-extractor for `filePath`. Returns `null`
 * when the path is not a Dysflow-specific shape ÔÇö the framework
 * extractor's `extract()` then no-ops and the base language owns the file.
 */
function pickSubExtractor(filePath: string): {
  kind: 'form' | 'manifest' | 'sequence';
  run: (filePath: string, source: string) => { nodes: any[]; edges: any[]; unresolvedReferences: any[]; errors: any[] };
} | null {
  if (detectVbaFormFile(filePath)) {
    return {
      kind: 'form',
      run: (fp, src) => new VbaFormExtractor(fp, src).extract(),
    };
  }
  if (isVbaTestSequenceFile(filePath)) {
    return {
      kind: 'sequence',
      run: (fp, src) => new VbaTestSequenceExtractor(fp, src).extract(),
    };
  }
  if (isVbaTestManifestFile(filePath)) {
    return {
      kind: 'manifest',
      run: (fp, src) => new VbaTestManifestExtractor(fp, src).extract(),
    };
  }
  return null;
}

export const dysflowExportResolver: FrameworkResolver = {
  name: 'dysflow-export',
  // Scoped to VBA ÔÇö these file shapes are only meaningful on top of a VBA
  // source tree (Dysflow's SaveAsText form / .form.txt / .report.txt /
  // tests.<slice>.json / sequences/*.json are all VBA artifacts).
  languages: ['vba'],

  /**
   * Project-level detection: the project "uses Dysflow export" iff at
   * least one VBA tree-sitter-detected file matches a Dysflow file shape
   * AND the project's `codegraph.json` hasn't opted out via
   * `vba.dysflowExport: false` (issue #154). The opt-out is checked
   * FIRST so a project that carries legacy `.form.txt`/`.report.txt`
   * files (or test manifests from a different system) is not classified
   * as "using Dysflow" just because one of those files happens to live in
   * the tree ÔÇö they'd be tracked as just a `file` node instead.
   *
   * The detection runs once at index start (via `detectFrameworks`), so a
   * `.form.txt`/`.report.txt`/manifest/sequence present anywhere in the
   * project tree qualifies it. When the opt-out is on, the framework is
   * never registered as "detected" ÔÇö its `extract()` is never called.
   */
  detect(context: ResolutionContext): boolean {
    try {
      if (!loadDysflowExportConfig(context.getProjectRoot())) {
        return false;
      }
    } catch {
      // Config load failure (very rare) ÔÇö treat as opt-out so we don't
      // emit Dysflow-specific nodes against an unverified config.
      return false;
    }
    try {
      const files = context.getAllFiles();
      for (const f of files) {
        if (
          detectVbaFormFile(f) ||
          isVbaTestSequenceFile(f) ||
          isVbaTestManifestFile(f)
        ) {
          return true;
        }
      }
      return false;
    } catch {
      return false;
    }
  },

  /**
   * Per-file extraction. Dispatches to the matching sub-extractor when the
   * file is one of the three Dysflow shapes; returns an empty result for
   * any other VBA file (the framework `extract()` is invoked on every
   * detected-framework file, not just the ones this resolver owns, so
   * returning empty for `.bas`/`.cls` is correct).
   */
  extract(filePath: string, content: string): FrameworkExtractionResult {
    const sub = pickSubExtractor(filePath);
    if (!sub) {
      return { nodes: [], references: [] };
    }
    const r = sub.run(filePath, content);
    return {
      nodes: r.nodes,
      references: r.unresolvedReferences,
    };
  },

  /**
   * Reference resolution hook. The 3 Dysflow shapes emit `references`-kind
   * UnresolvedReferences whose targets are sibling `.cls` files (form
   * binding) or `Test_*` VBA function nodes (manifest + sequence). The
   * shared `ReferenceResolver` already binds these by name + language +
   * provenance, so this framework resolver doesn't need a custom
   * `resolve()` step ÔÇö the standard resolution pipeline handles it.
   * Declared as a no-op so the resolver's `claimsReference` / `resolve`
   * contract is consistent.
   */
  resolve(_ref: UnresolvedRef, _context: ResolutionContext): ResolvedRef | null {
    return null;
  },

  /**
   * The Dysflow references point at sibling-`.cls` basenames (form
   * binding) or `Test_*` VBA function names (manifest + sequence). The
   * shared name-based resolver can find them on its own; we don't need
   * a synthetic claim. Kept off so the resolver only intervenes at
   * extraction time.
   */
};
