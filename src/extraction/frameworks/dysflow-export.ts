/**
 * Dysflow-export framework resolver (issue #154).
 *
 * Marks a project as "this is a Dysflow export" so `dysflow-export` shows up
 * in the detected-framework list alongside Express, React, Spring and the
 * rest. Detection is path-shape based: at least one `.form.txt` /
 * `.report.txt` / `tests.*.json` / `sequences/*.json` in the tree, and no
 * `vba.dysflowExport: false` opt-out in `codegraph.json`.
 *
 * PER-FILE EXTRACTION DOES NOT LIVE HERE (issue #314). The three Dysflow
 * sub-extractors — `VbaFormExtractor`, `VbaTestManifestExtractor` and
 * `VbaTestSequenceExtractor` — are dispatched by the VBA branch of
 * `extractFromSource` in `../tree-sitter.ts`, gated on the `dysflowExport`
 * flag that `loadDysflowExportConfig()` reads from the project config. That
 * ladder is the ONLY dispatch site, and it runs on every path into
 * extraction, including the ones that pass no `frameworkNames` at all
 * (single-file re-index, library callers, the extraction unit tests).
 *
 * This module deliberately exposes no `extract()`. It used to carry one that
 * re-ran the same three sub-extractors, which meant every Dysflow artifact in
 * a real project was parsed twice per index — once by the `tree-sitter.ts`
 * ladder and once by the framework loop at the end of `extractFromSource` —
 * and every manifest/sequence entry landed twice in `unresolved_refs`. Adding
 * an `extract()` back here re-introduces that duplication; extend the
 * `tree-sitter.ts` ladder instead.
 *
 * Opt-out contract: `vba.dysflowExport: false` turns BOTH halves off. This
 * resolver's `detect()` returns `false` so the project is not classified as a
 * Dysflow export, and the same config flag makes the `tree-sitter.ts` ladder
 * emit a `file`-only node for the form/report/manifest/sequence shapes.
 */
import { FrameworkResolver, UnresolvedRef, ResolvedRef, ResolutionContext } from '../../resolution/types';
import { detectVbaFormFile, isVbaTestManifestFile, isVbaTestSequenceFile } from '../grammars';
import { loadDysflowExportConfig } from '../../project-config';

export const dysflowExportResolver: FrameworkResolver = {
  name: 'dysflow-export',
  // Scoped to VBA — these file shapes are only meaningful on top of a VBA
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
   * the tree — they'd be tracked as just a `file` node instead.
   *
   * The detection runs once at index start (via `detectFrameworks`), so a
   * `.form.txt`/`.report.txt`/manifest/sequence present anywhere in the
   * project tree qualifies it.
   */
  detect(context: ResolutionContext): boolean {
    try {
      if (!loadDysflowExportConfig(context.getProjectRoot())) {
        return false;
      }
    } catch {
      // Config load failure (very rare) — treat as opt-out so we don't
      // classify the project against an unverified config.
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
   * Reference resolution hook. The 3 Dysflow shapes emit `references`-kind
   * UnresolvedReferences whose targets are sibling `.cls` files (form
   * binding) or `Test_*` VBA function nodes (manifest + sequence). The
   * shared `ReferenceResolver` already binds these by name + language +
   * provenance, so this framework resolver doesn't need a custom
   * `resolve()` step — the standard resolution pipeline handles it.
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
   * detection time.
   */
};
