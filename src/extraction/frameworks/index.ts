/**
 * Extraction-side framework resolvers.
 *
 * `src/extraction/frameworks/` is the home for `FrameworkResolver`s whose
 * primary contribution is per-file extraction (nodes + references), as
 * opposed to `src/resolution/frameworks/` whose resolvers are primarily
 * reference resolvers. The two registries are kept separate so a future
 * "extraction-only" framework (e.g. another format-specific generator)
 * doesn't drag in the full resolution pipeline, and so the namespacing
 * matches the layer the resolver actually plugs into.
 *
 * Currently only the Dysflow export resolver lives here ÔÇö the 3
 * Dysflow-specific extractors (form/report, test manifest, test
 * sequence) were lifted out of `tree-sitter.ts` into this `FrameworkResolver`
 * shape so the project can opt out via `codegraph.json`
 * (`vba.dysflowExport: false`, issue #154). New extraction-side frameworks
 * (e.g. a future "terraform state" generator) should be added here.
 */
export { dysflowExportResolver } from './dysflow-export';
