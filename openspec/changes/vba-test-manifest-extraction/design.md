# Design: VBA Test Manifest Extraction

## Technical Approach

Add a standalone `VbaTestManifestExtractor` (sibling of `VbaFormExtractor` /
`SqlQueryExtractor`) that parses Dysflow test manifest JSON and links each
registered test atom to its manifest metadata. It emits **no** duplicate node
for the test procedure — the `Test_*` `function` node already exists from the
VBA module extraction. Instead the manifest emits `UnresolvedReference`s that
the `ReferenceResolver` binds to those existing function nodes by bare name,
exactly as `DoCmd.OpenQuery` (`synthesizedBy: 'vba-opens-query'`) binds to the
real `query` node today. The mapping symbol → test then falls out of the call
graph codegraph already has: `getCallers(X)` reaches the test atoms, and each
atom's incoming `vba-test-manifest` edge names the manifest + tags to run.

## Manifest Shape (ground truth)

From `00_GESTION_RIESGOS_staging/tests/tests.vba.*.json`:

```json
{
  "_comment": "optional",
  "tests": [
    { "name": "Presenter Happy 5 Anexos 5 Filas",
      "procedure": "Test_Presenter_Happy_5Anexos_5Filas",
      "expect": { "ok": true },
      "tags": ["presenter", "happy", "b2-punto-15"] }
  ]
}
```

`procedure` is required (the VBA atom). `name`, `tags`, `expect` are optional.
`sequences/*.json` are a DIFFERENT shape (`runnerPolicy` + `procedures` string
array) and are OUT OF SCOPE here (follow-up).

## Architecture Decisions

| Decision | Choice | Alternatives | Rationale |
|---|---|---|---|
| File detection | Basename regex `^tests(\.[\w-]+)*\.json$` **AND** parsed JSON has a top-level `tests` array whose items carry a string `procedure` | Match all `.json`; match by directory (`tests/`) | Filename alone would catch `tests.json` fixtures unrelated to VBA; content gate makes it precise and avoids `package.json`/`tsconfig.json`. |
| Node model | Reuse the existing `Test_*` `function` node; manifest emits only the `file` node + references | Emit a new `test`/`test-atom` node kind per entry | Avoids duplicating a node that already exists; keeps the graph single-source-of-truth for the procedure. |
| Edge model | `references` edge, `metadata.synthesizedBy: 'vba-test-manifest'`, `metadata.{testName,tags,manifestFile}` | Add a new `tests` EdgeKind | Mirrors how table/binding references are modeled as tagged `references`; no schema/enum churn. Revisit only if a distinct edge kind proves necessary for queries. |
| Resolution | `UnresolvedReference{ referenceName: procedure }` bound by name via the existing resolver path (as `vba-opens-query`) | Bind at extraction time | Extractors are per-file; the test module lives in another file. Name resolution is the resolver's job and already exists. |
| Failure mode | Malformed JSON / missing `procedure` → low-severity `ExtractionError`, zero references, never throw | Fail the file / crash | Consistent with every other extractor; a bad manifest must not break the index. |
| Unresolved procedure | Leave the reference unresolved (dangling) | Drop silently | A manifest naming a renamed/absent procedure is real drift — surfacing it structurally mirrors `validate_manifest`. |

## Data Flow

```text
tests.vba.<slice>.json
  -> isVbaTestManifestFile(path)  (grammars.ts)  gates routing
  -> tree-sitter.ts dispatch -> new VbaTestManifestExtractor(filePath, source)
       -> JSON.parse (guarded) -> file node
       -> for each tests[] entry with a string `procedure`:
            UnresolvedReference{
              referenceName: <procedure>,
              referenceKind: 'references',
              metadata: { synthesizedBy: 'vba-test-manifest', testName, tags, manifestFile }
            }
  -> ReferenceResolver binds referenceName -> Test_* function node (by name)
       -> references edge: manifest file node -> Test_* function node (metadata carried)
  -> queries:
       codegraph_explore(Test_*)  -> shows manifest + tags
       getCallers(productionSymbol) -> reaches Test_* atoms whose incoming
                                       vba-test-manifest edge names manifest+tags
```

## File Changes

| File | Action | Description |
|---|---|---|
| `src/extraction/vba-test-manifest-extractor.ts` | **New** | `VbaTestManifestExtractor` class: parse guarded JSON, emit file node + `UnresolvedReference`s. |
| `src/extraction/grammars.ts` | Modify | Add `isVbaTestManifestFile(path)`; mark these files as source files so the orchestrator indexes them. |
| `src/extraction/tree-sitter.ts` | Modify | Add dispatch branch (~line 6585) routing manifest files to `VbaTestManifestExtractor` **before** the generic `.json` skip. |
| `src/resolution/*` (name-matcher / import-resolver) | Modify | Resolve `synthesizedBy: 'vba-test-manifest'` references by bare procedure name to `function` nodes (reuse/extend the `vba-opens-query` name-binding path). |
| `src/mcp/server-instructions.ts` | Modify | One line: test manifests link `Test_*` atoms to their manifest + tags. |
| `__tests__/extraction-vba-test-manifest.test.ts` | **New** | RED-first coverage (see tasks). |
| `CHANGELOG.md` | Modify | `[Unreleased]` New Features entry. |
| `~/.claude/skills/vba-source-impact/SKILL.md` | Modify (out-of-repo) | `related_tests` now filled from the graph in one call. |
| `~/.claude/skills/vba-run-tests/SKILL.md` | Modify (out-of-repo) | Note codegraph can name the manifest/tags to run for a changed symbol. |

## Interfaces / Contracts

```ts
// grammars.ts
export function isVbaTestManifestFile(filePath: string): boolean;
// true iff basename matches /^tests(\.[\w-]+)*\.json$/i
// (content gate applied inside the extractor)

// vba-test-manifest-extractor.ts
export class VbaTestManifestExtractor {
  constructor(filePath: string, source: string);
  extract(): ExtractionResult; // file node + UnresolvedReference[] (+ errors)
}
```

Reference contract (per test entry):
- `referenceName`: the `procedure` string (bare VBA proc name).
- `metadata.synthesizedBy`: `'vba-test-manifest'`.
- `metadata.testName`: the entry `name` (or `procedure` when absent).
- `metadata.tags`: `string[]` (empty when absent).
- `metadata.manifestFile`: the manifest file path.

Resolved edge: `kind: 'references'`, source = manifest file node, target =
`Test_*` `function` node, metadata carried through.

## Risks & Mitigations

| Risk | Mitigation |
|---|---|
| Over-matching non-manifest `.json` | Filename regex + content shape gate (must have `tests: [{procedure:string}]`). |
| Test modules not in the index → all references dangle | Documented; degrade gracefully; the staging repo indexes its test modules, so validate there. |
| Malformed / partial manifest | Guarded `JSON.parse`; low-severity error; emit what parses. |
| A new EdgeKind tempting scope creep | Explicitly reuse `references`; only revisit with evidence. |
| Watcher churn on manifest edits | Manifests are small and change rarely; the incremental `sync` path already handles per-file re-extraction. |

## Validation

On `00_GESTION_RIESGOS_staging` (20+ manifests, real):
1. Deterministic probe: index; confirm each `tests.vba.*.json` emits a file node
   + N resolved `vba-test-manifest` edges to `Test_*` nodes; count stable on
   re-index; no non-manifest `.json` produces edges.
2. Workflow probe: pick a production symbol; `getCallers` → covering `Test_*` →
   read incoming `vba-test-manifest` edges → obtain `{testName, tags,
   manifestFile}` with zero JSON grep.
