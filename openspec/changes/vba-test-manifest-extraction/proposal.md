# Proposal: feat(vba): parse Dysflow VBA test manifests into the graph

## Intent

Let codegraph-vba parse Dysflow's VBA **test manifests** (`tests.vba*.json` and
sibling `tests.*.json`) so the knowledge graph links each registered test atom
to its manifest (test name, tags, source manifest file). Combined with the call
edges codegraph already extracts from the test modules, this answers a frequent
TDD-workflow question in **one graph query** instead of a manual grep across
dozens of JSON files:

> "I'm changing production symbol `X`. Which manifest(s) / tags do I run to
> cover it?"

This is the read/map counterpart to Dysflow's run/validate role: **Dysflow runs
manifests (`test_vba`) and validates them (`validate_manifest`); codegraph maps
which tests relate to a changed symbol.** Neither Dysflow tool maps
symbol → test today, because that mapping lives in the call graph — codegraph's
home turf.

## Problem

- A manifest entry maps a human `name` → a VBA `procedure` (a `Test_*` atom) +
  `tags`. The tags and the *which-manifest-to-run* fact live **only in the
  JSON**, never in the VBA code.
- The `vba-source-impact` skill already wants to emit
  `related_tests: [{ name, procedure, tag }]`. Today codegraph finds the
  `Test_*` procedures via call edges (giving `procedure`), but the agent must
  then **grep the 20+ manifest files** to recover `name`, `tag`, and the
  manifest path to pass to `test_vba`.
- Dysflow's `validate_manifest` / `test_vba` never expose symbol → test
  coverage; they consume the manifest as a run list.

## Scope

### In Scope
- A new standalone extractor `VbaTestManifestExtractor` for manifest JSON files.
- Filename + content gating so ONLY Dysflow test manifests are parsed (never
  `package.json`, `tsconfig.json`, etc.).
- Emit a `file` node for the manifest and, per test entry, a `references` edge
  (tagged `synthesizedBy: 'vba-test-manifest'`) toward the named `procedure`,
  carrying `metadata.testName` and `metadata.tags` and `metadata.manifestFile`.
- Resolve those references to the existing `Test_*` `function` nodes by bare
  name (reuse the `UnresolvedReference` → resolver path that `DoCmd.OpenQuery`
  already uses).
- Docs/skills alignment: `server-instructions.ts`, `vba-source-impact`,
  `vba-run-tests`, and a CHANGELOG entry (this change alters agent-visible
  behavior).

### Out of Scope
- Running or validating tests (Dysflow `test_vba` / `validate_manifest` own it).
- Inferring which production symbols a test covers beyond the call edges
  codegraph already extracts from the test module bodies.
- Propagating tags onto production symbols.
- `sequences/*.json` orchestration files (different shape: `runnerPolicy` +
  `procedures` string array) — deferred to a follow-up (see tasks / sub-issue).

## Capabilities

### New Capabilities
- `vba-test-manifest-extraction`: parse Dysflow VBA test manifests and link
  registered test atoms to their manifest metadata and (via existing call
  edges) to the production symbols they exercise.

### Modified Capabilities
None (additive — no change to existing VBA code/form/SQL extraction).

## Approach

1. Add manifest detection to the extraction file-type routing: a file whose
   basename matches `tests(.<segment>)*.json` AND whose parsed JSON has a
   top-level `tests` array of objects carrying a string `procedure`.
2. `VbaTestManifestExtractor.extract()`:
   - Emit the manifest `file` node.
   - For each `tests[]` entry with a `procedure`, push an `UnresolvedReference`
     to that procedure name, tagged `synthesizedBy: 'vba-test-manifest'`, with
     `metadata: { testName, tags, manifestFile }`.
   - Tolerate malformed JSON: catch, emit a low-severity extraction error, emit
     zero references (never crash the index).
3. The `ReferenceResolver` binds each unresolved test reference to the real
   `Test_*` `function` node by name, producing a `references` edge
   (manifest file → test function). Unresolved (renamed/missing procedure)
   references degrade gracefully — surfaced as a dangling test, which is exactly
   the drift `validate_manifest` warns about.
4. `codegraph_explore` then surfaces, for any `Test_*`, its manifest + tags; and
   for a production symbol `X`, `getCallers(X)` reaches test atoms whose
   incoming `vba-test-manifest` edge names the manifest + tags to run.
5. Align docs/skills so the capability is discoverable and the impact skill can
   fill `related_tests` in one call.

## Success Criteria

- Given a changed production symbol, an agent obtains `{ testName, procedure,
  tags, manifestFile }` for the covering tests in one codegraph query (no JSON
  grep).
- Parsing is precise: no non-manifest `.json` file produces nodes/edges.
- No regression in existing extraction suites; malformed manifests never crash
  the index.
- Validated on the real `00_GESTION_RIESGOS_staging` repo (20+ manifests).
