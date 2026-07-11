# Tasks: VBA Test Manifest Extraction

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | 250–380 (new extractor + routing + resolver hook + tests) |
| 400-line budget risk | Low–Medium |
| Chained PRs recommended | Optional (2 slices if kept lean) |
| Suggested split | Slice A: extractor + routing + tests; Slice B: resolver binding + docs/skills |
| Delivery strategy | ask-on-risk |

Decision needed before apply: No (single PR fits; split only if it grows)

### Suggested Work Units → Sub-issues

| Unit | Goal | Maps to sub-issue |
|------|------|-------------------|
| 1 | File detection + extractor skeleton (file node, guarded parse) | SUB-1 |
| 2 | Emit `vba-test-manifest` references with metadata (testName/tags/manifestFile) | SUB-2 |
| 3 | Resolver binds references → `Test_*` function nodes by name | SUB-3 |
| 4 | Docs/skills alignment (server-instructions, vba-source-impact, vba-run-tests, CHANGELOG) | SUB-4 |
| 5 | Validate on `00_GESTION_RIESGOS_staging` (deterministic + workflow probes) | SUB-5 |
| 6 | (Deferred) `sequences/*.json` support | SUB-6 |

## Phase 1: RED — Failing Tests (new `__tests__/extraction-vba-test-manifest.test.ts`)

- [ ] 1.1 Manifest recognized → file node emitted (`tests.vba.smoke.json` shape).
- [ ] 1.2 `package.json` / `tsconfig.json` → no nodes/references.
- [ ] 1.3 `tests.*.json` without a `tests[{procedure}]` array → no references.
- [ ] 1.4 Three entries → three `vba-test-manifest` UnresolvedReferences with `testName`/`tags`/`manifestFile`.
- [ ] 1.5 Entry without `name`/`tags` → `testName === procedure`, `tags === []`.
- [ ] 1.6 Malformed JSON → low-severity error, zero references, no throw.
- [ ] 1.7 (Resolver, integration) manifest + module with `Test_X_RunAll` → resolved `references` edge file→function with metadata.
- [ ] 1.8 (Resolver) missing procedure → reference stays unresolved, no edge.

## Phase 2: GREEN — Implementation

- [ ] 2.1 `grammars.ts`: add `isVbaTestManifestFile(path)`; mark as source file.
- [ ] 2.2 `src/extraction/vba-test-manifest-extractor.ts`: guarded `JSON.parse`; file node; per-entry `UnresolvedReference` with metadata; content-shape gate.
- [ ] 2.3 `tree-sitter.ts`: dispatch branch to `VbaTestManifestExtractor` before the generic `.json` skip.
- [ ] 2.4 Resolver (`name-matcher.ts` / `import-resolver.ts`): bind `synthesizedBy: 'vba-test-manifest'` references to `function` nodes by bare name (reuse the `vba-opens-query` path); keep unresolved when absent.
- [ ] 2.5 Confirm `copy-assets`/build unaffected (pure TS + JSON; no wasm).

## Phase 3: REFACTOR

- [ ] 3.1 Extract the content-shape guard into a small pure helper; keep the extractor a thin orchestrator (project convention).

## Phase 4: Docs / Skills alignment (REQUIRED — capability change)

- [ ] 4.1 `src/mcp/server-instructions.ts`: one line — test manifests link `Test_*` atoms to manifest + tags.
- [ ] 4.2 `CHANGELOG.md` `[Unreleased]` New Features entry.
- [ ] 4.3 `~/.claude/skills/vba-source-impact/SKILL.md`: `related_tests` filled from the graph in one call.
- [ ] 4.4 `~/.claude/skills/vba-run-tests/SKILL.md`: codegraph names the manifest/tags to run for a changed symbol.

## Phase 5: Validation

- [ ] 5.1 Deterministic probe on `00_GESTION_RIESGOS_staging`: per-manifest file node + resolved edges; node/edge count stable on re-index; zero non-manifest `.json` output.
- [ ] 5.2 Workflow probe: production symbol → `getCallers` → covering `Test_*` → `{testName,tags,manifestFile}` with zero JSON grep.
- [ ] 5.3 Full `npm test` + `tsc` green; no regression in existing suites.

## Phase 6 (Deferred): sequences/*.json

- [ ] 6.1 Separate change — parse `sequences/*.json` (`runnerPolicy` + `procedures[]`) into a suite/sequence node linking its procedures. Do NOT bundle here.
