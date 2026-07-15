# Changelog

All notable changes to CodeGraph are documented here. Each entry also ships as
a [GitHub Release](https://github.com/colbymchenry/codegraph/releases) tagged
`vX.Y.Z`, which is where most people will look.

This project follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/)
and adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- Added `codegraph_index` MCP tool (CLI subprocess wrapper around `codegraph index`). Idempotent rebuild; off by default; enable via `CODEGRAPH_MCP_TOOLS=explore,index`.

### Fixes

- Orphan npm staging dirs from any previous `npm install -g` failure (including pre-fix-era upgrades) are now cleaned automatically on every install via a `postinstall` script. No more EPERM noise on the next upgrade; no more manual cleanup of `.codegraph-vba-<HASH>` leftovers.

## [1.7.1] - 2026-07-13

### Documentation

- The post-extraction VBA stub resolver's actual contract is now documented for consumers. `metadata.repointDecision` carries one of `reponted-to-real`, `declined-runtime`, `declined-ambiguous`, or `declined-not-found` — consumers detecting "missing callees" must filter on `repointDecision='declined-not-found'`, NOT on the raw `stub=true` count (which is dominated by runtime-object noise from `DAO.*`, `fso.*`, etc.). The original round-5 prompt's `stub_true_count < 500` acceptance criterion was replaced by the `declined-not-found` filter. Reference: `docs/vba-stub-repoint-decision.md`. (#115)

## [1.7.0] - 2026-07-13

### Fixes

- `codegraph-vba upgrade` no longer leaves the install stuck on the previous version when npm hits a stale cache, EPERM mid-install, or a leftover orphan staging dir. The upgrade now sweeps orphaned `.codegraph-vba-<HASH>` dirs from npm's global `node_modules` before every run, passes `--prefer-online` to bypass the local npm cache, verifies the installed `package.json` actually reflects the target version after npm reports success, and falls back to installing directly from the registry tarball URL when it doesn't — so a silent stale install becomes visible (and recoverable in the same command) instead of leaving `codegraph --version` reporting a phantom version.
- VBA's `unresolved_refs` table now reports the syntactic shape of each unresolved reference — call sites, form-property reads and writes, `DoCmd.OpenQuery` arguments, and bare identifiers each get their own row kind — so a SQL filter for "missing callee" stops drowning in DAO-field and form-control noise. The legacy `references` kind is preserved for any path the round does not reclassify, so SQL filters that key on it keep working. (#108)
- VBA's post-extraction call-stub resolver now correctly declines runtime-object call stubs (`DAO.*`, `fso.*`, `ListBox.AddItem`, `Collection.Add`, `err.*`, `VBA.*`, …) instead of pointing them at themselves; a user class or module that happens to share a runtime-object name is still linked. Stubs now carry a `repointDecision` field on their edge metadata so consumers can tell a runtime-object decline apart from a genuinely-missing callee. (#110, supersedes #109)

## [1.6.2] - 2026-07-12

### New Features

- VBA SQL table references now record whether the code reads or writes each table — derived from the SQL verb (`SELECT`/`FROM` and `JOIN` are reads; `INSERT`, `UPDATE`, and `DELETE` targets are writes; a form's `RecordSource`/`RowSource` binding is a read) — so you can ask which procedures *write* a table versus only read it. (#87)
- VBA object variables assigned from a factory function in the same module (`Set x = CreateThing()` where `CreateThing` returns a class) are now typed from that function's return type, so calls like `x.DoWork` connect to the factory's class instead of dead-ending — more complete call graphs and impact analysis for factory-style Access code. (#89)
- Dysflow VBA test manifests (`tests.*.json`) are now indexed: each registered `Test_*` procedure is linked to its manifest, so you can ask which tests cover a changed symbol — and get the test names and tags to run — straight from the graph, without grepping the manifest files. (#91)
- Dysflow VBA test sequences under `sequences/*.json` are now indexed: each `Test_*` procedure listed in a sequence's `procedures[]` is linked back to the sequence file, with the sequence's `runnerPolicy`, position, and origin carried on the edge — so you can ask which sequence exercises a procedure and what runner policy applies, straight from the graph. The `runnerPolicy` + `procedures` shape is the only one supported today; strict-sequence (`executionUnits`) and slices (`slices[]` + submanifests) shapes are deferred to a follow-up. (#97)
- Indexing an Access/VBA project is faster on large exports: the extractor no longer re-splits the source into lines once per concern. Five independent line-by-line sweeps now share a single pre-tokenized line array and a single walker dispatches each line to its classifier, so extraction of a ~2900-line `.cls` runs ~10% faster at the same accuracy as before. (#83)

### Fixes

- VBA conditional-compilation expressions (`#If … Then`) with mixed arithmetic and comparison, e.g. `#If 2147483647 + 1 = -2147483648 Then`, are now evaluated correctly. The Pratt parser previously had no binary additive level, so expressions mixing `+`/`-` with `=` silently failed to parse and the inactive branch was indexed. (#84)

## [1.6.0] - 2026-07-11

### Changed

- **Synced with upstream `colbymchenry/codegraph` v1.4.1** (107 upstream commits merged into the fork). All VBA/Access extraction — its own files, resolvers, and 9 test suites — was preserved unchanged; conflicts occurred only at the integration seams and were resolved keeping the fork's VBA behavior while adopting upstream's improvements.

### Fixed

- Closed ArkTS test database handles before removing temporary directories, preventing Windows `EBUSY`/`EPERM` teardown failures.
