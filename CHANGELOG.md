# Changelog

All notable changes to CodeGraph are documented here. Each entry also ships as
a [GitHub Release](https://github.com/colbymchenry/codegraph/releases) tagged
`vX.Y.Z`, which is where most people will look.

This project follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/)
and adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### New Features

- VBA extraction rules can now explicitly inspect masked source, original source, or both while preserving procedure, type-block, and enum-block boundaries. (#165)
- The VBA extraction pipeline now has a published performance report on the real `00_VBA_TOOLKIT_BENCH` corpus. Set `CODEGRAPH_VBA_TIMING=2 codegraph index` to capture per-stage timings (preprocess / classifiers / walk) for every `.cls` and `.bas`, and compare across runs to spot slow files. `docs/vba-extraction-perf.md` shows the medians across three runs for the named fixtures, a corpus-level breakdown by stage, and an honest note on the v1.6.2 baseline. (#166)

### Fixes

- VBA extraction now refuses to load if any per-concern rule table is empty, so an accidentally-emptied classifier fails the import instead of silently dropping an entire symbol family. (#164)
- VBA `Public Const` and `Private Const` symbols now retain their declared `As` type alongside the value, with untyped constants reported as `Variant`. (#167)

## [1.11.0] - 2026-07-18

### New Features

- VBA extraction is now driven by a declarative rule table. Every `Sub` / `Function` / `Property`, `Implements`, `Dim`, `Const`, `Enum`, `Event`, `Type`, and `Declare` declaration is matched by an explicit `VbaExtractionRule` with a stable `id`, plain-English description, and isolated `emit` body — no more giant inline `if/else` cascades inside the per-concern sweepers. The new test suite enforces the table shape and a non-empty invariant per concern, so an accidentally-emptied rule set fails loudly at module load instead of silently dropping a whole symbol family. Pure structural refactor: zero behavior change, the full existing VBA test suite passes unchanged. (#153)
- The VBA extraction pipeline now reports per-stage wall-clock timing on demand, gated by the `CODEGRAPH_VBA_TIMING` env var. Set `CODEGRAPH_VBA_TIMING=1` for a per-file block on stderr (preprocess / classifier / walk stages, plus the inner conditional-compilation lexer+parser), or `=2` to add a per-process aggregate across the whole index run. With the env var unset the default path stays at zero overhead — no Map allocations, no Map writes — so existing `codegraph index` runs are unaffected. Useful for diagnosing why a particular `.bas` or `.cls` is slow to extract, or for tuning the extraction budget on a large Access project. (#156)
- VBA event-handler relationships are now materialized in the graph at index time. A `WithEvents m_X As ClassName` binding in a form combined with a matching `m_X_<EventName>` handler Sub is now connected to the `RaiseEvent <EventName>` site via a single `event-handler` edge, so `codegraph_explore` reaches the handler in one call instead of the three-hop walk the vba-event-tracer skill used to repeat on every query. Projects with no WithEvents bindings are unaffected; the FORMS-* test suite reports zero new edges. (#150)
- The Dysflow-specific VBA extractors — form/report SaveAsText, test manifests, and test sequences — are now opt-out-able. If your `.bas`/`.cls` files happen to live next to legacy `.form.txt`/`.report.txt` files (or test-manifest JSON from a different system) that you don't want expanded into the graph, set `vba.dysflowExport: false` in your project's `codegraph.json`; the Dysflow file types are then tracked as just a `file` node, while the rest of the VBA pipeline keeps behaving exactly as before. The Dysflow extractors also now live behind a `FrameworkResolver` you can discover alongside the other frameworks. (#154)
- A spike report on using the `tree-sitter-vbnet` grammar as a primary parser for VBA is now available. It classifies the per-construct failure modes of a representative Dysflow-exported corpus against the grammar (probes cover procedures, classes, `Dim`, `Implements`, `WithEvents`, `RaiseEvent`, `DoCmd`, `With`, control flow, and string-literal SQL), and includes a go/no-go recommendation for the next phase. (#155)

## [1.10.0] - 2026-07-16

### New Features

- Code that reads or writes form controls via `Me.` now shows up in control impact analysis. (#140)
- Form events that call a function directly (`=MyFunction()`) are now part of the graph, so those flows no longer dead-end. (#137)
- Subforms now link to the forms they embed, so parent-to-subform flows and "who embeds this form" queries work. (#136)
- Access form controls now link to the table columns they display, so column-rename impact reaches the control level. (#135)
- Access form and report controls now expose their section membership and direct layout containment, making UI structure available to graph queries. (#134)

### Fixes

- Helper Subs with underscores in form code-behind no longer appear as phantom form controls. (#139)
- Following `DoCmd.OpenForm` now reaches the opened form's code instead of stopping at a placeholder. (#138)
- Access controls with long property blocks are now indexed reliably, and report layouts use the same dedicated kind as opened-report references. (#134)

## [1.9.0] - 2026-07-15

### Added

- Added `codegraph_uninit` MCP tool (CLI subprocess wrapper around `codegraph uninit`). Destructive; off by default; enable via `CODEGRAPH_MCP_TOOLS=explore,uninit`.

### New Features

- Added `codegraph_affected` MCP tool (CLI subprocess wrapper around `codegraph affected`). Off by default; enable via `CODEGRAPH_MCP_TOOLS=explore,affected`. (#127)
- Added `codegraph_index` MCP tool (CLI subprocess wrapper around `codegraph index`). Idempotent rebuild; off by default; enable via `CODEGRAPH_MCP_TOOLS=explore,index`. (#129)

## [1.8.0] - 2026-07-14

### New Features

- Added the opt-in `codegraph_init` MCP lifecycle tool (CLI subprocess wrapper around `codegraph init`). Off by default; enable via `CODEGRAPH_MCP_TOOLS=explore,init`. (#121)
- Added the opt-in `codegraph_sync` MCP lifecycle tool for synchronizing a project's full index, including project selection, quiet mode, process output, and exit-code reporting. (#122)
- Added the opt-in, read-only `codegraph_query` MCP tool for structured CLI symbol queries. (#123)

## [1.7.3] - 2026-07-13

### Fixes

- Orphan npm staging dirs from any previous `npm install -g` failure (including pre-fix-era upgrades) are now cleaned automatically on every install via a `postinstall` script. No more EPERM noise on the next upgrade; no more manual cleanup of `.codegraph-vba-<HASH>` leftovers.
- VBA `RaiseEvent` sites are no longer graphed for events with too many raise sites in the same file. Events with names like `AfterUpdate`, `Click`, or `Open` can be raised from hundreds of sites in a single form, producing edges that drown out the meaningful ones; the graph now stamps those event nodes with `metadata.highFanout` and a count, and drops the noisy edges, so the event itself (and its handler linkage) stays reachable while the noise is suppressed. The threshold defaults to 50 and is configurable via `vba.maxRaiseFanout` in `codegraph.json`. (#152)

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
[1.7.3]: https://github.com/colbymchenry/codegraph/releases/tag/v1.7.3
[1.8.0]: https://github.com/colbymchenry/codegraph/releases/tag/v1.8.0
[1.9.0]: https://github.com/colbymchenry/codegraph/releases/tag/v1.9.0
[1.10.0]: https://github.com/colbymchenry/codegraph/releases/tag/v1.10.0
[1.11.0]: https://github.com/colbymchenry/codegraph/releases/tag/v1.11.0
