# Changelog

All notable changes to CodeGraph are documented here. Each entry also ships as
a [GitHub Release](https://github.com/ardelperal/codegraph-vba/releases) tagged
`vX.Y.Z`, which is where most people will look.

This project follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/)
and adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]


## [1.17.2] - 2026-09-12

### Fixes

- Fork documentation now describes v1.17.1 capabilities and pins the integrated upstream v1.4.1 lineage, replacing stale version, feature-parity, and implementation-plan claims. (#312)
- Dysflow exports are no longer read twice. Every form and report layout, test manifest and test sequence was being parsed once by the VBA dispatcher and again by the Dysflow framework hook, which made a full index do double the work on those files and left duplicate pending references behind. Turning Dysflow expansion off with `vba.dysflowExport: false` is also honoured again in every case. (#314)
- A test manifest or test sequence saved with a UTF-8 byte-order mark is read correctly again. Until now such a file was rejected outright, so every test it registered quietly vanished from the graph with nothing but a warning to show for it — and on Windows that mark is what PowerShell writes by default. (#316)
- Saved Access queries now appear in the graph no matter what order Dysflow writes them. The `queries.json` manifest is what marks the `.sql` files beside it as saved queries, and a manifest arriving after its queries left them unindexed until the next manual sync. Editing only the manifest — renaming a query, adding an entry — now re-indexes too. (#318)
- Agents are now told about three Access surfaces the index already carried but never advertised: saved queries and the tables they touch, Dysflow test sequences alongside the manifests, and the table-and-column structure from the Access ERD export. An agent that was not told these exist fell back to reading the files by hand. (#320)
- Table and column names from an Access structure export keep their accents. The export is read with the same encoding rules as the rest of an Access project, so a table like `TbSituación` is no longer recorded under a corrupted name — which had also stopped it from lining up with the queries and code that reference it, leaving its column list attached to nothing. (#322)

## [1.17.1] - 2026-09-12

### Fixes

- Behavior evidence now bounds traversal work as well as returned results, preventing branched call graphs and duplicate paths from exhausting the request budget. Incomplete traversal is reported explicitly. (#309)
- Handlers can now be selected by form or report layout filename as well as layout name, including expression functions in separate modules. Ambiguous matches remain explicit. (#310)
- Publishing a release no longer reports failure for a release that shipped correctly. The final check that each package reached the npm registry gave up after a little over three minutes and threw away its own last wait, while a package of this size can take longer than that to become available — so a perfectly good release was marked red. It now waits proportionately and checks once more at the end, while a package that genuinely never arrives still fails the release. (#307)

## [1.17.0] - 2026-09-12

### New Features

- Access projects can now ask what a single control actually does and get a structured answer back: which procedure runs on the event, the call paths under it, and the tables it reads or writes along the way — assembled in one read instead of joined together by hand. Available from the library as `getBehaviorEvidence` and over MCP as `codegraph_behavior_evidence`. (#299)
- That answer is deliberately careful about what it does not know. A control name that exists on more than one form is refused with the candidates listed rather than answered for the wrong form, branches are kept apart instead of being strung into one sequence, and anything the index could not resolve — or a budget that cut the answer short — is reported alongside it. An empty list of tables means nothing is recorded, never that the code is harmless. (#299)
- A repeatable acceptance check for Access answers: `npm run acceptance:vba` indexes a fresh copy of a small checked-in Access export and asks it the questions a consumer really asks — which handler runs on this control, which tables it reaches, which form a query affects when nothing binds it to the screen — then compares the answers against expectations written by hand from the source. It fails on a missing answer and equally on an unrelated one, such as a second form's button of the same name. The same command can be pointed at an authorized copy of your own export, locally, without anything leaving your machine. (#301)

### Fixes

- The Access guidance agents actually receive — the README and the instructions the MCP server sends on connect — now describes the form and report model the indexer really produces, instead of the older one it outgrew. It also states the parts that were never written down: an event binding is stored in one direction only, a control belongs to the form that contains it (so the same button name on two forms is two different answers), and the most common way of calling a procedure in VBA does not look like a call. (#300)
- The Access documentation now says plainly that CodeGraph reads the exported source tree, not the live Access file: nothing here proves a handler ran, and a missing relationship is missing evidence rather than proof that nothing happens. The worked example in the README is executed as a test against a checked-in fixture, so following it cannot silently stop working. (#300)
- The event-tracing guidance now separates custom VBA events from Access control events, which are wired differently and were easy to confuse. (#300)
- Tracing an Access form control now reaches its event handler and everything that handler calls. Starting a trace at a button used to come back empty, because the trace followed the wiring between a control and its handler in the wrong direction, and because the common way of calling a procedure in VBA — writing its name on a line by itself — was not being counted as a call at all. (#298)
- Impact analysis for a saved query now names the forms and reports affected by it even when nothing binds the query to the screen directly, so a query reached only through code no longer looks harmless. Ownership of a control is taken from the form that contains it, so two forms with a button of the same name no longer answer for each other. (#298)
- A trace no longer pads its answer with relationships that are not steps in the code's execution, such as what contains a procedure or which table it reads, and an unknown starting point now says so instead of returning an empty result that reads like "nothing calls this". (#298)
- Release notes published to GitHub now carry every section of the changelog. A section whose heading was more than one word — such as the list of new features — could be dropped from the published notes without any warning, which is how v1.16.0 shipped with its entire feature list missing. Those entries have been restored to the changelog. (#296)
- Each version heading in the changelog now links to this project's own release instead of a different project's, where the tag does not exist. (#296)
- Publishing a release now refuses to run from anywhere but the project's main branch, so the changelog update it makes can no longer be left behind on a temporary branch and lost. (#296)

## [1.16.0] - 2026-09-03

### New Features

- Module-level variables are now part of the graph, so shared state can be traced to the procedures that read and write it. (#251)
- A form's own lifecycle handlers — what runs when it opens, loads or closes — now connect to the form in the graph, the same way a button's click handler already did. (#247)
- VBA procedures now carry their real shape in the graph — whether something is a Sub, a Function or a property getter or setter, what it returns, and which parameters it takes, are optional or are passed by reference. (#250)
- SQL that creates, alters or drops a table now counts as a write to that table, and a query pointed at another Access database file with `IN "…"` now shows that file in the graph, so cross-backend access is finally visible. (#256)
- Access code that runs a macro, opens a table, exports a spreadsheet or emails a report now records which object it names, so those relationships show up in the graph instead of disappearing. (#254)
- Closing a form or report from code is now part of the graph, so a form's full lifecycle — who opens it and who closes it — is visible in one place. (#246)
- Opening a recordset or a query definition by its saved-query name now links the calling procedure to that query, so you can follow a procedure through the query it runs all the way to the tables it touches. (#253)
- Every parameter a VBA procedure declares is now its own symbol in the graph, linked to the procedure that owns it and to the class it is typed as, so you can ask which procedures take a given class and follow it from there. (#257)
- SQL that a form or control is bound to at runtime — assigned to `RowSource`, `RecordSource`, `ControlSource`, `Filter` or `OrderBy` — now shows the tables it reads, whether the statement is written inline or built up in a variable first. (#252)
- Access domain functions such as `DLookup` and `DCount` now record the table or saved query they read, so a procedure that touches data only through them no longer looks like it touches no data at all. (#255)
- Procedures now record how they handle errors — whether they have a handler, silently suppress errors, or have no protection at all — so unguarded code paths can be found without reading every module. (#259)
- An Access project that ships a generated structure export of its backend now gets real tables in the graph, with their columns, and linked tables point at the external database file they actually live in — so a table found from a query is the same table that carries its fields. (#257)
- Work a procedure does only when something goes wrong is now marked as such, and each error handler records whether it records the message, shows it to the user, re-raises it, or does more than one of those — so a procedure's failure path can be told apart from its normal one. (#260)
- Each error handler in Access code is now its own symbol you can search for and jump to, linked to the procedure that routes errors to it, so a handler can be found and followed directly instead of only asking whether a procedure has one. (#263)
- The variables this kind of Access project uses to pass error messages between objects are now recognised as such, so an error can be followed from where it happened to where the user sees it. (#261)
- New guidance shows how to find Access code that runs without error handling, including procedures with more than five executable statements that touch SQL tables, `DoCmd` targets, filesystem functions, or exact `Kill`/`Open`/`Close` statements, plus suppression scopes left open, missing handler labels, and where errors reach the user. (#262)

### Changed

- Project-scoped CodeGraph resources can now be released non-interactively through the CLI or the opt-in MCP tool before deleting a worktree, without stopping unrelated projects. (#234)
- VBA extractor comments now describe the behavior they guard, and the rule-table invariant also validates the procedure pre-walk so future drift fails loudly. (#215)
- VBA classifiers now share one rule dispatcher with consistent scan modes, structural gates, counting, and declarative terminal behavior, preventing future rule tables from silently bypassing their documented preconditions. (#216)
- VBA extraction now exposes only the classifier factories used by the shared source walker, removing obsolete compatibility entry points and other unused public helpers. (#217)

### Tests

- Removed documentation-prose and tautological assertions while retaining behavioral rule-table validation. (#214)

### Fixes

- Initializing or indexing a project now reliably releases its database and parser workers after both successful runs and unexpected failures, preventing stale Windows file locks from blocking worktree removal. (#241)
- A VBA module that declares its own variable named `Error` no longer reports every `On Error GoTo` line as a read of it, so error-handling statements stop showing up as data access. (#292)
- VBA: a call written with the `Call` keyword, or with an argument list, is now reported as a call rather than as an ambiguous bare-identifier read, so a genuinely missing procedure is no longer filtered out by the constant-lookup rules meant for plain identifier reads. A bare name with no `Call` keyword and no arguments still counts as an identifier read, because it really can be a constant. (#265)
- A form or report whose file was saved under a different name than the module itself carries now gets its event handlers and control references wired up, instead of quietly coming through with none of them. (#249)
- VBA projects that reach their database through a named accessor — one per backend, for example — no longer lose the tables those queries touch, and recordsets opened with parentheses are now tracked like every other query. Projects whose accessors follow their own naming can list them in the project configuration. (#244)
- Calls into Access and VBA built-ins no longer create thousands of phantom symbols that show up in search results and node counts. (#245)
- Class setup and teardown routines in Access class modules are now recognised; the previous check looked for a constructor spelling that only exists in VB.NET, so it never matched real VBA code. (#248)
- VBA references mentioned only inside messages, logs, and other string literals no longer create false form, query, or temporary-variable relationships. (#209)
- VBA SQL extraction now ignores reserved words exposed by dynamic table-name concatenation instead of emitting misleading table references. (#203)
- VBA continued statements are now parsed as one logical line while retaining their original source locations, restoring class references, qualified calls, procedure headers, and API declarations split with line continuations. (#202)
- VBA SQL variables now stay within their procedure while module-level SQL remains available as a fallback, preventing same-named variables in separate procedures from creating false table references. (#204)
- VBA: colon-separated single-line procedures now close their scope on the same physical line and scan calls in their inline body, preventing later module constants from being misclassified and making procedure tracking consistent regardless of header placement. (#208)
- VBA: a local variable declared with `Dim` inside a procedure now keeps that procedure's type for the lifetime of the procedure instead of being silently overwritten by the last `Dim` of the same name anywhere in the file, so a `item.Guardar` call inside `Sub AltaProducto` (where `Dim item As Producto` was declared) no longer resolves to the `Cliente` class declared in `Sub AltaCliente` later in the same file. (#205)
- VBA conditional-compilation directives no longer silently delete a branch: an unterminated `#If` now emits a warning and restores the affected lines, and expressions the lexer cannot handle (string literals, floating-point, unknown syntax) are treated as `could-not-evaluate` and kept active by default. (#206)
- VBA `Me.Name` and `Me!Field` references inside comparisons, `MsgBox` arguments, and `IIf` expressions are now classified as reads (`property-get` / `bang-get`) instead of writes, while statement-form assignments such as `Me.Name = "X"` and `Me!Field = 0` still classify as writes; all three reference-classification branches now share one direct-assignment predicate. (#211)
- VBA extraction no longer returns a graph fragment with a dangling edge when a `.bas` / `.cls` parse throws: malformed source still produces a `parse_error` and a valid (possibly partial) result, but edges that were being held pending module/class attribution are now dropped on the throw path instead of leaking into the returned graph. (#213)
- VBA `Declare`, `Event`, `Enum`, and `Type` header lines no longer leak phantom parameter-type references into the graph or pollute the qualified-call gate with parameter names, so the symbols these headers describe (real project classes, Win32 APIs, custom events, UDTs) keep their declared shape instead of being mistaken for project-local variables. (#207)

## [1.15.0] - 2026-07-20

### New Features

- External integrations now have a stable headless CLI contract and a generated index-schema reference, making it safer to build tools on CodeGraph without depending on undocumented internals. (#200)

## [1.14.0] - 2026-07-20

### Fixes

- Projects built with an older engine are now correctly flagged for re-indexing instead of silently going stale, and `codegraph status --json` exposes a `reindexReasons` array so scripts can detect a missed update without parsing the prose warning. (#189)
- VBA intrinsic constants such as `vbCrLf`, `vbTab`, `vbYesNo`, and `vbExclamation`, plus DAO option flags such as `dbFailOnError`, `dbSeeChanges`, `dbReadOnly`, and `dbAppendOnly`, no longer appear as failed references when they appear as bare identifiers; user-defined symbols that happen to share one of these names still resolve normally. (#188)
- VBA statement-form built-in calls such as `MsgBox "…"`, `DoEvents`, and `Shell "calc.exe"` are now classified as runtime calls instead of failing, while user-defined symbols that happen to share a built-in name still resolve normally. (#192)
- VBA array parameters declared with `ByRef name() As Type` or `ByVal name() As Type`, including continued declarations, no longer flag indexed accesses inside their procedure as unresolved calls, while same-named genuine calls elsewhere still surface normally. (#190)
- The human-readable `codegraph status` output no longer reports an index as healthy when a re-index is recommended; the clean-health sentence now reads "No source changes detected" and the re-index hint names the structured reasons, matching the existing JSON contract. (#193)

### Documentation

- The seven `reference_kind` literals that appear in VBA `unresolved_refs` (`unqualified-ident`, `calls`, `qualified-call`, `member-with`, `property-get`, `property-set`, `references`) are now documented in one place alongside the `declined-runtime` semantics and the v1.13.0 noise-ratio benchmarks on the `00_VBA_TOOLKIT_BENCH` corpus, so audit scripts can filter actionable "missing callees" without re-reading the resolver source. Reference: `docs/vba-reference-kinds.md`. (#191)

## [1.13.0] - 2026-07-18

### New Features

- VBA standard-library calls such as `CStr`, `CLng`, `Nz`, and `IsNull` are now classified as runtime calls instead of missing user-defined callees. (#181)

### Fixes

- VBA control-flow keywords are no longer emitted as unresolved unqualified identifiers. (#179)
- Unresolved VBA function and procedure calls now use the single canonical `calls` reference kind. (#180)
- VBA array index expressions such as `values(index)` are no longer mistaken for function calls. (#182)

## [1.12.0] - 2026-07-18

### New Features

- A new `codegraph stats vba-rules [--json]` command exposes the VBA extractor's available rules for debugging and automated health checks. (#168)
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
[1.7.3]: https://github.com/ardelperal/codegraph-vba/releases/tag/v1.7.3
[1.8.0]: https://github.com/ardelperal/codegraph-vba/releases/tag/v1.8.0
[1.9.0]: https://github.com/ardelperal/codegraph-vba/releases/tag/v1.9.0
[1.10.0]: https://github.com/ardelperal/codegraph-vba/releases/tag/v1.10.0
[1.11.0]: https://github.com/ardelperal/codegraph-vba/releases/tag/v1.11.0
[1.12.0]: https://github.com/ardelperal/codegraph-vba/releases/tag/v1.12.0
[1.13.0]: https://github.com/ardelperal/codegraph-vba/releases/tag/v1.13.0
[1.14.0]: https://github.com/ardelperal/codegraph-vba/releases/tag/v1.14.0
[1.15.0]: https://github.com/ardelperal/codegraph-vba/releases/tag/v1.15.0
[1.16.0]: https://github.com/ardelperal/codegraph-vba/releases/tag/v1.16.0
[1.17.0]: https://github.com/ardelperal/codegraph-vba/releases/tag/v1.17.0
[1.17.1]: https://github.com/ardelperal/codegraph-vba/releases/tag/v1.17.1
[1.17.2]: https://github.com/ardelperal/codegraph-vba/releases/tag/v1.17.2
