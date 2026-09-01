# VBA Node Discovery — Implementation Plan

> **Repo:** `ardelperal/codegraph-vba` (fork of `colbymchenry/codegraph`)
> **Baseline:** v1.15.0 + unreleased #202–#217 on `main`
> **Home in the repo:** copy this file to `docs/vba-node-discovery-plan.md` on the first branch that implements it.
> **Source of the findings:** the read-only audit run on 2026-09-01 against
> `00_EXPEDIENTES`, `00_GESTION_RIESGOS` and `HPS_SOLICITUDES`.
> **Status of the code today:** nothing in this plan has been implemented. Every task below is open.

## Issue tracker

Filed against `ardelperal/codegraph-vba` on 2026-09-01. Epic: **#264**.

| Task | Issue | Task | Issue |
|---|---|---|---|
| T0 coverage probe | #242 | T9 module variables | #251 |
| T1 options object | #243 | T10 runtime data binding | #252 |
| T2 SQL receivers | #244 | T11 saved query names | #253 |
| T3 runtime allowlists | #245 | T12 remaining DoCmd verbs | #254 |
| T4 DoCmd.Close | #246 | T13 domain functions | #255 |
| T5 form lifecycle events | #247 | T14 error handling | *superseded — see below* |
| T6 class initializer | #248 | T15 SQL clause coverage | #256 |
| T7 VB_Name binding | #249 | T16 parameters / fields | #257 |
| T8 procedure signature | #250 | T17 statement-form call kind | #265 |

Every issue is self-contained: an implementer does not need this document to work one of them.
All 24 are labelled `status:approved`, and every one carries its own implementation study — none defers a design decision to the implementer. #265 was filed while designing E3; it fixes the statement-form call classification that E3 depends on.

---

## 0. How to use this document

This plan is written so an AI agent can pick up any single task, implement it end to end,
and stop — without having read the audit, and without needing to re-derive the evidence.

**Rules of engagement:**

1. **One task = one GitHub issue = one PR.** Never bundle two task IDs into one PR. The
   tasks are sized so each PR stays reviewable (< 400 changed lines where possible; if a
   task exceeds that, split it along the seam named in its *Split if oversized* note).
2. **Do the tasks in ID order within a wave.** Cross-wave order is fixed: W0 → W1 → W2 → W3.
   `T1` is a prerequisite for `T2`; `T0` is a prerequisite for everything.
3. **Every task must land with its own tests and a `CHANGELOG.md` entry.** No exceptions.
4. **Before you write code for a task, read the two files it names under *Anchors*.**
   Line numbers are from the baseline commit and will drift — search for the named symbol,
   don't trust the number.
5. **Do not bump `package.json`'s version.** Do not run `npm publish`, `git push` or
   `git tag`. Write the files and hand the commands to the maintainer.
6. **Measure before and after** using the harness from `T0`. A task whose *Acceptance*
   section names a number is not done until that number is reproduced and pasted into the PR body.

**Issue and PR conventions (from the user's global config and the repo's `CLAUDE.md`):**

- Issue first, always. The PR body carries `Closes #NNN`.
- Conventional commits. **Never** add `Co-Authored-By` or any AI attribution.
- `CHANGELOG.md` entries go under `## [Unreleased]`, grouped `### New Features` / `### Fixes`,
  one plain-language sentence per bullet, **no internal file paths, no symbol names, no
  node/edge counts** — those live in the PR body, not in the changelog.
- Merge gate: `npm test` green **and** GitHub Actions CI green. No other gate.

---

## 1. Context an implementer needs

### 1.1 What the VBA extraction surface is

~7,900 lines across 24 modules. The shape is a **thin orchestrator + declarative rule tables**:

| File | Role |
|---|---|
| `src/extraction/vba-extractor.ts` | Orchestrator. Preprocess → one pre-walk (procedures) → one main walk (5 classifiers) → finalize. Owns `VBA_RULE_TABLES` and the load-time `validateVbaRuleTables` invariant. |
| `src/extraction/vba/context.ts` | `VbaExtractorContext` — all inter-line state: procedure stack, `localProcs`, `localVarTypeMap` (proc-scoped), `functionReturnTypes`, `sqlVariables`, pending edges. Helpers: `emitReference()`, `findOrCreateFunctionNodeId()`, `setLocalVarTypeInScope()`. |
| `src/extraction/vba/rules.ts` | `VbaExtractionRule` shape, `defineRule`, `runRules` dispatcher (scan mode, structural gate, count, terminal). |
| `src/extraction/vba/procedures.ts` | `PROC_RE` rule. Emits `function` nodes, `contains` edges, control `event-handler` edges. |
| `src/extraction/vba/declarations.ts` | `Event`, `Type`/`End Type`, type members, `Declare` (Win32 API). |
| `src/extraction/vba/dims.ts` | `Dim`/`Public`/`Private`/`Global`/`Static` declarations, `WithEvents`. Emits **type-reference edges only — no variable nodes**. |
| `src/extraction/vba/enums-consts.ts` | `Enum` blocks, members, `Const`. Also tracks proc start/end for Const scoping. |
| `src/extraction/vba/implements.ts` | `Implements IFoo`. |
| `src/extraction/vba/call-sweep.ts` | The per-line driver for calls + SQL + controls + DoCmd + TempVars + `With` stack. Calls out to the modules below. |
| `src/extraction/vba/calls.ts` | Call-site scanning, synthetic stub `function` nodes, `raises-event`. |
| `src/extraction/vba/controls.ts` | `Me.<Ctl>` / `Me!<Ctl>` and `Forms!<Form>` sweeps. |
| `src/extraction/vba/docmd.ts` | `DoCmd.OpenForm` / `OpenReport` / `OpenQuery` only. |
| `src/extraction/vba/sql-wrapper.ts` | SQL-in-strings sweep, wrapper regexes, `sql = sql & "…"` accumulation. |
| `src/extraction/vba/tempvars.ts` | `TempVars("X")` references. |
| `src/extraction/vba/event-synth.ts` | `WithEvents` handler synthesis. |
| `src/extraction/vba-form-extractor.ts` | `.form.txt` / `.report.txt`. Emits `form-layout`, `form-instance-control`, `property`, and `RecordSource`/`RowSource` data bindings. Never emits code nodes. |
| `src/extraction/vba-preprocess.ts` | Comment strip → line-continuation join → conditional-compilation evaluation. All line-count preserving. |
| `src/extraction/sql-table-scan.ts` | Shared `FROM`/`JOIN`/`INTO`/`UPDATE` table scanner + reserved-word reject list. **The single source of truth — never re-implement it.** |
| `src/resolution/vba-runtime-objects.ts` | `RUNTIME_OBJECTS`, `VBA_STDLIB_FUNCTIONS`, `DAO_ENUM_VALUES`, `VBA_INTRINSIC_CONSTANTS`, `classifyVbaReferenceAsRuntime()`. |
| `src/resolution/index.ts` | Post-extraction stub resolver — `repointDecision` stamping. |

### 1.2 Config threading (you will need this for T1/T2)

`codegraph.json` → `vba.*` is parsed in `src/project-config.ts` (`extractVbaTargets`,
`loadVbaConfig`) and threaded to the extractor through **five** hops:

```
project-config.ts  loadVbaConfig()
  → extraction/index.ts        reads vbaConfig, picks maxRaiseFanout
    → extraction/parse-pool.ts ParseTask fields (vbaTargets, maxRaiseFanout)
      → extraction/parse-worker.ts  message shape → extractFromSource(...)
        → extraction/tree-sitter.ts extractFromSource(filePath, source, language,
                                    frameworkNames, vbaTargets, maxRaiseFanout, dysflowExport)
          → new VbaExtractor(filePath, source, vbaTargets, maxRaiseFanout)
```

That chain is already **seven positional parameters** deep. `T1` exists to stop it growing.

### 1.3 Doctrine to preserve

- **Silent beats wrong.** An emitted-but-wrong edge costs more than a missing one. Where a
  name is genuinely ambiguous, emit an `UnresolvedReference` with a typed `referenceKind`
  and let the resolver decide — never guess an edge.
- **Partial coverage is worse than none** (`CLAUDE.md`). If a flow needs two hops to be
  useful, ship both hops or neither.
- **Every new emitter ships with its own gate** (allowlist / reject list) in the same PR.
- `provenance: 'heuristic'` + `metadata.synthesizedBy: '<tag>'` on every synthesized edge.

### 1.4 `EXTRACTION_VERSION`

`src/extraction/extraction-version.ts` — currently `25`. Bump it **once per wave, in the
last PR of that wave**, not per task. Over-bumping turns the re-index hint into noise. Each
wave's *Definition of done* names whether that wave earns a bump.

---

## 2. Wave 0 — the measurement harness

Everything else depends on being able to say "this changed by N". Do this first.

### T0 — Check in the VBA coverage probe

**Goal.** A committed script that runs the extractors over a directory tree and prints a
stable, diffable coverage report. It is how every later task proves its acceptance number.

**Why it must be committed.** The audit's numbers came from a throwaway script. If the next
agent can't re-run it byte-identically, no later task can prove a regression or a win.

**Files.**
- Create `scripts/vba-coverage-probe.mjs`.
- Add to `package.json` scripts: `"probe:vba": "npx tsx scripts/vba-coverage-probe.mjs"`.
- Document it in `docs/vba-extraction-perf.md` under a new "Coverage probe" heading
  (that doc already owns VBA measurement tooling).

**Implementation.**

The probe imports `VbaExtractor`, `VbaFormExtractor` and `SqlQueryExtractor` directly —
no index, no SQLite, no CLI. It walks the given roots, dispatches by extension, and
aggregates. Output is JSON on `--json`, a human table otherwise.

Report these fields, and only these (adding fields later is fine; removing them breaks
comparability):

```
files                     total files dispatched, by extension
nodesByKind               { [kind]: count }
declaredProcedures        function nodes WITHOUT metadata.stub
stubProcedures            function nodes WITH metadata.stub
stubTargetsTop            top 30 stub node names by count
edgesByKind               { [kind]: count }
edgesBySynthesizer        { [metadata.synthesizedBy]: count }
unresolvedByKind          { [referenceKind]: count }
sqlTablesReferenced       distinct table names reached via vba-sql-table
formsReferenced           distinct form names reached via opens-form / form refs
proceduresWithNoOutgoing  declared procedures with zero edges and zero unresolved refs
errors                    extraction errors by code
```

> **Critical detail — the stub discriminator.** A synthesized call target and a declared
> procedure are both `kind: 'function'` and both carry `visibility: 'public'`. The **only**
> discriminator is `metadata.stub === true` (set in `vba/calls.ts`). Without it the corpus
> reads 9,972 procedures instead of 3,840 — a 2.6× inflation that will silently corrupt
> every later measurement.

**Tests.** `__tests__/vba-coverage-probe.test.ts` — run the probe over a temp dir holding
three tiny fixtures (one `.bas`, one `.cls`, one `.form.txt`) and assert the counts are
exact. Follow the existing convention: `fs.mkdtempSync`, real files, cleanup in `afterEach`.

**Acceptance.** Running the probe over the three corpus projects reproduces the audit
baseline in §7.1 exactly. Paste the output into the PR body — it becomes the reference
baseline every later task diffs against.

**Changelog.** None (internal tooling, no user-facing behaviour). This is the one task
exempt from the changelog rule; say so in the PR body.

---

## 3. Wave 1 — precision and cheap recall

No new node kinds, no new edge kinds, no schema change. Highest value per unit of risk.

### T1 — Replace the positional extractor parameters with an options object

**Goal.** Introduce `VbaExtractionOptions` so later tasks can add a knob without touching
five files each time.

**Why now.** `T2` needs a new config knob. Threading an 8th positional parameter through
`extractFromSource` → `parse-pool` → `parse-worker` → `tree-sitter` → constructor is how
this chain got to seven. Pay the refactor once, here, before three more tasks need it.

**Anchors.**
- `src/extraction/tree-sitter.ts` — `export function extractFromSource(` (~line 6533) and the
  `new VbaExtractor(...)` call (~line 6679).
- `src/extraction/parse-pool.ts` — the `ParseTask` interface (~line 54) and the postMessage
  payload (~line 354).
- `src/extraction/parse-worker.ts` — the `parentPort.on('message', …)` handler shape (~line 58).
- `src/extraction/index.ts` — `~line 1603` (reads `vbaConfig`) and `~line 1723` (dispatch).
- `src/extraction/vba-extractor.ts` — the constructor (~line 200).

**Implementation.**

1. Define in `src/extraction/vba/options.ts` (new leaf module — keeps `types.ts` free of
   extraction internals):

   ```ts
   export interface VbaExtractionOptions {
     targets?: Record<string, boolean>;
     maxRaiseFanout?: number;
     sqlWrappers?: readonly string[];   // added by T2, declared here now
   }
   ```

2. `VbaExtractor` gains a 3rd parameter `options: VbaExtractionOptions = {}` and keeps the
   existing 3rd/4th positional parameters working via an overload **for one release only**,
   marked `@deprecated`. Every in-repo call site moves to the object form in this PR.
3. `extractFromSource` gains a trailing `vbaOptions?: VbaExtractionOptions` and keeps
   `vbaTargets` / `maxRaiseFanout` positional for back-compat, merging them into the object
   (object wins on conflict).
4. `ParseTask` and the worker message carry `vbaOptions` as one field. Note: the worker
   boundary is `structuredClone`-based — the object must stay a plain data object with no
   functions, no RegExp, no Map.

**Tests.** `__tests__/extraction-vba-options.test.ts`:
- object form and legacy positional form produce byte-identical `ExtractionResult` on the
  same source (deep-equal minus `durationMs`);
- an option set on the object survives the worker round-trip (assert through
  `parse-pool` with a real worker, mirroring `extraction-vba-event-fanout.test.ts`'s setup);
- `maxRaiseFanout` still gates as before (regression guard for #152).

**Acceptance.** Zero behaviour change. Probe output identical to the T0 baseline, field for field.

**Risk.** Low, but it touches the hot parse path. If the worker round-trip test is awkward,
that is a signal you have put a non-clonable value in the object — fix the value, don't skip the test.

**Changelog.** None (pure refactor, no user-visible change). State that in the PR body.

---

### T2 — Configurable SQL execution receivers + the parenthesised literal form

**Goal.** Stop dropping SQL that flows through a database accessor whose name does not end
in `db`, and stop dropping the `OpenRecordset("SELECT …")` call form.

**Evidence (measured, do not re-derive).**

The receiver regex requires the identifier to *end* at `db`:
`/\b(?:\p{L}[\p{L}\p{N}_]*)?db\b(?:\(\))?\.(?:OpenRecordset|Execute)/`

```
MATCH  getdb().Execute strSQL
MATCH  p_db.OpenRecordset sql
MATCH  CurrentDb.Execute sql
MISS   getdbHPS().OpenRecordset strSQL
MISS   getdbExpedientes().OpenRecordset sql
MISS   getdbLanzadera().OpenRecordset sql
MISS   dbUse.OpenRecordset sql
MISS   m_dbLanzadera.OpenRecordset sql
MISS   dbToUse.Execute sql
MISS   qdf.Execute                      (DAO QueryDef)
```

Share of execution receivers missed: **36%** in `00_EXPEDIENTES` + `00_GESTION_RIESGOS`,
**26%** in `HPS_SOLICITUDES`. The missed names are the per-backend accessors, so the loss
is concentrated on cross-database traffic.

Separately, the literal-SQL regexes require whitespace after the method name
(`\.OpenRecordset\s+"`), so the paren form falls through both the literal path and the
variable path. Verified on one sampler module through the real extractor:

```
CAPTURED  getdb().Execute "DELETE FROM TbTemporal"     → references TbTemporal
DROPPED   getdb().OpenRecordset("SELECT * FROM TbX")   → no table reference at all
```

**Anchors.** `src/extraction/vba/sql-wrapper.ts`: `SQL_WRAPPERS` (~line 23),
`SQL_VAR_EXEC_RE` (~line 34), `SQL_VAR_DOCMD_RUNSQL_RE` (~line 58), `scanSqlInLine` (~line 232).
Config: `src/project-config.ts` `extractVbaTargets` (~line 279) and `loadVbaConfig` (~line 435).

**Implementation.**

1. **Config knob.** `codegraph.json` → `vba.sqlWrappers`: an array of receiver *name patterns*.
   Parse it in `project-config.ts` next to `maxRaiseFanout`, with the same validation shape
   (wrong type → `logWarn` + ignore, never throw). Thread it through the `VbaExtractionOptions`
   object from `T1`.

   Accept two entry forms and nothing else — keep the surface small:
   - a bare identifier prefix, e.g. `"getdb"` → matches `getdb`, `getdbHPS`, `getdbExpedientes`, with or without `()`;
   - an explicit `receiver.method` pair, e.g. `"cnn.Execute"`.

   **Do not accept raw regex from config.** A user-supplied regex in the hot per-line path is
   a catastrophic-backtracking foot-gun and an unbounded review surface.

2. **Defaults.** When `vba.sqlWrappers` is absent, use a built-in list that is a strict
   superset of today's behaviour:
   `db`, `getdb`, `CurrentDb`, `DBEngine`, plus DAO `QueryDef` receivers (`qd`, `qdf`) and the
   ADO pair (`Connection.Execute`, `Recordset.Open`). Build the per-file RegExp **once per
   extractor instance**, not per line — cache it on the context.

3. **Receiver matching.** Replace the `…db\b` shape with prefix matching:
   `^(?:m_|p_)?(?:<prefix>)\p{L}*$` evaluated against the captured receiver identifier,
   case-insensitively. This is what turns `getdbHPS`, `dbUse`, `dbToUse` and `m_dbLanzadera`
   into matches.

4. **Paren form.** Add an optional `\(` to the three literal regexes:
   `\.OpenRecordset\s*\(?\s*"` — and make the same change for `Execute`. Verify the closing
   paren is not consumed into the literal body; the `((?:[^"]|"")*)` group already stops at
   the closing quote.

5. **Guard.** Reuse `scanSqlTables` from `sql-table-scan.ts` unchanged. Do not add a second
   table regex. The reserved-word reject list is the gate this task ships with.

**Split if oversized.** Land the paren-form fix (step 4) as its own PR first — it is ~10 lines
and independently valuable — then the config knob.

**Tests.** `__tests__/extraction-vba-sql-wrappers.test.ts`:
- every `MISS` line in the evidence block above now produces a `vba-sql-table` reference;
- every `MATCH` line still produces exactly the same reference it did before (no double-emission);
- `getdb().OpenRecordset("SELECT * FROM TbX")` emits a reference to `TbX`;
- a custom `vba.sqlWrappers: ["conn"]` in `codegraph.json` makes `conn.Execute "SELECT * FROM T"` resolve, and does **not** disable the defaults;
- `vba.sqlWrappers: 42` logs a warning and falls back to defaults without throwing;
- a receiver that merely *contains* `db` in the middle of an unrelated word (`dbg.Print`) does **not** match.

**Acceptance.**
- `sqlTablesReferenced` in the probe rises on all three corpus projects. Record the exact before/after per project.
- `edgesBySynthesizer['vba-sql-table']` rises; **no other synthesizer's count changes.**
- Zero new entries in `unresolvedByKind` — this task adds edges, it must not add noise.
- Spot-check 10 new table references by hand against the source line. All 10 must be real tables. If any is a SQL keyword, the reject list has a hole — fix that before merging.

**Changelog (Fixes).** "VBA projects that reach their database through a named accessor —
one per backend, for example — no longer lose the tables those queries touch, and recordsets
opened with parentheses are now tracked like every other query."

---

### T3 — Reconcile the two runtime-object lists

**Goal.** Stop creating stub nodes for VBA/DAO runtime members that the resolver already
knows can never be user code.

**Evidence.** 6,132 stub `function` nodes against 3,840 declared procedures across the
corpus — 61% of all function nodes. Top targets:

```
VBA.DoEvents 2,711 · Collection.Add 1,250 · fso.GetFileName 91 · fso.FileExists 75
ListBox.AddItem 67 · ListBox.Selected 49 · ComboBox.AddItem 42 · DBEngine.Workspaces 23
```

~70% of stubs name a runtime member. `RUNTIME_OBJECTS` in
`src/resolution/vba-runtime-objects.ts` **already lists** `VBA`, `fso`, `ListBox`,
`Collection`, `DBEngine` and stamps the resulting edges `repointDecision: 'declined-runtime'`.
The extractor's own `RUNTIME_RECEIVER_BLACKLIST` (`src/extraction/vba/constants.ts`) does
not. Two lists, written at different times, never reconciled.

**Why it matters at the node level.** Edge consumers can filter on `repointDecision`.
Symbol search, `codegraph query` results and node counts cannot — an agent reads past
thousands of fake symbols.

**Anchors.** `src/extraction/vba/constants.ts` `RUNTIME_RECEIVER_BLACKLIST` (~line 100);
`src/resolution/vba-runtime-objects.ts` `RUNTIME_OBJECTS` (~line 33), `isRuntimeObject` (~line 59).

**Implementation.**

1. **Layering.** `src/extraction/` must not import from `src/resolution/`. Move the canonical
   set to a new leaf module `src/extraction/vba/runtime-objects.ts`, and have
   `src/resolution/vba-runtime-objects.ts` re-export from it so every existing resolver
   import keeps working unchanged. (Note: `docs/vba-stub-repoint-decision.md` already refers
   to the allowlist as living under `src/extraction/vba/` — that reference is stale today and
   this task makes it true. Fix the path in that doc in the same PR.)
2. Keep **one** exported set plus the existing predicates. `RUNTIME_RECEIVER_BLACKLIST`
   becomes a derived view of it, not a second literal.
3. In `vba/calls.ts`, gate stub-node creation on `isRuntimeObject(receiverType)` **before**
   `generateNodeId`. Emit the `UnresolvedReference` exactly as today so the resolver's
   `declined-runtime` accounting is unchanged — this task removes *nodes*, not *references*.

**Tests.** `__tests__/extraction-vba-runtime-node-gate.test.ts`:
- `Call VBA.DoEvents` emits an unresolved reference and **zero** nodes;
- `Set c = New Collection : c.Add x` emits zero stub nodes for `Collection.Add`;
- a genuine user class with the same member name (`MiClase.Add`) still emits its stub;
- the re-export keeps `import { RUNTIME_OBJECTS } from '../src/resolution/vba-runtime-objects'` working.

**Acceptance.**
- `stubProcedures` drops by **≥ 3,800** on the corpus (the `VBA.*` + `Collection.*` + `fso.*` + `ListBox.*` + `ComboBox.*` + `DBEngine.*` share alone).
- `declaredProcedures` is **unchanged at 3,840**. If it moves at all, the gate is eating real procedures — stop and fix.
- `unresolvedByKind` is unchanged. This task must not change what the resolver sees.
- `edgesByKind['calls']` drops by the same count as the removed stubs (the edges pointed at those nodes).

**Risk.** Medium. This deletes nodes that existing consumer SQL may join on. Check
`docs/vba-stub-repoint-decision.md`'s published consumer queries still return sensible
results and update the doc's expected-shape section in the same PR.

**Changelog (Fixes).** "Calls into Access and VBA built-ins no longer create thousands of
phantom symbols that show up in search results and node counts."

---

### T4 — `DoCmd.Close acForm, "<Name>"` → a reference to the form

**Goal.** Close the form-lifecycle loop that `opens-form` opened.

**Evidence.** `DoCmd.Close` appears **586** times in the corpus. ~150 of those name a form
as a string literal; the rest use `Me.Name` (out of scope — that is a runtime value).
Meanwhile `DoCmd.OpenReport` and `DoCmd.OpenQuery` — two of the three verbs the extractor
does model — have **zero** occurrences. The modelled verbs are not the called verbs.

Corpus verb distribution: `Hourglass` 4,037 · `Close` 586 · `OpenForm` 254 ·
`ShowToolbar` 12 · `Quit` 10 · `Maximize` 2 · `Restore` 1.

**Anchors.** `src/extraction/vba/docmd.ts`: `DOCMD_OPEN_DISPATCH` (~line 60),
`scanDoCmdOpenCalls` (~line 93), `emitOpensStubEdge` (~line 129). Wired from
`src/extraction/vba/call-sweep.ts` (~line 409).

**Implementation.**

1. Add a `CLOSE_ARG_RE` capturing the second argument of
   `DoCmd.Close acForm|acReport, ("Name"|Identifier)`. Only act on the **string-literal**
   form and only for `acForm` / `acReport`; anything else (a variable, `Me.Name`, a bare
   `DoCmd.Close`) is skipped silently. This is the "silent beats wrong" call — a form name
   held in a variable is not statically known.
2. Reuse `emitOpensStubEdge`'s stub-creation path so the target is the same
   `form-layout` / `report-layout` node `opens-form` already points at. **Do not create a
   second stub shape** — a duplicate stub with a different id is worse than no edge.
3. Edge kind: **`references`**, not a new `closes-form` kind. Wave 1 adds no edge kinds.
   Tag it `metadata.synthesizedBy: 'vba-closes-form'` and `metadata.targetFormName`. If the
   maintainer later wants a first-class `closes-form` kind, the synthesizer tag is the seam.

**Tests.** `__tests__/extraction-vba-docmd-close.test.ts`:
- `DoCmd.Close acForm, "FormExpediente", acSaveNo` → one `references` edge tagged `vba-closes-form` to the `Form_FormExpediente` stub;
- `DoCmd.Close acForm, Me.Name, acSaveNo` → **no** edge;
- bare `DoCmd.Close` → no edge;
- `DoCmd.Close acReport, "Rpt_X"` → target uses the `Report_` prefix convention;
- an `OpenForm` and a `Close` naming the same form in one file converge on **one** stub node with two distinct edges.

**Acceptance.** `edgesBySynthesizer['vba-closes-form']` ≈ 150 on the corpus.
`nodesByKind['form-layout']` rises by **0** (the stubs already exist from `opens-form`) or by
a small number for forms only ever closed, never opened — verify each of those by hand.

**Changelog (New Features).** "Closing a form from code is now part of the graph, so a form's
full lifecycle — who opens it and who closes it — is visible in one place."

---

### T5 — Form and report lifecycle events → `event-handler` edge

**Goal.** Make "what runs when this form opens" a graph query instead of a naming convention
the reader has to know.

**Evidence.** 115 form-level handlers in the corpus: `Form_Open` 61, `Form_Load` 45,
`Form_Unload` 3, `Form_Timer` 2, `Form_Close` 2, `Form_Resize` 1, `Form_Activate` 1.
None of them produces an edge today — `vba/procedures.ts` deliberately skips them, because a
form-level event fires on the form object, not on a control, and the control-handler path
would synthesize a bogus `form-instance-control` node named `Form`.

**Anchors.** `src/extraction/vba/procedures.ts` — the `parseEventHandlerName` block and the
`Form_*` skip comment. Event-name allowlist: `src/extraction/vba/events.ts`
(`ACCESS_EVENT_NAMES`, `isAccessEventName`).

**Implementation.**

1. When the handler's control part is exactly `Form` or `Report` **and** the file is
   `Form_*.cls` / `Report_*.cls`, target the sibling **`form-layout` / `report-layout`** node
   instead of a `form-instance-control`. Emit the local stub with the same deterministic id
   the form extractor will produce, exactly as the control path does today — the
   `INSERT OR REPLACE` convergence is the existing mechanism, reuse it.
2. Gate on `isAccessEventName(eventName)`. `Form_Load` is an event; a class method that
   happens to be called `Form_Helper` is not.
3. Keep the edge kind `event-handler` and set `metadata.eventName`. Add
   `metadata.scope: 'form'` so consumers can separate form-level from control-level handlers
   without parsing the name.

**Tests.** `__tests__/extraction-vba-form-level-events.test.ts`:
- `Private Sub Form_Load()` in `Form_X.cls` → `event-handler` edge to the `Form_X` layout node, `eventName: 'Load'`, `scope: 'form'`;
- `Private Sub Report_Open()` in `Report_Y.cls` → the report layout node;
- `Private Sub Form_Helper()` → **no** edge (not an Access event name);
- a control handler `cmdSave_Click` is unchanged — same node id, same edge, `scope` absent or `'control'`;
- the same source in a file **not** named `Form_*.cls` emits no form-level edge (pre-existing guard preserved; see T7).

**Acceptance.** `edgesByKind['event-handler']` rises by **≈115** (from 695 to ≈810) on the
corpus. `nodesByKind['form-instance-control']` rises by **0** — if it moves, a bogus `Form`
control node is being created and the task is wrong.

**Changelog (New Features).** "A form's own lifecycle handlers — what runs when it opens,
loads or closes — now connect to the form in the graph, the same way a button's click
handler already did."

---

### T6 — Fix the class-initializer detection

**Goal.** Detect the construct VBA actually has.

**Evidence.** `vba-extractor.ts` sets `metadata.hasClassInitializer` when a `.cls` declares a
procedure named `New`. VBA has no `Sub New` — that is VB.NET. Access classes use
`Private Sub Class_Initialize()` and `Class_Terminate()`.
Corpus: **`Sub New` → 0 occurrences. `Class_Initialize`/`Class_Terminate` → 2.**
The flag can never be true on real Access code.

**Anchors.** `src/extraction/vba-extractor.ts` — the `procs.some((p) => p.name === 'New')`
block (~line 365). Tests that pin the wrong construct: `__tests__/extraction-vba.test.ts`
lines ~100, ~110, ~113 (REQ-CODE-3).

**Implementation.**

1. Detect `Class_Initialize` → `metadata.hasClassInitializer = true`,
   `metadata.initializerName = 'Class_Initialize'`.
2. Add `metadata.hasClassTerminator` / `terminatorName` for `Class_Terminate` — same cost,
   and destructor presence is the other half of the lifecycle question.
3. Keep accepting `New` as a fallback so the historical behaviour is not *removed*, just
   corrected. Note in the code comment why it can never fire on Access.
4. Update the REQ-CODE-3 tests to assert the real construct. **Rename the test cases**, do
   not silently rewrite their bodies — a reviewer must see the construct changed.

**Tests.** Update `__tests__/extraction-vba.test.ts`:
- `Private Sub Class_Initialize()` in a `.cls` → `hasClassInitializer === true`, `initializerName === 'Class_Initialize'`;
- `Private Sub Class_Terminate()` → `hasClassTerminator === true`;
- a `.cls` with neither → both unset;
- a `.bas` with `Class_Initialize` → unset (only classes have a lifecycle).

**Acceptance.** Two class nodes in the corpus gain the flag. Small, but it converts a
documented capability from never-true to true.

**Changelog (Fixes).** "Class setup and teardown routines in Access class modules are now
recognised; the previous check looked for a construct that only exists in VB.NET."

---

### T7 — Fall back to `VB_Name` when binding a form to its code-behind

**Goal.** Remove a silent failure mode.

**Evidence.** Event-handler synthesis and the `Me.<Control>` sweep both gate on
`path.basename(filePath)` starting with `Form_` / `Report_`. Verified by running one source
file under two names: as `Form_FormExpediente.cls` it produced an `event-handler` edge and
two control references; as `sample2.cls` — same bytes, same `Attribute VB_Name = "Form_FormExpediente"` —
it produced **none**. The file still parses and still emits its procedures, so nothing looks broken.

**Anchors.** `src/extraction/vba/procedures.ts` (the `codeBehindExt` basename check);
`src/extraction/vba/controls.ts` `siblingLayoutPath()` (~line 50).

**Implementation.**

1. `siblingLayoutPath()` and the `codeBehindExt` check both consult the file basename first
   (unchanged, fast path), then fall back to `ctx.classNamePrefix` — which already holds the
   resolved `VB_Name` for `.cls` files.
2. The sibling path is still derived from the **file** path (that is where the sibling lives
   on disk); only the *decision to bind* falls back to `VB_Name`.
3. When the two disagree, set `metadata.bindingSource: 'vb-name'` on the synthesized edges so
   the mismatch is diagnosable rather than invisible.

**Tests.** `__tests__/extraction-vba-form-binding-fallback.test.ts`:
- identical source under `Form_X.cls` and under `renamed.cls` (with `VB_Name = "Form_X"`) produce the same set of `event-handler` edges;
- the mismatch case carries `bindingSource: 'vb-name'`;
- a plain service class (`InformeRiesgoPDFServicio.cls`, no `Form_` prefix, no matching `VB_Name`) still produces **zero** control stubs — this is the guard that stopped ~550 spurious nodes; it must not regress.

**Acceptance.** Corpus counts unchanged (Dysflow exports keep name and file in sync). This
task buys robustness, not recall — say so plainly in the PR body rather than inventing a number.

**Changelog (Fixes).** "Form code-behind is now matched to its layout by the module's own
name as well as its filename, so a renamed export no longer silently loses its event wiring."

### Wave 1 — definition of done

- All of T1–T7 merged.
- Probe re-run on all three corpus projects; the before/after table is pasted into the last PR.
- **Bump `EXTRACTION_VERSION` to 26** in the last PR of the wave — T2, T4 and T5 materially
  change which edges exist, so existing indexes should be rebuilt.
- No node kind and no edge kind was added.

---

## 4. Wave 2 — the symbols that are missing

New nodes, but only using kinds that already exist in `NODE_KINDS` / `EdgeKind`.

### T8 — Procedure kind and signature metadata

**Goal.** Make a `Sub`, a `Function`, a `Property Get` and a `Property Let` distinguishable,
and give a call site an arity to check against.

**Evidence.** Every procedure becomes `kind: 'function'` with **no metadata at all** — the
node has no `metadata` field. `ProcInfo` already computes `kind: 'sub' | 'function' | 'property'`
internally and throws it away. Corpus: 900 `Property Get` + 207 `Property Let`/`Set`. A Get
and its matching Let are two nodes with the same name and same qualifiedName, separable only
by line number.

`PROC_RE` also captures nothing about the parameter list; `parseArrayParameters` and
`parseReturnType` exist but feed only internal maps.

**Anchors.** `src/extraction/vba/procedures.ts` — the `fnNode` literal; `PROC_RE` in
`src/extraction/vba/constants.ts`; `ProcInfo` in `src/extraction/vba/context.ts`.

**Implementation.**

1. Add to the `function` node's `metadata`:
   ```
   procKind      'sub' | 'function' | 'property'
   accessor      'get' | 'let' | 'set'        (property only)
   isStatic      boolean
   isFriend      boolean                      (visibility folds to 'public'; keep the fact)
   returnType    string                       (functions and Property Get)
   params        [{ name, type, byRef, optional, isArray, hasDefault }]
   arity         { required: number, total: number }
   ```
2. Parse the parameter list from the declaration line **after** line-continuation joining —
   the preprocessor guarantees a multi-line signature arrives as one logical line, and #202
   already fixed the location mapping. Do not write a second continuation joiner.
3. `ParamArray` sets `arity.total = Infinity`; serialise it as `null` in metadata, since JSON
   has no `Infinity` and this crosses the worker boundary.
4. **Do not emit `parameter` nodes** in this task. Metadata first; nodes only if a consumer
   asks (see T16).

**Tests.** `__tests__/extraction-vba-procedure-metadata.test.ts` — one case per row of the
metadata table above, plus:
- `Property Get X` and `Property Let X` in the same class produce two nodes with different `accessor` values;
- a continued signature (`_` at end of line) parses identically to the one-line form;
- `Optional ByRef msg As String = ""` yields `optional: true, byRef: true, hasDefault: true`;
- `ParamArray args() As Variant` yields `arity.total === null`.

**Acceptance.** `declaredProcedures` unchanged at 3,840. Every one of them carries `procKind`.
Spot-check the 20 largest classes: no `Property` misclassified as `sub`.

**Changelog (New Features).** "Procedures now carry their real shape in the graph — whether
something is a Sub, a Function or a property accessor, what parameters it takes and which are
optional — so tools can tell a property's getter from its setter."

---

### T9 — `variable` nodes for module-level state

**Goal.** Make shared mutable state visible. "Who reads `gblConn`?" should be answerable.

**Evidence.** 2,902 module-level `Public`/`Private`/`Global` declarations across the corpus,
and **zero** nodes. `vba/dims.ts` emits type-reference edges and populates `localVarTypeMap`,
but never a node. This is the largest uncovered coupling channel between modules.

**Anchors.** `src/extraction/vba/dims.ts` — the `dim-decl` rule's emit body;
`ctx.currentVarTypeProcKey` (the existing module-vs-procedure scope discriminator, from #205).

**Implementation.**

1. Emit a `variable` node **only when `ctx.currentVarTypeProcKey === 'module'`** — that is,
   only for module-level declarations. Procedure locals stay out of the graph; tracking every
   local is the node-explosion failure mode `CLAUDE.md` warns about, and the def-use frontier
   the upstream project deliberately left uncovered.
2. Node shape: `kind: 'variable'`, `name`, `qualifiedName: '<Module>.<name>'`, `visibility`
   from the declaration keyword, `metadata: { declaredType, isArray, isWithEvents, isConst: false }`.
3. `contains` edge from the module/class node — reuse the `pendingModuleOrClassSource`
   mechanism, since the module node is created lazily after the walk.
4. **Read/write references.** In the call sweep, an identifier that matches a known
   module-level variable name emits an `UnresolvedReference` with
   `referenceKind: 'property-get'` / `'property-set'` (reuse the existing direct-assignment
   predicate `isDirectAssignment` from `controls.ts` — do not write a second one) and
   `metadata.synthesizedBy: 'vba-module-var'`.
5. **Gate.** Only names already registered as module-level variables in this file are
   considered. Never scan for arbitrary identifiers — that is how you get 10,000 false references.

**Split if oversized.** Nodes + `contains` in PR 1; read/write references in PR 2. But do not
stop after PR 1 and call the feature done: a variable node with no references is a symbol
nobody can trace, which is the "partial coverage is worse than none" trap. Both PRs land in the same wave.

**Tests.** `__tests__/extraction-vba-module-variables.test.ts`:
- `Public gblConn As DAO.Database` at module level → one `variable` node + `contains` edge, and the existing `references` edge to `DAO` is **still emitted once** (no double-count);
- `Dim x As Long` inside a `Sub` → **no** node;
- a module variable and a procedure local with the same name: the local does not produce a node, and a read inside that procedure does **not** produce a module-var reference (this is the #205 scoping rule; it must hold);
- `Public WithEvents m_Form As Form_X` → variable node with `isWithEvents: true`, and the existing `subscribes-event` edge unchanged;
- `x = gblConn` → `property-get`; `gblConn = Nothing` → `property-set`.

**Acceptance.**
- `nodesByKind['variable']` ≈ 2,900 on the corpus (the exact figure is whatever the census says after the module-level filter; record it).
- `declaredProcedures` and every other node kind unchanged.
- `proceduresWithNoOutgoing` **drops** — procedures that only touch globals were previously invisible leaves.
- Precision spot-check: sample 15 read/write references, all 15 must be genuine accesses to that module's variable.

**Changelog (New Features).** "Module-level variables are now part of the graph, so shared
state can be traced to the procedures that read and write it."

---

### T10 — SQL assigned to a binding property at runtime

**Goal.** Scan the SQL this codebase actually binds, which is assigned in code, not stored in
the layout file.

**Evidence.** 161 `.RowSource =` assignments in `.cls`/`.bas` across the corpus. The form
files themselves are mostly unbound: `RecordSource` appears **0** times in
`00_EXPEDIENTES`'s `.form.txt` files, `ControlSource` **0**. The static `.form.txt` sweep
(`vba-record-source` / `vba-row-source`) is therefore scanning the empty half of the problem.

**Anchors.** `src/extraction/vba/sql-wrapper.ts` `scanSqlInLine` (~line 232), wired from
`call-sweep.ts` (~line 306). The `.form.txt` side that already does this correctly:
`src/extraction/vba-form-extractor.ts` (~line 604 onward).

**Implementation.**

1. Add a rule for `<anything>.(RowSource|RecordSource|ControlSource|Filter|OrderBy)\s*=` and
   feed the right-hand side through the **same** pipeline the wrapper path uses:
   `collectConcatFragments` → `scanSqlTables`. Non-literal operands already become the `?`
   sentinel, so `"SELECT * FROM " & tabla` correctly yields no table.
2. Also resolve the variable form (`.RowSource = strSQL`) through the existing `sqlVariables`
   map — the accumulate semantics from #13 and the procedure scoping from #204 apply unchanged.
3. Tag `metadata.synthesizedBy: 'vba-row-source-dynamic'` (distinct from the static
   `vba-row-source`) so the two provenances stay separable in audits.
4. A value that is **not** SQL (`.RowSource = "Value List"` content, a bare query name) is
   handled by T11, not here. If `scanSqlTables` returns nothing and the value is a bare
   identifier, leave it for T11 rather than guessing.

**Tests.** `__tests__/extraction-vba-dynamic-binding.test.ts`:
- `Me.cbo.RowSource = "SELECT Id FROM TbUsuarios"` → table reference to `TbUsuarios`;
- `Me.RecordSource = strSQL` where `strSQL` was built earlier in the same procedure → same table set as the literal form;
- `Me.Filter = "Id = 1"` → **no** table reference (a filter is not a FROM clause);
- `.RowSource = "SELECT " & campo & " FROM " & tabla` → no reference (both operands dynamic), and specifically **not** a reference to `FROM` or `WHERE`.

**Acceptance.** `edgesBySynthesizer['vba-row-source-dynamic']` > 0 on the corpus;
`sqlTablesReferenced` rises. **No new reserved-word captures** — grep the new table names for
any entry in `SQL_RESERVED_TABLE_TOKENS`; there must be none.

**Changelog (New Features).** "Combo boxes, list boxes and forms whose data source is built
in code now show which tables they read, not just the ones wired in the form designer."

---

### T11 — Saved queries as call targets

**Goal.** Connect the `query` nodes the SQL extractor already emits to the code that runs them.

**Evidence.** `SqlQueryExtractor` emits one `query` node per `queries/<Name>.sql` and links it
to its tables. But the only code-side path to a query node is `DoCmd.OpenQuery` — which has
**zero** occurrences in the corpus. `getdb().OpenRecordset("nombreConsulta")`,
`QueryDefs("nombreConsulta")` and `RowSource = "nombreConsulta"` all name a saved query and
produce nothing.

**Anchors.** `src/extraction/sql-query-extractor.ts` (target node shape and id derivation);
`src/extraction/vba/docmd.ts` `scanDoCmdOpenQuery` (~line 193) — the existing
`vba-opens-query` unresolved-reference pattern to copy.

**Implementation.**

1. When a wrapper receives a **string literal that is not SQL** — no `SELECT`/`INSERT`/
   `UPDATE`/`DELETE` verb, and it parses as a bare Access object name — emit an
   `UnresolvedReference` with `referenceKind: 'dao-query'` (the literal already exists in the
   `ReferenceKind` union) and `metadata.synthesizedBy: 'vba-query-name'`.
2. Same for `QueryDefs("X")` and for the T10 binding path when the value is a bare identifier.
3. Resolution happens in `src/resolution/index.ts`, matching the reference name against
   `query` nodes by name. **If no query node matches, decline** — do not fall back to
   creating a table placeholder. A name that matches neither a query nor a table is a
   `failed` reference, and that is the correct, actionable outcome.

**End-to-end requirement.** This task is only useful if the full path
`procedure → query → table` connects. Verify with a real project that has a `queries/`
directory and assert the two hops in one test. A one-hop version is the half-bridged flow
`CLAUDE.md` forbids.

**Tests.** `__tests__/extraction-vba-query-names.test.ts` + a resolution test:
- `getdb().OpenRecordset("qryPendientes")` → `dao-query` reference, resolved to the `query` node when `queries/qryPendientes.sql` is indexed;
- the same call when no such query exists → `failed`, not a synthetic table;
- `getdb().OpenRecordset("SELECT * FROM T")` still takes the SQL path, not the query path (verb detection wins);
- traversal from the calling procedure reaches the query's tables in two hops.

**Acceptance.** On a corpus project with saved queries, every `queries/*.sql` that is called
from code has at least one inbound edge. List the ones that don't — those are genuinely dead
queries and are a useful finding in the PR body.

**Changelog (New Features).** "Saved Access queries now connect to the code that runs them,
so a query's impact can be traced from the procedure that opens it all the way to the tables
it touches."

### Wave 2 — definition of done

- T8–T11 merged, including both halves of T9.
- **Bump `EXTRACTION_VERSION` to 27** in the last PR — a new node kind is in the output.
- Node-explosion check: total node count rises by roughly the expected variable count and
  nothing else. An unexplained rise means a gate is missing.
- `docs/vba-reference-kinds.md` updated with the new `referenceKind` usages and a refreshed
  noise-ratio table.

---

## 5. Wave 3 — the larger commitments

Do not start these until Waves 1 and 2 are merged and measured. Each one either adds a kind
or has a much bigger surface.

### T12 — The remaining `DoCmd` verbs

Cover `RunMacro`, `RunSQL` (statement form beyond the current wrapper), `TransferSpreadsheet`,
`TransferText`, `TransferDatabase`, `OutputTo`, `SendObject`, `RunCommand`, `ApplyFilter`,
`CopyObject`, `DeleteObject`, `Rename`, `OpenTable`, `SelectObject`, `BrowseTo`.

**Zero of these appear in the current corpus** — they are standard Access, not this
organisation's Access. That is exactly why they rank here and not in Wave 1. Build them as
one table-driven dispatch (extend `DOCMD_OPEN_DISPATCH` into a general
`DOCMD_OBJECT_DISPATCH` keyed by verb → argument position → target kind) rather than fifteen
bespoke regexes. `RunMacro` needs a decision from the maintainer: a new `macro` node kind, or
a `references` edge to an unresolved name. Ask before adding the kind.

### T13 — Domain aggregate functions

`DLookup`, `DCount`, `DSum`, `DMax`, `DMin`, `DAvg`, `DFirst`, `DLast`. The second argument is
a table or query name. Zero uses in this corpus, heavy use in typical Access code, and cheap
once T11's "is this a query name or a table name?" resolution exists. Reuse it; do not write a
second resolver.

### T14 — Error-handling structure → **superseded**

**This task has been analysed separately and replaced. See `vba-node-discovery-plan`'s sibling
document `vba-error-handling-plan.md` (repo home: `docs/vba-error-handling-plan.md`), tasks
E1–E6. Do not implement T14 as described here.**

The design note this task asked for was written, and it reached a different conclusion than the
sketch below assumed. Summary of what changed:

- Labels are **not** worth a node kind here: 96.5% of the corpus's 3,912 labels are the same
  handler label (`errores`) doing the same job, and only 136 are real control flow.
- Error propagation in this codebase does **not** travel through VBA's error mechanism — only
  16 handlers of 3,774 re-raise. ~74% write the message into a module-level variable
  (`m_Error` / `p_Error` / `Me.Error`), which **T9 already models**.
- The replacement design adds **zero node kinds and zero edge kinds** — only `errorPolicy`
  metadata on the procedure node, an `inErrorHandler` flag on edges emitted from a handler
  region, and an `errorChannel` flag on the T9 references.

The original sketch, kept for the record: 4,029 `On Error GoTo`, 4,237 `GoTo <label>`,
850 `Resume` — a new label node kind and a `handles-error` edge kind. That design survives as
**E6**, explicitly blocked pending maintainer sign-off and a query that the cheaper model
cannot serve.

### T15 — SQL clause coverage

Extend `sql-table-scan.ts` with `CREATE TABLE`, `ALTER TABLE`, `DROP TABLE`, and the
Access-specific `IN "otra.accdb"` clause that points a query at an external database file —
a cross-backend edge with no representation today. Small, self-contained, and it touches the
one module every SQL path shares, so it needs careful reserved-word regression testing.

### T16 — `parameter` nodes and table field members

Only worth doing if this graph starts feeding migration tooling (an Access-to-web blueprint,
for instance). Until then it is node inflation for its own sake. T8's signature metadata
already answers most signature questions without the nodes.

---

## 6. Guardrails — how not to make the graph worse

The fork's recent history (#181 → #216) is almost entirely **precision** work. Everything in
this plan pushes toward **recall**, and each task is a fresh chance to emit a confident wrong
edge. Three rules, applied to every PR:

1. **One emitter, one gate, one PR.** No new emitter merges without its allowlist or reject
   list in the same change. The `vba-sql-table` path is the model: the emitter and
   `SQL_RESERVED_TABLE_TOKENS` live together.
2. **Measure both directions.** `docs/vba-reference-kinds.md` fixes per-kind `failed` rates as
   a floor. A change that *lowers* a kind's failed count is **not automatically a win** — it
   may mean a classifier swallowed a real callee. Every PR reports: nodes by kind before/after,
   edges by synthesizer before/after, unresolved by kind before/after. Extend that doc's table
   with node counts so recall work is as visible as precision work.
3. **Ambiguity goes to the resolver, not to a guess.** Where a name could be several things,
   emit an `UnresolvedReference` with a typed `referenceKind`. That mechanism already exists,
   is already documented, and already has consumer queries written against it. Do not invent
   a third path.

**Per-PR checklist:**

- [ ] Issue exists; PR body carries `Closes #NNN`
- [ ] `npm test` green locally; CI green
- [ ] New tests cover both the positive case **and** the case that must stay silent
- [ ] Probe re-run; before/after table pasted in the PR body
- [ ] Precision spot-check done by hand on ≥10 new edges, result stated
- [ ] `CHANGELOG.md` entry under `[Unreleased]`, user-facing wording, no internals
- [ ] `EXTRACTION_VERSION` bumped **only** if this is the wave's last PR and the wave earned it
- [ ] No `Co-Authored-By`, no version bump, no push/tag/publish

---

## 7. Appendix — the evidence

### 7.1 Probe baseline (2026-09-01, before any task)

Corpus: `00_EXPEDIENTES` + `00_GESTION_RIESGOS`, 355 files dispatched.

| Node kind | Count | | Edge kind | Count | | Unresolved kind | Count |
|---|--:|---|---|--:|---|---|--:|
| `function` | 9,972 | | `calls` | 8,322 | | `calls` | 7,679 |
| `property` | 2,668 | | `contains` | 6,248 | | `references` | 2,057 |
| `form-instance-control` | 2,485 | | `references` | 4,968 | | `member-with` | 1,488 |
| `class` | 1,577 | | `event-handler` | 695 | | `unqualified-ident` | 1,459 |
| `file` | 353 | | `opens-form` | 126 | | `qualified-call` | 399 |
| `enum_member` | 266 | | `raises-event` | 74 | | `property-get` | 215 |
| `form-layout` | 208 | | `type-member` | 69 | | `property-set` | 76 |
| `constant` | 178 | | `subscribes-event` | 56 | | | |
| `event` | 83 | | | | | | |
| `type_member` | 69 | | | | | | |
| `enum` | 47 | | | | | | |
| `declare` | 35 | | | | | | |
| `module` | 35 | | | | | | |
| `type` | 9 | | | | | | |

Derived: **3,840 declared procedures**, **6,132 stubs**, 210 real classes vs 1,367 synthetic
placeholders, 31.4% of all function nodes have an outgoing edge (81.5% when stubs are
excluded — the number to quote).

### 7.2 Construct census (both projects, source text)

| Construct | Count | Modelled today |
|---|--:|---|
| `On Error GoTo` | 4,029 | no |
| `GoTo <label>` | 4,237 | no |
| `DoCmd.Hourglass` | 4,037 | n/a (no object named) |
| `Sub`/`Function` declarations | 3,932 | yes |
| `Me.<ident>` | 15,551 | yes (form code-behind only) |
| Module-level variable declarations | 2,902 | **no node** |
| `Optional` parameters | 2,223 | **no** |
| `Resume` | 850 | no |
| `Property Get` | 900 | yes (kind not recorded) |
| `Property Let`/`Set` | 207 | yes (accessor not recorded) |
| SQL `SELECT` literals in code | 565 | partially |
| `DoCmd.Close` | 586 | **no** |
| `.OpenRecordset` | 480 | partially (receiver heuristic) |
| `DoCmd.OpenForm` | 254 | yes |
| `Const` declarations | 241 | yes |
| `TempVars` | 187 | yes |
| `.RowSource =` in code | 161 | **no** |
| `Event` declarations | 159 | yes |
| `RaiseEvent` | 147 | yes |
| `Form_*` lifecycle handlers | 115 | **no edge** |
| `WithEvents` | 107 | yes |
| `Declare` (Win32) | 88 | yes |
| bang references (`x!y`) | 79 | partially |
| `Enum` blocks | 47 | yes |
| `#If` / `#Const` | 46 | yes |
| `Type` blocks | 13 | yes |
| `CallByName` | 1 (+4 in HPS) | no |
| `DLookup` family | 0 | no |
| `Forms!X` / `Reports!X` | 0 | yes (unused here) |
| `DoCmd.OpenReport` / `OpenQuery` / `RunSQL` / `RunMacro` | 0 | partially (unused here) |
| `Application.Run` | 0 | no |
| `Implements` | 0 | yes (unused here) |

### 7.3 SQL receiver census

| Receiver | `.OpenRecordset` | `.Execute` | Matched by today's regex |
|---|--:|--:|---|
| `getdb()` | — | 138 | yes |
| `db` | 23 | 12 | yes |
| `dbUse` | 19 | 27 | **no** |
| `p_db` | 17 | 4 | yes |
| `qd` / `qdf` | 3 | 1 | **no** |
| `CurrentDb` / `CurrentDb()` | — | 7 | yes |
| `getdbRiesgos()` | — | 4 | **no** |
| `dbToUse` | — | 3 | **no** |
| `m_db` | 1 | — | yes |
| `m_dbLanzadera` | 1 | — | **no** |
| `getdbHPS()` (HPS) | 10 | 6 | **no** |
| `getdbLanzadera()` (HPS) | 6 | — | **no** |
| `getdbExpedientes()` (HPS) | 6 | — | **no** |
| `getdbCorreo()` (HPS) | 1 | — | **no** |

### 7.4 Provenance tags in use today

`vba-name-resolution` · `vba-new-binding` · `vba-set-new` · `vba-withevents` ·
`vba-event-handler` · `vba-me-control` · `vba-forms-bang` · `vba-tempvar` · `vba-sql-table` ·
`vba-form-binding` · `vba-record-source` · `vba-row-source` · `vba-source-object` ·
`vba-expression-handler` · `vba-opens-form` · `vba-opens-report` · `vba-opens-query` ·
`vba-qualified-call-unresolved` · `vba-paren-call-unresolved` ·
`vba-statement-call-unresolved` · `vba-with-member-unresolved` · `vba-test-manifest` ·
`vba-test-sequence`

Tags this plan adds: `vba-closes-form` (T4) · `vba-module-var` (T9) ·
`vba-row-source-dynamic` (T10) · `vba-query-name` (T11).

### 7.5 Reproducing the audit from scratch

```bash
git clone https://github.com/ardelperal/codegraph-vba.git repo
cd repo && npm install
npx tsx scripts/vba-coverage-probe.mjs \
  C:/00repos/codigo/00_EXPEDIENTES/src \
  C:/00repos/codigo/00_GESTION_RIESGOS/src \
  --json > baseline.json
```

The extractor classes are importable standalone — no index, no SQLite, no CLI needed.
