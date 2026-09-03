# VBA Error Handling — Analysis and Implementation Plan

> **Repo:** `ardelperal/codegraph-vba`
> **Baseline:** v1.15.0 + unreleased #202–#217 on `main`
> **Home in the repo:** copy to `docs/vba-error-handling-plan.md` on the first branch that implements it.
> **Relationship to the main plan:** this document *replaces* task **T14** in
> `vba-node-discovery-plan.md`. That plan deferred error handling and said "design note first,
> maintainer sign-off on the kinds, prove it on one project". This is that design note, and it
> reaches a different conclusion than T14 assumed — read §4 before writing any code.
> **Evidence:** 4,817 procedures across `00_EXPEDIENTES`, `00_GESTION_RIESGOS`, `HPS_SOLICITUDES`.
> **Status:** nothing here is implemented. Every task is open.

## Issue tracker

Filed against `ardelperal/codegraph-vba` on 2026-09-01. Epic: **#264**.

| Task | Issue | Status |
|---|---|---|
| E1 error metrics in the probe | #258 | approved |
| E2 `errorPolicy` on the procedure node | #259 | approved |
| E3 mark edges inside a handler | #260 | landed |
| E4 recognise the error channel | #261 | approved — needs #251 |
| E5 publish the queries | #262 | approved |
| E6 label nodes + `handles-error` edges | #263 | approved — full design in the issue: `label` kind, `handles-error` kind, +30% nodes budgeted, no call re-parenting |

---

## 1. Why error handling deserves its own document

It is the largest single construct in the corpus — 4,029 `On Error GoTo` statements against
3,932 procedure declarations — and it is the one where the naive model is most obviously wrong.

The naive model is "a line label is a node, `GoTo` is an edge". Applied here that produces
~3,900 label nodes and ~4,200 jump edges: an entire second graph, the same size as the
procedure graph, encoding almost no information. 96.5% of those labels are the *same label*
doing the *same job* in every procedure.

**The conclusion this document reaches: do not build an error graph. Annotate the graph you
already have.** §4 makes that case with the measurements behind it, §5 is the task breakdown
that follows from it.

---

## 2. What this codebase actually does

Not what the Microsoft docs describe — what the census found.

### 2.1 The house protocol

Nearly every procedure follows one shape:

```vba
Public Function Registrar() As Boolean
    On Error GoTo errores          ' 3,850 sites — always this label name
    m_Error = ""
    ...
    If m_Error <> "" Then Err.Raise 1000   ' 5,152 sites — always this number
    ...
    Exit Function
errores:
    If Err.Number <> 1000 Then     ' 4,898 Err.Number reads, nearly all this comparison
        p_Error = "El método Registrar ha devuelto el error: " & Err.Description
    End If
End Function
```

Read that carefully, because three project-specific facts fall out of it and every design
decision below depends on them:

1. **`errores` is a house convention, not a name to discover.** 3,789 of 3,850 `On Error GoTo`
   statements target a label called `errores`; 3,735 procedures define it. The remaining ~60
   use `TestFail`, `utc_ErrorHandling`, `ManejoError`, `ErrorHandler` — test scaffolding and
   two vendored libraries (`JsonConverter`, `UtcConverter`).

2. **`1000` is a sentinel, not an error code.** 5,152 of 5,241 `Err.Raise` calls raise the bare
   literal `1000`. The guard `If Err.Number <> 1000` in the handler means *"an inner procedure
   already wrote a human-readable message into the error channel — do not overwrite it with
   this frame's generic one."* It is a hand-rolled exception chain built on one magic number.

3. **The real error propagation channel is not VBA's.** Only **16** handlers out of 3,774
   re-raise with `Err.Raise`. About **2,788 (74%)** write the message into an object field —
   `m_Error`, `p_Error`, `Me.Error` — and return normally. The caller then checks that field.
   VBA's own error mechanism is used to *unwind* one frame; the message travels through a
   module-level variable.

That third point is the hinge of this entire document. **Error propagation in this codebase is
a data flow through module-level variables, and `vba-node-discovery-plan.md` task T9 already
models module-level variables with read/write references.** The error flow becomes traceable
as a side effect of T9 — this plan only has to *label* those reads and writes as the error channel.

### 2.2 The census

| Construct | Count | Note |
|---|--:|---|
| Procedures | 4,817 | across the three projects |
| …with `On Error GoTo <label>` | 3,777 | **78.4%** |
| …with only `On Error Resume Next` | 227 | 4.7% |
| …**with no protection at all** | **813** | **16.9%** — see §3.1 |
| `On Error GoTo <label>` statements | 3,850 | 337 procedures have more than one |
| `On Error Resume Next` | 584 | |
| `On Error GoTo 0` (reset) | 38 | **~546 Resume-Next scopes are never closed** |
| Line labels defined | 3,912 | |
| …that are error handler targets | 3,776 | **96.5%** |
| …that are pure control flow | 136 | 3.5% — `siguiente`, `salir`, `fin`, `Teardown` |
| `GoTo <label>` (all forms) | 4,237 | dominated by `On Error GoTo`; free-form jumps are rare |
| Dangling `GoTo` targets | **1** | a label referenced but not defined in its procedure |
| `Resume Next` | 585 | essentially all of them are part of `On Error Resume Next` |
| `Resume <label>` | 3 | |
| Bare `Resume` | 0 | |
| `Err.Raise` | 5,241 | 5,152 are the literal `Err.Raise 1000` |
| `Err.Number` | 4,898 | almost always `<> 1000` |
| `Err.Description` | 3,731 | |
| `Err.Clear` | 31 | |
| `Err.Source` | 1 | |

> These are the original hand-rolled figures. Task **E1** re-measured them with the committed
> probe; where the two disagree the probe wins, and the reconciliation table lives in E1 below.

### 2.3 What the handlers do

3,774 handler bodies, classified by their dominant statement. **These are the probe's
figures** (`npm run probe:vba`), not the original hand-rolled estimate — see the note below.

| Behaviour | Count | Share |
|---|--:|--:|
| Writes the error channel only (`m_Error` / `p_Error` / `g_Error` / `Me.Error`) and returns | 2,681 | 71.0% |
| Writes the channel **and** displays (`mixed`) | 921 | 24.4% |
| `MsgBox` / `Debug.Print` only — terminal, mostly form code-behind | 52 | 1.4% |
| Re-raises with `Err.Raise` only | 14 | 0.4% |
| No recognised signal (`unknown`) | 106 | 2.8% |

Non-exclusive signal totals, which is what the earlier estimate was really counting:

| Signal | Handlers |
|---|--:|
| writes the error channel | 3,602 |
| displays | 971 |
| re-raises | 16 |

> **The caveat this section used to carry is now resolved.** The original classifier put 889
> bodies in the channel bucket outright and 1,899 in an "other" bucket, and a 12-sample audit
> of "other" found 12 of 12 were `p_Error = ...` writes its regex missed; the reported 2,788
> (74%) was the sum of two buckets rather than a measurement. The probe's channel matcher takes
> its names from a configurable list (`--error-channel`, default `m_Error` / `p_Error` /
> `g_Error` / `Error`), so the real figure is **3,602 handlers touching the channel** — the old
> estimate under-counted by ~814. The 970/16 display and re-raise figures were, as stated,
> essentially exact: the probe reads 971 and 16.
>
> `mixed` is now its own bucket rather than being resolved by evaluation order, which is where
> most of the difference between the exclusive and non-exclusive tables comes from: 921 handlers
> both record the error and show it.

### 2.4 What the extractor does with all this today

**Nothing — and, importantly, nothing wrong.** Verified by running this through the real extractor:

```vba
Public Sub Guardar()
    On Error GoTo errores
    Call Escribir
    Exit Sub
errores:
    If Err.Number <> 1000 Then m_Error = "Guardar: " & Err.Description
End Sub
```

Output: one `function` node, one `contains` edge, one unresolved reference for `Escribir`.
The label line `errores:` produces no node, no edge and — this is the part worth knowing —
**no spurious unresolved reference**. `On Error GoTo errores` does not leak `errores` into the
identifier scan either.

So this is a clean slate: there is nothing to fix first, and nothing that will start
double-counting when the feature lands.

---

## 3. The questions worth answering

A feature that adds nodes has to earn them. These are the questions this codebase cannot
currently answer any other way, ordered by how much they are worth.

### 3.1 "Which risky procedures have no error handling?"

**813 procedures (16.9%) have no `On Error` at all.** That alone is not actionable — many are
one-line property accessors where a handler would be noise. Narrowing:

| Filter | Count |
|---|--:|
| Unprotected procedures | 813 |
| …longer than 5 statements | **310** |
| …that also touch the database, `DoCmd`, the filesystem or a type conversion | **185** |

Sample of what falls out — all from one class, all write paths, all unprotected in a codebase
where 78% of procedures are protected:

```
ExpedienteOperaciones.cls :: RegistrarAnualidad      (37 loc)
ExpedienteOperaciones.cls :: RegistrarAnexo          (37 loc)
ExpedienteOperaciones.cls :: RegistrarComercial      (27 loc)
ExpedienteOperaciones.cls :: RegistrarCPV            (26 loc)
ExpedienteOperaciones.cls :: RegistraResponsable     (25 loc)
```

This is the highest-value question in the list. It is a real defect class, it is invisible to
grep (you cannot grep for an *absence* scoped to a procedure body), and the answer is a short
list a human can act on.

### 3.2 "Which `On Error Resume Next` scopes are never closed?"

584 `On Error Resume Next` against 38 `On Error GoTo 0`. Roughly **546 scopes suppress every
error to the end of the procedure**. Some of that is deliberate (`Resume Next` around an
optional delete), most is not. Same shape of answer as 3.1: a list, scoped, actionable.

### 3.3 "Where does an error surface to the user?"

The handler behaviour split (§2.3) is the answer, but only if it is attached to the procedure
in the graph. Then "which user-facing form displays the error raised in `Riesgo.Guardar`"
becomes a traversal: follow the calls until a procedure whose handler is a `MsgBox`.

### 3.4 "What does this procedure do *only* when things go wrong?"

Today every call made inside a handler is attributed to the enclosing procedure, exactly like a
call on the happy path. So `Riesgo.Guardar → MsgBox` and `Riesgo.Guardar → Escribir` look
identical in the graph, when one runs always and the other runs only on failure. Marking the
handler region turns every existing call edge into error-aware data at almost no cost.

### 3.5 Questions deliberately **not** in scope

- **"Which errors can reach here?"** Not statically knowable in VBA, and the sentinel `1000`
  would make any answer meaningless anyway.
- **"What is the control flow inside a procedure?"** That is a CFG. CodeGraph is a symbol
  graph. Out of scope, permanently.
- **A node per `Err.Raise` site.** 5,152 of them, all raising the same number. Zero information
  per node.

---

## 4. The model — and what it rejects

### 4.1 The decision

**Annotate the existing nodes and edges. Do not create a parallel error graph.**

Three reasons, in order of weight:

1. **The information is per-procedure, not per-label.** Every question in §3 is of the form
   "which *procedures* …". The label is an implementation detail of the answer, not the subject
   of it. A node whose every instance is called `errores` and sits at the bottom of its
   procedure is a node that carries one bit — *"a handler exists"* — which is a boolean field.

2. **Error propagation here is variable data flow, and T9 already models it.** Once
   `m_Error` / `p_Error` are `variable` nodes with typed read/write references, the chain
   `Riesgo.Guardar writes m_Error → Form_X.cmdGuardar_Click reads m_Error` exists in the graph.
   This plan needs to *tag* those references, not invent a second mechanism for them.

3. **The repo's own doctrine says so.** `CLAUDE.md`: *"tracking every local would explode the
   graph"* — the reason the def-use frontier was left uncovered upstream. 3,900 label nodes is
   the same failure with a different name.

### 4.2 What gets added

Nothing in `NODE_KINDS`. Nothing in `EdgeKind`. Only metadata:

**On the `function` node** — a single `errorPolicy` object:

```
errorPolicy: {
  protection:       'handler' | 'resume-next' | 'none'
  handlerLabel:     string | null      // 'errores' in 98% of cases
  handlerStartLine: number | null      // first line after the label
  handlerEndLine:   number | null      // the procedure's End Sub/Function/Property
  behavior:         'channel' | 'display' | 'reraise' | 'mixed' | 'unknown' | null
  handlerCount:     number             // >1 means the procedure swaps handlers mid-body
  resumeNextOpen:   boolean            // On Error Resume Next never closed by On Error GoTo 0
  danglingTarget:   string | null      // On Error GoTo <label> where <label> is never defined
}
```

**On every edge and unresolved reference emitted from inside a handler region** — one flag:

```
metadata.inErrorHandler: true
```

**On reads and writes of a configured error-channel variable** (needs T9):

```
metadata.errorChannel: true
```

That is the entire model. It is three fields and one object, and it answers all four questions
in §3.

### 4.3 The alternative, and when to revisit it

The rejected design is a `label` node kind plus a `handles-error` edge kind. It is not wrong in
principle — it is wrong *for this corpus at this time*, because 96.5% of labels are one
convention and nothing queries them individually.

**Revisit it if, and only if, one of these becomes true:**

- A consumer needs to query handler *bodies* as independent units (for example, "find every
  handler that calls the logging helper" — which E3's `inErrorHandler` flag also answers, so
  the bar is a query that flag genuinely cannot serve).
- A project appears in the corpus where labels are used as real control flow. Today that is
  136 labels out of 3,912 — an order of magnitude below the threshold where a node kind pays
  for itself.

Task **E6** holds this option open, and it stays blocked on maintainer sign-off.

---

## 5. Tasks

Same conventions as the main plan: one task = one issue = one PR, tests and a `CHANGELOG.md`
entry with every one, `npm test` plus CI green as the only merge gate, conventional commits,
no `Co-Authored-By`, no version bump, no push or tag.

**Dependency:** E4 requires `T9` (module-level `variable` nodes) from `vba-node-discovery-plan.md`.
E1–E3 and E5 are independent of the main plan and can land in any order relative to it.

---

### E1 — Extend the coverage probe with error metrics

**Goal.** Make every number in §2 reproducible before anything depends on them, and fix the
one measurement this analysis left approximate.

**Files.** `scripts/vba-coverage-probe.mjs` (created by main-plan `T0` — if `T0` has not landed,
this task creates the probe as specified there and adds the error section on top).

**Implementation.** Add an `errorHandling` block to the probe output:

```
proceduresTotal
proceduresByProtection      { handler, resumeNext, none }
unprotectedOverFiveLoc
unprotectedTouchingIo       (OpenRecordset | Execute | DoCmd. | CreateObject | file ops | conversions)
handlersByBehavior          { channel, display, reraise, mixed, unknown }
handlerLabelNames           { [label]: count }
resumeNextOpenScopes
danglingGotoTargets         [{ file, procedure, label }]
proceduresWithMultipleHandlers
```

**Get the behaviour classifier right this time.** The channel regex must cover the real
variants — `m_Error`, `p_Error`, `g_Error`, `Me.Error`, `<obj>.Error` — as a configurable list,
not a hardcoded one; §2.3's caveat exists because a narrow regex under-counted by 1,899. Classify
`mixed` when a body both writes the channel and displays, rather than letting evaluation order
decide silently.

**Tests.** `__tests__/vba-coverage-probe-errors.test.ts` — a fixture module with one procedure
of each protection class and one of each handler behaviour; assert every counter exactly.

**Acceptance.** The probe reproduces §2.2 within the stated tolerance, and reports the exact
channel-vs-display split that §2.3 could only approximate. Paste the output in the PR body;
it supersedes the numbers in this document, and this document should be updated to match.

**Changelog.** None (internal tooling). Say so in the PR body.

**Status — landed, and the corpus re-measured.** The `errorHandling` block ships in
`scripts/vba-coverage-probe.mjs`, covered by `__tests__/vba-coverage-probe-errors.test.ts`.
Reproduce every figure in §2.2 and §2.3 with:

```
npm run probe:vba -- <expedientes>/src <gestion-riesgos>/src <hps-solicitudes>/src
```

The probe agrees with §2.2 on the headline counts and diverges on four, which the probe's
figures now supersede because it is the tested, committed classifier and the census was not:

| Metric | §2.2 census | Probe | Note |
|---|--:|--:|---|
| Procedures | 4,817 | 4,817 | exact |
| with `On Error GoTo` | 3,777 | 3,774 | within 3 |
| only `On Error Resume Next` | 227 | 227 | exact |
| no protection | 813 | 816 | within 3 |
| unprotected and > 5 statements | 310 | 305 | statement-line counting differs slightly |
| ...of those, touching I/O | 185 | 103 | **diverges** — the probe's I/O marker list is narrower and explicit |
| line labels defined | 3,912 | 3,911 | within 1 |
| ...pure control flow | 136 | 137 | within 1 |
| `Resume Next` scopes never closed | ~546 | 534 | probe tracks open scopes per procedure rather than subtracting totals |
| procedures with more than one handler | 337 | 47 | **diverges** — the census counted `On Error` *statements* per procedure, the probe counts `On Error GoTo <label>` sites only |
| dangling `GoTo` targets | 1 | 0 | **diverges** — the probe resolves labels per procedure scope; no procedure in the corpus has an unresolved target |
| `Err.Raise` | 5,241 | 5,189 | within 1% |
| `Err.Raise 1000` sentinel | 5,152 | 5,133 | within 1% |

The three divergences are definitional, not defects on either side; each is documented in the
probe's docblock so a later reader can tell which question a number answers.

---

### E2 — `errorPolicy` on the procedure node

**Goal.** Answer §3.1 and §3.2 with a graph query.

**Anchors.**
- `src/extraction/vba/procedures.ts` — the `procedure` rule's emit; where the `function` node
  literal is built (it currently has **no** `metadata` field at all — main-plan `T8` also adds
  one; if `T8` landed first, extend its object rather than replacing it).
- `src/extraction/vba/context.ts` — `ProcInfo` and the procedure stack; this is where the
  per-procedure state must accumulate.
- `src/extraction/vba/call-sweep.ts` — the per-line driver; the new classifier hooks in beside
  the existing sweeps (~line 298 onward).
- `src/extraction/vba/constants.ts` — `PROCEDURE_END_RE`, which already handles the
  colon-separated single-line procedure form (#208). Reuse it; do not write a second one.

**Implementation.**

1. New classifier `src/extraction/vba/errors.ts` exporting a `RULES` table, registered in
   `VBA_RULE_TABLES` in `vba-extractor.ts` and added to `REQUIRED_DISPATCH_TABLES` so the
   load-time invariant covers it.

2. Rules, all `scan: 'masked'` so a string literal containing `"On Error GoTo x"` cannot trigger
   them (this is exactly what #209 fixed for other sweeps — inherit that discipline):

   | Rule id | Pattern | Effect on the open `ProcInfo` |
   |---|---|---|
   | `on-error-label` | `On\s+Error\s+GoTo\s+(\w+)` where the target is not `0`/`-1` | record label, `handlerCount++`, `protection = 'handler'` |
   | `on-error-resume-next` | `On\s+Error\s+Resume\s+Next` | `protection = 'resume-next'` unless already `'handler'`; open a resume-next scope |
   | `on-error-reset` | `On\s+Error\s+GoTo\s+0` | close the open resume-next scope |
   | `line-label` | `^\s*(\w+):\s*$` | record the definition; if it matches a recorded handler label, mark the handler region open at the next line |

3. At procedure end (the existing `proc-end` boundary in `enums-consts.ts` already tracks it —
   consume that signal, do not add a second end detector): close the handler region, set
   `handlerEndLine`, resolve `danglingTarget` by checking every recorded `On Error GoTo` label
   against the recorded definitions, and set `resumeNextOpen` from whether a scope is still open.

4. Attach the resulting `errorPolicy` to the function node.

**Precision rules — the gates this task ships with:**

- A label defined but never targeted by an `On Error GoTo` is **control flow, not a handler**.
  It sets nothing. (136 such labels in the corpus; misclassifying them would corrupt §3.3.)
- A `.form.txt` or `.report.txt` never reaches this classifier — the form extractor owns those
  and emits no code nodes.
- `On Error GoTo -1` is valid VBA (clears the current error) and appears 0 times here. Treat it
  as a reset like `GoTo 0`, and add a fixture so the zero-occurrence case is still pinned.

**Tests.** `__tests__/extraction-vba-error-policy.test.ts`:
- handler / resume-next / none — one procedure each, exact `protection` value;
- `handlerStartLine` and `handlerEndLine` bracket exactly the handler body;
- a label defined but never targeted → `protection: 'none'`, no handler region;
- two `On Error GoTo` statements in one procedure → `handlerCount: 2`;
- `On Error Resume Next` then `On Error GoTo 0` → `resumeNextOpen: false`; without the reset → `true`;
- `On Error GoTo noExiste` with no such label → `danglingTarget: 'noExiste'`;
- a colon-separated single-line procedure with a handler parses correctly (regression guard for #208);
- **`s = "On Error GoTo errores"` inside a string literal sets nothing** (regression guard for #209);
- a `Property Get` with a handler gets the same treatment as a `Sub`.

**Acceptance.**
- `errorPolicy.protection` distribution matches E1's census, measured: 3,774 / 227 / 816.
- Node count is **unchanged**. Edge count is **unchanged**. This task adds zero nodes and zero edges.
- `danglingTarget` is non-null for **0** procedures in the corpus. This line previously predicted
  exactly 1, from the pre-probe hand census; E1's reconciliation table above already records that
  the committed probe measures 0. The corpus does contain a handler literally named `noExiste`,
  but it is defined in its own procedure and is therefore correctly not dangling — almost
  certainly what the hand census mistook for an unresolved target. A future project that DOES
  have one should still surface it, so keep the field and name any hit in the PR body.

**Changelog (New Features).** "Procedures now record how they handle errors — whether they have
a handler, suppress errors silently, or have no protection at all — so unguarded code paths can
be found without reading every module."

---

### E3 — Mark what happens inside the handler

**Goal.** Answer §3.4: separate "runs always" from "runs only on failure".

**Depends on:** E2 (the handler region boundaries).

**Anchors.** `src/extraction/vba/call-sweep.ts` — the per-line dispatch, which already knows the
current procedure from `ctx` and now knows the handler region from E2. Every emitter it drives
(`scanCallSites`, `scanMeControlReferences`, `scanSqlInLine`, `scanDoCmdOpenCalls`,
`scanFormsBang`, `sweepTempVars`) inherits the flag from one place.

**Implementation.**

1. Expose `ctx.inErrorHandler(lineNum): boolean` derived from the current procedure's
   `handlerStartLine`/`handlerEndLine`.
2. Set `metadata.inErrorHandler = true` on every edge and every `UnresolvedReference` emitted
   while that is true. Set it **at the single point where edges and references are pushed**, not
   in each emitter — the context already centralises `emitReference()`, and the edge pushes are
   few enough to route through one helper. Six emitters each remembering to set a flag is six
   places to forget it.
3. Derive `errorPolicy.behavior` here rather than in E2, since this is where handler-body
   content is already being classified: `display` when the body's calls include `MsgBox`,
   `channel` when it writes a configured error-channel variable, `reraise` on `Err.Raise`,
   `mixed` when more than one applies.

**Note on a dependency worth checking first.** Statement-form calls (`Call Escribir`,
`Escribir 1, 2`) are currently classified `unqualified-ident`, while
`docs/vba-reference-kinds.md` documents `calls` as covering "statement-form Sub call" with
`Call LimpiaBuffer` as the example. Handler bodies are almost entirely statement-form calls
(`MsgBox "…"`), so this task's `behavior` classification sits directly on that disagreement.
**Resolve which one is intended before implementing** — either the doc or the classifier is
wrong, and this task should not encode the ambiguity.

**Tests.** `__tests__/extraction-vba-error-handler-region.test.ts`:
- a call before `Exit Sub` has no flag; the same call after the handler label has `inErrorHandler: true`;
- a procedure with no handler flags nothing;
- an SQL statement inside a handler still emits its table reference, now flagged;
- `behavior` resolves to `display` / `channel` / `reraise` / `mixed` on four fixtures;
- the region ends at `End Sub` — a call in the *next* procedure is not flagged (off-by-one guard).

**Acceptance.** Roughly 970 procedures show `behavior: 'display'` and roughly 2,788 show
`'channel'` (E1's exact figures are the target). Total edges and unresolved references are
**unchanged** — this task only adds a field to existing rows.

**Changelog (New Features).** "Calls that only run when something goes wrong are now marked as
such, so a procedure's error path can be told apart from its normal one."

**Status — landed (#260).** The dependency question above is resolved: **#265** made the two
unambiguous statement-call forms (`Call X`, `X arg1, arg2`) classify as `calls`, so the
`behavior` classifier reads real calls rather than bare-identifier reads.

`metadata.inErrorHandler` is stamped at ONE place — `closeErrorPolicy` in
`src/extraction/vba/errors.ts`, over the procedure's own slice of `ctx.edges` /
`ctx.unresolvedReferences` — rather than in each of the six emitters, so a new emitter inherits
it. `ctx.inErrorHandler(lineNum)` exposes the same predicate to anything that needs it during
the walk.

`behavior` mirrors the probe's classifier by hand — the same channel-write matcher, the same
display-call list, the same `Err.Raise` test, the same statement splitter with its `If … Then`
guard strip, and the same "`mixed` means more than one signal" rule. The probe cannot import the
extractor (it must also run against a shipped `dist/`), so the two copies are kept in sync by
hand: change one, change the other.
`__tests__/extraction-vba-error-handler-region.test.ts` runs BOTH classifiers over one fixture
and asserts they agree procedure by procedure, so a fork fails there rather than in a corpus
re-measurement nobody runs.

Measured over the three corpus projects, the extractor's distribution is **identical to the
probe's**, procedure by procedure — 0 disagreements over all 4,817:

| `behavior` | Procedures |
|---|--:|
| `channel` | 2,681 |
| `mixed` | 921 |
| `unknown` | 106 |
| `display` | 52 |
| `reraise` | 14 |
| `null` (no handler: 227 resume-next + 816 unprotected) | 1,043 |

This is the **exclusive** table of §2.3, not the raw signal totals — the ~970 `display` /
~2,788 `channel` figures the task was written against were the non-exclusive counts
(3,602 channel / 971 display / 16 re-raise as raw signals), which the probe already superseded.

Totals are unchanged, as required: 26,089 nodes, 29,521 edges and 26,755 unresolved references
on this branch and on `origin/main`, with identical `nodesByKind` / `edgesByKind` /
`unresolvedByKind` breakdowns. 57 edges and 2,823 unresolved references gained the flag.

One shape the corpus turned out NOT to have: a handler label carrying a trailing statement on
its own line (`errores: MsgBox "x"`) occurs **0 times** in 3,774 handlers. The `behavior`
signals read that trailing statement anyway, because the probe does; the `inErrorHandler` flag
uses `[handlerStartLine, handlerEndLine]` — the region published on the node — so an edge
emitted on the label's own line would not be flagged. On this corpus the two can never disagree.

---

### E4 — Recognise the error channel

**Goal.** Make cross-procedure error propagation traceable — the thing that actually carries the
message in this codebase.

**Depends on:** main-plan **T9** (module-level `variable` nodes with read/write references).
Do not start before T9 has merged.

**Implementation.**

1. Config knob `vba.errorChannel`: an array of variable-name patterns, defaulting to
   `["m_Error", "p_Error", "g_Error", "Error"]`. Threaded through the `VbaExtractionOptions`
   object from main-plan `T1`. Same validation shape as `vba.sqlWrappers`; **no user-supplied
   regex** (same reasoning as T2 — this runs per line).
2. When a read or write reference targets a variable matching the channel, set
   `metadata.errorChannel: true` alongside the existing `property-get` / `property-set` kind.
3. Nothing else. No new edge kind. The propagation chain is
   `procedure --property-set(errorChannel)--> m_Error --property-get(errorChannel)--> caller`,
   and both hops already exist once T9 has landed.

**End-to-end requirement.** Verify on a real class that the full chain
`inner procedure writes channel → outer procedure reads channel → form displays it` is
traversable in the graph. A one-hop version is the half-bridged flow `CLAUDE.md` forbids —
if the second hop does not connect, the task is not done.

**Tests.** `__tests__/extraction-vba-error-channel.test.ts`:
- `m_Error = "x"` inside a handler → `property-set`, `errorChannel: true`, `inErrorHandler: true`;
- `If m_Error <> "" Then` in a caller → `property-get`, `errorChannel: true`;
- a variable named `ErrorCount` does **not** match (the pattern is a name, not a substring);
- a custom `vba.errorChannel: ["lastFailure"]` matches that name and keeps the defaults;
- the traversal test above.

**Acceptance.** On `Riesgo.cls` / `Edicion.cls` / `ExpedienteOperaciones.cls` — the three
heaviest error-channel users — every handler that writes the channel produces a flagged
reference, and at least one complete write→read→display chain is traversable. Show the chain in
the PR body.

**Changelog (New Features).** "The variables this kind of Access project uses to pass error
messages between objects are now recognised as such, so an error can be followed from where it
happened to where the user sees it."

---

### E5 — Surface the four questions

**Goal.** A feature nobody can query is a feature nobody has. E2–E4 put the data in the graph;
this task makes it reachable.

**Implementation.** Documented SQL in a new `docs/vba-error-handling.md` (the consumer-facing
sibling of this plan, in the same style as `docs/vba-reference-kinds.md`), covering:

1. **Unprotected risky procedures** — `errorPolicy.protection = 'none'`, more than five
   statements, and at least one outgoing edge to a SQL table, a `DoCmd` target or a filesystem call.
2. **Open `Resume Next` scopes** — `errorPolicy.resumeNextOpen = true`.
3. **Where an error surfaces** — procedures with `behavior = 'display'`, and what reaches them.
4. **Dangling handler targets** — `errorPolicy.danglingTarget IS NOT NULL`.

Then decide with the maintainer whether any of these deserves a CLI surface (`codegraph vba-lint`
or a `stats` extension). **Ask — do not build a command unprompted.** `CLAUDE.md` is explicit
that CodeGraph provides code context, not product requirements, and a new command is product.

If the MCP tools should mention any of this to agents, the single source of truth is
`src/mcp/server-instructions.ts`. Nothing else.

**Tests.** Run each documented query against a temp index built from fixtures and assert the
row counts. Follow `db-vba-call-stub-queries.test.ts`, which does exactly this for the stub
resolver's published queries.

**Acceptance.** Every query in the doc runs against a real index and returns the expected shape.
Query 1 returns ≈185 rows on the corpus; hand-check the top 10 and confirm each is genuinely
unprotected — if any has a handler the extractor missed, E2 has a bug.

**Changelog (New Features).** "New guidance shows how to find Access code that runs without
error handling, and where errors end up being shown to the user."

---

### E6 — `label` nodes and `handles-error` edges — **blocked, do not implement**

Held open deliberately. This is the design E2 rejected, and it stays rejected until a query
appears that `inErrorHandler` genuinely cannot serve (§4.3 sets the bar).

If it is ever unblocked, it needs, in this order: a written statement of the query that forces
it; maintainer sign-off on the new `NodeKind` and `EdgeKind`; and a node-budget forecast
(≈3,900 nodes and ≈4,200 edges on this corpus, roughly doubling the symbol count). Ship it on
one project and measure before rolling it out.

**Do not implement E6 as part of E1–E5. Do not implement it because it seems more complete.**

---

## 6. Guardrails specific to this feature

The main plan's three rules apply unchanged. Three more, particular to error handling:

1. **A label is not a handler until something targets it.** 136 labels in the corpus are pure
   control flow. Classifying them as handlers would poison every count in §3 and produce a
   confidently wrong answer to "does this procedure handle errors" — the worst failure mode
   available here.

2. **Never infer that an unprotected procedure is a bug.** The graph reports the fact; the
   human decides. A one-line property accessor with no handler is correct code. This is why
   §3.1's query filters on size *and* on I/O — and why the changelog wording says "can be found",
   not "are flagged".

3. **This feature adds no nodes and no edges. If a PR in E1–E5 changes the node or edge count,
   something is wrong.** That invariant is cheap to assert and it is the single best protection
   against this task quietly turning into E6.

**Per-PR checklist:** the main plan's checklist, plus:

- [ ] Node count and edge count unchanged (paste both, before and after)
- [ ] String-literal guard test present (a `"On Error GoTo x"` inside a string sets nothing)
- [ ] Colon-separated single-line procedure fixture passes (#208 regression)
- [ ] `EXTRACTION_VERSION` — bump once, in the last PR of E1–E5. Metadata on existing nodes
      still changes what a re-index would produce, so the wave earns one bump; each individual
      task does not.

---

## 7. Appendix

### 7.1 Handler label names (all occurrences)

| Label | As `On Error` target | Defined |
|---|--:|--:|
| `errores` | 3,789 | 3,735 |
| `TestFail` | 19 | 18 |
| `utc_ErrorHandling` | 15 | 15 |
| `ManejoError` | 13 | 13 |
| `ErrorHandler` | 4 | 3 |
| `salir` | 3 | 5 |
| `errMem`, `noExiste`, `filasInvalidas`, `Manejador`, `ErrorEliminar`, `AjustarTamaño_TratamientoErrores` | 1 each | — |

Pure control-flow labels (never an `On Error` target): `siguiente` (61), `siguienteRiesgo` (7),
`Teardown` (7), `siguienteEdicion` (6), `fin` (4), `SiguienteRiesgoExterno` (3),
`SiguienteRiesgoExt` (3), `siguienteProyecto` (2), `recalcular` (2), and 41 others.

### 7.2 The most common handler bodies

Normalised, string literals elided:

```
850  If Err.Number <> 1000 Then m_Error = "…" & vbNewLine & Err.Description End If Me.Error = m_Error
609  If Err.Number <> 1000 Then p_Error = "…" & Err.Description End If
540  If Err.Number <> 1000 Then m_Error = "…" & Err.Number & vbNewLine & "…" & Err.Description
510  If Err.Number <> 1000 Then p_Error = "…" & vbNewLine & Err.Description End If
291  DoCmd.Hourglass False  If Err.Number <> 1000 Then m_Error = "…" & Err.Number & vbNewLine & …
194  If Err.Number <> 1000 Then m_Error = "…" & vbNewLine & Err.Description End If m_Error = Me.Error
```

Note the fifth: `DoCmd.Hourglass False` before recording. Cleanup-then-record is a distinct
handler shape and E3's `behavior` classifier must not be confused by the leading `DoCmd` call.

### 7.3 Reproducing the census

The scan is procedure-scoped — it tracks `On Error` statements, label definitions and `GoTo`
targets per procedure body, which is why it can answer questions grep cannot. Task **E1** turns
it into the committed probe; until then, the shape is:

```
for each .bas/.cls file
  for each procedure (PROC_RE … PROCEDURE_END_RE)
    collect: On Error GoTo <label> | On Error Resume Next | On Error GoTo 0
             label definitions (^\s*\w+:\s*$)
             GoTo targets
             handler body (label line → End Sub/Function/Property)
    classify: protection, behavior, resumeNextOpen, danglingTarget
```

Strip trailing comments before matching, and run against the **preprocessed** source once this
is inside the extractor — line continuations must be joined first or a handler split across
lines is misread.
