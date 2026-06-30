# SDD spec — vba-handler-backtrace (Phase 1 skill)

**Status**: BLOCKED on `ardelperal/codegraph-vba#10` (round-2 broad: requires Event + Type + Declare modeling).
**Phase**: 1 (Bisturí puro) of `ardelperal/VBA_TOOLKIT_BENCH#1`.
**Issue tracker**: https://github.com/ardelperal/VBA_TOOLKIT_BENCH/issues/10

## Goal

Ship the `vba-handler-backtrace` bisturí skill: given a control event (form Load, button Click), trace back through the wiring chain (control event → handler sub → called helpers → DAO calls → SQL).

## Pre-flight (must be true before ship)

- [ ] `ardelperal/codegraph-vba#10` closed (round-2 broad — Event + Type + Declare modeling)

Partial unblock: `ardelperal/codegraph-vba#11` (Event P1 alone) ships Event modeling — sufficient for top-of-chain tracing (control event → handler sub) but NOT for the full chain that includes Type/Declare-aware helpers.

## Contract

### Input

```json
{
  "form": string — nombre del form (Form_FormX),
  "event": string — nombre del evento (cmdSave_Click, Form_Load, etc.),
  "depth": number — opcional, default 5 (chain depth to traverse)
}
```

### Output

```json
{
  "form": "Form_FormX",
  "event": "cmdSave_Click",
  "chain": [
    { "step": 1, "kind": "form_event", "target": "Form_FormX", "handler_sub": "Form_FormX.cmdSave_Click", "line": 234 },
    { "step": 2, "kind": "helper_call", "target": "Form_FormX", "sub": "PersistRecord", "line": 267 },
    { "step": 3, "kind": "dao_call", "target": "CurrentDb", "method": "Execute", "line": 281, "sql_hint": "INSERT INTO ..." }
  ],
  "depth_reached": 3,
  "warnings": [
    { "reason_code": "MAX_DEPTH", "message": "chain truncated at depth 5; rerun with depth=10" }
  ]
}
```

### mcp_tools_used

- `codegraph-vba_codegraph_explore` (form-event edges, handler sub edges, DAO call edges) — REQUIRED, blocked
- `dysflow_export_modules` (source snapshot)

### Acceptance criteria

- Latency: <2s en proyectos con 300+ archivos
- Chain completeness: cada step es un `form_event`, `helper_call`, o `dao_call` (no pasos ciegos)
- SQL hint: si la DAO call es Execute/SQL/QueryDef, extrae el SQL fragment (first 200 chars)
- Type-aware tracing: si el helper toma un `Type` parameter, el step incluye `type_refs`
- Bisturí: 1 input pattern, 1 structured output
- Cycle detection: si A → B → A, marca `cycle_detected: true` y corta

### Failure modes

- Form no encontrado → hard-fail
- Event handler no encontrado → warning, devuelve chain vacío
- Max depth reached → warning, parcial chain visible
- Type declaration missing → Type segment del chain es null (fallback OK)

## RED tests (will fail until codegraph-vba ships round-2 broad)

### Test 1: cmdSave_Click → persistRecord → CurrentDb.Execute

**Input**: `{ "form": "Form_FormRiesgosGestionRiesgo", "event": "cmdSave_Click" }`

**Expected**: chain 3 steps: form_event → helper_call (`PersistRecord`) → dao_call (`CurrentDb.Execute` con sql_hint).

**Pass**: chain completo en <=3 steps, includes sql_hint.

### Test 2: Form_Load → InitForm chain con UDT

**Input**: `{ "form": "Form_FormX", "event": "Form_Load" }` (Form_Load llama InitForm pasando una UDT `TInitParams`)

**Expected**: chain 3 steps, segundo step incluye `type_refs: [{ name: "TInitParams", module: "modTypes" }]`

**Pass**: Type ref rastreado.

### Test 3: Cycle detection (A → B → A)

**Input**: `{ "form": "Form_FormX", "event": "cmdRefresh_Click" }` (Refresh llama ReloadAll, ReloadAll llama Refresh — cycle)

**Expected**: chain de 2 steps + `cycle_detected: true`

**Pass**: skill no cae en loop infinito.

## Tasks

1. Wait for codegraph-vba#10 to close.
2. Implement SKILL.md + references/examples.md.
3. Update skill-registry.md.
4. Update VBA_TOOLKIT_BENCH#10 with checkmarks.
5. GREEN verification.
6. ARCHIVE: engram observation.

## References

- Plan: `plans/plan-ambicioso-dysflow-codegraph-vba-2026-06-29.md` (Phase 1.3)
- Upstream graphs: `ardelperal/codegraph-vba#10` (round-2 broad — required for full chain)
- Partial unblock: `ardelperal/codegraph-vba#11` (Event P1 — sufficient for top-of-chain only)
- Cross-ref: VBA_TOOLKIT_BENCH#1 (Phase 1 epic)
