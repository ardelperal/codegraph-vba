# SDD spec — vba-event-tracer (Phase 1 skill)

**Status**: BLOCKED on `ardelperal/codegraph-vba#11` (Event declarations P1) or `ardelperal/codegraph-vba#10` (round-2 broad).
**Phase**: 1 (Bisturí puro) of `ardelperal/VBA_TOOLKIT_BENCH#1`.
**Issue tracker**: https://github.com/ardelperal/VBA_TOOLKIT_BENCH/issues/9

## Goal

Ship the `vba-event-tracer` bisturí skill: trace `Event ... RaiseEvent` declarations and find every subscriber / handler across the project.

## Pre-flight (must be true before ship)

- [ ] `ardelperal/codegraph-vba#10` closed (round-2 broad unblocks Event + Type + Declare)
- OR
- [ ] `ardelperal/codegraph-vba#11` closed (Event declarations P1 alone, fast path)

The fast path via #11 ships Event without waiting for Type/Declare. The skill only requires Event modeling to function — Type/Declare enable richer queries but aren't required for the core "trace events" functionality.

## Contract

### Input

```json
{
  "event_name": string — nombre del evento o qualified class.event,
  "or": { "class_name": string — alternativa: trazar todos los eventos de una clase },
  "access_path": string — ruta al .accdb (requerido)
}
```

### Output

```json
{
  "event_declarations": [
    { "module": "Form_FormX", "line": 25, "signature": "Public Event SomethingChanged(ByVal NewValue As String)" }
  ],
  "raise_sites": [
    { "module": "Form_FormX", "line": 142, "context": "RaiseEvent SomethingChanged(\"new\")" }
  ],
  "handlers": [
    { "form": "Form_FormY", "module": "Form_FormY.cls", "handler": "Form_FormY_Class_SomethingChanged", "via": "WithEvents" }
  ],
  "warnings": [
    { "reason_code": "NO_RAISERS", "message": "event is declared but never raised" }
  ]
}
```

### mcp_tools_used

- `codegraph-vba_codegraph_explore` (event node kinds, RaiseEvent edges, WithEvents edges) — REQUIRED, blocked currently
- `dysflow_export_modules` (snapshot read-only of source)
- `dysflow_dysflow_doctor` (verify project structure)

### Acceptance criteria

- Latency: <2s en proyectos con 300+ archivos
- Coverage: 0% false positives — events en strings o comments NO aparecen
- Event declaration match: signature exacta, args resueltos
- Multi-class scoping: si event_name es qualified, scope a la clase; sino show all classes
- Forms handlers: detecta `WithEvents` + handler wiring por convención VBA
- Bisturí: 1 input pattern, 1 structured output

### Failure modes

- Event no encontrado → hard-fail `EVENT_NOT_FOUND`
- Event ambiguo (mismo nombre en múltiples classes) → `EVENT_AMBIGUOUS` + pedir qualified input
- Source race (event added mid-trace) → snapshot + retry

## RED tests (will fail until codegraph-vba ships Event modeling)

### Test 1: Event con 3 raise sites + 2 handlers

**Input**: `Form_FormX.SomethingChanged`

**Expected**: 1 declaration, 3 raise_sites, 2 handlers (one WithEvents + one dynamic via Set X = New ClassY)

**Pass**: 0% false positives on the raise_sites (no string-literal false positives).

### Test 2: Event declarado pero nunca raised

**Input**: `Form_FormX.OrphanEvent`

**Expected**: 1 declaration, 0 raise_sites, 0 handlers + warning `NO_RAISERS`

**Pass**: warns about unused event.

### Test 3: Multi-class scoping

**Input**: `SomethingChanged` (sin qualifier)

**Expected**: lista todos los eventos con ese nombre en todas las classes + pide desambiguación si >1

**Pass**: handles ambiguity gracefully.

## Tasks

1. Wait for codegraph-vba#10 or #11 to close (pre-flight dependency).
2. Implement `SKILL.md` at `~/.config/opencode/skills/vba-event-tracer/SKILL.md`.
3. Implement `references/examples.md` with 4-6 RED test cases.
4. Update `~/.config/opencode/skills/.atl/skill-registry.md` to add the row.
5. Update VBA_TOOLKIT_BENCH#9 with checkmarks.
6. GREEN verification against `Gestion_Riesgos.accdb` (read-side of TDD).
7. ARCHIVE: engram observation.

## References

- Plan: `plans/plan-ambicioso-dysflow-codegraph-vba-2026-06-29.md` (Phase 1.3)
- Upstream graphs: `ardelperal/codegraph-vba#10` (round-2 broad) and `#11` (Event P1)
- Cross-ref: VBA_TOOLKIT_BENCH#1 (Phase 1 epic)
