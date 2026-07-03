# SDD spec — vba-sql-impact (Phase 1 skill)

**Status**: BLOCKED on `ardelperal/codegraph-vba#9` (round-3: `.sql` saved queries modeling).
**Phase**: 1 (Bisturí puro) of `ardelperal/VBA_TOOLKIT_BENCH#1`.
**Issue tracker**: https://github.com/ardelperal/VBA_TOOLKIT_BENCH/issues/11

## Goal

Ship the `vba-sql-impact` bisturí skill: given a saved query (`.sql` file or Access QueryDef), trace which code calls it, which forms bind to it, and which tables/columns it touches.

## Pre-flight (must be true before ship)

- [ ] `ardelperal/codegraph-vba#9` closed (round-3 — `.sql` saved queries modeled as graph nodes)

## Contract

### Input

```json
{
  "query_name": string — nombre del saved query,
  "or": { "sql_fragment": string — alternativa: pegar un fragmento SQL y buscar matches }
}
```

### Output

```json
{
  "query_name": "Qry_RiesgosActivos",
  "query_definition": { "sql": "SELECT r.id, r.descripcion FROM tblRiesgos r WHERE r.estado = 'ACTIVO'", "type": "Select" },
  "callers": [
    { "module": "Form_FormX", "line": 67, "context": "Set rs = CurrentDb.OpenRecordset(\"Qry_RiesgosActivos\")" }
  ],
  "form_bindings": [
    { "form": "Form_FormRiesgosGestionRiesgo", "control": "lstRiesgos", "record_source": "Qry_RiesgosActivos" }
  ],
  "tables_touched": [
    { "table": "tblRiesgos", "columns": ["id", "descripcion", "estado"] }
  ],
  "downstream_impact": [
    { "table": "tblRiesgos", "reason": "query reads from this table" }
  ]
}
```

### mcp_tools_used

- `codegraph-vba_codegraph_explore` (query node kinds, query-to-callers edges, query-to-table edges) — REQUIRED, blocked
- `dysflow_export_modules` (snapshot del source para identificar OpenRecordset literal strings)
- `dysflow_dysflow_query_execute` (read del QueryDef SQL)

### Acceptance criteria

- Latency: <2s en proyectos con 300+ archivos
- Callers completeness: detecta tanto OpenRecordset literal string como QueryDef reference
- Form bindings: detecta `record_source`, `rowsource`, `control_source` que apunten a la query
- Tables touched: parse del SQL → identifica tablas + columnas referenciadas (named resolution; "FROM tblRiesgos r WHERE r.estado" → tblRiesgos.estado)
- Downstream impact: lista las tablas cuyo schema change rompería esta query

### Failure modes

- Query no encontrada → hard-fail `QUERY_NOT_FOUND`
- SQL fragment match ambiguo (>1 query matchea) → hard-fail `AMBIGUOUS_FRAGMENT` + lista candidatos
- Parse SQL fails (SQL no estándar) → warning, devuelve lo que pudo parsear + lista tablas no resueltas

## RED tests (will fail until codegraph-vba ships round-3)

### Test 1: Qry_RiesgosActivos con 2 callers + 1 form binding + tabla tblRiesgos

**Input**: `Qry_RiesgosActivos`

**Expected**: 2 callers (uno OpenRecordset literal, uno QueryDef reference), 1 form binding (`Form_FormRiesgosGestionRiesgo.lstRiesgos.record_source`), tables_touched: [tblRiesgos].

**Pass**: detection comprehensive across all 3 connection types.

### Test 2: SQL fragment match

**Input**: `' SELECT id FROM tblRiesgos '`

**Expected**: lista queries que matchean (Qry_RiesgosActivos probablemente) + confidence score.

**Pass**: fragment matching works.

## Tasks

1. Wait for codegraph-vba#9 to close.
2. Implement SKILL.md + references/examples.md.
3. Update skill-registry.md.
4. Update VBA_TOOLKIT_BENCH#11 with checkmarks.
5. GREEN verification against Gestion_Riesgos.accdb.
6. ARCHIVE: engram observation.

## References

- Plan: `plans/plan-ambicioso-dysflow-codegraph-vba-2026-06-29.md` (Phase 1.3)
- Upstream graphs: `ardelperal/codegraph-vba#9` (round-3 — required)
- Cross-ref: VBA_TOOLKIT_BENCH#1 (Phase 1 epic)
