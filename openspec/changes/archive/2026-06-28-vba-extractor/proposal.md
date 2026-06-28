# Proposal: VBA / Access Language Support (regex extractors)

## Intent

Dysflow exports Access/VBA source trees (`.bas`, `.cls`, `.form.txt`, `.report.txt`). Codegraph has no VBA language; these files index as `unknown` and agents fall back to Read/Grep. This change adds `vba` via two regex extractors (Path B), restoring structural retrieval without regressing any existing language.

## Scope

### In Scope
- `VbaExtractor` (`.bas`/`.cls`): `Sub`/`Function`/`Property`, classes, `Sub New`, `Implements`, `Dim As`, `WithEvents`, calls, SQL tables in strings.
- `VbaFormExtractor` (`.form.txt`/`.report.txt`): controls → `property` nodes + one `references` edge to sibling `.cls`. **Zero** `function`/`sub`/`module` nodes from form UI.
- Wire into `extractFromSource()`; update `LANGUAGES`/`EXTENSION_MAP`/support flags/display name in `src/types.ts` + `src/extraction/grammars.ts`; teach the split in `src/mcp/server-instructions.ts`.
- Vitest: `__tests__/extraction-vba.test.ts`, `__tests__/extraction-vba-form.test.ts`.

### Out of Scope
- Legacy `.frm`/`.dsr` (obs #14703). Tree-sitter grammar (Path A — not chosen). Refactor of `EXTRACTORS` map.

## Capabilities

### New Capabilities
- `vba-code-extraction`: regex extraction of `.bas`/`.cls`. Emits `function`/`class`/`module`/`variable` nodes and `calls`/`extends`/`implements`/`references`/`contains` edges. Heuristic edges carry `provenance: 'heuristic'` + `metadata.synthesizedBy`.
- `vba-form-ui-extraction`: regex extraction of `.form.txt`/`.report.txt`. Emits only `property` nodes per control + one `references` edge to sibling `.cls`. MUST NOT emit `function`/`sub`/`module` nodes.

### Modified Capabilities
- None.

## Approach

Mirror `DfmExtractor`. Two classes mechanically separated so the `.cls`/`.form.txt` split is enforced by construction. Synthetic edges use `Edge.metadata.synthesizedBy`. No schema migration.

### Decisions baked in (explore open questions)
1. **Two extractors** — enforces the `.cls`/`.form.txt` rule by construction.
2. **`Sub New` only** as canonical class initializer.
3. **`Implements IFoo`** → `EdgeKind.implements` (already upstream).
4. **`Dim x As Foo.Bar`** → best-effort `references` to `Foo` with `synthesizedBy: 'vba-name-resolution'`; silent when unresolved.
5. **`WithEvents m_X As Form_Foo`** → `references` to `Form_Foo` with `synthesizedBy: 'vba-withevents'`.
6. **SQL in strings** → `FROM`/`INTO`/`UPDATE <table>` table names → `references` with `synthesizedBy: 'vba-sql-table'`.

## Affected Areas

| Area | Impact |
|------|--------|
| `src/extraction/vba-extractor.ts` | New |
| `src/extraction/vba-form-extractor.ts` | New |
| `src/extraction/tree-sitter.ts` | Modified |
| `src/extraction/grammars.ts` | Modified |
| `src/types.ts` | Modified |
| `src/mcp/server-instructions.ts` | Modified |
| `__tests__/extraction-vba.test.ts` | New |
| `__tests__/extraction-vba-form.test.ts` | New |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| False-positive edges (SQL in comments) | Med | Comment-string adversarial tests |
| `Sub New` ambiguity with other `New` | Low | Word-boundary regex + `NewClase` test |
| Heuristic edges pollute `codegraph_explore` | Med | `synthesizedBy` inline; precision spot-check |
| Non-VBA regression | Low | Same checkout both builds; counts MUST match (obs #14702) |

## Rollback Plan

Revert the merge commit. All edits additive; no other language's behavior changes. Drop `.codegraph/` index in VBA projects post-revert.

## Dependencies

Already shipped upstream (obs #14694): `Edge.provenance`, `Edge.metadata`, `EdgeKind.implements`. No external deps.

## Success Criteria

- [ ] `npm test` green; covers `Sub`/`Function`/`Property`, `Sub New`, `Implements`, `Dim As`, `WithEvents`, SQL tables, plus "no function/sub/module from `.form.txt`" assertion.
- [ ] `codegraph_explore` on a VBA fixture returns expected symbols; heuristic edges render `synthesizedBy` inline.
- [ ] Non-VBA regression: codegraph checkout identical non-VBA counts vs baseline (obs #14702).
- [ ] `server-instructions.ts` names the `.cls`/`.form.txt` split.