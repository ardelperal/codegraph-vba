# Design: Issue 40 Qualified Calls

## Technical Approach

Implement one receiver eligibility gate inside `VbaExtractor` and use it for both qualified paren-form calls (`Receiver.Member(...)`) and qualified statement-form calls (`Receiver.Member args`). The scanner already builds `localVarTypeMap` from `Dim` and `WithEvents` declarations before call scanning, so the helper can decide whether a receiver is a project-class local variable, a declared non-project local variable, or an undeclared candidate module name without adding a global module lookup.

## Architecture Decisions

| Decision | Choice | Alternatives considered | Rationale |
|---|---|---|---|
| Centralize qualified-call eligibility | Add private `shouldProcessQualifiedCall(receiverName: string): boolean` near `isLocalProjectClassVar` / `resolveReceiverType` | Keep separate checks in `scanCallSites` and `sweepCallsAndSql` | One predicate prevents the current drift between paren and statement scanners. |
| Preserve receiver type resolution | Keep `resolveReceiverType` unchanged and call it only after eligibility passes | Fold resolution into the new gate | Eligibility and target naming are separate concerns: class locals resolve to their type, undeclared module candidates keep the raw receiver. |
| Keep heuristic node shape stable | Continue emitting `function` stub nodes and `calls` edges tagged `metadata.synthesizedBy: 'vba-name-resolution'` | Introduce a new node/edge kind for module candidates | Existing post-extraction resolution expects this stub contract. |

## Data Flow

```text
VBA source
  -> preprocessing / comment and string handling
  -> sweepDimsAndWithEvents builds localVarTypeMap
  -> scanCallSites / detectQualifiedStatementCall
  -> shouldProcessQualifiedCall(receiver)
      -> project-class local: emit resolved Type.Member
      -> declared external/primitive local: skip silently
      -> undeclared receiver: emit raw Receiver.Member candidate
```

## File Changes

| File | Action | Description |
|---|---|---|
| `src/extraction/vba-extractor.ts` | Modify | Add `shouldProcessQualifiedCall`; use it before qualified paren-form emission and in the qualified statement-form branch. Update comments that currently say undeclared statement receivers are silent. |
| `__tests__/extraction-vba.test.ts` | Modify | Add failing tests for primitive/external paren-form suppression and module-name statement-form emission. Update the existing undeclared statement receiver test to match the new module-candidate rule. |
| `openspec/specs/vba-code-extraction/spec.md` | Modify | Archive-time merge target for the delta requirement already defined by this change. |

## Interfaces / Contracts

```ts
private shouldProcessQualifiedCall(receiverName: string): boolean;
```

Contract:
- Returns `true` when `receiverName` is a declared local project-class variable.
- Returns `false` when `receiverName` is declared locally but is qualified (`DAO.Recordset`) or primitive (`Long`, `String`, etc.).
- Returns `true` when `receiverName` is not declared locally, treating it as a candidate module name.

## NodeKind / Edge Mapping

Eligible qualified calls continue to create a synthetic `Node` with `kind: 'function'`, `metadata.stub: true`, and a heuristic `calls` edge with `metadata.synthesizedBy: 'vba-name-resolution'`. No nodes or edges are emitted for declared external/primitive local receivers.

## Testing Strategy

| Layer | What to Test | Approach |
|---|---|---|
| Unit | Qualified paren-form primitive/local external receivers are silent | Add Vitest cases around `nCount.ToString()` and/or `DAO.Recordset` paren-form calls. |
| Unit | Qualified statement-form module receivers emit heuristic calls | Add a `.bas` fixture snippet with `modUtils.Foo arg` and no local `modUtils` declaration. |
| Regression | Existing class-local receiver behavior remains intact | Keep current `m_AROp.Eliminar` / `m_NCOp.Registrar` tests green. |

## Migration / Rollout

No migration required. This only changes extraction-time heuristic emission.

## Open Questions

None.
