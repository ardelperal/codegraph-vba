/**
 * Dim / typed-declaration sweep (REQ-CODE-6) and WithEvents sweep
 * (REQ-CODE-7). Emits `references` edges for non-primitive declared types,
 * populates `localVarTypeMap` so the call sweep can gate qualified calls, and
 * emits `subscribes-event` edges for `WithEvents` fields.
 */
import { Edge } from '../../types';
import { generateNodeId } from '../tree-sitter-helpers';
import { PRIMITIVE_TYPES } from './constants';
import { VbaExtractorContext, VbaClassifier } from './context';

/**
 * Check that a line is a variable declaration and NOT a Sub/Function/
 * Property/Const/WithEvents header (those have their own sweeps).
 * Fix 1+3 (Issues #1,#3): replaced the old single-match DIM_QUAL_RE /
 * DIM_UNQUAL_RE pair with a prefix-check + global scan that handles
 * `As New <Type>`, multi-variable `Dim a As Foo, b As Bar`, and all
 * visibility keywords in one pass.
 * Issue #47: now also accepts `Global` (module-level typed instance) and
 * `Static` (procedure-local retention modifier) so they emit the same
 * `references` edge and `localVarTypeMap` registration as their `Dim`
 * siblings today. The negative lookahead is unchanged: `Const` is still
 * routed to `sweepEnumsAndConsts`.
 */
const DIM_DECL_PREFIX_RE =
  /^\s*(?:Dim|Private|Public|Global|Static)\s+(?!(?:Function|Sub|Property|Const|WithEvents)\b)/i;

/**
 * Globally scan all `identifier As [New] TypePart1[.TypePart2]` on a
 * variable declaration line. Run with /g after confirming DIM_DECL_PREFIX_RE.
 *
 * Groups: (1) variable name, (2) bracketed outer type, (3) unbracketed
 * outer type, (4) bracketed inner type (if qualified), (5) unbracketed
 * inner type. The variable name is always bare (`Dim` cannot declare a
 * bracketed variable). The TYPE position accepts BOTH bracketed names
 * with spaces (e.g. `[Clase Con Espacios]`) and bare identifiers — the
 * bracketed capture wins when present. Only one of (2)/(3) and one of
 * (4)/(5) is ever populated per match.
 * `(?:New\s+)?` consumes the VBA auto-instantiation keyword so it is
 * never captured as the type name (Fix 1).
 *
 * Issue #54: extends the type alternative to accept `[Name With Spaces]`
 * so `Dim x As [Clase Con Espacios]` emits a `references` edge to
 * `Clase Con Espacios` (brackets unwrapped). The unwrap is applied in
 * the sweep loop by picking the bracketed capture group when present.
 */
const DIM_ALL_VARS_RE =
  /\b(\p{L}[\p{L}\p{N}_]*)\s+As\s+(?:New\s+)?(?:\[([^\]]+)\]|(\p{L}[\p{L}\p{N}_]*))(?:\.(?:\[([^\]]+)\]|(\p{L}[\p{L}\p{N}_]*)))?/giu;

/**
 * Bare-declared variable capture for the `Dim|Private|Public|Global|Static`
 * prefix. Captures (1) the variable name. Used to register bare `Dim x`
 * (no `As` clause) and explicit-primitive `Dim x As Long|String|...`
 * declarations into `localVarTypeMap` so the type tracking is consistent
 * across all three Dim shapes:
 *
 *   `Dim x`                → outer = 'variant' (VBA default)
 *   `Dim x As Variant`     → outer = 'variant'  (PRIMITIVE_TYPES member)
 *   `Dim x As Long`        → outer = 'long'     (PRIMITIVE_TYPES member)
 *   `Dim x As Foo`         → outer = 'foo'      (project class — non-primitive)
 *
 * Antigravity audit Task 3: the previous `DIM_ALL_VARS_RE` only matched
 * the `... As <Type>` form, so a bare `Dim x` was invisible to
 * `isLocalProjectClassVar` / `scanCallSites` and `x.Method(1)` produced
 * a dead-end `calls` edge to a stub named `x.Method` that no resolver
 * could repoint. Registering bare Dim with `outer = 'variant'` closes
 * the gate, so `scanCallSites` skips ONLY when the receiver is mapped
 * as a primitive — leaving the "undeclared receiver → stub → resolver
 * repoints" path intact for cross-module qualified calls like
 * `modUtils.Foo(1)` (`modUtils` is not in `localVarTypeMap`).
 */
const BARE_DIM_VAR_RE =
  /^\s*(?:Dim|Private|Public|Global|Static)\s+(\p{L}[\p{L}\p{N}_]*)\s*(?:,|$|\b)/iu;

/** `WithEvents m_X As Form_Foo` — Dim/Private/Public/Global/Static prefix is optional. */
const WITHEVENTS_RE =
  /^\s*(?:(?:Dim|Private|Public|Global|Static)\s+)?WithEvents\s+\p{L}[\p{L}\p{N}_]*\s+As\s+(\p{L}[\p{L}\p{N}_]*)/iu;

/**
 * Issue #83: factory for the Dim / WithEvents classifier. Stateless per-line.
 */
export function createDimsClassifier(): VbaClassifier {
  return {
    name: 'dims',
    count: 0,
    classifyLine(line, i, ctx) {
      const lineNum = i + 1;

      // Issue #151: tag every `localVarTypeMap` entry written by this
      // classifier with the procedure scope that owns it, so the call
      // sweep can clear proc-local entries at `End Sub` /
      // `End Function` / `End Property` and a `Dim x As MyClassA` in
      // `Sub First()` does not leak into `Sub Second()`.
      const dimsProcScope: 'module' | string = ctx.procStack.length > 0
        ? String(ctx.procStack[ctx.procStack.length - 1]!)
        : 'module';

      // Fix 1 + Fix 3 (Issues #1, #3): replace the old single-match
      // DIM_QUAL_RE / DIM_UNQUAL_RE pair with a global scan that handles
      // `As New <Type>`, multi-variable `Dim a As Foo, b As Bar`, and
      // qualified `Dim x As Foo.Bar` in a single pass.
      if (DIM_DECL_PREFIX_RE.test(line)) {
        DIM_ALL_VARS_RE.lastIndex = 0;
        let m: RegExpExecArray | null;
        while ((m = DIM_ALL_VARS_RE.exec(line)) !== null) {
          const varName = m[1] ?? '';
          // Issue #54: DIM_ALL_VARS_RE groups (2) and (3) are alternative
          // captures for the same outer-type position — the bracketed
          // alternative wins when present, the bare one otherwise. The
          // captured value is already unwrapped (the `[...]` is consumed
          // by the regex, group (2) holds the inner content). Same shape
          // for the inner type at groups (4)/(5).
          const outerType = m[2] ?? m[3] ?? '';
          const innerType = m[4] ?? m[5] ?? '';

          // Fix 2 (Issue #2): populate the local var type map so that
          // `sweepCallsAndSql` can gate qualified statement-form calls.
          if (varName && outerType) {
            ctx.localVarTypeMap.set(varName.toLowerCase(), {
              outer: outerType,
              qualified: !!innerType,
              procScope: dimsProcScope,
            });
          }

          if (innerType) {
            // Qualified form (`Dim x As Foo.Bar`) — emit reference to the
            // outer type `Foo` (same behaviour as the old DIM_QUAL_RE path).
            if (outerType) {
              ctx.emitReference(outerType, lineNum, 0, 'vba-name-resolution');
              this.count++;
            }
          } else {
            // Unqualified form (`Dim x As SomeType`, `Dim x As New SomeType`)
            // — emit reference only when the type is not a primitive or keyword.
            if (outerType && !PRIMITIVE_TYPES.has(outerType.toLowerCase())) {
              ctx.emitReference(outerType, lineNum, 0, 'vba-name-resolution');
              this.count++;
            }
          }
        }

        // Antigravity audit Task 3: bare `Dim x` (no `As` clause) and
        // explicit-primitive `Dim x As <primitive>` declarations must
        // still register `x` in `localVarTypeMap` so the qualified-call
        // site scan in `scanCallSites` can gate the dead-end stub that
        // no resolver could ever repoint.
        //
        // Captures the FIRST variable name only. Multi-variable bare
        // Dim (e.g. `Dim a, b, c` without an `As` clause) is rare in
        // real Dysflow fixtures and is intentionally NOT tracked here —
        // those bare variables fall back to the undeclared-receiver
        // path, which is the conservative choice. Skip if the typed-form
        // loop already populated the entry.
        const bm = BARE_DIM_VAR_RE.exec(line);
        if (bm) {
          const varName = bm[1] ?? '';
          if (varName) {
            const key = varName.toLowerCase();
            if (!ctx.localVarTypeMap.has(key)) {
              // Look for an `As <Type>` continuation on the same line so the
              // outer type matches the existing typed-form behaviour. If
              // absent, the variable is implicit `Variant` per VBA semantics.
              const asRe = /\bAs\s+(\p{L}[\p{L}\p{N}_]*)/iu;
              const asMatch = asRe.exec(line);
              const outer = asMatch ? (asMatch[1] ?? '').toLowerCase() : 'variant';
              ctx.localVarTypeMap.set(key, {
                outer,
                qualified: false,
                procScope: dimsProcScope,
              });
            }
          }
        }
      }

      // WithEvents declarations — handled by their own regex; also populate
      // the local var type map for completeness.
      const weMatch = WITHEVENTS_RE.exec(line);
      if (weMatch) {
        const formType = weMatch[1] ?? '';
        if (formType) {
          // Extract the variable name from the WithEvents line for the map.
          const weVarM = /^\s*(?:(?:Dim|Private|Public|Global|Static)\s+)?WithEvents\s+(\p{L}[\p{L}\p{N}_]*)/iu.exec(line);
          const weVarName = weVarM?.[1] ?? '';
          if (weVarName) {
            ctx.localVarTypeMap.set(weVarName.toLowerCase(), {
              outer: formType,
              qualified: false,
              withEvents: true,
              variableName: weVarName,
              procScope: dimsProcScope,
            });
          }
          ctx.emitReference(formType, lineNum, 0, 'vba-withevents');
          const targetId = generateNodeId(ctx.filePath, 'class', formType, 0);
          const subscriberEdge: Edge = {
            source: ctx.moduleOrClassNode?.id ?? '',
            target: targetId,
            kind: 'subscribes-event',
            provenance: 'heuristic',
            metadata: {
              synthesizedBy: 'vba-withevents',
              variableName: weVarName || undefined,
            },
            line: lineNum,
            column: 0,
          };
          ctx.edges.push(subscriberEdge);
          if (!ctx.moduleOrClassNode) {
            ctx.pendingModuleOrClassSource.push(subscriberEdge);
          }
          this.count++;
        }
      }
    },
  };
}

/**
 * Backward-compat wrapper (see procedures.ts). Returns the classifier's
 * `count` so the orchestrator can decide `hasAnySymbols`.
 */
export function sweepDimsAndWithEvents(ctx: VbaExtractorContext, src: string): number {
  const cls = createDimsClassifier();
  const lines = src.split('\n');
  for (let i = 0; i < lines.length; i++) {
    cls.classifyLine(lines[i] ?? '', i, ctx);
  }
  return cls.count;
}
