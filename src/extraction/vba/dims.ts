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
import { defineRule, matchRule, VbaExtractionRule } from './rules';

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
 * Capture the variable name from a WithEvents line (used to populate
 * `localVarTypeMap` and to stamp `metadata.variableName` on the
 * references edge — see `withevents-decl` rule's description for the
 * full rationale tied to issue #150).
 */
const WITHEVENTS_VAR_RE =
  /^\s*(?:(?:Dim|Private|Public|Global|Static)\s+)?WithEvents\s+(\p{L}[\p{L}\p{N}_]*)/iu;

/** Helper regex for the bare-Dim fallback path — extracts an `As <Type>` continuation. */
const DIM_AS_TYPE_RE = /\bAs\s+(\p{L}[\p{L}\p{N}_]*)/iu;

function isArrayDeclaration(line: string, variableName: string): boolean {
  const escaped = variableName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return new RegExp(`\\b${escaped}\\s*\\(`, 'iu').test(line);
}

/**
 * Issue #153: the declarative rule table for the Dim / WithEvents
 * concern. Two rules:
 *
 *  - `dim-decl`       — match a `Dim|Private|Public|Global|Static
 *                       <var> [As [New] <TypePart>[.<TypePart>]][, …]`
 *                       declaration; populate `localVarTypeMap` and
 *                       emit one `references` edge per non-primitive
 *                       type (or per qualified outer type).
 *  - `withevents-decl`— match a `[Dim|Private|Public|Global|Static]
 *                       WithEvents <var> As <FormType>` declaration;
 *                       populate `localVarTypeMap` with `withEvents: true`,
 *                       emit a `references` edge (synthesizedBy:
 *                       `vba-withevents`) and a `subscribes-event` edge.
 *
 * The two rules are dispatched independently and BOTH can fire on the
 * same line in principle (a `WithEvents` line does NOT match the
 * `Dim|Private|...` prefix because the negative lookahead excludes it).
 * The `dim-decl` rule's `count` hook reports the number of
 * non-primitive `references` edges it emitted, so a multi-variable
 * `Dim a As Foo, b As Bar` line adds 2 to the classifier count, and a
 * bare `Dim x` line adds 0 (the bare-Dim path is a `localVarTypeMap`
 * only side effect, no graph edges).
 *
 * No inter-line state — each rule is self-contained.
 */
export const RULES: readonly VbaExtractionRule<unknown>[] = [
  defineRule({
    id: 'dim-decl',
    description:
      'Match a `Dim|Private|Public|Global|Static <var> [As [New] <Type>][, …]` declaration; populate `localVarTypeMap` and emit one `vba-name-resolution` `references` edge per non-primitive type (or per qualified outer type).',
    pattern: DIM_DECL_PREFIX_RE,
    count: (result) => (result as { edges: number }).edges,
    emit: (_m, ctx, line, lineNum) => {
      // Re-run the prefix check inside the emit (the pattern is a
      // RegExp without /g so `.test()` is enough and idempotent). This
      // keeps the rule table's contract "pattern matches → emit fires"
      // while the emit body is the same code that used to live inline.
      let edgesEmitted = 0;
      if (!DIM_DECL_PREFIX_RE.test(line)) return null;
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
            isArray: isArrayDeclaration(line, varName),
          });
        }

        if (innerType) {
          // Qualified form (`Dim x As Foo.Bar`) — emit reference to the
          // outer type `Foo` (same behaviour as the old DIM_QUAL_RE path).
          if (outerType) {
            ctx.emitReference(outerType, lineNum, 0, 'vba-name-resolution');
            edgesEmitted++;
          }
        } else {
          // Unqualified form (`Dim x As SomeType`, `Dim x As New SomeType`)
          // — emit reference only when the type is not a primitive or keyword.
          if (outerType && !PRIMITIVE_TYPES.has(outerType.toLowerCase())) {
            ctx.emitReference(outerType, lineNum, 0, 'vba-name-resolution');
            edgesEmitted++;
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
            const asMatch = DIM_AS_TYPE_RE.exec(line);
            const outer = asMatch ? (asMatch[1] ?? '').toLowerCase() : 'variant';
            ctx.localVarTypeMap.set(key, {
              outer,
              qualified: false,
              isArray: isArrayDeclaration(line, varName),
            });
          }
        }
      }
      return { edges: edgesEmitted };
    },
  }),
  defineRule({
    id: 'withevents-decl',
    description:
      'Match a `[Dim|Private|Public|Global|Static] WithEvents <var> As <FormType>` declaration; populate `localVarTypeMap` with `withEvents: true`, emit a `vba-withevents` `references` edge AND a `subscribes-event` edge (the `variableName` metadata is preserved on both edges so the issue #150 event-handler synthesis pass can locate the handler Sub after the resolver repoints the references edge).',
    pattern: WITHEVENTS_RE,
    emit: (m, ctx, line, lineNum) => {
      const formType = m[1] ?? '';
      if (!formType) return null;
      // Extract the variable name from the WithEvents line for the map.
      const weVarM = WITHEVENTS_VAR_RE.exec(line);
      const weVarName = weVarM?.[1] ?? '';
      if (weVarName) {
        ctx.localVarTypeMap.set(weVarName.toLowerCase(), {
          outer: formType,
          qualified: false,
          withEvents: true,
          variableName: weVarName,
        });
      }
      // Stamp `variableName` onto the `references` edge's metadata so
      // the post-extraction event-handler synthesis pass (#150) can
      // locate the `m_<var>_<event>` handler Sub even after
      // `resolveVbaReferenceStubs` repoints this edge to the real
      // class node AND CASCADE-deletes the companion
      // `subscribes-event` edge along with the synthetic class stub.
      // Without this, the variable name would be lost once the
      // resolver runs, and the synthesis pass would have no way to
      // bind the `WithEvents` binding to its handler.
      ctx.emitReference(formType, lineNum, 0, 'vba-withevents', undefined, {
        ...(weVarName ? { variableName: weVarName } : {}),
      });
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
      return { name: formType };
    },
  }),
];

/**
 * Issue #83: factory for the Dim / WithEvents classifier.
 *
 * The body walks the declarative `RULES` table (Issue #153). The two
 * rules are independent: `dim-decl` handles typed declarations
 * (REJECTED if the line is a `WithEvents` because the prefix
 * negative-lookahead excludes `WithEvents`), `withevents-decl`
 * handles WithEvents. No inter-line state.
 */
export function createDimsClassifier(): VbaClassifier {
  return {
    name: 'dims',
    count: 0,
    classifyLine(line, i, ctx) {
      const lineNum = i + 1;
      for (const rule of RULES) {
        const m = matchRule(rule.pattern, line);
        if (!m) continue;
        const result = rule.emit(m, ctx, line, lineNum);
        if (result !== null && result !== undefined) {
          this.count += rule.count ? rule.count(result as never) : 1;
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
