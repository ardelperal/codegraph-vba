/**
 * Error-policy sweep (issue #259, task E2 of `docs/vba-error-handling-plan.md`,
 * extended by issue #260 / task E3).
 *
 * Records how each procedure handles errors as an `errorPolicy` object on the
 * `function` node the procedures sweep already emitted. It adds **zero node
 * kinds and zero edge kinds**, and emits neither a node nor an edge nor an
 * unresolved reference — the node and edge totals of an indexed project are
 * byte-for-byte what they were before this classifier existed. That is the
 * deliberate conclusion of §4 of the plan: 96.5% of the corpus's 3,911 line
 * labels are the same label (`errores`) doing the same job, so a label node
 * would carry exactly one bit — "a handler exists" — which is a boolean field.
 *
 * Issue #260 keeps that invariant while making the handler region visible on
 * the graph the extractor already builds: every edge and every unresolved
 * reference emitted from inside a handler gets `metadata.inErrorHandler:
 * true`, so `Riesgo.Guardar -> MsgBox` (runs only on failure) stops looking
 * identical to `Riesgo.Guardar -> Escribir` (runs always). It is a field
 * added to existing rows — no row is created, dropped or reordered.
 *
 * ## What it answers
 *
 *   - "Which risky procedures have no error handling at all?" (§3.1) —
 *     `protection: 'none'`, which cannot be grepped: an ABSENCE scoped to a
 *     procedure body has no text to match.
 *   - "Which `On Error Resume Next` scopes are never closed?" (§3.2) —
 *     `resumeNextOpen: true`.
 *   - "Where does an error surface to the user?" (§3.3) — `behavior`.
 *   - "What does this procedure do ONLY when things go wrong?" (§3.4) — the
 *     edges and references carrying `metadata.inErrorHandler`.
 *
 * ## The precision gate
 *
 * **A label defined but never targeted by an `On Error GoTo` is control flow,
 * not a handler.** 137 of the corpus's labels are exactly that (`siguiente`,
 * `salir`, `fin`, `Teardown`). Treating one as a handler produces a
 * confidently wrong answer to "does this procedure handle errors", which is
 * the worst failure mode available here — so a label only opens a handler
 * region when its name is in `targets`, the set of recorded `On Error GoTo`
 * destinations.
 *
 * ## Boundaries — one detector, three consumers
 *
 * The procedure body's START comes from `ctx.functionNodeByStartLine`, which
 * the procedures pre-walk fully populates before this classifier sees its
 * first line. No second `PROC_RE` dispatch: "is this line a procedure
 * declaration" is already answered by the node the extractor emitted, exactly
 * as `scripts/vba-coverage-probe.mjs` derives its own procedure starts.
 *
 * The body's END is `PROCEDURE_END_RE` from `./constants` — the SAME constant
 * `enums-consts.ts`'s `proc-end` rule dispatches and `call-sweep.ts` tests.
 * The plan is explicit about this ("consume that signal, do not add a second
 * end detector"; "Reuse it; do not write a second one"), and it is what makes
 * the colon-separated single-line procedure form work here for free (#208).
 * Unlike its two sibling consumers this one tests the MASKED line, so a
 * string literal containing `": End Sub"` cannot close a procedure early
 * (#209 discipline, applied consistently across this whole file).
 *
 * ## Agreement with the probe
 *
 * `scripts/vba-coverage-probe.mjs` measured the corpus census this task is
 * accepted against, so the two classifiers must agree. `LINE_LABEL_RE` and
 * `NOT_A_LABEL` below are therefore the probe's, not the simplified
 * `^\s*(\w+):\s*$` sketched in the issue: the probe accepts the
 * trailing-statement form (`errores: MsgBox "x"`), guards the
 * `Foo bar:=1` named-argument form, guards statement keywords that can head a
 * colon-terminated line, and accepts accented label names. Keeping the strict
 * sketch instead would have classified every trailing-statement handler as a
 * dangling target and disagreed with the census on both `handlerStartLine`
 * and `danglingTarget`. The probe cannot import this module (it is a
 * standalone ESM script that must also run against a shipped `dist/`), so the
 * two copies are kept in sync by hand — change one, change the other.
 */
import { PROCEDURE_END_RE } from './constants';
import { CompiledErrorChannel, defaultErrorChannel } from './error-channel';
import { maskStringContent } from './text-utils';
import {
  VbaClassifier,
  VbaExtractorContext,
  VbaErrorHandlerSignals,
  VbaErrorPolicy,
  VbaErrorPolicyState,
} from './context';
import { defineRule, runRules, VbaExtractionRule } from './rules';

/**
 * `On Error GoTo <label|0|-1>`. The numeric alternative is captured on
 * purpose so the rule can tell a handler from a reset instead of the label
 * rule silently swallowing `0` (which `\w+` would happily match).
 */
const ON_ERROR_GOTO_RE =
  /\bOn\s+Error\s+GoTo\s+(-?\d+|\p{L}[\p{L}\p{N}_]*)/iu;

/** Global twin of {@link ON_ERROR_GOTO_RE} — every site on one line. */
const ON_ERROR_GOTO_G = new RegExp(ON_ERROR_GOTO_RE.source, 'giu');

/** `On Error Resume Next` — opens a suppression scope. */
const ON_ERROR_RESUME_NEXT_RE = /\bOn\s+Error\s+Resume\s+Next\b/i;

/** Global twin of {@link ON_ERROR_RESUME_NEXT_RE}. */
const ON_ERROR_RESUME_NEXT_G = new RegExp(ON_ERROR_RESUME_NEXT_RE.source, 'gi');

/**
 * `On Error GoTo 0` (disable the handler) and `On Error GoTo -1` (clear the
 * current error). `-1` is valid VBA and appears ZERO times in this corpus;
 * the plan asks for it to be treated as a reset anyway, with a fixture, so
 * the zero-occurrence case stays pinned rather than rotting untested.
 *
 * The trailing `\b` is what keeps `On Error GoTo 01` — a numeric VBA line
 * number, therefore a handler label — out of this rule.
 */
const ON_ERROR_RESET_RE = /\bOn\s+Error\s+GoTo\s+(0|-1)\b/i;

/** Global twin of {@link ON_ERROR_RESET_RE}. */
const ON_ERROR_RESET_G = new RegExp(ON_ERROR_RESET_RE.source, 'gi');

/**
 * A line-label definition. VBA allows a statement to follow on the same line
 * (`errores: MsgBox "x"`), so the trailing-code form is accepted.
 *
 * `(?!=)` keeps the named-argument form (`Foo bar:=1`) out; {@link NOT_A_LABEL}
 * keeps `Case 1:` / `Do:` / `Else:` out. Kept identical to the probe's
 * `LINE_LABEL_RE`.
 */
const LINE_LABEL_RE = /^\s*(\p{L}[\p{L}\p{N}_]*)\s*:(?!=)/u;

/**
 * Statement keywords that can head a colon-terminated line but name no label.
 * Kept identical to the probe's `NOT_A_LABEL`.
 */
const NOT_A_LABEL = new Set([
  'case', 'do', 'else', 'end', 'exit', 'for', 'if', 'loop', 'next', 'select',
  'then', 'wend', 'while', 'with', 'rem', 'call', 'set', 'let', 'dim',
]);

/* ─────────────────── handler behaviour (issue #260, task E3) ───────────── */

/**
 * The error-propagation channel this corpus actually uses: a module-level or
 * object field the failing procedure writes and the caller reads. §2.3 of the
 * plan measures 3,602 handlers touching it against 16 that re-raise — VBA's
 * own error mechanism unwinds one frame, the MESSAGE travels through one of
 * these variables.
 *
 * Names only, never substrings, so `ErrorCount` cannot match `Error`.
 *
 * Issue #261 (task E4) moved this list — and the write matcher built from it —
 * into `./error-channel`, so the two consumers can no longer fork: this
 * classifier asks "is this statement a channel WRITE?", `module-vars.ts` asks
 * "is this variable the channel?", and both now read one compiled object that
 * `codegraph.json` → `vba.errorChannel` extends. Re-exported from here because
 * this module was its home and the probe-agreement suite imports it.
 */
export { DEFAULT_ERROR_CHANNEL_NAMES } from './error-channel';

/**
 * Calls that make an error visible to a human. `MsgBox` is the form-code
 * shape, `Debug.Print` the developer-only one; §2.3 counts both as "display",
 * so this list — identical to the probe's `DEFAULT_DISPLAY_CALLS` — does too.
 */
export const DEFAULT_DISPLAY_CALLS: readonly string[] = ['MsgBox', 'Debug.Print'];

function escapeRe(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** `MsgBox` / `Debug.Print`, anywhere in the statement. The probe's twin. */
const DISPLAY_CALL_RES: readonly RegExp[] = DEFAULT_DISPLAY_CALLS.map(
  (call) =>
    new RegExp(`\\b${call.split('.').map(escapeRe).join('\\s*\\.\\s*')}\\b`, 'i'),
);

/** `Err.Raise` — the frame re-throws. The probe's twin. */
const ERR_RAISE_RE = /\bErr\s*\.\s*Raise\b/i;

/**
 * Split a masked line into statements and strip any leading `If … Then` /
 * `ElseIf … Then` / `Else` guard, so the house shape
 * `If Err.Number <> 1000 Then p_Error = "…"` presents `p_Error = "…"` to the
 * channel matcher as a statement in its own right.
 *
 * This is also what makes the corpus's fifth most common handler shape —
 * `DoCmd.Hourglass False` cleanup first, the guarded channel write second —
 * classify as `channel`: the signals are collected per statement over the
 * WHOLE region, so a leading call cannot shadow a later write.
 *
 * Kept identical to the probe's `statementsOf`.
 */
function statementsOf(maskedLine: string): string[] {
  return maskedLine.split(':').map((part) => {
    const guard = /^\s*(?:(?:If|ElseIf)\b.*?\bThen\b|Else\b)/i.exec(part);
    return guard ? part.slice(guard[0].length) : part;
  });
}

/**
 * The three signals one masked line (or line fragment) fires, if any.
 *
 * `errorChannel` is the compiled, config-aware channel (issue #261). A project
 * that adds `lastFailure` to `codegraph.json` → `vba.errorChannel` gets the
 * handlers writing it classified `channel` HERE too, not merely its references
 * flagged in `module-vars.ts` — one list, two consumers, no drift. With no
 * config the list is the four defaults, so every pre-#261 classification, and
 * therefore the probe agreement, is unchanged.
 */
function signalsOf(
  maskedText: string,
  errorChannel: CompiledErrorChannel,
): VbaErrorHandlerSignals | null {
  const channel = statementsOf(maskedText).some((statement) =>
    errorChannel.writeRes.some((re) => re.test(statement)),
  );
  const display = DISPLAY_CALL_RES.some((re) => re.test(maskedText));
  const reraise = ERR_RAISE_RE.test(maskedText);
  if (!channel && !display && !reraise) return null;
  return { channel, display, reraise };
}

/**
 * Record what one masked line contributes to the open body's handler
 * behaviour: the whole line under {@link VbaErrorPolicyState.signalsByLine},
 * and — when the line defines a label — its trailing statement separately
 * under `labelRestSignals`, so a region opened by `errores: MsgBox "x"` reads
 * the `MsgBox` without reading the label.
 *
 * No-op outside a procedure body.
 */
function recordHandlerSignals(
  state: VbaErrorPolicyState | null,
  maskedLine: string,
  lineNum: number,
  errorChannel: CompiledErrorChannel,
): void {
  if (!state) return;

  const whole = signalsOf(maskedLine, errorChannel);
  if (whole) state.signalsByLine.set(lineNum, whole);

  const label = LINE_LABEL_RE.exec(maskedLine);
  if (!label) return;
  const name = label[1] ?? '';
  if (!name || NOT_A_LABEL.has(name.toLowerCase())) return;
  const rest = signalsOf(maskedLine.slice(label[0].length), errorChannel);
  if (rest) state.labelRestSignals.set(lineNum, rest);
}

/** A fresh accumulator for a procedure body opening at `startLine`. */
export function newErrorPolicyState(
  startLine: number,
  edgeMark: number,
  refMark: number,
): VbaErrorPolicyState {
  return {
    startLine,
    protection: 'none',
    targets: new Map(),
    definedLabels: new Map(),
    handlerCount: 0,
    lastScopeEvent: null,
    openedTarget: null,
    handlerStartLine: null,
    edgeMark,
    refMark,
    signalsByLine: new Map(),
    labelRestSignals: new Map(),
  };
}

/**
 * Record an `On Error Resume Next` (`opens`) or an `On Error GoTo 0|-1`
 * (`!opens`) at a source position, keeping only the LAST one.
 *
 * Position-ordered rather than counted because the three `On Error` rules are
 * independent table entries: on a colon-separated line carrying both forms,
 * whichever rule the dispatcher happens to reach first must not decide the
 * outcome. `resumeNextOpen` is then "the last scope event opened a scope",
 * which is exactly the probe's `openResumeNextScopes > 0`.
 */
function noteScopeEvent(
  state: VbaErrorPolicyState,
  line: number,
  column: number,
  opens: boolean,
): void {
  const previous = state.lastScopeEvent;
  if (
    previous !== null &&
    (previous.line > line ||
      (previous.line === line && previous.column > column))
  ) {
    return;
  }
  state.lastScopeEvent = { line, column, opens };
}

/**
 * Issue #259: the declarative rule table for the error-policy concern. All
 * four rules are `scan: 'masked'` and `requires: 'inside-procedure'`.
 *
 * `masked` is the #209 discipline: `s = "On Error GoTo errores"` is a string
 * assignment, and a classifier that reads it as a handler would report a
 * protected procedure that is not protected. `inside-procedure` is what keeps
 * module-level noise out — there is no error policy without a procedure to
 * attach it to.
 *
 * The three `On Error` rules re-scan their own line globally instead of
 * acting on the dispatcher's single match, so a colon-separated single-line
 * procedure carrying two `On Error GoTo` statements counts two.
 */
export const RULES: readonly VbaExtractionRule<unknown>[] = [
  defineRule({
    id: 'on-error-label',
    description:
      'Match `On Error GoTo <label>` where the target is not `0` / `-1`; record the label as a handler target, bump `handlerCount`, and set `protection` to `handler`.',
    pattern: ON_ERROR_GOTO_RE,
    requires: 'inside-procedure',
    scan: 'masked',
    emit: (_m, ctx, line) => {
      const state = ctx.vbaErrorPolicy;
      if (!state) return null;
      let handlers = 0;
      ON_ERROR_GOTO_G.lastIndex = 0;
      let match: RegExpExecArray | null;
      while ((match = ON_ERROR_GOTO_G.exec(line)) !== null) {
        const target = match[1] ?? '';
        // `0` and `-1` — and ONLY those two — are resets, owned by the
        // `on-error-reset` rule. Everything else is a handler label, including
        // a numeric VBA line number (`On Error GoTo 100`), which is exactly
        // how the probe classifies it. This corpus has 38 `GoTo 0`, no `-1`
        // and no line-number target, so today the two agree vacuously here;
        // from now on they agree by construction.
        if (target === '0' || target === '-1') continue;
        const key = target.toLowerCase();
        if (!state.targets.has(key)) state.targets.set(key, target);
        state.handlerCount += 1;
        state.protection = 'handler';
        handlers += 1;
      }
      return handlers > 0 ? { handlers } : null;
    },
    count: (result) => (result as { handlers: number }).handlers,
  }),
  defineRule({
    id: 'on-error-resume-next',
    description:
      'Match `On Error Resume Next`; open a suppression scope and set `protection` to `resume-next` unless an `On Error GoTo <label>` already made it `handler`.',
    pattern: ON_ERROR_RESUME_NEXT_RE,
    requires: 'inside-procedure',
    scan: 'masked',
    emit: (_m, ctx, line, lineNum) => {
      const state = ctx.vbaErrorPolicy;
      if (!state) return null;
      let scopes = 0;
      ON_ERROR_RESUME_NEXT_G.lastIndex = 0;
      let match: RegExpExecArray | null;
      while ((match = ON_ERROR_RESUME_NEXT_G.exec(line)) !== null) {
        noteScopeEvent(state, lineNum, match.index, true);
        scopes += 1;
      }
      if (scopes === 0) return null;
      // A real handler outranks blanket suppression: a procedure that does
      // both is reported as protected, matching the probe's precedence.
      if (state.protection !== 'handler') state.protection = 'resume-next';
      return { scopes };
    },
    count: (result) => (result as { scopes: number }).scopes,
  }),
  defineRule({
    id: 'on-error-reset',
    description:
      'Match `On Error GoTo 0` or `On Error GoTo -1`; close the open `On Error Resume Next` scope. `-1` clears the current error rather than disabling the handler, but for this field the two are the same event.',
    pattern: ON_ERROR_RESET_RE,
    requires: 'inside-procedure',
    scan: 'masked',
    emit: (_m, ctx, line, lineNum) => {
      const state = ctx.vbaErrorPolicy;
      if (!state) return null;
      let resets = 0;
      ON_ERROR_RESET_G.lastIndex = 0;
      let match: RegExpExecArray | null;
      while ((match = ON_ERROR_RESET_G.exec(line)) !== null) {
        noteScopeEvent(state, lineNum, match.index, false);
        resets += 1;
      }
      return resets > 0 ? { resets } : null;
    },
    count: (result) => (result as { resets: number }).resets,
  }),
  defineRule({
    id: 'line-label',
    description:
      'Match a `<Label>:` definition; record it, and — ONLY when an `On Error GoTo` already named it — open the handler region on the next line. A label nobody targets is control flow, not a handler.',
    pattern: LINE_LABEL_RE,
    requires: 'inside-procedure',
    scan: 'masked',
    emit: (match, ctx, _line, lineNum) => {
      const state = ctx.vbaErrorPolicy;
      if (!state) return null;
      const name = match[1] ?? '';
      if (!name || NOT_A_LABEL.has(name.toLowerCase())) return null;
      const key = name.toLowerCase();
      if (!state.definedLabels.has(key)) state.definedLabels.set(key, lineNum);
      // The precision gate. `targets` holds only `On Error GoTo` destinations,
      // so a `siguiente:` / `salir:` / `Teardown:` label leaves the policy
      // untouched and the procedure keeps whatever protection it earned.
      if (state.handlerStartLine === null && state.targets.has(key)) {
        state.handlerStartLine = lineNum + 1;
        state.openedTarget = key;
      }
      return { label: name };
    },
  }),
];

/**
 * Fold the open accumulator into the procedure's `metadata.errorPolicy` and
 * clear it. `endLine` is the procedure's terminating `End` line — the handler
 * region's closing boundary.
 *
 * Called from three places: the `End Sub` boundary (the normal path), a new
 * procedure declaration while a body is somehow still open (malformed VBA
 * with no `End`), and end-of-file. All three converge on the same shape, so a
 * truncated module still records a policy instead of silently dropping it.
 */
export function closeErrorPolicy(
  ctx: VbaExtractorContext,
  endLine: number,
): void {
  const state = ctx.vbaErrorPolicy;
  if (!state) return;
  ctx.vbaErrorPolicy = null;

  // The handler region is "the EARLIEST label definition in this body that an
  // `On Error GoTo` names" — the probe's `handlerLabelDefs[0]`, since it
  // collects definitions in line order. The `line-label` rule already opened
  // the region for the common case (label defined after the `On Error GoTo`
  // that targets it), but two orderings need the whole body to be known before
  // they resolve: a label defined BEFORE its `On Error GoTo`, and a body whose
  // first targeted definition is only targeted later. Re-resolving here is
  // what keeps this classifier and the probe from disagreeing on
  // `handlerStartLine`; the rule's incremental open is confirmed by it.
  let best: { key: string; line: number } | null = null;
  for (const [key, line] of state.definedLabels) {
    if (!state.targets.has(key)) continue;
    if (best === null || line < best.line) best = { key, line };
  }
  if (best) {
    state.handlerStartLine = best.line + 1;
    state.openedTarget = best.key;
  }

  const firstTarget = state.targets.keys().next().value ?? null;
  const labelKey = state.openedTarget ?? firstTarget;
  const danglingKey =
    [...state.targets.keys()].find((key) => !state.definedLabels.has(key)) ??
    null;

  const policy: VbaErrorPolicy = {
    protection: state.protection,
    handlerLabel: labelKey ? (state.targets.get(labelKey) ?? null) : null,
    handlerStartLine: state.handlerStartLine,
    handlerEndLine: state.handlerStartLine === null ? null : endLine,
    // Issue #260 (task E3): derived below, from the signals collected over
    // the handler region this function has just finished resolving.
    behavior: classifyBehavior(state, best, endLine),
    handlerCount: state.handlerCount,
    resumeNextOpen: state.lastScopeEvent?.opens ?? false,
    danglingTarget: danglingKey
      ? (state.targets.get(danglingKey) ?? null)
      : null,
  };

  // Issue #260: THE single stamping point. Every edge and every unresolved
  // reference this procedure body emitted is already in `ctx`, so one pass
  // over its own slice of the two accumulators flags all of them — instead
  // of six emitters each remembering a flag, which is six places to forget
  // it. New emitters inherit the behaviour for free.
  markErrorHandlerRegion(ctx, state, policy);

  const node = ctx.functionNodeByStartLine.get(state.startLine);
  if (!node) return;
  if (!node.metadata) node.metadata = {};
  node.metadata.errorPolicy = policy;
}

/**
 * Issue #260: fold the per-line signals collected over the whole body down to
 * the handler region, and name the result.
 *
 * Only a `handler`-protected procedure gets a `behavior` at all: an
 * `On Error Resume Next` body has no handler to describe, and neither does an
 * unprotected one, so both stay `null`. A `handler` procedure whose region is
 * empty — or whose target label is never defined — is `unknown`, not `null`:
 * "it handles errors and does nothing recognisable with them" is a different
 * fact from "it does not handle errors". Both rules are the probe's.
 *
 * `mixed` is what MORE THAN ONE signal means. The three are collected
 * independently precisely so evaluation order cannot silently pick a winner —
 * §2.3 of the plan reports 921 handlers that both record the error AND show
 * it, and the pre-probe classifier hid every one of them behind whichever
 * test it happened to run first.
 */
function classifyBehavior(
  state: VbaErrorPolicyState,
  region: { key: string; line: number } | null,
  endLine: number,
): VbaErrorPolicy['behavior'] {
  if (state.protection !== 'handler') return null;

  const signals: VbaErrorHandlerSignals = {
    channel: false,
    display: false,
    reraise: false,
  };
  if (region) {
    const merge = (part: VbaErrorHandlerSignals | undefined): void => {
      if (!part) return;
      signals.channel ||= part.channel;
      signals.display ||= part.display;
      signals.reraise ||= part.reraise;
    };
    // The label's own line contributes ONLY its trailing statement — the
    // `errores: MsgBox "x"` one-liner form — because the label itself is not
    // handler code. Every later line contributes in full.
    merge(state.labelRestSignals.get(region.line));
    for (const [line, part] of state.signalsByLine) {
      if (line > region.line && line <= endLine) merge(part);
    }
  }

  const hit = (['channel', 'display', 'reraise'] as const).filter(
    (name) => signals[name],
  );
  if (hit.length === 0) return 'unknown';
  if (hit.length === 1) return hit[0]!;
  return 'mixed';
}

/**
 * Issue #260: stamp `metadata.inErrorHandler: true` on every edge and every
 * unresolved reference this procedure emitted from inside its handler region.
 *
 * Scoped two ways, and both matter:
 *
 *   - by POSITION in the accumulators (`edgeMark` / `refMark`, taken when the
 *     body opened), so the scan is per-procedure rather than per-file and a
 *     module with N procedures stays linear rather than quadratic;
 *   - by LINE against `[handlerStartLine, handlerEndLine]`, the exact region
 *     published on the node. A row with no line — a `contains` edge, say — is
 *     structural, not emitted from a statement, and is never flagged.
 *
 * The flag is only ever set to `true`; its ABSENCE is the "not in a handler"
 * encoding, so this adds a key to a minority of rows instead of a `false` to
 * every one of them. Nothing here creates, drops or reorders a row.
 */
function markErrorHandlerRegion(
  ctx: VbaExtractorContext,
  state: VbaErrorPolicyState,
  policy: VbaErrorPolicy,
): void {
  const start = policy.handlerStartLine;
  const end = policy.handlerEndLine;
  if (start === null || end === null) return;

  for (let i = state.edgeMark; i < ctx.edges.length; i++) {
    const edge = ctx.edges[i];
    if (!edge || edge.line === undefined) continue;
    if (edge.line < start || edge.line > end) continue;
    if (!edge.metadata) edge.metadata = {};
    edge.metadata.inErrorHandler = true;
  }

  for (let i = state.refMark; i < ctx.unresolvedReferences.length; i++) {
    const ref = ctx.unresolvedReferences[i];
    if (!ref) continue;
    if (ref.line < start || ref.line > end) continue;
    if (!ref.metadata) ref.metadata = {};
    ref.metadata.inErrorHandler = true;
  }
}

/**
 * Issue #259: factory for the error-policy classifier.
 *
 * `count` stays at 0 for the life of the classifier, deliberately. The
 * orchestrator derives `hasAnySymbols` from the sum of every classifier's
 * `count`, and this one emits no symbols at all — it only annotates nodes the
 * procedures sweep already created. Letting it count would make a file whose
 * only content is an `On Error` line look like a file with symbols.
 */
export function createErrorPolicyClassifier(): VbaClassifier {
  let lastLineNum = 0;
  return {
    name: 'errorPolicy',
    count: 0,
    classifyLine(line, i, ctx) {
      const lineNum = i + 1;
      lastLineNum = lineNum;
      const maskedLine = maskStringContent(line);

      // Procedure START: the pre-walk already emitted a function node for
      // this line, so no second `PROC_RE` dispatch is needed.
      if (ctx.functionNodeByStartLine.has(lineNum)) {
        closeErrorPolicy(ctx, lineNum - 1);
        ctx.vbaErrorPolicy = newErrorPolicyState(
          lineNum,
          ctx.edges.length,
          ctx.unresolvedReferences.length,
        );
      }

      runRules(RULES, ctx, line, maskedLine, lineNum, {
        'inside-procedure': ctx.vbaErrorPolicy !== null,
      });

      // Issue #260: collect this line's handler signals for the OPEN body.
      // Collected for the whole body, not just the region: the region's first
      // line is only certain once the body has been fully read, so
      // `closeErrorPolicy` does the filtering. Sparse — a line with no signal
      // stores nothing, which is the overwhelming majority of them.
      recordHandlerSignals(
        ctx.vbaErrorPolicy,
        maskedLine,
        lineNum,
        ctx.errorChannel ?? defaultErrorChannel(),
      );

      // Procedure END. Deliberately independent of the start branch above: a
      // colon-separated single-line procedure carries both markers (#208).
      if (ctx.vbaErrorPolicy !== null && PROCEDURE_END_RE.test(maskedLine)) {
        closeErrorPolicy(ctx, lineNum);
      }
    },
    finalize(ctx) {
      closeErrorPolicy(ctx, lastLineNum);
    },
  };
}
