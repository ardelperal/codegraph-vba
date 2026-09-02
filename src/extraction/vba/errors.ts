/**
 * Error-policy sweep (issue #259, task E2 of `docs/vba-error-handling-plan.md`).
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
 * ## What it answers
 *
 *   - "Which risky procedures have no error handling at all?" (§3.1) —
 *     `protection: 'none'`, which cannot be grepped: an ABSENCE scoped to a
 *     procedure body has no text to match.
 *   - "Which `On Error Resume Next` scopes are never closed?" (§3.2) —
 *     `resumeNextOpen: true`.
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
import { maskStringContent } from './text-utils';
import {
  VbaClassifier,
  VbaExtractorContext,
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

/** A fresh accumulator for a procedure body opening at `startLine`. */
export function newErrorPolicyState(startLine: number): VbaErrorPolicyState {
  return {
    startLine,
    protection: 'none',
    targets: new Map(),
    definedLabels: new Map(),
    handlerCount: 0,
    lastScopeEvent: null,
    openedTarget: null,
    handlerStartLine: null,
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
    // Task E3 derives this, where the handler body's calls are already being
    // classified. Emitting a guess here would be fabrication.
    behavior: null,
    handlerCount: state.handlerCount,
    resumeNextOpen: state.lastScopeEvent?.opens ?? false,
    danglingTarget: danglingKey
      ? (state.targets.get(danglingKey) ?? null)
      : null,
  };

  const node = ctx.functionNodeByStartLine.get(state.startLine);
  if (!node) return;
  if (!node.metadata) node.metadata = {};
  node.metadata.errorPolicy = policy;
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
        ctx.vbaErrorPolicy = newErrorPolicyState(lineNum);
      }

      runRules(RULES, ctx, line, maskedLine, lineNum, {
        'inside-procedure': ctx.vbaErrorPolicy !== null,
      });

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
