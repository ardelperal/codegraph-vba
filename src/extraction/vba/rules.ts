/**
 * VbaExtractionRule — the declarative rule shape every VBA classifier
 * exports as a `RULES: VbaExtractionRule[]` constant. Issue #153.
 *
 * Before this refactor the per-concern sweepers under `src/extraction/vba/`
 * encoded their rules as inline `if (RE.test(line)) { ... }` branches
 * inside `classifyLine`. Each branch was hard to enumerate, impossible
 * to unit-test in isolation, and easy to silently drop when refactoring
 * the body of a sweep. The rule table here promotes the pattern →
 * emit-mapping to a typed object so:
 *
 *   1. Every rule has a stable `id` (used in error messages, logs, and
 *      as a stable handle for tools that consume the rule table).
 *   2. The pattern is explicit (single RegExp or RegExp[]) — the
 *      dispatcher knows exactly what to test, no "branched control flow
 *      around the regex".
 *   3. The emit function is the ONLY side effect for the rule — pure
 *      given (match, ctx, line, lineNum) → T | null. That makes the
 *      rule trivially testable in isolation.
 *   4. The `count?` hook lets a rule report "N symbols emitted" without
 *      sharing an accumulator with the orchestrator (each rule owns
 *      its own count semantics).
 *
 * Inter-line state (procedure stack, With stack, SQL variables) STILL
 * lives on `VbaExtractorContext` — the rule table is per-line. The
 * `classifyLine` orchestrator inside each classifier is the "shell"
 * that maintains the state machine and walks the rule table.
 */
import type { VbaExtractorContext } from './context';

/**
 * A single per-line rule the orchestrator dispatches.
 *
 * @typeParam T - The shape `emit` returns. Most rules return
 *   `void | null` because they push directly onto `ctx.nodes` /
 *   `ctx.edges` (the canonical VBA sweep idiom). Rules that
 *   synthesize a specific symbol can narrow `T` to the symbol's
 *   shape so `count?` and downstream consumers get a typed handle.
 *
 * Fields:
 *
 *  - `id`        Stable, human-readable identifier (e.g. `'implements'`,
 *                `'procedure'`, `'dim'`, `'sql-in-strings'`). The rule
 *                table's primary key — every id within one file must be
 *                unique. Used in error messages and in the test suite
 *                that pins the table's shape.
 *  - `description` One-line plain-English summary. Surfaced in tooling
 *                that introspects the table.
 *  - `pattern`   The regex (or regex alternatives) the dispatcher tests
 *                against the per-line source. Either a single `RegExp`
 *                or a non-empty `RegExp[]` (a rule is matched when ANY
 *                of the alternatives matches). The dispatcher runs
 *                `.exec()` for single regexes and iterates for arrays.
 *  - `requires?` Optional structural gate from the closed
 *                `VbaRuleRequirement` union. The orchestrator is
 *                expected to honour this — keeping the gate declarative
 *                lets the rule own its own preconditions instead of
 *                scattering `if` checks into the rule body.
 *  - `scan?`     `'masked'` = run on the string-literal-masked line
 *                (so call patterns inside `"..."` are ignored), the
 *                default. `'unmasked'` = run on the original line (so
 *                SQL patterns inside `"..."` are caught). `'both'` tries
 *                the masked line first, then the original line if needed.
 *  - `emit`      The per-match side effect. Receives the
 *                `RegExpExecArray` from the pattern, the shared
 *                `VbaExtractorContext`, the original (or masked) line,
 *                and the 1-based line number. Returns the symbol it
 *                emitted (used by `count?`) or `null` when the match
 *                didn't apply (the caller is expected to skip
 *                `count?` in that case).
 *  - `count?`    Optional counter. When the rule emits exactly one
 *                symbol per successful match the orchestrator can
 *                treat `1` as the implicit count; the hook is for
 *                rules whose emit can fan out (e.g. a multi-variable
 *                `Dim a As Foo, b As Bar` line that produces two
 *                `references` edges from one match).
 *  - `terminal?` Stop dispatch after this rule matches, without coupling
 *                dispatcher control flow to a rule id.
 *
 * The `count` parameter is typed as `unknown` to keep the
 * `VbaExtractionRule<T>` shape covariant in `T` — narrowing `T`
 * (e.g. to `{ edges: number }` for the dim-decl fan-out rule) does
 * NOT then force `count` to take that narrower type, which would
 * break the `VbaExtractionRule<unknown>[]` aggregate. Rule
 * authors cast inside their `count` body.
 */
export interface VbaExtractionRule<T = unknown> {
  readonly id: string;
  readonly description: string;
  readonly pattern: RegExp | RegExp[];
  readonly requires?: VbaRuleRequirement;
  readonly scan?: 'masked' | 'unmasked' | 'both';
  readonly terminal?: boolean;
  readonly emit: (
    match: RegExpMatchArray,
    ctx: VbaExtractorContext,
    line: string,
    lineNum: number,
  ) => T | null;
  readonly count?: (result: unknown) => number;
}

export type VbaRuleRequirement =
  | 'inside-procedure'
  | 'inside-type-block'
  | 'outside-type-block'
  | 'inside-enum-block'
  | 'outside-enum-block';

export type VbaRuleGates = Partial<Record<VbaRuleRequirement, boolean>>;

/**
 * Helper to build a `VbaExtractionRule<T>` with a single RegExp.
 * Most rules are 1-line declarations; this collapses the boilerplate
 * to one place.
 */
export function defineRule<T = unknown>(
  spec: Omit<VbaExtractionRule<T>, 'pattern'> & { pattern: RegExp },
): VbaExtractionRule<T> {
  return spec;
}

/**
 * Helper to build a `VbaExtractionRule<T>` whose `pattern` is an array
 * of alternative regexes. The dispatcher treats the rule as matched
 * when ANY alternative matches; it runs the alternatives in order and
 * uses the first match's `RegExpExecArray` for `emit`.
 */
export function defineRuleAlternatives<T = unknown>(
  spec: Omit<VbaExtractionRule<T>, 'pattern'> & { pattern: RegExp[] },
): VbaExtractionRule<T> {
  if (spec.pattern.length === 0) {
    throw new Error(
      `defineRuleAlternatives: rule "${spec.id}" has an empty pattern[]`,
    );
  }
  return spec;
}

/**
 * Run a single rule's `pattern` against `line` and return the first
 * match's `RegExpMatchArray`, or `null` when no alternative matches.
 * `pattern` is `RegExp | RegExp[]` (an array means "match any of
 * these alternatives"); this helper normalizes both into the same
 * `RegExpMatchArray | null` shape the dispatcher's `emit` expects.
 *
 * Every classifier's `classifyLine` walks the declarative `RULES`
 * table via this helper. Centralizing the RegExp/RegExp[] branching
 * here keeps the per-classifier dispatcher loops trivial and avoids
 * each classifier re-implementing the same boilerplate.
 */
export function matchRule(
  pattern: RegExp | RegExp[],
  line: string,
): RegExpMatchArray | null {
  if (Array.isArray(pattern)) {
    for (const re of pattern) {
      const m = re.exec(line);
      if (m) return m;
    }
    return null;
  }
  return pattern.exec(line);
}

export function matchRuleForScan(
  rule: VbaExtractionRule,
  line: string,
  maskedLine: string,
): { match: RegExpMatchArray; line: string } | null {
  const candidates = rule.scan === 'unmasked'
    ? [line]
    : rule.scan === 'both'
      ? [maskedLine, line]
      : [maskedLine];

  for (const candidate of candidates) {
    const match = matchRule(rule.pattern, candidate);
    if (match) return { match, line: candidate };
  }
  return null;
}

/** Dispatch a rule table with consistent scan, gate, count, and terminal semantics. */
export function runRules(
  rules: readonly VbaExtractionRule[],
  ctx: VbaExtractorContext,
  line: string,
  maskedLine: string,
  lineNum: number,
  gates: VbaRuleGates,
): number {
  let count = 0;
  for (const rule of rules) {
    if (rule.requires && gates[rule.requires] !== true) continue;
    const matched = matchRuleForScan(rule, line, maskedLine);
    if (!matched) continue;
    const result = rule.emit(matched.match, ctx, matched.line, lineNum);
    if (result !== null && result !== undefined) {
      count += rule.count ? rule.count(result) : 1;
    }
    if (rule.terminal) break;
  }
  return count;
}
