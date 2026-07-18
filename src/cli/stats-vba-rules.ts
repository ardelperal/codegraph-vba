/**
 * `codegraph stats vba-rules` — issue #168.
 *
 * Pure shape function over `VBA_RULE_TABLES` (the orchestrator's
 * aggregated view of every per-concern classifier's declarative
 * rule table, declared in `src/extraction/vba-extractor.ts`).
 *
 * The output reports **static rule counts** (the size of each
 * classifier's `RULES` array), NOT per-rule runtime emission counts.
 * The orchestrator does NOT track per-rule emission counts at the
 * rule-table level today — the only instrumentation in the VBA
 * pipeline is `vba-timing.ts` (per-stage wall-clock timings, gated
 * by `CODEGRAPH_VBA_TIMING`) and the `VbaClassifier.count`
 * per-classifier total. Neither attribute "rule X fired N times",
 * so the field name `ruleCount` keeps the static-only semantics
 * explicit. Anything resembling a runtime emission count would be
 * fabricated — keep it that way until the orchestrator actually
 * records it.
 *
 * Pattern serialization:
 *  - single `RegExp` → its `.source` (a string)
 *  - `RegExp[]` (alternatives) → `string[]` of `.source`, preserving
 *    declaration order. Matches `defineRuleAlternatives` semantics:
 *    a rule matches when ANY alternative matches, in array order.
 *
 * Concern order follows the insertion order of the `VBA_RULE_TABLES`
 * object literal — the orchestrator exports it as `{ implements,
 * procedures, declarations, dims, 'enums-consts', 'call-sweep' }`.
 * The CLI command surfaces that exact order so the JSON output is
 * stable run-to-run (no JSON-key sorting surprises) and matches the
 * dispatch order the orchestrator documents in
 * `src/extraction/vba-extractor.ts:REQUIRED_DISPATCH_TABLES`.
 */

import { VBA_RULE_TABLES } from '../extraction/vba-extractor';
import type { VbaExtractionRule } from '../extraction/vba/rules';

/**
 * One rule as serialized for `codegraph stats vba-rules`.
 *
 * `pattern` is intentionally a `string | string[]` (NOT a `RegExp`):
 * `JSON.stringify` would reduce a `RegExp` to `{}`. Serializing the
 * `.source` keeps the output copy-pasteable for debugging
 * ("what does this rule actually match against?") without losing
 * the structure that distinguishes single-RegExp rules from
 * alternatives.
 */
export interface StatsVbaRule {
  readonly id: string;
  readonly description: string;
  /** Single RegExp → its `.source`. RegExp[] → array of `.source`s in order. */
  readonly pattern: string | string[];
  /** Optional structural gate from the rule (e.g. `'class'`, `'module'`). */
  readonly requires?: string;
  /** Optional scan mode (e.g. `'masked'`, `'unmasked'`, `'both'`). */
  readonly scan?: string;
}

export interface StatsVbaConcern {
  readonly concern: string;
  readonly ruleCount: number;
  readonly rules: readonly StatsVbaRule[];
}

export interface StatsVbaRulesOutput {
  readonly concerns: readonly StatsVbaConcern[];
  readonly totalRules: number;
}

/**
 * Flatten a single rule's pattern to a `string | string[]`. A
 * single RegExp serializes to its `.source`; an array of RegExps
 * (alternatives) serializes to an array of `.source`s preserving
 * declaration order — the dispatcher matches in that same order.
 */
function serializePattern(pattern: VbaExtractionRule['pattern']): string | string[] {
  if (Array.isArray(pattern)) {
    return pattern.map((re) => re.source);
  }
  return pattern.source;
}

export function buildStatsVbaRules(): StatsVbaRulesOutput {
  const concerns: StatsVbaConcern[] = [];
  let totalRules = 0;
  for (const [concern, rules] of Object.entries(VBA_RULE_TABLES)) {
    concerns.push({
      concern,
      ruleCount: rules.length,
      rules: rules.map((rule) => {
        return {
          id: rule.id,
          description: rule.description,
          pattern: serializePattern(rule.pattern),
          ...(rule.requires !== undefined ? { requires: rule.requires } : {}),
          ...(rule.scan !== undefined ? { scan: rule.scan } : {}),
        };
      }),
    });
    totalRules += rules.length;
  }
  return { concerns, totalRules };
}
