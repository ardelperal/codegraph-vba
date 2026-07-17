/**
 * VBA declarative rule table (Issue #153).
 *
 * The 5 per-concern classifiers under `src/extraction/vba/` used to encode
 * their extraction rules as inline `if (RE.test(line))` branches inside
 * `classifyLine`. Each branch was hard to enumerate and impossible to
 * unit-test in isolation. This test pins the new declarative shape:
 *
 *   - Every classifier exports a `RULES: VbaExtractionRule[]` constant.
 *   - Every rule has a unique `id` (no duplicate ids within one table).
 *   - Every rule has a non-empty `pattern` (RegExp or RegExp[]).
 *   - Every rule has a non-null `emit` function.
 *   - The orchestrator (`vba-extractor.ts`) imports each table and
 *     asserts it is non-empty — an accidentally-empty table must fail
 *     loudly at module load, not silently drop a whole concern.
 *
 * Adding a new test rule should require touching 1 file: the new rule's
 * classifier file (1 entry in the RULES array) + this test (1 id in the
 * declared unique-id list, optional).
 */
import { describe, it, expect } from 'vitest';
import {
  RULES as IMPLEMENTS_RULES,
  createImplementsClassifier,
} from '../src/extraction/vba/implements';
import {
  RULES as PROCEDURES_RULES,
  createProceduresClassifier,
} from '../src/extraction/vba/procedures';
import {
  RULES as DECLARATIONS_RULES,
  createEventsTypesDeclaresClassifier,
} from '../src/extraction/vba/declarations';
import {
  RULES as DIMS_RULES,
  createDimsClassifier,
} from '../src/extraction/vba/dims';
import {
  RULES as ENUM_CONSTS_RULES,
  createEnumsConstsClassifier,
} from '../src/extraction/vba/enums-consts';
import {
  RULES as CALL_SWEEP_RULES,
  createCallsAndSqlClassifier,
} from '../src/extraction/vba/call-sweep';
import {
  VBA_RULE_TABLES,
  validateVbaRuleTables,
} from '../src/extraction/vba-extractor';
import type { VbaExtractionRule } from '../src/extraction/vba/rules';

/**
 * Assert a single rule has the canonical declarative shape:
 *  - non-empty string `id`
 *  - non-empty string `description`
 *  - `pattern` is a RegExp or a non-empty RegExp[] (mixed accepted)
 *  - `emit` is a function
 *
 * Pure shape check — does NOT call `emit` (that needs a full
 * `VbaExtractorContext` and may be order-sensitive). Behavioral
 * coverage stays in `__tests__/extraction-vba.test.ts`.
 */
function assertRuleShape(rule: unknown, contextLabel: string): void {
  expect(rule, `${contextLabel} should be an object`).toBeTypeOf('object');
  expect(rule, `${contextLabel} should not be null`).not.toBeNull();

  const r = rule as Partial<VbaExtractionRule>;
  expect(typeof r.id, `${contextLabel}.id should be a string`).toBe('string');
  expect(
    r.id && r.id.length > 0,
    `${contextLabel}.id should be non-empty`,
  ).toBe(true);

  expect(
    typeof r.description,
    `${contextLabel}.description should be a string`,
  ).toBe('string');
  expect(
    r.description && r.description.length > 0,
    `${contextLabel}.description should be non-empty`,
  ).toBe(true);

  // `pattern` must be present and either a RegExp or a non-empty array of RegExp.
  const pattern = r.pattern as unknown;
  const isRegExp = pattern instanceof RegExp;
  const isRegExpArray =
    Array.isArray(pattern) &&
    pattern.length > 0 &&
    pattern.every((p) => p instanceof RegExp);
  expect(
    isRegExp || isRegExpArray,
    `${contextLabel}.pattern should be a RegExp or a non-empty RegExp[] (got ${typeof pattern})`,
  ).toBe(true);

  expect(
    typeof r.emit,
    `${contextLabel}.emit should be a function`,
  ).toBe('function');
}

/**
 * Assert that every rule id within one table is unique. Catches
 * "two rules with the same id" copy-paste regressions.
 */
function assertUniqueIds(
  rules: readonly VbaExtractionRule[],
  tableLabel: string,
): void {
  const seen = new Set<string>();
  for (const rule of rules) {
    expect(
      seen.has(rule.id),
      `${tableLabel} contains duplicate rule id "${rule.id}"`,
    ).toBe(false);
    seen.add(rule.id);
  }
}

describe('VBA declarative rule table (Issue #153)', () => {
  describe('every classifier exports a non-empty RULES array with the right shape', () => {
    const cases: ReadonlyArray<{
      label: string;
      rules: readonly VbaExtractionRule[];
    }> = [
      { label: 'implements.ts', rules: IMPLEMENTS_RULES },
      { label: 'procedures.ts', rules: PROCEDURES_RULES },
      { label: 'declarations.ts', rules: DECLARATIONS_RULES },
      { label: 'dims.ts', rules: DIMS_RULES },
      { label: 'enums-consts.ts', rules: ENUM_CONSTS_RULES },
      { label: 'call-sweep.ts', rules: CALL_SWEEP_RULES },
    ];

    for (const c of cases) {
      it(`${c.label} exports a non-empty RULES array`, () => {
        expect(Array.isArray(c.rules), `${c.label}.RULES should be an array`).toBe(
          true,
        );
        expect(
          c.rules.length,
          `${c.label}.RULES should be non-empty`,
        ).toBeGreaterThan(0);
      });

      it(`every rule in ${c.label} has the canonical shape`, () => {
        expect(c.rules.length).toBeGreaterThan(0);
        for (const rule of c.rules) {
          assertRuleShape(rule, `${c.label}.RULES entry`);
        }
      });

      it(`every rule id in ${c.label} is unique`, () => {
        expect(c.rules.length).toBeGreaterThan(0);
        assertUniqueIds(c.rules, c.label);
      });
    }
  });

  describe('orchestrator aggregates and validates the 5 rule tables', () => {
    it('VBA_RULE_TABLES exposes all 5 classifier tables keyed by concern name', () => {
      expect(VBA_RULE_TABLES).toBeDefined();
      // The aggregator must reference each of the 5 classifier files by the
      // same name the per-classifier export uses, so a test that imports
      // both sides can assert they match.
      expect(VBA_RULE_TABLES.implements).toBe(IMPLEMENTS_RULES);
      expect(VBA_RULE_TABLES.procedures).toBe(PROCEDURES_RULES);
      expect(VBA_RULE_TABLES.declarations).toBe(DECLARATIONS_RULES);
      expect(VBA_RULE_TABLES.dims).toBe(DIMS_RULES);
      expect(VBA_RULE_TABLES['enums-consts']).toBe(ENUM_CONSTS_RULES);
      expect(VBA_RULE_TABLES['call-sweep']).toBe(CALL_SWEEP_RULES);
    });

    it('validateVbaRuleTables() returns ok for the live tables', () => {
      // The orchestrator runs this on module load — calling it again here
      // is a no-op assertion that the live tables remain non-empty.
      const result = validateVbaRuleTables();
      expect(result.ok).toBe(true);
      if (!result.ok) {
        // Surface the empty tables for diagnosis if the assert ever fails.
        throw new Error(
          `validateVbaRuleTables reported empty tables: ${result.empty.join(', ')}`,
        );
      }
    });

    it('validateVbaRuleTables() flags a missing table when one is removed', () => {
      // Defensive: build a fake table set with one concern emptied, and
      // confirm the validator returns `{ ok: false, empty: [...] }`. This
      // pins the validator's contract independent of the live data.
      const empty = { ...VBA_RULE_TABLES, dims: [] };
      const result = validateVbaRuleTables(empty);
      expect(result.ok).toBe(false);
      expect(result.empty).toContain('dims');
    });
  });

  describe('factory functions still exist and still produce classifiers', () => {
    // Backward-compat guard: every `create*Classifier()` factory must
    // keep its existing signature so `vba-extractor.ts` and any external
    // consumers (tests, downstream tools) that imported the factory
    // before this refactor keep compiling.
    it('createImplementsClassifier returns a VbaClassifier', () => {
      const c = createImplementsClassifier();
      expect(c.name).toBe('implements');
      expect(typeof c.classifyLine).toBe('function');
      expect(c.count).toBe(0);
    });

    it('createProceduresClassifier returns a VbaClassifier', () => {
      const c = createProceduresClassifier();
      expect(c.name).toBe('procedures');
      expect(typeof c.classifyLine).toBe('function');
      expect(c.count).toBe(0);
    });

    it('createEventsTypesDeclaresClassifier returns a VbaClassifier', () => {
      const c = createEventsTypesDeclaresClassifier();
      expect(c.name).toBe('eventsTypesDeclares');
      expect(typeof c.classifyLine).toBe('function');
      expect(c.count).toBe(0);
    });

    it('createDimsClassifier returns a VbaClassifier', () => {
      const c = createDimsClassifier();
      expect(c.name).toBe('dims');
      expect(typeof c.classifyLine).toBe('function');
      expect(c.count).toBe(0);
    });

    it('createEnumsConstsClassifier returns a VbaClassifier', () => {
      const c = createEnumsConstsClassifier();
      expect(c.name).toBe('enumsConsts');
      expect(typeof c.classifyLine).toBe('function');
      expect(c.count).toBe(0);
    });

    it('createCallsAndSqlClassifier(lines) returns a VbaClassifier', () => {
      // The calls classifier takes the pre-split lines array — a
      // structural argument the orchestrator owns. Pass a small fixture.
      const c = createCallsAndSqlClassifier(['']);
      expect(c.name).toBe('callsAndSql');
      expect(typeof c.classifyLine).toBe('function');
      expect(c.count).toBe(0);
    });
  });
});
