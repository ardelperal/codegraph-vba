import { describe, expect, it, vi } from 'vitest';

import type { VbaExtractorContext } from '../src/extraction/vba/context';
import {
  defineRule,
  runRules,
  type VbaExtractionRule,
} from '../src/extraction/vba/rules';

const ctx = {} as VbaExtractorContext;

describe('VBA rule dispatcher', () => {
  it('honours scan mode for every rule and accumulates emitted counts', () => {
    const maskedEmit = vi.fn(() => ({ count: 2 }));
    const unmaskedEmit = vi.fn(() => ({ count: 3 }));
    const rules: VbaExtractionRule[] = [
      defineRule({ id: 'masked', description: 'masked', pattern: /Call/, emit: maskedEmit, count: (r) => (r as { count: number }).count }),
      defineRule({ id: 'unmasked', description: 'unmasked', pattern: /secret/, scan: 'unmasked', emit: unmaskedEmit, count: (r) => (r as { count: number }).count }),
    ];

    expect(runRules(rules, ctx, 'Call "secret"', 'Call ""', 1, {})).toBe(5);
    expect(maskedEmit).toHaveBeenCalledWith(expect.anything(), ctx, 'Call ""', 1);
    expect(unmaskedEmit).toHaveBeenCalledWith(expect.anything(), ctx, 'Call "secret"', 1);
  });

  it('honours structural gates and declarative terminal rules without inspecting ids', () => {
    const skipped = vi.fn(() => ({}));
    const terminal = vi.fn(() => ({}));
    const afterTerminal = vi.fn(() => ({}));
    const rules: VbaExtractionRule[] = [
      defineRule({ id: 'any-id', description: 'gated', pattern: /x/, requires: 'inside-procedure', emit: skipped }),
      defineRule({ id: 'not-a-special-id', description: 'terminal', pattern: /x/, terminal: true, emit: terminal }),
      defineRule({ id: 'later', description: 'later', pattern: /x/, emit: afterTerminal }),
    ];

    expect(runRules(rules, ctx, 'x', 'x', 1, { 'inside-procedure': false })).toBe(1);
    expect(skipped).not.toHaveBeenCalled();
    expect(terminal).toHaveBeenCalledOnce();
    expect(afterTerminal).not.toHaveBeenCalled();
  });
});
