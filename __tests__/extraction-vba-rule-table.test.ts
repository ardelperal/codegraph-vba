import { describe, expect, it } from 'vitest';
import {
  VBA_RULE_TABLES,
  validateVbaRuleTables,
} from '../src/extraction/vba-extractor';

describe('VBA rule-table validation', () => {
  it('accepts the live rule tables', () => {
    expect(validateVbaRuleTables()).toEqual({ ok: true, empty: [] });
  });

  it('reports an empty rule table', () => {
    const tables = { ...VBA_RULE_TABLES, dims: [] };

    expect(validateVbaRuleTables(tables)).toEqual({
      ok: false,
      empty: ['dims'],
    });
  });
});
