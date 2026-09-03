import { describe, expect, it } from 'vitest';
import { EXTRACTION_VERSION } from '../src/extraction/extraction-version';

describe('VBA error-handling extraction wave version', () => {
  it('requires existing indexes to be rebuilt for the completed metadata and edge wave', () => {
    expect(EXTRACTION_VERSION).toBe(27);
  });
});
