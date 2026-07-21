import { existsSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';

describe('test suite hygiene', () => {
  it('does not keep documentation-prose test files', () => {
    const obsoleteTests = [
      'extraction-vba-extraction-perf-doc.test.ts',
      'spike-vbnet-as-vba.test.ts',
    ];

    expect(
      obsoleteTests.filter((file) => existsSync(join(__dirname, file))),
    ).toEqual([]);
  });
});
