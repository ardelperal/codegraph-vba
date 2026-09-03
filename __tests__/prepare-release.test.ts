/**
 * Unit tests for `scripts/prepare-release.mjs`.
 *
 * The script reads CHANGELOG.md and package.json from `process.cwd()`,
 * so the tests run it via `node` in a temp directory after staging
 * those files. Real script, real fs — keeps the test honest about what
 * the workflow will actually do.
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { execFileSync } from 'node:child_process';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as os from 'node:os';

const SCRIPT = path.resolve(__dirname, '..', 'scripts', 'prepare-release.mjs');
const EXTRACTOR = path.resolve(__dirname, '..', 'scripts', 'extract-release-notes.mjs');

function run(cwd: string, ...args: string[]) {
  const out = execFileSync('node', [SCRIPT, ...args], { cwd, encoding: 'utf8' });
  return out.trim();
}

function setup(changelog: string, version = '1.2.3') {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'prepare-release-'));
  fs.writeFileSync(path.join(dir, 'CHANGELOG.md'), changelog);
  fs.writeFileSync(path.join(dir, 'package.json'), JSON.stringify({ name: 'x', version }));
  return dir;
}

const HEADER = `# Changelog

Some intro.

`;

/** Read back the `## [1.2.3]` block of the CHANGELOG the script just wrote. */
function result123(cwd: string) {
  const result = fs.readFileSync(path.join(cwd, 'CHANGELOG.md'), 'utf8');
  return result.split('## [1.2.3]')[1].split('## [1.2.2]')[0];
}

describe('prepare-release.mjs', () => {
  let dir: string;
  afterEach(() => {
    if (dir && fs.existsSync(dir)) fs.rmSync(dir, { recursive: true, force: true });
  });

  describe('Case A: [version] block does not yet exist', () => {
    it('renames [Unreleased] to [version] - <today> and adds a fresh empty [Unreleased]', () => {
      dir = setup(
        HEADER +
          `## [Unreleased]\n\n### Added\n- New feature foo\n- New feature bar\n\n### Fixed\n- Fixed thing\n\n## [1.2.2] - 2026-01-01\n\n### Added\n- Old entry\n`,
      );
      const out = run(dir);
      expect(out).toMatch(/renamed \[Unreleased\] to \[1\.2\.3\]/);
      const result = fs.readFileSync(path.join(dir, 'CHANGELOG.md'), 'utf8');

      // [Unreleased] is now empty and at the top.
      expect(result).toMatch(/## \[Unreleased\]\n\n\n## \[1\.2\.3\]/);
      // [1.2.3] gets a date.
      expect(result).toMatch(/## \[1\.2\.3\] - \d{4}-\d{2}-\d{2}/);
      // Promoted content lives under [1.2.3].
      const v123Section = result.split('## [1.2.3]')[1].split('## [1.2.2]')[0];
      expect(v123Section).toContain('### Added');
      expect(v123Section).toContain('- New feature foo');
      expect(v123Section).toContain('- New feature bar');
      expect(v123Section).toContain('### Fixed');
      expect(v123Section).toContain('- Fixed thing');
      // [1.2.2] is intact.
      expect(result).toContain('## [1.2.2] - 2026-01-01');
      expect(result).toContain('- Old entry');
    });
  });

  describe('Case B: [version] already exists AND [Unreleased] has content', () => {
    it('merges Unreleased sub-sections into the matching [version] sub-sections', () => {
      // The v0.9.5 scenario verbatim: sparse [0.9.5] with two Fixed
      // entries, full [Unreleased] above it with Added + more Fixed.
      dir = setup(
        HEADER +
          `## [Unreleased]\n\n### Added\n- Big feature 1\n- Big feature 2\n\n### Fixed\n- Watcher fix\n- Worktree fix\n\n## [1.2.3] - 2026-02-02\n\n### Fixed\n- Old fix A\n- Old fix B\n\n## [1.2.2] - 2026-01-01\n`,
      );
      const out = run(dir);
      expect(out).toMatch(/merged \d+ Unreleased entries/);
      const result = fs.readFileSync(path.join(dir, 'CHANGELOG.md'), 'utf8');

      // [Unreleased] is emptied.
      const unrelSection = result.split('## [Unreleased]')[1].split('## [1.2.3]')[0];
      expect(unrelSection.trim()).toBe('');

      // [1.2.3] now has BOTH the original Fixed entries AND the
      // Unreleased Fixed entries, plus the new Added sub-section.
      const v123Section = result.split('## [1.2.3]')[1].split('## [1.2.2]')[0];
      expect(v123Section).toContain('### Added');
      expect(v123Section).toContain('- Big feature 1');
      expect(v123Section).toContain('- Big feature 2');
      expect(v123Section).toContain('### Fixed');
      expect(v123Section).toContain('- Old fix A');
      expect(v123Section).toContain('- Old fix B');
      expect(v123Section).toContain('- Watcher fix');
      expect(v123Section).toContain('- Worktree fix');
      // Date on [1.2.3] is preserved (we don't re-stamp it).
      expect(result).toContain('## [1.2.3] - 2026-02-02');
    });

    it('appends sub-sections that exist only in [Unreleased] to the [version] block', () => {
      dir = setup(
        HEADER +
          `## [Unreleased]\n\n### Security\n- CVE patch\n\n## [1.2.3] - 2026-02-02\n\n### Fixed\n- Old fix\n`,
      );
      run(dir);
      const result = fs.readFileSync(path.join(dir, 'CHANGELOG.md'), 'utf8');
      const v123 = result.split('## [1.2.3]')[1];
      expect(v123).toContain('### Fixed');
      expect(v123).toContain('- Old fix');
      expect(v123).toContain('### Security');
      expect(v123).toContain('- CVE patch');
    });
  });

  describe('Case C: [Unreleased] has no entries', () => {
    it('is a no-op when [Unreleased] is empty', () => {
      dir = setup(HEADER + `## [Unreleased]\n\n## [1.2.3] - 2026-02-02\n\n### Fixed\n- thing\n`);
      const before = fs.readFileSync(path.join(dir, 'CHANGELOG.md'), 'utf8');
      const out = run(dir);
      expect(out).toMatch(/nothing to do/);
      const after = fs.readFileSync(path.join(dir, 'CHANGELOG.md'), 'utf8');
      expect(after).toBe(before);
    });

    it('is a no-op when [Unreleased] has only sub-section headings with no bullets', () => {
      dir = setup(
        HEADER + `## [Unreleased]\n\n### Added\n\n### Fixed\n\n## [1.2.3] - 2026-02-02\n`,
      );
      const before = fs.readFileSync(path.join(dir, 'CHANGELOG.md'), 'utf8');
      const out = run(dir);
      expect(out).toMatch(/nothing to do/);
      const after = fs.readFileSync(path.join(dir, 'CHANGELOG.md'), 'utf8');
      expect(after).toBe(before);
    });
  });

  describe('idempotency', () => {
    it('running twice produces the same output as running once', () => {
      dir = setup(
        HEADER +
          `## [Unreleased]\n\n### Added\n- Thing A\n\n## [1.2.2] - 2026-01-01\n\n### Added\n- Old\n`,
      );
      run(dir); // first run promotes
      const afterFirst = fs.readFileSync(path.join(dir, 'CHANGELOG.md'), 'utf8');
      const out2 = run(dir); // second run should be a no-op
      const afterSecond = fs.readFileSync(path.join(dir, 'CHANGELOG.md'), 'utf8');
      expect(out2).toMatch(/nothing to do/);
      expect(afterSecond).toBe(afterFirst);
    });
  });

  describe('version source', () => {
    it('reads the target version from package.json by default', () => {
      dir = setup(HEADER + `## [Unreleased]\n\n### Added\n- x\n`, '9.9.9');
      run(dir);
      const result = fs.readFileSync(path.join(dir, 'CHANGELOG.md'), 'utf8');
      expect(result).toContain('## [9.9.9]');
    });

    it('accepts an explicit version argument that overrides package.json', () => {
      dir = setup(HEADER + `## [Unreleased]\n\n### Added\n- x\n`, '9.9.9');
      run(dir, '5.5.5');
      const result = fs.readFileSync(path.join(dir, 'CHANGELOG.md'), 'utf8');
      expect(result).toContain('## [5.5.5]');
      expect(result).not.toContain('## [9.9.9]');
    });
  });

  describe('link reference', () => {
    it('appends a `[version]: https://...` link reference at EOF when promoting (Case A)', () => {
      dir = setup(HEADER + `## [Unreleased]\n\n### Added\n- x\n\n## [1.2.2] - 2026-01-01\n`);
      run(dir);
      const result = fs.readFileSync(path.join(dir, 'CHANGELOG.md'), 'utf8');
      expect(result).toContain(
        '[1.2.3]: https://github.com/ardelperal/codegraph-vba/releases/tag/v1.2.3',
      );
    });

    it('appends a link reference when merging into an existing [version] (Case B)', () => {
      dir = setup(
        HEADER + `## [Unreleased]\n\n### Added\n- new\n\n## [1.2.3] - 2026-02-02\n\n### Fixed\n- prior\n`,
      );
      run(dir);
      const result = fs.readFileSync(path.join(dir, 'CHANGELOG.md'), 'utf8');
      expect(result).toContain(
        '[1.2.3]: https://github.com/ardelperal/codegraph-vba/releases/tag/v1.2.3',
      );
    });

    it('does not double-add an existing link reference', () => {
      const ref = '[1.2.3]: https://github.com/ardelperal/codegraph-vba/releases/tag/v1.2.3';
      dir = setup(
        HEADER +
          `## [Unreleased]\n\n### Added\n- x\n\n## [1.2.2] - 2026-01-01\n\n${ref}\n`,
      );
      run(dir);
      const result = fs.readFileSync(path.join(dir, 'CHANGELOG.md'), 'utf8');
      const occurrences = result.split(ref).length - 1;
      expect(occurrences).toBe(1);
    });
  });

  // Regression suite for #296. Every fixture above uses the old
  // single-word Keep-a-Changelog vocabulary, which is why a `\w+`-only
  // sub-section pattern shipped green while the real CHANGELOG had
  // already moved to `### New Features` / `### Breaking Changes`.
  describe("multi-word sub-section headings (the repo's real vocabulary)", () => {
    it('carries every New Features entry into [version] when merging (Case B)', () => {
      dir = setup(
        HEADER +
          `## [1.2.3] - 2026-02-02

## [Unreleased]

### New Features
- Feature one
- Feature two

### Fixes
- A fix

## [1.2.2] - 2026-01-01
`,
      );
      const out = run(dir);
      expect(out).toContain('merged 3');

      const v123 = result123(dir);
      expect(v123).toContain('### New Features');
      expect(v123).toContain('- Feature one');
      expect(v123).toContain('- Feature two');
      expect(v123).toContain('### Fixes');
      expect(v123).toContain('- A fix');
    });

    it('reproduces the v1.16.0 shape: New Features first, above an empty pre-created block', () => {
      // The exact input that lost 16 bullets: a pre-created empty
      // [version] block, and `### New Features` as the FIRST heading in
      // [Unreleased] — so it landed in `leading`, which Case B ignored.
      dir = setup(
        HEADER +
          `## [1.2.3] - 2026-02-02

## [Unreleased]

### New Features
- Headline feature

### Changed
- A change

### Fixes
- A fix

## [1.2.2] - 2026-01-01
`,
      );
      run(dir);

      const v123 = result123(dir);
      expect(v123).toContain('- Headline feature');
      expect(v123).toContain('- A change');
      expect(v123).toContain('- A fix');
    });

    it('merges into a matching multi-word heading instead of duplicating it', () => {
      dir = setup(
        HEADER +
          `## [Unreleased]

### New Features
- Late feature

## [1.2.3] - 2026-02-02

### New Features
- Early feature

## [1.2.2] - 2026-01-01
`,
      );
      run(dir);

      const v123 = result123(dir);
      expect(v123).toContain('- Early feature');
      expect(v123).toContain('- Late feature');
      expect(v123.split('### New Features').length - 1).toBe(1);
    });

    it('still merges the single-word headings the older entries use', () => {
      // Must stay silent: the fix widens the pattern, it must not change
      // how the legacy vocabulary behaves.
      dir = setup(
        HEADER +
          `## [Unreleased]

### Added
- New thing

## [1.2.3] - 2026-02-02

### Added
- Old thing

## [1.2.2] - 2026-01-01
`,
      );
      run(dir);

      const v123 = result123(dir);
      expect(v123).toContain('- Old thing');
      expect(v123).toContain('- New thing');
      expect(v123.split('### Added').length - 1).toBe(1);
    });

    it('carries over entries written before any heading rather than dropping them', () => {
      // Defence in depth: whatever the heading vocabulary becomes, a
      // bullet with no recognised heading above it must survive.
      dir = setup(
        HEADER +
          `## [Unreleased]

- Heading-less entry

### Fixes
- A fix

## [1.2.3] - 2026-02-02

### Fixes
- Prior fix

## [1.2.2] - 2026-01-01
`,
      );
      const out = run(dir);
      expect(out).toContain('merged 2');

      const v123 = result123(dir);
      expect(v123).toContain('- Heading-less entry');
      expect(v123).toContain('- Prior fix');
      expect(v123).toContain('- A fix');
    });

    it('leaves [Unreleased] empty so the next release cannot republish these entries', () => {
      // The second half of #296: entries left behind in [Unreleased]
      // after a release get published a second time by the next one.
      dir = setup(
        HEADER +
          `## [Unreleased]

### New Features
- Shipped once

## [1.2.3] - 2026-02-02

### Fixes
- Prior
`,
      );
      run(dir);

      const result = fs.readFileSync(path.join(dir, 'CHANGELOG.md'), 'utf8');
      const unrel = result.split('## [Unreleased]')[1].split('## [1.2.3]')[0];
      expect(unrel).not.toContain('- Shipped once');
      expect(unrel.trim()).toBe('');
    });

    it('the published notes carry the New Features section end to end', () => {
      // Closes the loop through the script the workflow actually pipes
      // into `gh release create --notes-file`.
      dir = setup(
        HEADER +
          `## [1.2.3] - 2026-02-02

## [Unreleased]

### New Features
- Headline feature

### Fixes
- A fix

## [1.2.2] - 2026-01-01
`,
      );
      run(dir);

      const notes = execFileSync('node', [EXTRACTOR, '1.2.3'], {
        cwd: dir,
        encoding: 'utf8',
      });
      expect(notes).toContain('### New Features');
      expect(notes).toContain('- Headline feature');
      expect(notes).toContain('- A fix');
    });
  });

  describe('extractor integration', () => {
    it('the resulting [version] block is what extract-release-notes.mjs would surface', () => {
      // Run prepare, then extract — confirm the output contains all the
      // promoted entries.
      dir = setup(
        HEADER +
          `## [Unreleased]\n\n### Added\n- Feature A\n- Feature B\n\n### Fixed\n- Bug fix\n\n## [1.2.2] - 2026-01-01\n`,
      );
      run(dir);

      const extractor = path.resolve(__dirname, '..', 'scripts', 'extract-release-notes.mjs');
      const notes = execFileSync('node', [extractor, '1.2.3'], { cwd: dir, encoding: 'utf8' });
      expect(notes).toContain('### Added');
      expect(notes).toContain('Feature A');
      expect(notes).toContain('Feature B');
      expect(notes).toContain('### Fixed');
      expect(notes).toContain('Bug fix');
    });
  });
});
