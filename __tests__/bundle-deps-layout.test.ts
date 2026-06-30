/**
 * Regression test for #28: the bundled CLI must ship a FLAT, symlink-free
 * node_modules.
 *
 * The bundle (scripts/build-bundle.sh) stages production deps and then archives
 * the tree (a .zip for Windows). pnpm's default `node-linker=isolated` builds a
 * `.pnpm/` virtual store reached through symlinks, so a transitive dep like
 * `@clack/core` (pulled in by `@clack/prompts`, which the CLI's `init` imports)
 * never lands at top-level `node_modules/@clack/core`. The Windows .zip step
 * flattens the top-level `@clack/prompts` symlink into a real directory but
 * leaves `@clack/core` buried in `.pnpm/`, so Node fails with
 * `Cannot find package '@clack/core'` and `init` dies before creating
 * `.codegraph/`.
 *
 * This test extracts the real production-deps install command from
 * build-bundle.sh and runs it against the actual package.json + pnpm-lock.yaml,
 * then asserts the resulting tree is flat (transitive deps hoisted, zero
 * symlinks) — the property that survives archiving. Real script command, real
 * pnpm, real fs: red before the fix, green after.
 */
import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import { execSync } from 'node:child_process';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as os from 'node:os';

const ROOT = path.resolve(__dirname, '..');
const BUNDLE_SCRIPT = path.join(ROOT, 'scripts', 'build-bundle.sh');

/** Pull the `pnpm install --prod ...` invocation out of the bundle recipe. */
function extractProdInstallArgs(): string {
  const script = fs.readFileSync(BUNDLE_SCRIPT, 'utf8');
  // Anchor on the real command (it carries --ignore-scripts) so we don't match
  // the prose mention of `pnpm install --prod` in the recipe's comments.
  const match = script.match(/pnpm install --prod --ignore-scripts[^\n>)]*/);
  if (!match) throw new Error('could not find `pnpm install --prod` in build-bundle.sh');
  return match[0].trim();
}

/** Count symlinks anywhere under a directory tree. */
function countSymlinks(dir: string): number {
  let count = 0;
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isSymbolicLink()) {
      count++;
      continue; // don't descend into symlinks
    }
    if (entry.isDirectory()) count += countSymlinks(full);
  }
  return count;
}

describe('bundle production-deps layout (#28)', () => {
  let stage: string;
  let installed = false;

  beforeAll(() => {
    stage = fs.mkdtempSync(path.join(os.tmpdir(), 'cg-bundle-deps-'));
    fs.copyFileSync(path.join(ROOT, 'package.json'), path.join(stage, 'package.json'));
    fs.copyFileSync(path.join(ROOT, 'pnpm-lock.yaml'), path.join(stage, 'pnpm-lock.yaml'));
    // Run the exact command the bundle recipe uses. Warm pnpm store (the suite's
    // own install already populated it) → no network, fast.
    execSync(extractProdInstallArgs(), { cwd: stage, stdio: 'ignore' });
    installed = true;
  }, 180_000);

  afterAll(() => {
    if (stage && fs.existsSync(stage)) fs.rmSync(stage, { recursive: true, force: true });
  });

  it('hoists transitive deps to a real top-level directory (@clack/core resolvable)', () => {
    expect(installed).toBe(true);
    const corePkg = path.join(stage, 'node_modules', '@clack', 'core', 'package.json');
    expect(fs.existsSync(corePkg)).toBe(true);
    // It must be a real directory, not a symlink into a .pnpm store.
    expect(fs.lstatSync(path.join(stage, 'node_modules', '@clack', 'core')).isSymbolicLink()).toBe(false);
  });

  it('produces a symlink-free node_modules that survives archiving', () => {
    expect(installed).toBe(true);
    expect(countSymlinks(path.join(stage, 'node_modules'))).toBe(0);
  });
});
