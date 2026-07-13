#!/usr/bin/env node
/**
 * npm postinstall hook for @aroman22/codegraph-vba.
 *
 * Runs after every `npm install -g @aroman22/codegraph-vba@...` (any version),
 * regardless of which binary triggered the upgrade. Sweeps orphan staging dirs
 * (`.<pkg>-<HASH>`) from the global npm root so a subsequent `npm install`
 * doesn't trip EBUSY/EPERM on a locked `node.exe` left behind by a previous
 * interrupted upgrade. Idempotent + best-effort; never fatal to install.
 *
 * The same logic lives in `src/upgrade/index.ts:cleanupOrphanStagings` for the
 * runtime upgrade path, but a postinstall hook is the right place for THIS
 * concern because:
 *
 * 1. It runs in the staging dir BEFORE the new binary is on PATH, so it can
 *    use the new cleanup code even when the triggering upgrade came from an
 *    older binary (PR #114 only added pre-cleanup to the upgrade command;
 *    pre-fix upgrades can't clean pre-fix orphans).
 * 2. `npm install -g` runs postinstall by default (the `--ignore-scripts`
 *    flag is opt-in; users who set it explicitly are on their own).
 * 3. The script is small + dependency-free (only `fs`, `path`) so it has
 *    minimal cost when there's nothing to clean.
 *
 * Mirrors `src/upgrade/index.ts:cleanupOrphanStagings` semantics — if that
 * file changes, update this in lockstep.
 */
'use strict';

const fs = require('fs');
const path = require('path');

// Match the package name used in src/upgrade/index.ts:REPO. Any update
// to one must mirror the other.
const PKG_PREFIX = '.codegraph-vba-';

/**
 * Walk up from `startDir` until we hit a `node_modules/` ancestor; return its
 * parent (the npm global root). Mirrors resolveNpmGlobalRoot in
 * src/upgrade/index.ts. Cross-platform via path.dirname walking — works on
 * POSIX + Windows with forward/back slashes.
 */
function resolveNpmGlobalRoot(startDir) {
  let dir = path.resolve(startDir);
  const root = path.parse(dir).root;
  for (;;) {
    if (path.basename(dir) === 'node_modules') return path.dirname(dir);
    const parent = path.dirname(dir);
    if (parent === dir) return null;
    dir = parent;
    if (dir === root) {
      // Reached filesystem root without finding a node_modules ancestor —
      // probably an unusual install layout. Bail.
      return null;
    }
  }
}

/**
 * Best-effort: remove every `<PKG_PREFIX><HASH>` dir under the global
 * node_modules. Returns the count actually removed.
 */
function cleanupOrphanStagings(globalRoot) {
  if (!fs.existsSync(globalRoot)) return 0;
  let entries;
  try {
    entries = fs.readdirSync(globalRoot);
  } catch {
    return 0;
  }
  let removed = 0;
  for (const entry of entries) {
    if (!entry.startsWith(PKG_PREFIX)) continue;
    const orphan = path.join(globalRoot, entry);
    try {
      fs.rmSync(orphan, { recursive: true, force: true });
      removed++;
    } catch {
      // Locked file inside (e.g., node.exe still running) — skip; the next
      // install will retry.
    }
  }
  return removed;
}

function main() {
  // Walk up from this script's location to find the global node_modules
  // ancestor. npm guarantees the postinstall runs from the package's install
  // dir, so `<here>/node_modules/...` resolves up to the npm global root.
  const globalRoot = resolveNpmGlobalRoot(__dirname);
  if (!globalRoot) {
    // No npm global root reachable — unusual layout. Skip silently.
    return;
  }
  const cleaned = cleanupOrphanStagings(globalRoot);
  if (cleaned > 0) {
    process.stderr.write(
      `[codegraph-vba postinstall] removed ${cleaned} orphan staging dir${cleaned === 1 ? '' : 's'}\n`
    );
  }
}

main();
