#!/usr/bin/env node
/**
 * npm postinstall hook for @aroman22/codegraph-vba.
 *
 * Runs after every `npm install -g @aroman22/codegraph-vba@...` (any version),
 * regardless of which binary triggered the upgrade. Sweeps orphan staging dirs
 * (`.<pkg>-<HASH>`) from the global npm root so a subsequent `npm install`
 * doesn't trip EBUSY/EPERM on a locked `node.exe` left behind by a previous
 * interrupted upgrade.
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
 * Find the npm global root: walk up from `startDir` until we reach a
 * directory whose direct children include `.codegraph-vba-*` entries. That
 * directory is the npm global root (the staging dirs and the final install's
 * `node_modules/` all live under it).
 *
 * We can't just look for `node_modules/` because the script may run from
 * INSIDE a staging dir (e.g. `<root>/.codegraph-vba-stage123/node_modules/@aroman22/codegraph-vba/scripts/`),
 * where walking up finds the staging's own `node_modules/` first — that's the
 * red herring. The "has `.codegraph-vba-*` children" check reliably
 * identifies the real global root.
 */
function resolveNpmGlobalRoot(startDir) {
  let dir = path.resolve(startDir);
  const root = path.parse(dir).root;
  for (;;) {
    // Skip `.codegraph-vba-<HASH>` staging dirs themselves — we don't want
    // to check inside them (the staging dir's own name matches, but it's
    // not the global root).
    if (path.basename(dir).startsWith(PKG_PREFIX)) {
      dir = path.dirname(dir);
      continue;
    }
    // Found a directory whose children include `.codegraph-vba-*` entries?
    // That's the npm global root (it contains the staging dirs as siblings
    // of the final install's `node_modules/`).
    try {
      const entries = fs.readdirSync(dir);
      if (entries.some((e) => e.startsWith(PKG_PREFIX))) return dir;
    } catch {
      /* not readable; keep walking */
    }
    const parent = path.dirname(dir);
    if (parent === dir) return null;
    dir = parent;
    if (dir === root) return null;
  }
}

/**
 * Walk up from `startDir` and return the first ancestor whose basename
 * starts with `.codegraph-vba-`. Used to identify the staging dir we're
 * part of, so the cleanup can skip it (deleting the dir we're running
 * from under npm would race with npm's atomic move).
 */
function findCurrentStagingAncestor(startDir) {
  let dir = path.resolve(startDir);
  const root = path.parse(dir).root;
  for (;;) {
    if (path.basename(dir).startsWith(PKG_PREFIX)) return dir;
    const parent = path.dirname(dir);
    if (parent === dir) return null;
    dir = parent;
    if (dir === root) return null;
  }
}

/**
 * Best-effort: remove every `.codegraph-vba-<HASH>` dir under `globalRoot`,
 * skipping the staging that this script is currently running from. Returns
 * the count actually removed.
 */
function cleanupOrphanStagings(globalRoot, currentStaging) {
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
    if (currentStaging && orphan === currentStaging) continue;
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
  const globalRoot = resolveNpmGlobalRoot(__dirname);
  if (!globalRoot) {
    // No npm global root reachable — unusual layout. Skip silently.
    return;
  }
  const currentStaging = findCurrentStagingAncestor(__dirname);
  const cleaned = cleanupOrphanStagings(globalRoot, currentStaging);
  if (cleaned > 0) {
    process.stderr.write(
      `[codegraph-vba postinstall] removed ${cleaned} orphan staging dir${cleaned === 1 ? '' : 's'}\n`
    );
  }
}

main();
