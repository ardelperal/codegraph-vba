import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as os from 'node:os';
import { execFileSync } from 'node:child_process';
import {
  detectInstallMethod,
  deriveInstallDir,
  parseSemver,
  compareVersions,
  isUpdateAvailable,
  normalizeVersion,
  stripV,
  parseLatestTagFromLocation,
  reindexAdvisory,
  runUpgrade,
  verifyResolvedVersion,
  buildWindowsUpgradeScript,
  findOrphanStagings,
  cleanupOrphanStagings,
  readInstalledPackageVersion,
  resolveNpmGlobalRoot,
  NPM_PACKAGE,
  type InstallMethod,
  type UpgradeDeps,
} from '../src/upgrade';
import { EXTRACTION_VERSION } from '../src/extraction/extraction-version';
import { CodeGraph } from '../src';

// ---------------------------------------------------------------------------
// detectInstallMethod — structural detection from the running file's path
// ---------------------------------------------------------------------------

describe('detectInstallMethod', () => {
  // A bundle exists if a vendored node + launcher sit next to lib/.
  function bundleExists(present: Set<string>) {
    return (p: string) => present.has(p.replace(/\\/g, '/'));
  }

  it('detects a unix bundle and derives the install dir from the versions/ layout', () => {
    const root = '/home/u/.codegraph/versions/v0.9.9';
    const filename = `${root}/lib/dist/bin/codegraph.js`;
    const present = new Set([`${root}/node`, `${root}/bin/codegraph`, '/home/u/.codegraph']);
    const m = detectInstallMethod({
      filename,
      platform: 'linux',
      cwd: '/home/u/project',
      exists: bundleExists(present),
    });
    expect(m).toEqual({
      kind: 'bundle',
      os: 'unix',
      bundleRoot: root,
      installDir: '/home/u/.codegraph',
    });
  });

  it('detects a windows bundle and derives the install dir from current\\', () => {
    const root = 'C:/Users/u/AppData/Local/codegraph/current';
    const filename = `${root}/lib/dist/bin/codegraph.js`;
    const present = new Set([`${root}/node.exe`, `${root}/bin/codegraph.cmd`]);
    const m = detectInstallMethod({
      filename,
      platform: 'win32',
      cwd: 'C:/Users/u/project',
      exists: bundleExists(present),
    }) as Extract<InstallMethod, { kind: 'bundle' }>;
    expect(m.kind).toBe('bundle');
    expect(m.os).toBe('windows');
    // win32 path math emits backslashes; compare separator-independently.
    expect(m.installDir?.replace(/\\/g, '/')).toBe('C:/Users/u/AppData/Local/codegraph');
  });

  it('detects a global npm install', () => {
    const filename = '/usr/local/lib/node_modules/@colbymchenry/codegraph/dist/bin/codegraph.js';
    const m = detectInstallMethod({
      filename,
      platform: 'linux',
      cwd: '/home/u/project',
      exists: () => false,
    });
    expect(m).toEqual({ kind: 'npm', scope: 'global' });
  });

  it('detects a local (project) npm install as local', () => {
    const cwd = '/home/u/project';
    const filename = `${cwd}/node_modules/@colbymchenry/codegraph/dist/bin/codegraph.js`;
    const m = detectInstallMethod({ filename, platform: 'linux', cwd, exists: () => false });
    expect(m).toEqual({ kind: 'npm', scope: 'local' });
  });

  it('detects an npx run from the _npx cache', () => {
    const filename = '/home/u/.npm/_npx/abc123/node_modules/@colbymchenry/codegraph/dist/bin/codegraph.js';
    const m = detectInstallMethod({ filename, platform: 'linux', cwd: '/home/u', exists: () => false });
    expect(m).toEqual({ kind: 'npx' });
  });

  // The npm thin-installer's per-platform package IS a complete bundle
  // (vendored node + bin/ launcher) sitting inside node_modules. The layout
  // sniff must not win over the node_modules path check, or `upgrade` curls
  // install.sh into ~/.codegraph — a second install that loses the PATH race
  // to npm's shim, so `codegraph -v` stays on the old version forever.
  it('detects the npm thin-installer platform package as npm, not bundle', () => {
    const root = '/usr/local/lib/node_modules/@colbymchenry/codegraph/node_modules/@colbymchenry/codegraph-linux-x64';
    const filename = `${root}/lib/dist/bin/codegraph.js`;
    const present = new Set([`${root}/node`, `${root}/bin/codegraph`]);
    const m = detectInstallMethod({
      filename,
      platform: 'linux',
      cwd: '/home/u/project',
      exists: bundleExists(present),
    });
    expect(m).toEqual({ kind: 'npm', scope: 'global' });
  });

  it('detects a project-local thin-installer platform package as npm local', () => {
    const cwd = '/home/u/project';
    const root = `${cwd}/node_modules/@colbymchenry/codegraph/node_modules/@colbymchenry/codegraph-darwin-arm64`;
    const filename = `${root}/lib/dist/bin/codegraph.js`;
    const present = new Set([`${root}/node`, `${root}/bin/codegraph`]);
    const m = detectInstallMethod({ filename, platform: 'darwin', cwd, exists: bundleExists(present) });
    expect(m).toEqual({ kind: 'npm', scope: 'local' });
  });

  it('still detects an npx run when the cached platform package has the bundle layout', () => {
    const root = '/home/u/.npm/_npx/abc123/node_modules/@colbymchenry/codegraph/node_modules/@colbymchenry/codegraph-linux-x64';
    const filename = `${root}/lib/dist/bin/codegraph.js`;
    const present = new Set([`${root}/node`, `${root}/bin/codegraph`]);
    const m = detectInstallMethod({ filename, platform: 'linux', cwd: '/home/u', exists: bundleExists(present) });
    expect(m).toEqual({ kind: 'npx' });
  });

  it('detects a source checkout via sibling package.json + .git', () => {
    const repo = '/home/u/dev/codegraph';
    const filename = `${repo}/dist/bin/codegraph.js`;
    const present = new Set([`${repo}/package.json`, `${repo}/.git`]);
    const m = detectInstallMethod({
      filename,
      platform: 'darwin',
      cwd: repo,
      exists: bundleExists(present),
    });
    expect(m).toEqual({ kind: 'source', root: repo });
  });

  it('returns unknown for an unrecognized layout', () => {
    const m = detectInstallMethod({
      filename: '/opt/weird/place/codegraph.js',
      platform: 'linux',
      cwd: '/tmp',
      exists: () => false,
    });
    expect(m.kind).toBe('unknown');
  });
});

describe('deriveInstallDir', () => {
  it('unix: returns the dir above versions/', () => {
    expect(deriveInstallDir('/a/b/.codegraph/versions/v1.2.3', 'unix', () => true)).toBe('/a/b/.codegraph');
  });
  it('unix: null when not under versions/', () => {
    expect(deriveInstallDir('/a/b/somewhere', 'unix', () => true)).toBeNull();
  });
  it('windows: returns the parent of current\\', () => {
    expect(deriveInstallDir('C:/x/codegraph/current', 'windows', () => true)?.replace(/\\/g, '/')).toBe('C:/x/codegraph');
  });
  it('windows: null when basename is not current', () => {
    expect(deriveInstallDir('C:/x/codegraph/v1', 'windows', () => true)).toBeNull();
  });
});

// ---------------------------------------------------------------------------
// version helpers
// ---------------------------------------------------------------------------

describe('version helpers', () => {
  it('parseSemver handles v-prefix and prerelease', () => {
    expect(parseSemver('v1.2.3')).toEqual({ major: 1, minor: 2, patch: 3, pre: null });
    expect(parseSemver('1.2.3-rc.1')).toEqual({ major: 1, minor: 2, patch: 3, pre: 'rc.1' });
    expect(parseSemver('not-a-version')).toBeNull();
  });

  it('compareVersions orders correctly incl. prerelease < release', () => {
    expect(compareVersions('1.0.1', '1.0.0')).toBeGreaterThan(0);
    expect(compareVersions('1.0.0', '1.1.0')).toBeLessThan(0);
    expect(compareVersions('v2.0.0', '2.0.0')).toBe(0);
    expect(compareVersions('1.0.0-rc.1', '1.0.0')).toBeLessThan(0);
  });

  it('isUpdateAvailable compares, and falls back to string-inequality for unparseable', () => {
    expect(isUpdateAvailable('0.9.8', '0.9.9')).toBe(true);
    expect(isUpdateAvailable('0.9.9', '0.9.9')).toBe(false);
    expect(isUpdateAvailable('0.9.9', '0.9.8')).toBe(false);
    // dev sentinel can't parse → any difference means "update available"
    expect(isUpdateAvailable('0.0.0-unknown', '0.9.9')).toBe(true);
  });

  it('normalizeVersion / stripV round-trip', () => {
    expect(normalizeVersion('0.9.9')).toBe('v0.9.9');
    expect(normalizeVersion('v0.9.9')).toBe('v0.9.9');
    expect(stripV('v0.9.9')).toBe('0.9.9');
    expect(stripV('0.9.9')).toBe('0.9.9');
  });

  it('parseLatestTagFromLocation extracts the tag from a releases redirect', () => {
    expect(parseLatestTagFromLocation('https://github.com/colbymchenry/codegraph/releases/tag/v0.9.9')).toBe('v0.9.9');
    expect(parseLatestTagFromLocation('https://github.com/o/r/releases/tag/v1.2.3?foo=bar')).toBe('v1.2.3');
    expect(parseLatestTagFromLocation(undefined)).toBeNull();
    expect(parseLatestTagFromLocation('https://github.com/o/r/releases')).toBeNull();
  });

  it('reindexAdvisory mentions the refresh commands', () => {
    const a = reindexAdvisory();
    expect(a).toContain('codegraph-vba sync');
    expect(a).toContain('codegraph-vba index -f');
  });

  it('buildWindowsUpgradeScript targets the right asset per arch and renames-not-deletes the exe', () => {
    const arm = buildWindowsUpgradeScript('C:\\cg\\current', 'v1.2.3', 'arm64');
    // Asset names carry the `codegraph-vba-` prefix (build-bundle.sh +
    // release.yml publish `codegraph-vba-<target>.zip`); the old
    // `codegraph-<target>.zip` URL 404s.
    expect(arm).toContain('releases/download/v1.2.3/codegraph-vba-win32-arm64.zip');
    // The extracted inner dir is also `codegraph-vba-<target>`.
    expect(arm).toContain('codegraph-vba-win32-arm64');
    expect(arm).toContain("$dest='C:\\cg\\current'");
    expect(arm).toContain('Rename-Item'); // never Remove-Item on the locked exe
    expect(arm).not.toMatch(/Remove-Item[^;]*\$dest'?\s*;/); // doesn't delete current\
    const x64 = buildWindowsUpgradeScript('C:\\cg\\current', 'v1.2.3', 'x64');
    expect(x64).toContain('codegraph-vba-win32-x64.zip');
  });
});

// ---------------------------------------------------------------------------
// runUpgrade orchestration — mocked side-effects
// ---------------------------------------------------------------------------

interface Calls {
  runs: Array<{ cmd: string; args: string[]; env?: NodeJS.ProcessEnv }>;
  captures: Array<{ cmd: string; args: string[] }>;
  logs: string[];
  errors: string[];
}

function makeDeps(
  overrides: Partial<UpgradeDeps> & { method: InstallMethod; currentVersion: string },
  runExit = 0
): { deps: UpgradeDeps; calls: Calls } {
  const calls: Calls = { runs: [], captures: [], logs: [], errors: [] };
  const deps: UpgradeDeps = {
    currentVersion: overrides.currentVersion,
    method: overrides.method,
    resolveLatest: overrides.resolveLatest ?? (async () => 'v0.9.9'),
    run: (cmd, args, env) => {
      calls.runs.push({ cmd, args, env });
      return runExit;
    },
    // Default probe: spawn fails → 'inconclusive'. Tests that exercise the
    // post-upgrade version check override this.
    capture: (cmd, args) => {
      calls.captures.push({ cmd, args });
      return overrides.capture ? overrides.capture(cmd, args) : null;
    },
    hasCommand: overrides.hasCommand ?? ((c) => c === 'curl'),
    log: (m) => calls.logs.push(m),
    warn: (m) => calls.logs.push(m),
    error: (m) => calls.errors.push(m),
    platform: overrides.platform ?? 'linux',
    // Forward the post-install verification + pre-cleanup hooks so tests
    // can drive them without touching the filesystem.
    installedPackageVersion: overrides.installedPackageVersion,
    orphanCleanupOverride: overrides.orphanCleanupOverride,
  };
  return { deps, calls };
}

/** Decode a `-EncodedCommand` base64 (UTF-16LE) payload back to its script. */
function decodeEncodedCommand(args: string[]): string {
  const i = args.indexOf('-EncodedCommand');
  if (i < 0) throw new Error('no -EncodedCommand in args');
  return Buffer.from(args[i + 1]!, 'base64').toString('utf16le');
}

describe('runUpgrade', () => {
  it('does nothing when already up to date', async () => {
    const { deps, calls } = makeDeps({ method: { kind: 'npm', scope: 'global' }, currentVersion: '0.9.9' });
    const code = await runUpgrade({}, deps);
    expect(code).toBe(0);
    expect(calls.runs).toHaveLength(0);
    expect(calls.logs.join('\n')).toMatch(/up to date/i);
  });

  it('--check reports an available update without running anything', async () => {
    const { deps, calls } = makeDeps({
      method: { kind: 'npm', scope: 'global' },
      currentVersion: '0.9.8',
    });
    const code = await runUpgrade({ check: true }, deps);
    expect(code).toBe(0);
    expect(calls.runs).toHaveLength(0);
    expect(calls.logs.join('\n')).toMatch(/update is available/i);
  });

  it('unix bundle: runs the installer via sh with the derived install dir', async () => {
    const { deps, calls } = makeDeps({
      method: { kind: 'bundle', os: 'unix', bundleRoot: '/h/.codegraph/versions/v0.9.8', installDir: '/h/.codegraph' },
      currentVersion: '0.9.8',
    });
    const code = await runUpgrade({}, deps);
    expect(code).toBe(0);
    expect(calls.runs).toHaveLength(1);
    expect(calls.runs[0].cmd).toBe('sh');
    expect(calls.runs[0].args[0]).toBe('-c');
    expect(calls.runs[0].args[1]).toContain('curl -fsSL');
    expect(calls.runs[0].args[1]).toContain('| sh');
    expect(calls.runs[0].env?.CODEGRAPH_INSTALL_DIR).toBe('/h/.codegraph');
    expect(calls.logs.join('\n')).toMatch(/codegraph-vba sync/); // re-index advisory printed
  });

  it('unix bundle: falls back to wget, and errors when neither downloader exists', async () => {
    const { deps, calls } = makeDeps({
      method: { kind: 'bundle', os: 'unix', bundleRoot: '/h/.codegraph/versions/v0.9.8', installDir: null },
      currentVersion: '0.9.8',
      hasCommand: () => false,
    });
    const code = await runUpgrade({}, deps);
    expect(code).toBe(1);
    expect(calls.runs).toHaveLength(0);
    expect(calls.errors.join('\n')).toMatch(/curl nor wget/i);
  });

  it('windows bundle: runs a synchronous in-place (rename + extract) powershell upgrade', async () => {
    const { deps, calls } = makeDeps({
      method: { kind: 'bundle', os: 'windows', bundleRoot: 'C:/x/codegraph/current', installDir: 'C:/x/codegraph' },
      currentVersion: '0.9.8',
      platform: 'win32',
    });
    const code = await runUpgrade({}, deps);
    expect(code).toBe(0);
    expect(calls.runs).toHaveLength(1);
    expect(calls.runs[0].cmd).toBe('powershell.exe');
    const decoded = decodeEncodedCommand(calls.runs[0].args);
    // Downloads the right asset, renames the locked exe aside, copies over current\.
    expect(decoded).toContain('releases/download/v0.9.9/codegraph-vba-win32-');
    expect(decoded).toContain('Rename-Item');
    expect(decoded).toContain('node.exe.old-');
    expect(decoded).toContain('Copy-Item');
  });

  it('windows bundle: a non-zero installer exit is a failure', async () => {
    const { deps, calls } = makeDeps(
      {
        method: { kind: 'bundle', os: 'windows', bundleRoot: 'C:/x/codegraph/current', installDir: 'C:/x/codegraph' },
        currentVersion: '0.9.8',
        platform: 'win32',
      },
      1
    );
    const code = await runUpgrade({}, deps);
    expect(code).toBe(1);
    expect(calls.errors.join('\n')).toMatch(/exited with code/i);
  });

  it('npm global: shells out to npm install -g @pkg@latest', async () => {
    const { deps, calls } = makeDeps({
      method: { kind: 'npm', scope: 'global' },
      currentVersion: '0.9.8',
    });
    const code = await runUpgrade({}, deps);
    expect(code).toBe(0);
    expect(calls.runs[0].cmd).toBe('npm');
    expect(calls.runs[0].args).toEqual(['install', '-g', '--prefer-online', `${NPM_PACKAGE}@latest`]);
  });

  it('npm on win32 routes through cmd.exe (a direct npm.cmd spawn EINVALs on modern Node)', async () => {
    const { deps, calls } = makeDeps({
      method: { kind: 'npm', scope: 'global' },
      currentVersion: '0.9.8',
      platform: 'win32',
    });
    await runUpgrade({}, deps);
    expect(calls.runs[0].cmd).toBe('cmd.exe');
    expect(calls.runs[0].args.slice(0, 3)).toEqual(['/d', '/s', '/c']);
    expect(calls.runs[0].args[3]).toBe(`npm install -g --prefer-online ${NPM_PACKAGE}@latest`);
  });

  it('npm: a pinned version is passed through as @<version>', async () => {
    const { deps, calls } = makeDeps({
      method: { kind: 'npm', scope: 'global' },
      currentVersion: '0.9.9',
    });
    await runUpgrade({ version: '0.9.8' }, deps);
    // npm spec carries no leading "v". `--prefer-online` bypasses the local
    // npm cache so a stale 1.6.3 tarball can't shadow the freshly-published 1.7.0.
    expect(calls.runs[0].args).toEqual(['install', '-g', '--prefer-online', `${NPM_PACKAGE}@0.9.8`]);
  });

  it('npm: surfaces a non-zero exit as failure', async () => {
    const { deps, calls } = makeDeps(
      { method: { kind: 'npm', scope: 'global' }, currentVersion: '0.9.8' },
      1
    );
    const code = await runUpgrade({}, deps);
    expect(code).toBe(1);
    expect(calls.errors.join('\n')).toMatch(/npm exited/i);
  });

  it('npx: nothing to upgrade', async () => {
    const { deps, calls } = makeDeps({ method: { kind: 'npx' }, currentVersion: '0.9.8' });
    const code = await runUpgrade({}, deps);
    expect(code).toBe(0);
    expect(calls.runs).toHaveLength(0);
    expect(calls.logs.join('\n')).toMatch(/nothing to upgrade/i);
  });

  it('source: runs git pull, then installs and builds, returning 0 on success', async () => {
    // Use a real tmp dir because production chdir's into `method.root` —
    // a fake path like `/dev/codegraph` would throw on Windows and on any
    // platform where the directory doesn't exist.
    const root = fs.mkdtempSync(path.join(os.tmpdir(), 'cg-upgrade-source-'));
    try {
      const { deps, calls } = makeDeps({
        method: { kind: 'source', root },
        currentVersion: '0.9.8',
      });
      const code = await runUpgrade({}, deps);
      expect(code).toBe(0);
      // git pull + package-manager install + package-manager run build.
      expect(calls.runs.length).toBeGreaterThan(0);
      expect(calls.logs.join('\n')).toMatch(/git pull/);
      expect(calls.runs.some((r) => r.cmd === 'git' && r.args[0] === 'pull')).toBe(true);
    } finally {
      fs.rmSync(root, { recursive: true, force: true });
    }
  });

  it('source: triggers upgrade when versions match but git hashes differ', async () => {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), 'cg-upgrade-source-hash-'));
    try {
      execFileSync('git', ['init', '-q'], { cwd: root });
      execFileSync('git', ['config', 'user.name', 'test'], { cwd: root });
      execFileSync('git', ['config', 'user.email', 'test@test.com'], { cwd: root });
      execFileSync('git', ['config', 'commit.gpgsign', 'false'], { cwd: root });
      fs.writeFileSync(path.join(root, 'dummy'), '1');
      execFileSync('git', ['add', '.'], { cwd: root });
      execFileSync('git', ['commit', '-m', 'commit1'], { cwd: root });
      
      const localHash = execFileSync('git', ['rev-parse', 'HEAD'], { encoding: 'utf-8', cwd: root }).trim();
      
      fs.writeFileSync(path.join(root, 'dummy'), '2');
      execFileSync('git', ['commit', '-a', '-m', 'commit2'], { cwd: root });
      const targetHash = execFileSync('git', ['rev-parse', 'HEAD'], { encoding: 'utf-8', cwd: root }).trim();
      
      execFileSync('git', ['reset', '--hard', localHash], { cwd: root });
      execFileSync('git', ['branch', 'origin/main', targetHash], { cwd: root });
      
      const { deps, calls } = makeDeps({
        method: { kind: 'source', root },
        currentVersion: '1.2.0',
      });
      deps.resolveLatest = async () => '1.2.0';

      const code = await runUpgrade({}, deps);
      expect(code).toBe(0);
      expect(calls.runs.some((r) => r.cmd === 'git' && r.args[0] === 'pull')).toBe(true);
    } finally {
      fs.rmSync(root, { recursive: true, force: true });
    }
  });
});

// ---------------------------------------------------------------------------
// Post-upgrade self-heal of installed agent surfaces
// ---------------------------------------------------------------------------

describe('post-upgrade refresh of installed agent surfaces', () => {
  it('runs `codegraph install --refresh` via the NEW binary after a successful npm upgrade', async () => {
    const { deps, calls } = makeDeps({
      method: { kind: 'npm', scope: 'global' },
      currentVersion: '0.9.8',
      hasCommand: (cmd) => cmd === 'codegraph',
    });
    const code = await runUpgrade({}, deps);
    expect(code).toBe(0);
    // The refresh is spawned AFTER the binary swap, so the fresh install
    // (with the current templates) does the writing — not this process.
    const last = calls.runs[calls.runs.length - 1];
    expect(last?.cmd).toBe('codegraph');
    expect(last?.args).toEqual(['install', '--refresh']);
  });

  it('runs the Windows .cmd launcher through cmd.exe', async () => {
    const { deps, calls } = makeDeps({
      method: { kind: 'npm', scope: 'global' },
      currentVersion: '0.9.8',
      platform: 'win32',
      hasCommand: (cmd) => cmd === 'codegraph',
    });
    const code = await runUpgrade({}, deps);
    expect(code).toBe(0);
    const last = calls.runs[calls.runs.length - 1];
    expect(last?.cmd).toBe('cmd.exe');
    expect(last?.args).toEqual(['/d', '/s', '/c', 'codegraph install --refresh']);
  });

  it('skips the refresh when `codegraph` is not resolvable on PATH', async () => {
    const { deps, calls } = makeDeps({
      method: { kind: 'npm', scope: 'global' },
      currentVersion: '0.9.8',
      // default hasCommand resolves only curl
    });
    const code = await runUpgrade({}, deps);
    expect(code).toBe(0);
    expect(calls.runs.filter((r) => r.cmd === 'codegraph')).toHaveLength(0);
  });

  it('a failing refresh warns but does not fail the upgrade', async () => {
    const { deps, calls } = makeDeps({
      method: { kind: 'npm', scope: 'global' },
      currentVersion: '0.9.8',
      hasCommand: (cmd) => cmd === 'codegraph',
    });
    deps.run = (cmd, args, env) => {
      calls.runs.push({ cmd, args, env });
      return cmd === 'codegraph' ? 1 : 0;
    };
    const code = await runUpgrade({}, deps);
    expect(code).toBe(0);
    expect(calls.logs.join('\n')).toMatch(/install --refresh/);
  });

  it('does not run after a failed upgrade', async () => {
    const { deps, calls } = makeDeps(
      {
        method: { kind: 'npm', scope: 'global' },
        currentVersion: '0.9.8',
        hasCommand: (cmd) => cmd === 'codegraph',
      },
      1
    );
    const code = await runUpgrade({}, deps);
    expect(code).toBe(1);
    expect(calls.runs.filter((r) => r.cmd === 'codegraph')).toHaveLength(0);
  });

  it('respects the CODEGRAPH_NO_INSTALL_REFRESH kill-switch', async () => {
    process.env.CODEGRAPH_NO_INSTALL_REFRESH = '1';
    try {
      const { deps, calls } = makeDeps({
        method: { kind: 'npm', scope: 'global' },
        currentVersion: '0.9.8',
        hasCommand: (cmd) => cmd === 'codegraph',
      });
      const code = await runUpgrade({}, deps);
      expect(code).toBe(0);
      expect(calls.runs.filter((r) => r.cmd === 'codegraph')).toHaveLength(0);
    } finally {
      delete process.env.CODEGRAPH_NO_INSTALL_REFRESH;
    }
  });

  it('skips the refresh when the version probe says a stale install shadows the new one', async () => {
    const { deps, calls } = makeDeps({
      method: { kind: 'npm', scope: 'global' },
      currentVersion: '0.9.8',
      hasCommand: (cmd) => cmd === 'codegraph',
      capture: () => ({ code: 0, stdout: '0.9.8\n' }), // PATH still serves the OLD version
    });
    const code = await runUpgrade({}, deps);
    expect(code).toBe(0);
    // Spawning `codegraph install --refresh` would execute the shadowed stale
    // binary — the exact staleness the refresh exists to heal.
    expect(calls.runs.filter((r) => r.cmd === 'codegraph')).toHaveLength(0);
    expect(calls.logs.join('\n')).toMatch(/run `codegraph install --refresh` once the PATH is fixed/);
  });
});

// ---------------------------------------------------------------------------
// Post-upgrade version probe — does the PATH-resolved `codegraph` serve the
// version we just installed, in THIS terminal?
// ---------------------------------------------------------------------------

describe('post-upgrade version probe', () => {
  const npmGlobal = { method: { kind: 'npm', scope: 'global' } as InstallMethod, currentVersion: '0.9.8' };

  it('match: confirms the same terminal already serves the new version', async () => {
    const { deps, calls } = makeDeps({
      ...npmGlobal,
      hasCommand: (c) => c === 'codegraph',
      capture: () => ({ code: 0, stdout: '0.9.9\n' }),
    });
    const code = await runUpgrade({}, deps);
    expect(code).toBe(0);
    expect(calls.captures).toEqual([{ cmd: 'codegraph', args: ['--version'] }]);
    const out = calls.logs.join('\n');
    expect(out).toMatch(/now reports v0\.9\.9/);
    expect(out).not.toMatch(/Open a new terminal/);
  });

  it('mismatch: warns that a shadowing install is still serving the old version', async () => {
    const { deps, calls } = makeDeps({
      ...npmGlobal,
      hasCommand: (c) => c === 'codegraph',
      capture: () => ({ code: 0, stdout: '0.9.8\n' }),
    });
    const code = await runUpgrade({}, deps);
    expect(code).toBe(0); // the upgrade itself succeeded — warn, don't fail
    const out = calls.logs.join('\n');
    expect(out).toMatch(/still reports an older version/);
    expect(out).toMatch(/shadowing/);
    expect(out).toMatch(/which -a codegraph/);
  });

  it('inconclusive: falls back to the soft new-terminal hint when codegraph is not on PATH', async () => {
    const { deps, calls } = makeDeps(npmGlobal); // hasCommand resolves only curl
    const code = await runUpgrade({}, deps);
    expect(code).toBe(0);
    expect(calls.captures).toHaveLength(0);
    expect(calls.logs.join('\n')).toMatch(/Open a new terminal/);
  });

  it('inconclusive: a failing or unparsable probe never warns about shadowing', async () => {
    const { deps, calls } = makeDeps({
      ...npmGlobal,
      hasCommand: (c) => c === 'codegraph',
      capture: () => ({ code: 0, stdout: 'something went wrong\n' }),
    });
    const code = await runUpgrade({}, deps);
    expect(code).toBe(0);
    const out = calls.logs.join('\n');
    expect(out).not.toMatch(/shadowing/);
    expect(out).toMatch(/Open a new terminal/);
  });

  it('parses the last non-empty line, so a runtime warning above the version is harmless', () => {
    const { deps } = makeDeps({
      ...npmGlobal,
      hasCommand: (c) => c === 'codegraph',
      capture: () => ({ code: 0, stdout: '(node:1) ExperimentalWarning: blah\nv0.9.9\n\n' }),
    });
    expect(verifyResolvedVersion('v0.9.9', deps)).toBe('match');
  });

  it('routes the probe through cmd.exe on Windows (.cmd launcher)', async () => {
    const { deps, calls } = makeDeps({
      ...npmGlobal,
      platform: 'win32',
      hasCommand: (c) => c === 'codegraph' || c === 'npm.cmd',
      capture: () => ({ code: 0, stdout: '0.9.9\r\n' }),
    });
    const code = await runUpgrade({}, deps);
    expect(code).toBe(0);
    expect(calls.captures).toEqual([{ cmd: 'cmd.exe', args: ['/d', '/s', '/c', 'codegraph --version'] }]);
    expect(calls.logs.join('\n')).toMatch(/now reports v0\.9\.9/);
  });

  it('skips the probe for npm-local installs — PATH serves a different copy', async () => {
    const { deps, calls } = makeDeps({
      method: { kind: 'npm', scope: 'local' },
      currentVersion: '0.9.8',
      hasCommand: (c) => c === 'codegraph',
      capture: () => ({ code: 0, stdout: '0.9.7\n' }),
    });
    const code = await runUpgrade({}, deps);
    expect(code).toBe(0);
    expect(calls.captures).toHaveLength(0);
    expect(calls.logs.join('\n')).not.toMatch(/shadowing/);
  });

  it('does not probe after a failed upgrade', async () => {
    const { deps, calls } = makeDeps(
      { ...npmGlobal, hasCommand: (c) => c === 'codegraph', capture: () => ({ code: 0, stdout: '0.9.9\n' }) },
      1
    );
    const code = await runUpgrade({}, deps);
    expect(code).toBe(1);
    expect(calls.captures).toHaveLength(0);
  });
});

// ---------------------------------------------------------------------------
// Re-index staleness — real index, real metadata stamp
// ---------------------------------------------------------------------------

describe('index extraction-version stamp / isIndexStale', () => {
  let dir: string;

  beforeEach(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'cg-upgrade-stamp-'));
  });
  afterEach(() => {
    fs.rmSync(dir, { recursive: true, force: true });
  });

  it('stamps the current extraction version on full index and is not stale', async () => {
    fs.writeFileSync(path.join(dir, 'a.ts'), 'export function hello() { return 1; }\n');
    const cg = await CodeGraph.init(dir, { index: false });
    // No index yet → not stale (nothing to refresh).
    expect(cg.isIndexStale()).toBe(false);

    await cg.indexAll();
    const info = cg.getIndexBuildInfo();
    expect(info.extractionVersion).toBe(EXTRACTION_VERSION);
    expect(typeof info.version).toBe('string');
    expect(cg.isIndexStale()).toBe(false);
    cg.destroy();
  });

  it('flags an index stamped by an older extraction version as stale', async () => {
    fs.writeFileSync(path.join(dir, 'a.ts'), 'export function hello() { return 1; }\n');
    const cg = await CodeGraph.init(dir, { index: false });
    await cg.indexAll();

    // Simulate an index built by an older engine.
    (cg as unknown as { queries: { setMetadata(k: string, v: string): void } }).queries.setMetadata(
      'indexed_with_extraction_version',
      String(EXTRACTION_VERSION - 1)
    );
    expect(cg.isIndexStale()).toBe(true);
    cg.destroy();
  });

  // Issue #189: at v1.13.0 the EXTRACTION_VERSION constant stayed at 24 even
  // though parser PRs #188 / #190 / #192 materially changed extraction output.
  // Every index stamped with the previous constant is now silently stale —
  // `isIndexStale()` returned false because `24 < 24` is false. Pin the stamp
  // to the v1.12.x constant value (24) and assert the bump surfaces it as stale.
  it('flags an index stamped with the previous extraction-version constant (24) as stale after the bump to 25 (#189)', async () => {
    fs.writeFileSync(path.join(dir, 'a.ts'), 'export function hello() { return 1; }\n');
    const cg = await CodeGraph.init(dir, { index: false });
    await cg.indexAll();

    // Simulate an index built by the previous engine (constant 24). With the
    // bumped engine (constant 25), `24 < 25` is true → stale.
    (cg as unknown as { queries: { setMetadata(k: string, v: string): void } }).queries.setMetadata(
      'indexed_with_extraction_version',
      '24'
    );
    expect(cg.getIndexBuildInfo().extractionVersion).toBe(24);
    expect(cg.isIndexStale()).toBe(true);
    // The runtime signal behind `codegraph status --json > reindexReasons`
    // must surface `'extraction-version'` so scripts can detect a missed bump
    // without parsing a prose warning.
    expect(cg.getReindexReasons()).toContain('extraction-version');
    cg.destroy();
  });

  it('returns no reindex reasons for a fresh full index', async () => {
    fs.writeFileSync(path.join(dir, 'a.ts'), 'export function hello() { return 1; }\n');
    const cg = await CodeGraph.init(dir, { index: false });
    await cg.indexAll();
    expect(cg.isIndexStale()).toBe(false);
    expect(cg.getReindexReasons()).toEqual([]);
    cg.destroy();
  });
});

// ---------------------------------------------------------------------------
// Orphan staging cleanup — npm leaves `.codegraph-vba-<HASH>` dirs behind when
// a global upgrade is interrupted (EBUSY on a locked node.exe, EPERM during
// cleanup, etc.). Those orphans block subsequent upgrades and can hold a stale
// partial copy that confuses `npm view`. The cleanup must never touch the
// live install dir.
// ---------------------------------------------------------------------------

describe('findOrphanStagings', () => {
  it('returns paths to dirs starting with .codegraph-vba- in the global root', () => {
    const files = new Set([
      '/root/node_modules/.codegraph-vba-abc123',
      '/root/node_modules/.codegraph-vba-def456',
      '/root/node_modules/.codegraph-vba-win32-x64-789', // a real npm staging hash
      '/root/node_modules/@aroman22/codegraph-vba',      // the live install — NOT an orphan
      '/root/node_modules/some-other-package',           // unrelated — NOT an orphan
    ]);
    const orphans = findOrphanStagings(
      '/root/node_modules',
      // exists: the global root itself isn't in the children-only Set, so
      // match it explicitly (real fs.existsSync returns true for any path
      // that exists, including parents).
      (p) => p === '/root/node_modules' || files.has(p),
      (p) => Array.from(files)
        .filter((f) => f.startsWith(p + '/'))
        .map((f) => f.slice(p.length + 1).split('/')[0]!)
    );
    expect(orphans).toEqual([
      '/root/node_modules/.codegraph-vba-abc123',
      '/root/node_modules/.codegraph-vba-def456',
      '/root/node_modules/.codegraph-vba-win32-x64-789',
    ]);
  });

  it('returns [] when the global root does not exist', () => {
    expect(findOrphanStagings('/nope', () => false, () => [])).toEqual([]);
  });

  it('returns [] when the global root cannot be read', () => {
    expect(findOrphanStagings('/root', () => true, () => { throw new Error('EACCES'); })).toEqual([]);
  });
});

describe('cleanupOrphanStagings', () => {
  it('removes each orphan and reports the count', () => {
    const removed: string[] = [];
    const deps = makeDeps({
      method: { kind: 'npm', scope: 'global' },
      currentVersion: '0.9.9',
    }).deps;
    const result = cleanupOrphanStagings(deps, {
      globalRoot: '/root/node_modules',
      exists: () => true,
      readdir: () => ['.codegraph-vba-abc', '.codegraph-vba-def', '@aroman22', 'other'],
      rm: (p) => { removed.push(p); return true; },
    });
    expect(result).toBe(2);
    // Both staging dirs removed; the live install and unrelated dirs untouched.
    expect(removed).toEqual([
      '/root/node_modules/.codegraph-vba-abc',
      '/root/node_modules/.codegraph-vba-def',
    ]);
  });

  it('counts only successful removals when some rm calls fail (locked file etc.)', () => {
    const deps = makeDeps({
      method: { kind: 'npm', scope: 'global' },
      currentVersion: '0.9.9',
    }).deps;
    const result = cleanupOrphanStagings(deps, {
      globalRoot: '/root/node_modules',
      exists: () => true,
      readdir: () => ['.codegraph-vba-a', '.codegraph-vba-b'],
      rm: (p) => p.endsWith('.codegraph-vba-a'), // second one fails (locked)
    });
    expect(result).toBe(1);
  });

  it('returns 0 when the global root cannot be resolved (source checkout, bundle, npx)', () => {
    // When the upgrade module runs from a non-node_modules path (test runner,
    // bundle, source), resolveNpmGlobalRoot returns null and cleanup is a no-op.
    const deps = makeDeps({
      method: { kind: 'npm', scope: 'global' },
      currentVersion: '0.9.9',
    }).deps;
    expect(resolveNpmGlobalRoot(deps)).toBeNull();
    const result = cleanupOrphanStagings(deps, {
      // globalRoot NOT specified → falls back to resolveNpmGlobalRoot(deps) → null
      exists: () => true,
      readdir: () => { throw new Error('should not be called'); },
      rm: () => { throw new Error('should not be called'); },
    });
    expect(result).toBe(0);
  });
});

// ---------------------------------------------------------------------------
// Post-install version verification — npm can report "Installed v1.7.0" while
// leaving package.json at v1.6.3 (stale cache, EPERM mid-install, PATH shadow).
// Reading the installed package.json catches that.
// ---------------------------------------------------------------------------

describe('readInstalledPackageVersion', () => {
  it('returns the version from the installed package.json', () => {
    const deps = makeDeps({
      method: { kind: 'npm', scope: 'global' },
      currentVersion: '0.9.9',
    }).deps;
    const v = readInstalledPackageVersion(deps, {
      globalRoot: '/root',
      exists: () => true,
      readFile: () => JSON.stringify({ name: NPM_PACKAGE, version: '1.7.0' }),
    });
    expect(v).toBe('1.7.0');
  });

  it('returns null when package.json does not exist', () => {
    const deps = makeDeps({
      method: { kind: 'npm', scope: 'global' },
      currentVersion: '0.9.9',
    }).deps;
    const v = readInstalledPackageVersion(deps, {
      globalRoot: '/root',
      exists: () => false,
      readFile: () => { throw new Error('should not be called'); },
    });
    expect(v).toBeNull();
  });

  it('returns null on JSON parse error', () => {
    const deps = makeDeps({
      method: { kind: 'npm', scope: 'global' },
      currentVersion: '0.9.9',
    }).deps;
    const v = readInstalledPackageVersion(deps, {
      globalRoot: '/root',
      exists: () => true,
      readFile: () => '{not json',
    });
    expect(v).toBeNull();
  });

  it('returns null when version field is missing or wrong type', () => {
    const deps = makeDeps({
      method: { kind: 'npm', scope: 'global' },
      currentVersion: '0.9.9',
    }).deps;
    expect(readInstalledPackageVersion(deps, {
      globalRoot: '/root',
      exists: () => true,
      readFile: () => JSON.stringify({ name: NPM_PACKAGE }),
    })).toBeNull();
    expect(readInstalledPackageVersion(deps, {
      globalRoot: '/root',
      exists: () => true,
      readFile: () => JSON.stringify({ name: NPM_PACKAGE, version: 17 }),
    })).toBeNull();
  });
});

// ---------------------------------------------------------------------------
// End-to-end upgrade flow with stale-install detection + tarball fallback
// ---------------------------------------------------------------------------

describe('upgradeNpm post-install verification', () => {
  it('falls back to the dist tarball when package.json version does not match the target', async () => {
    // Simulate the silent stale-install failure mode: npm install -g ran
    // (mock returns code 0), but the package.json is still the previous
    // version. The upgrade must retry from the registry tarball URL.
    const { deps, calls } = makeDeps({
      method: { kind: 'npm', scope: 'global' },
      currentVersion: '0.9.8',
      installedPackageVersion: '0.9.8',
    });
    const code = await runUpgrade({}, deps);
    expect(code).toBe(0);
    // First run: regular npm install -g.
    expect(calls.runs[0].args.slice(0, 4)).toEqual(['install', '-g', '--prefer-online', `${NPM_PACKAGE}@latest`]);
    // Second run: tarball fallback — uses the RESOLVED version (0.9.9),
    // not the user-supplied 'latest' string, because the npm registry has
    // no `latest.tgz` literal (latest is a dist-tag).
    // The unscoped name comes from stripping the full `@<scope>/` prefix
    // (`codegraph-vba` from `@aroman22/codegraph-vba`), not just the `@`.
    expect(calls.runs[1].args).toEqual([
      'install', '-g', '--force',
      'https://registry.npmjs.org/@aroman22/codegraph-vba/-/codegraph-vba-0.9.9.tgz',
    ]);
    // User saw the warning + the retry. The "expected" version is the RESOLVED
    // target (0.9.9 from the mocked resolveLatest), not the literal string
    // "latest" the user passed — npm registry has no `latest.tgz` literal.
    expect(calls.logs.join('\n')).toMatch(/installed package\.json says 0\.9\.8, expected 0\.9\.9/);
    expect(calls.logs.join('\n')).toMatch(/Retrying from the npm registry tarball/);
  });

  it('does not retry when the installed version matches the target', async () => {
    // The `currentVersion: '0.9.8'` + `installedPackageVersion: '1.7.0'` combo
    // simulates a *successful* upgrade where the on-disk package.json now
    // reflects the target. We can't pin the target version precisely here
    // (the upgrade resolves `latest` from `resolveLatest`), but the post-
    // verify is skipped because `versionSpec === 'latest'` and the test
    // setup keeps currentVersion mismatched with installedPackageVersion so
    // the upgrade proceeds; the installedPackageVersion we set is non-null
    // so the verify step doesn't fall back to the warning branch either.
    const { deps, calls } = makeDeps({
      method: { kind: 'npm', scope: 'global' },
      currentVersion: '0.9.8',
      installedPackageVersion: '0.9.9', // matches the resolved latest
    });
    const code = await runUpgrade({}, deps);
    expect(code).toBe(0);
    // Only one npm run, no tarball fallback.
    expect(calls.runs).toHaveLength(1);
    expect(calls.runs[0].args.slice(0, 4)).toEqual(['install', '-g', '--prefer-online', `${NPM_PACKAGE}@latest`]);
    // No tarball-related log lines.
    expect(calls.logs.join('\n')).not.toMatch(/Retrying from the npm registry tarball/);
  });

  it('returns 1 and surfaces a clear error when the tarball fallback also fails', async () => {
    const { deps, calls } = makeDeps({
      method: { kind: 'npm', scope: 'global' },
      currentVersion: '0.9.8',
      installedPackageVersion: '0.9.8',
    }, /* runExit */ 1); // every run returns 1 → fallback also fails
    const code = await runUpgrade({}, deps);
    expect(code).toBe(1);
    expect(calls.errors.join('\n')).toMatch(/Tarball fallback also failed/);
  });

  it('runs pre-cleanup before the npm install call', async () => {
    let cleanupCalls = 0;
    const { deps, calls } = makeDeps({
      method: { kind: 'npm', scope: 'global' },
      currentVersion: '0.9.8',
      orphanCleanupOverride: () => { cleanupCalls++; return 1; },
      installedPackageVersion: '0.9.9', // match → no tarball fallback
    });
    await runUpgrade({}, deps);
    expect(cleanupCalls).toBe(1);
    expect(calls.logs.join('\n')).toMatch(/Removed 1 orphan npm staging dir/);
  });

  it('skips pre-cleanup for npm-local installs', async () => {
    let cleanupCalls = 0;
    const { deps } = makeDeps({
      method: { kind: 'npm', scope: 'local' },
      currentVersion: '0.9.8',
      orphanCleanupOverride: () => { cleanupCalls++; return 0; },
      installedPackageVersion: '0.9.9',
    });
    await runUpgrade({}, deps);
    expect(cleanupCalls).toBe(0);
  });

  it('does not warn or retry when installedPackageVersion is null (cannot read)', async () => {
    const { deps, calls } = makeDeps({
      method: { kind: 'npm', scope: 'global' },
      currentVersion: '0.9.8',
      installedPackageVersion: null,
    });
    const code = await runUpgrade({}, deps);
    expect(code).toBe(0);
    // Only the regular npm install; no tarball fallback.
    expect(calls.runs).toHaveLength(1);
    // Soft hint, not a scary warning.
    expect(calls.logs.join('\n')).toMatch(/Could not read installed package\.json/);
    expect(calls.logs.join('\n')).not.toMatch(/Retrying from the npm registry tarball/);
  });
});
