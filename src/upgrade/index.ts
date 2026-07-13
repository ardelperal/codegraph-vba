/**
 * `codegraph upgrade`
 *
 * Self-update for the CLI, whatever way it was installed:
 *
 *   - **bundle** — the self-contained runtime+app installed by `install.sh`
 *     (Linux/macOS) or `install.ps1` (Windows). Upgrading re-runs the SAME
 *     canonical installer script (single source of truth) so the download /
 *     version-resolution / PATH logic never drifts between first-install and
 *     upgrade.
 *   - **npm** — installed via `npm i -g @colbymchenry/codegraph`. Upgrading
 *     shells out to npm.
 *   - **npx** — ephemeral; nothing to upgrade (next `npx` fetches latest).
 *   - **source** — a git checkout running its own `dist/`; `git pull` + rebuild.
 *
 * Detection is structural (see `detectInstallMethod`): a bundle carries a
 * vendored `node` binary and a `bin/codegraph` launcher next to its `lib/`, so
 * we can recognize it from the running file's path without a marker file.
 *
 * Windows wrinkle: a running `node.exe` is locked and can't be deleted, so the
 * bundle's `current\` dir can't be overwritten in place by the process doing
 * the upgrade. We therefore spawn a DETACHED helper that waits for this
 * process to exit (releasing the lock), then runs `install.ps1`. This is the
 * conventional Windows self-update dance (rustup/nvm-windows do the same).
 */

import * as fs from 'fs';
import * as path from 'path';
import * as https from 'https';
import { spawnSync } from 'child_process';

export const REPO = 'ardelperal/codegraph-vba';
export const NPM_PACKAGE = '@aroman22/codegraph-vba';
const RAW_BASE = `https://raw.githubusercontent.com/${REPO}/main`;
export const INSTALL_SH_URL = `${RAW_BASE}/install.sh`;

// ---------------------------------------------------------------------------
// Install-method detection (pure — fully unit-testable via injected probes)
// ---------------------------------------------------------------------------

export type InstallMethod =
  | { kind: 'bundle'; os: 'unix' | 'windows'; bundleRoot: string; installDir: string | null }
  | { kind: 'npm'; scope: 'global' | 'local' }
  | { kind: 'npx' }
  | { kind: 'source'; root: string }
  | { kind: 'unknown'; reason: string };

export interface DetectInput {
  /** `__filename` of the running CLI module — `<…>/dist/bin/codegraph.js`. */
  filename: string;
  platform: NodeJS.Platform;
  cwd: string;
  /** Injectable existence probe (defaults to fs.existsSync) — for tests. */
  exists?: (p: string) => boolean;
}

function toPosix(p: string): string {
  return p.replace(/\\/g, '/');
}

/**
 * Where the bundle installer keeps its install root, derived from the bundle
 * dir so an upgrade reuses a custom `CODEGRAPH_INSTALL_DIR`. Returns null when
 * the layout isn't the one the installer creates (then the installer falls
 * back to its own default).
 *
 *   unix:    <installDir>/versions/<vX.Y.Z>   (bundleRoot)  → <installDir>
 *   windows: <installDir>\current             (bundleRoot)  → <installDir>
 */
export function deriveInstallDir(
  bundleRoot: string,
  os: 'unix' | 'windows',
  exists: (p: string) => boolean
): string | null {
  // Use the TARGET platform's path semantics (not the host's), so this is
  // deterministic when reasoning about a Windows layout from a POSIX host (CI)
  // and vice-versa. In production `os` always matches the running platform.
  const P = os === 'windows' ? path.win32 : path.posix;
  if (os === 'windows') {
    if (P.basename(bundleRoot).toLowerCase() === 'current') {
      return P.dirname(bundleRoot);
    }
    return null;
  }
  // unix: bundleRoot is <installDir>/versions/<version>
  const parent = P.dirname(bundleRoot);
  if (P.basename(parent) === 'versions') {
    const installDir = P.dirname(parent);
    return exists(installDir) ? installDir : P.dirname(parent);
  }
  return null;
}

export function detectInstallMethod(input: DetectInput): InstallMethod {
  const exists = input.exists ?? fs.existsSync;
  const isWin = input.platform === 'win32';
  // Path math keyed on the TARGET platform so detection is host-independent
  // (a Windows layout resolves correctly even when unit-tested on macOS/Linux).
  const P = isWin ? path.win32 : path.posix;
  const binDir = P.dirname(input.filename); // <…>/bin

  const norm = toPosix(input.filename);

  // Path-based checks come FIRST. The npm thin-installer's per-platform
  // package (@colbymchenry/codegraph-<platform>-<arch>) is itself a complete
  // bundle — vendored node + bin/ launcher — living inside node_modules, so
  // the layout sniff below would misread every npm install as a standalone
  // bundle. `upgrade` would then curl install.sh into ~/.codegraph: a SECOND
  // install that never wins the PATH race against npm's shim, leaving
  // `codegraph -v` permanently on the old version (the #1071 shadow,
  // self-inflicted). A path under node_modules is authoritative about HOW the
  // user installed, whatever the artifact inside looks like.

  // npx cache: <…>/_npx/<hash>/node_modules/@colbymchenry/codegraph/…
  // (checked before npm — the npx cache path also contains /node_modules/).
  if (norm.includes('/_npx/')) {
    return { kind: 'npx' };
  }

  // npm install (global or local): lives under a node_modules tree.
  if (norm.includes('/node_modules/')) {
    const underCwd = norm.startsWith(toPosix(P.resolve(input.cwd)) + '/');
    return { kind: 'npm', scope: underCwd ? 'local' : 'global' };
  }

  // Bundle: <root>/lib/dist/bin/codegraph.js → <root> is up 3 from bin/.
  // A bundle has a vendored node + a launcher script as siblings of lib/.
  const bundleRoot = P.resolve(binDir, '..', '..', '..');
  const vendoredNode = P.join(bundleRoot, isWin ? 'node.exe' : 'node');
  const launcher = P.join(bundleRoot, 'bin', isWin ? 'codegraph.cmd' : 'codegraph');
  if (exists(vendoredNode) && exists(launcher)) {
    const os = isWin ? 'windows' : 'unix';
    return { kind: 'bundle', os, bundleRoot, installDir: deriveInstallDir(bundleRoot, os, exists) };
  }

  // Source checkout: running <repo>/dist/bin/codegraph.js with a sibling .git.
  const repoRoot = P.resolve(binDir, '..', '..');
  if (exists(P.join(repoRoot, 'package.json')) && exists(P.join(repoRoot, '.git'))) {
    return { kind: 'source', root: repoRoot };
  }

  return { kind: 'unknown', reason: `unrecognized install layout at ${input.filename}` };
}

// ---------------------------------------------------------------------------
// Version helpers (pure)
// ---------------------------------------------------------------------------

export interface Semver {
  major: number;
  minor: number;
  patch: number;
  pre: string | null;
}

export function parseSemver(version: string): Semver | null {
  const m = /^v?(\d+)\.(\d+)\.(\d+)(?:-([0-9A-Za-z.-]+))?/.exec(version.trim());
  if (!m) return null;
  return {
    major: parseInt(m[1]!, 10),
    minor: parseInt(m[2]!, 10),
    patch: parseInt(m[3]!, 10),
    pre: m[4] ?? null,
  };
}

/** Returns >0 if a>b, <0 if a<b, 0 if equal. Throws on unparseable input. */
export function compareVersions(a: string, b: string): number {
  const sa = parseSemver(a);
  const sb = parseSemver(b);
  if (!sa || !sb) throw new Error(`cannot compare versions: "${a}" vs "${b}"`);
  if (sa.major !== sb.major) return sa.major - sb.major;
  if (sa.minor !== sb.minor) return sa.minor - sb.minor;
  if (sa.patch !== sb.patch) return sa.patch - sb.patch;
  // A prerelease is "less than" its release (1.0.0-rc < 1.0.0).
  if (sa.pre && !sb.pre) return -1;
  if (!sa.pre && sb.pre) return 1;
  if (sa.pre && sb.pre) return sa.pre < sb.pre ? -1 : sa.pre > sb.pre ? 1 : 0;
  return 0;
}

export function isUpdateAvailable(current: string, latest: string): boolean {
  try {
    return compareVersions(latest, current) > 0;
  } catch {
    // If either is unparseable (e.g. a dev "0.0.0-unknown"), treat differing
    // strings as "update available" so the user isn't stuck.
    return normalizeVersion(current) !== normalizeVersion(latest);
  }
}

/** `0.9.9` / `v0.9.9` → `v0.9.9` (release tags are v-prefixed). */
export function normalizeVersion(v: string): string {
  const t = v.trim();
  return t.startsWith('v') ? t : `v${t}`;
}

/** Strip a leading `v`: `v0.9.9` → `0.9.9`. */
export function stripV(v: string): string {
  const t = v.trim();
  return t.startsWith('v') ? t.slice(1) : t;
}

/**
 * Parse the release tag out of the `Location` header GitHub returns for
 * `/releases/latest` → `…/releases/tag/v0.9.9`. Pure so it's unit-tested.
 */
export function parseLatestTagFromLocation(location: string | undefined): string | null {
  if (!location) return null;
  const m = /\/releases\/tag\/([^/?#]+)/.exec(location);
  return m ? decodeURIComponent(m[1]!) : null;
}

// ---------------------------------------------------------------------------
// Latest-version resolution (network)
// ---------------------------------------------------------------------------

function httpsGet(
  url: string,
  headers: Record<string, string>,
  timeoutMs: number
): Promise<{ status: number; headers: Record<string, string | string[] | undefined>; body: string }> {
  return new Promise((resolve, reject) => {
    const req = https.get(url, { headers }, (res) => {
      let body = '';
      res.on('data', (c) => (body += c));
      res.on('end', () => resolve({ status: res.statusCode ?? 0, headers: res.headers, body }));
    });
    req.on('error', reject);
    req.setTimeout(timeoutMs, () => req.destroy(new Error(`request timed out after ${timeoutMs}ms`)));
  });
}

/**
 * Resolve the latest release tag (e.g. `v0.9.9`).
 *
 * Primary: read the redirect `Location` from `github.com/<repo>/releases/latest`
 * — same trick install.sh uses, because the unauthenticated GitHub API is
 * rate-limited to 60 req/h/IP and 403s on shared/cloud hosts (issue #325). The
 * redirect has no such limit. Fall back to the API only if the redirect can't
 * be read.
 */
export async function resolveLatestVersion(repo = REPO, timeoutMs = 12000): Promise<string> {
  try {
    const res = await httpsGet(
      `https://github.com/${repo}/releases/latest`,
      { 'User-Agent': 'codegraph-upgrade' },
      timeoutMs
    );
    const loc = res.headers.location;
    const tag = parseLatestTagFromLocation(Array.isArray(loc) ? loc[0] : loc);
    if (tag) return normalizeVersion(tag);
  } catch {
    /* fall through to API */
  }
  try {
    const res = await httpsGet(
      `https://api.github.com/repos/${repo}/releases/latest`,
      { 'User-Agent': 'codegraph-upgrade', Accept: 'application/vnd.github+json' },
      timeoutMs
    );
    const tag = JSON.parse(res.body)?.tag_name;
    if (typeof tag === 'string' && tag) return normalizeVersion(tag);
  } catch {
    /* fall through to error */
  }
  throw new Error(
    'could not resolve the latest version from GitHub. Check your network, or pin a version: `codegraph-vba upgrade <version>`.'
  );
}

// ---------------------------------------------------------------------------
// Orchestrator
// ---------------------------------------------------------------------------

export interface UpgradeOptions {
  /** Pin a specific version (positional arg or CODEGRAPH_VERSION). */
  version?: string;
  /** Report current vs latest, don't change anything. */
  check?: boolean;
  /** Reinstall even if already on the resolved version. */
  force?: boolean;
}

/** Injectable side-effects so the orchestrator stays unit-testable. */
export interface UpgradeDeps {
  currentVersion: string;
  method: InstallMethod;
  resolveLatest: (pin?: string) => Promise<string>;
  /** Run a command inheriting stdio; returns its exit code (-1 = spawn failed). */
  run: (cmd: string, args: string[], env?: NodeJS.ProcessEnv) => number;
  /** Run a command capturing stdout (nothing reaches the terminal); null = spawn failed. */
  capture: (cmd: string, args: string[]) => { code: number; stdout: string } | null;
  hasCommand: (cmd: string) => boolean;
  log: (msg: string) => void;
  warn: (msg: string) => void;
  error: (msg: string) => void;
  platform: NodeJS.Platform;
  /**
   * Override for the version read from the installed package.json. Used by
   * the post-install verification: if npm reports success but the on-disk
   * version differs from the target, the upgrade retries from the registry
   * tarball URL. `null` simulates "could not read the file" (no retry, no
   * warning); undefined falls back to actually reading the package.json.
   * Tests set this to simulate the silent stale-install failure mode without
   * touching the filesystem.
   */
  installedPackageVersion?: string | null;
  /**
   * Override for the orphan-staging sweep count. Tests use this to verify
   * pre-cleanup runs (and is skipped for non-global methods). undefined
   * triggers a real `readdir` + `rm` against the global node_modules tree.
   */
  orphanCleanupOverride?: () => number;
}

const c = {
  bold: (s: string) => `\x1b[1m${s}\x1b[0m`,
  dim: (s: string) => `\x1b[2m${s}\x1b[0m`,
  green: (s: string) => `\x1b[32m${s}\x1b[0m`,
  yellow: (s: string) => `\x1b[33m${s}\x1b[0m`,
  cyan: (s: string) => `\x1b[36m${s}\x1b[0m`,
};

/** The honest, additive re-index reminder shown after a successful upgrade. */
export function reindexAdvisory(): string {
  return [
    c.dim('Your existing project indexes keep working, but were built by the previous version.'),
    c.dim('To pick up this version’s extraction improvements, refresh each project:'),
    `  ${c.cyan('codegraph-vba sync')}        ${c.dim('# incremental, fast')}`,
    `  ${c.cyan('codegraph-vba index -f')}    ${c.dim('# full rebuild')}`,
    c.dim('(`codegraph-vba status` flags any index that predates the engine you’re running.)'),
  ].join('\n');
}

/**
 * Returns the process exit code (0 = success / nothing to do, 1 = failure).
 */
export async function runUpgrade(opts: UpgradeOptions, deps: UpgradeDeps): Promise<number> {
  const { currentVersion, method } = deps;

  // Resolve the target version (pinned or latest).
  let latest: string;
  try {
    latest = normalizeVersion(opts.version || (await deps.resolveLatest()));
  } catch (err) {
    deps.error(err instanceof Error ? err.message : String(err));
    return 1;
  }

  const currentDisplay = normalizeVersion(currentVersion);
  deps.log(`${c.bold('CodeGraph-VBA')}  current ${c.cyan(currentDisplay)}  ${opts.version ? 'target' : 'latest'} ${c.cyan(latest)}`);

  let updateAvailable = isUpdateAvailable(currentVersion, latest);

  if (method.kind === 'source') {
    try {
      const git = deps.platform === 'win32' ? 'git.exe' : 'git';
      const { execFileSync } = await import('child_process');
      // Fetch tags first to make sure we have the latest tags
      execFileSync(git, ['fetch', '--tags'], { stdio: 'ignore', cwd: method.root });

      const localHash = execFileSync(git, ['rev-parse', 'HEAD'], { encoding: 'utf-8', cwd: method.root, stdio: ['ignore', 'pipe', 'ignore'] }).trim();
      let targetHash = '';
      try {
        targetHash = execFileSync(git, ['rev-parse', latest], { encoding: 'utf-8', cwd: method.root, stdio: ['ignore', 'pipe', 'ignore'] }).trim();
      } catch {
        // Fall back to origin/main
        try {
          targetHash = execFileSync(git, ['rev-parse', 'origin/main'], { encoding: 'utf-8', cwd: method.root, stdio: ['ignore', 'pipe', 'ignore'] }).trim();
        } catch {
          // ignore
        }
      }

      if (localHash && targetHash && localHash !== targetHash) {
        updateAvailable = true;
      }
    } catch {
      // ignore git errors and fall back to version comparison
    }
  }

  if (opts.check) {
    if (updateAvailable) {
      deps.log(c.yellow(`An update is available: ${currentDisplay} → ${latest}`));
      deps.log(c.dim('Run `codegraph-vba upgrade` to install it.'));
    } else {
      deps.log(c.green(`You’re on the latest version (${currentDisplay}).`));
    }
    return 0;
  }

  if (!updateAvailable && !opts.force && !opts.version) {
    deps.log(c.green(`Already up to date (${currentDisplay}).`));
    deps.log(c.dim('Use `--force` to reinstall, or `codegraph-vba upgrade <version>` to change versions.'));
    return 0;
  }

  // Dispatch by install method. bundle/npm perform a real binary update, so
  // after they succeed we self-heal the front-load hook (below); npx/source/
  // unknown don't update anything here, so they return directly.
  let code: number;
  switch (method.kind) {
    case 'bundle':
      code = await (method.os === 'windows'
        ? upgradeWindowsBundle(method, latest, deps)
        : upgradeUnixBundle(method, opts.version ? latest : undefined, deps));
      break;
    case 'npm':
      // npm version specs have no leading "v" (`@0.9.8`, not `@v0.9.8` — the
      // latter resolves as a nonexistent dist-tag). The resolved target
      // (`latest`, with the `v` stripped) is passed separately so the
      // post-install verification compares against the actual target version
      // (not the literal string "latest").
      code = await upgradeNpm(method, opts.version ? stripV(latest) : 'latest', stripV(latest), deps);
      break;
    case 'npx':
      deps.log(c.green('npx always runs the latest version on demand — nothing to upgrade.'));
      deps.log(c.dim(`Force a fresh fetch with: npx ${NPM_PACKAGE}@latest`));
      return 0;
    case 'source': {
      deps.warn(`Running from a source checkout at ${method.root}.`);
      deps.log(c.dim('Updating source checkout via git pull and rebuilding…'));
      const originalCwd = process.cwd();
      try {
        process.chdir(method.root);
        
        const git = deps.platform === 'win32' ? 'git.exe' : 'git';
        deps.log(c.dim(`Running: git pull`));
        let code = deps.run(git, ['pull'], process.env);
        if (code !== 0) {
          deps.error('git pull failed.');
          return 1;
        }
        
        const pnpm = deps.platform === 'win32' ? 'pnpm.cmd' : 'pnpm';
        const npm = deps.platform === 'win32' ? 'npm.cmd' : 'npm';
        const hasPnpmLock = fs.existsSync(path.join(method.root, 'pnpm-lock.yaml'));
        const packageManager = hasPnpmLock ? pnpm : npm;
        
        deps.log(c.dim(`Running: ${packageManager} install`));
        code = deps.run(packageManager, ['install'], process.env);
        if (code !== 0) {
          deps.error(`${packageManager} install failed.`);
          return 1;
        }

        deps.log(c.dim(`Running: ${packageManager} run build`));
        code = deps.run(packageManager, ['run', 'build'], process.env);
        if (code !== 0) {
          deps.error(`${packageManager} run build failed.`);
          return 1;
        }
        
        deps.log('');
        deps.log(c.green('✓ Upgrade complete.'));
        return 0;
      } catch (err) {
        deps.error(`Upgrade failed: ${err instanceof Error ? err.message : String(err)}`);
        return 1;
      } finally {
        try {
          process.chdir(originalCwd);
        } catch {
          // ignore failures restoring CWD
        }
      }
    }
    default:
      deps.error(`Couldn’t determine how CodeGraph-VBA was installed (${method.reason}).`);
      deps.log(c.dim(`Reinstall manually — see https://github.com/${REPO}#install`));
      return 1;
  }

  // After a successful update, ensure the front-load prompt hook is wired for an
  // already-configured global Claude install — so existing users pick it up on
  // upgrade, not only on a fresh `install` (the hook config is version-agnostic,
  // so the still-running old binary can write it safely). Idempotent + gated on
  // an existing Claude config, and skipped entirely by the kill-switch. Never
  // fatal to the upgrade.
  if (code === 0) {
    let probe: VersionProbe = 'inconclusive';
    try {
      probe = reportResolvedVersion(latest, deps);
    } catch {
      /* an inconclusive probe must not fail the upgrade */
    }
    try {
      await selfHealPromptHook(deps);
    } catch {
      /* a hook-wiring hiccup must not fail the upgrade */
    }
    // The refresh executes whatever `codegraph` PATH resolves. If the probe
    // just proved that's a stale shadowed install, spawning it would rewrite
    // the agent surfaces with the very templates the refresh exists to heal —
    // skip, and point at the manual command for after the PATH is fixed.
    if (probe !== 'mismatch') {
      try {
        selfHealInstalledSurfaces(deps);
      } catch {
        /* a refresh hiccup must not fail the upgrade */
      }
    } else {
      deps.log(c.dim('Skipped refreshing agent instructions/config — run `codegraph install --refresh` once the PATH is fixed.'));
    }
  }
  return code;
}

type VersionProbe = 'match' | 'mismatch' | 'inconclusive';

/**
 * Prove the upgrade actually took: spawn the `codegraph` this terminal's PATH
 * resolves and compare its reported version to the target. Catches the silent
 * failure mode where ANOTHER install shadows the one we just upgraded (issue
 * #1071 — e.g. a stale `npm i -g` copy earlier on PATH than the bundle
 * launcher): the upgrade "succeeds" but `codegraph -v` — in this terminal and
 * every future one — keeps serving the old version. Exported for unit tests.
 */
export function verifyResolvedVersion(latest: string, deps: UpgradeDeps): VersionProbe {
  if (!deps.hasCommand('codegraph')) return 'inconclusive';
  // Windows installs expose codegraph through a .cmd launcher; Node can't
  // spawn .cmd files without a shell, so route through cmd.exe there.
  const probe = deps.platform === 'win32'
    ? deps.capture('cmd.exe', ['/d', '/s', '/c', 'codegraph --version'])
    : deps.capture('codegraph', ['--version']);
  if (!probe || probe.code !== 0) return 'inconclusive';
  // `codegraph --version` prints the bare version; take the last non-empty
  // line so a stray runtime warning above it can't spoil the parse.
  const reported = probe.stdout.trim().split(/\r?\n/).pop()?.trim() ?? '';
  if (!parseSemver(reported)) return 'inconclusive';
  return compareVersions(reported, latest) === 0 ? 'match' : 'mismatch';
}

/**
 * Log the outcome of the post-upgrade version probe. On a match the user
 * knows the current terminal is already serving the new version; on a
 * mismatch they get told exactly which stale install is hijacking their PATH
 * instead of discovering it via a mysteriously unchanged `codegraph -v`.
 * Inconclusive probes fall back to the old soft hint — never a scare on
 * setups we can't inspect (no `codegraph` on PATH yet, exotic wrappers).
 * Returns the probe result so the caller can gate the post-upgrade refresh
 * (which spawns the PATH-resolved binary) on it.
 */
function reportResolvedVersion(latest: string, deps: UpgradeDeps): VersionProbe {
  const { method } = deps;
  // A project-local npm install isn't served by PATH's `codegraph` (that
  // would be some other install) — a probe could only false-alarm.
  if (method.kind === 'npm' && method.scope === 'local') return 'inconclusive';
  const probe = verifyResolvedVersion(latest, deps);
  switch (probe) {
    case 'match':
      deps.log(c.green(`✓ \`codegraph\` on your PATH now reports ${latest} — this terminal is already using it.`));
      break;
    case 'mismatch':
      deps.warn(`Installed ${latest}, but the \`codegraph\` this terminal resolves still reports an older version.`);
      deps.log(c.dim('Another CodeGraph install earlier on your PATH is shadowing the one just upgraded.'));
      deps.log(c.dim('Find every copy with `which -a codegraph` (Windows: `where codegraph`) and remove or upgrade the stale one.'));
      break;
    case 'inconclusive':
      deps.log(c.dim('Open a new terminal if `codegraph --version` looks unchanged (PATH cache).'));
      break;
  }
  return probe;
}

/**
 * Refresh the agent surfaces previous installs wrote — the marker-fenced
 * instructions sections (CLAUDE.md / AGENTS.md / GEMINI.md), MCP entries,
 * legacy-hook cleanups — so they match the version that will serve them.
 * Unlike the prompt hook above, this content is NOT version-agnostic: the
 * templates are baked into the binary, so the still-running old process
 * would only rewrite its own stale copy — the exact staleness this heals.
 * We therefore spawn the freshly-installed binary (`codegraph install
 * --refresh`), which is refresh-only: agents never configured stay
 * untouched, and permission / prompt-hook choices are preserved. Gated on
 * `codegraph` being resolvable on PATH (an npm-local install isn't) and on
 * the kill-switch; never fatal to the upgrade.
 */
function selfHealInstalledSurfaces(deps: UpgradeDeps): void {
  if (process.env.CODEGRAPH_NO_INSTALL_REFRESH === '1') return;
  if (!deps.hasCommand('codegraph')) return;
  deps.log(c.dim('Refreshing agent instruction sections and config written by previous versions…'));
  // Windows installs expose codegraph through a .cmd launcher. Node cannot
  // spawn .cmd files directly without a shell, so route the constant command
  // through cmd.exe there (the same launcher a terminal would resolve).
  const code = deps.platform === 'win32'
    ? deps.run('cmd.exe', ['/d', '/s', '/c', 'codegraph install --refresh'])
    : deps.run('codegraph', ['install', '--refresh']);
  if (code !== 0) {
    deps.warn('Could not refresh the installed agent surfaces — run `codegraph install --refresh` manually.');
  }
}

/**
 * Wire the Claude `UserPromptSubmit` front-load hook on upgrade for an
 * already-configured global Claude install. No-op when Claude isn't configured,
 * when the hook is already present, or when the kill-switch is set.
 */
async function selfHealPromptHook(deps: UpgradeDeps): Promise<void> {
  if (process.env.CODEGRAPH_NO_PROMPT_HOOK === '1' || process.env.CODEGRAPH_PROMPT_HOOK === '0') return;
  const { claudeTarget, writePromptHookEntry } = await import('../installer/targets/claude');
  if (!claudeTarget.detect('global').alreadyConfigured) return;
  const res = writePromptHookEntry('global');
  if (res.action === 'created' || res.action === 'updated') {
    deps.log(
      c.dim('Enabled the CodeGraph-VBA front-load hook for Claude Code (structural prompts). Disable any time: CODEGRAPH_NO_PROMPT_HOOK=1'),
    );
  }
}

function upgradeUnixBundle(
  method: Extract<InstallMethod, { kind: 'bundle' }>,
  pinned: string | undefined,
  deps: UpgradeDeps
): number {
  const downloader = deps.hasCommand('curl')
    ? `curl -fsSL ${INSTALL_SH_URL}`
    : deps.hasCommand('wget')
      ? `wget -qO- ${INSTALL_SH_URL}`
      : null;
  if (!downloader) {
    deps.error('Neither curl nor wget is available to download the installer.');
    deps.log(c.dim(`Install curl, or run manually:  ${INSTALL_SH_URL} | sh`));
    return 1;
  }

  const env: NodeJS.ProcessEnv = { ...process.env };
  if (method.installDir) env.CODEGRAPH_INSTALL_DIR = method.installDir;
  if (pinned) env.CODEGRAPH_VERSION = pinned;

  deps.log(c.dim(`Running the installer (${downloader} | sh)…`));
  const code = deps.run('sh', ['-c', `${downloader} | sh`], env);
  if (code !== 0) {
    deps.error(`Installer exited with code ${code}.`);
    return 1;
  }
  deps.log('');
  // No "open a new terminal" hedge here — after the swap, runUpgrade probes
  // the PATH-resolved `codegraph --version` and reports the real outcome.
  deps.log(c.green('✓ Upgrade complete.'));
  deps.log(reindexAdvisory());
  return 0;
}

/** Build the in-place Windows upgrade script (exported for unit-testing). */
export function buildWindowsUpgradeScript(bundleRoot: string, version: string, arch: string): string {
  const target = `win32-${arch}`;
  // Asset names carry the `codegraph-vba-` prefix (build-bundle.sh +
  // release.yml publish `codegraph-vba-<target>.{zip,tar.gz}`). The old
  // `codegraph-<target>` name 404s — keep this in lockstep with the inner
  // dir below and with install.sh / install.ps1.
  const url = `https://github.com/${REPO}/releases/download/${version}/codegraph-vba-${target}.zip`;
  // Windows can't DELETE a running exe but CAN rename it, so we upgrade IN
  // PLACE: download → rename the locked node.exe aside → extract the new bundle
  // over current\. Synchronous, no detached helper (which dies under SSH/job
  // objects and has worse UX). The running process keeps its renamed node.exe
  // mapped; the NEXT `codegraph` invocation uses the new one. We can't reuse
  // install.ps1 here — it `Remove-Item`s current\, which fails on the locked exe.
  return [
    `$ErrorActionPreference='Stop'`,
    `$dest='${bundleRoot}'`,
    `$url='${url}'`,
    `Write-Host "Downloading $url"`,
    `$tmp=Join-Path $env:TEMP ('cg-up-'+[guid]::NewGuid().ToString('N'))`,
    `New-Item -ItemType Directory -Force -Path $tmp | Out-Null`,
    `$zip=Join-Path $tmp 'cg.zip'`,
    `Invoke-WebRequest -Uri $url -OutFile $zip`,
    `$stage=Join-Path $tmp 'stage'`,
    `Expand-Archive -Path $zip -DestinationPath $stage -Force`,
    `$inner=Join-Path $stage 'codegraph-vba-${target}'`,
    `$src=if(Test-Path $inner){$inner}else{$stage}`,
    `$node=Join-Path $dest 'node.exe'`,
    `if(Test-Path $node){Rename-Item -Path $node -NewName ('node.exe.old-'+[guid]::NewGuid().ToString('N')) -Force}`,
    `Copy-Item -Path (Join-Path $src '*') -Destination $dest -Recurse -Force`,
    `Get-ChildItem -Path $dest -Filter 'node.exe.old-*' -ErrorAction SilentlyContinue | ForEach-Object { try { Remove-Item $_.FullName -Force -ErrorAction Stop } catch {} }`,
    `Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue`,
    `Write-Host "Installed CodeGraph-VBA ${version} to $dest"`,
  ].join(';');
}

function upgradeWindowsBundle(
  method: Extract<InstallMethod, { kind: 'bundle' }>,
  latest: string,
  deps: UpgradeDeps
): number {
  const arch = process.arch === 'arm64' ? 'arm64' : 'x64';
  const script = buildWindowsUpgradeScript(method.bundleRoot, latest, arch);
  // -EncodedCommand (base64 UTF-16LE), NOT -Command: Node's Windows argv→command
  // -line quoting mangles a long multi-statement script, so PowerShell never
  // parses it. Encoding sidesteps all shell quoting — the canonical approach.
  const encoded = Buffer.from(script, 'utf16le').toString('base64');
  deps.log(c.dim(`Downloading and installing ${latest}…`));
  const code = deps.run('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', encoded]);
  if (code !== 0) {
    deps.error(`Installer exited with code ${code}.`);
    return 1;
  }
  deps.log('');
  // The running node.exe was renamed aside, so the version probe in
  // runUpgrade already exercises the NEW binary — no terminal hedge needed.
  deps.log(c.green('✓ Upgrade complete.'));
  deps.log(reindexAdvisory());
  return 0;
}

/**
 * How to invoke npm. On Windows npm is a .cmd batch file, which Node refuses
 * to spawn without a shell (EINVAL since the CVE-2024-27980 hardening) — a
 * direct `npm.cmd` spawn fails on every current Node, so route it through
 * cmd.exe, the same way the surface-refresh step invokes the .cmd launcher.
 * (Verified live on the Windows VM: `spawnSync('npm.cmd')` → EINVAL;
 * `cmd.exe /d /s /c npm …` → works.)
 */
export function npmInvocation(platform: NodeJS.Platform, npmArgs: string[]): { cmd: string; args: string[] } {
  if (platform === 'win32') {
    return { cmd: 'cmd.exe', args: ['/d', '/s', '/c', ['npm', ...npmArgs].join(' ')] };
  }
  return { cmd: 'npm', args: npmArgs };
}

function upgradeNpm(
  method: Extract<InstallMethod, { kind: 'npm' }>,
  versionSpec: string,
  targetVersion: string,
  deps: UpgradeDeps
): number {
  // PRE-CLEANUP: npm leaves orphan staging dirs when an upgrade is interrupted
  // (EBUSY on a locked node.exe, EPERM during cleanup, etc.). Those orphans can
  // hold a `dist/` + `node_modules/` with no `package.json`, lock the next
  // upgrade attempt, and eventually cause npm to report "Installed X" without
  // actually moving the package.json — a silent stale-cache failure mode we
  // hit on v1.7.0 (see issue write-up in obs 17859). Sweep them before every
  // upgrade. Best-effort; never fatal to the upgrade.
  if (method.scope === 'global') {
    const cleaned = deps.orphanCleanupOverride
      ? deps.orphanCleanupOverride()
      : cleanupOrphanStagings(deps);
    if (cleaned > 0) {
      deps.log(c.dim(`Removed ${cleaned} orphan npm staging dir${cleaned === 1 ? '' : 's'} before upgrade.`));
    }
  }

  // `--prefer-online` makes npm reach the registry instead of the local cache,
  // which catches up stale npm cache entries (the cause of the silent
  // "Installed v1.7.0 but package.json still 1.6.3" failure mode).
  const args = method.scope === 'global'
    ? ['install', '-g', '--prefer-online', `${NPM_PACKAGE}@${versionSpec}`]
    : ['install', '--prefer-online', `${NPM_PACKAGE}@${versionSpec}`];
  deps.log(c.dim(`Running: npm ${args.join(' ')}`));
  const inv = npmInvocation(deps.platform, args);
  const firstAttemptCode = deps.run(inv.cmd, inv.args, process.env);

  // RECOVERY PATH: if the first attempt failed outright OR npm reported
  // success but the installed package.json version is stale (cache, EPERM
  // mid-install, PATH shadow), retry from the dist tarball URL — which bypasses
  // the npm cache + staging dirs entirely. This is the user's escape hatch
  // when the regular upgrade fails for any reason; without it, the only
  // option is a manual `npm cache clean --force` they didn't ask for.
  let code = firstAttemptCode;
  let stale = false;
  if (method.scope === 'global') {
    const actual = deps.installedPackageVersion !== undefined
      ? deps.installedPackageVersion
      : readInstalledPackageVersion(deps);
    if (actual !== null && actual !== targetVersion) {
      stale = true;
      deps.warn(`npm reported success but installed package.json says ${actual}, expected ${targetVersion}.`);
    } else if (actual === null && firstAttemptCode === 0) {
      deps.warn('Could not read installed package.json to verify the upgrade. Run `codegraph-vba --version` to confirm.');
    } else if (firstAttemptCode !== 0) {
      deps.warn(`npm install failed (exit ${firstAttemptCode}); trying tarball fallback.`);
    }

    if (stale || firstAttemptCode !== 0) {
      deps.log(c.dim('Retrying from the npm registry tarball (bypasses cache + staging).'));
      const tarballCode = upgradeNpmFromTarball(targetVersion, deps);
      if (tarballCode === 0) {
        code = 0;
      } else {
        deps.error(
          `Tarball fallback also failed (exit ${tarballCode}). ` +
          `Retry with \`npm cache clean --force\`, then re-run.`
        );
        code = 1;
      }
    }
  }

  if (code !== 0) {
    deps.error(`npm exited with code ${code}.`);
    if (method.scope === 'global') {
      deps.log(c.dim('If this is a permissions error (EACCES), your global prefix needs sudo, or use a'));
      deps.log(c.dim('Node version manager (nvm/fnm) so global installs don’t require root.'));
    }
    return 1;
  }
  deps.log('');
  deps.log(c.green('✓ Upgrade complete.'));
  deps.log(reindexAdvisory());
  return 0;
}

/**
 * Resolve the npm global root for the current install. The upgrade module is
 * always running from inside its own install — `<root>/node_modules/@scope/pkg/
 * dist/upgrade/index.js` on both Windows and Unix — so we can derive the root
 * by walking up `__filename` to the first `node_modules` ancestor and taking
 * its parent. No need to shell out to `npm root -g` (which would consume a
 * capture slot in tests and is platform-dependent).
 *
 * Returns null when not running from a `node_modules` tree (e.g. source
 * checkout, npx cache, bundle). Callers treat null as "no globals to scan".
 */
export function resolveNpmGlobalRoot(_deps?: UpgradeDeps): string | null {
  const filename = __filename || '';
  const norm = filename.replace(/\\/g, '/');
  const m = /^(.*?)\/node_modules\//.exec(norm);
  if (!m) return null;
  return m[1] ?? null;
}

/**
 * Find npm orphan staging dirs in the global node_modules tree. An "orphan" is
 * a directory matching `.codegraph-vba-<HASH>` or `.codegraph-vba-win32-<ARCH>`
 * next to the live `@aroman22/codegraph-vba` install — npm creates these when
 * it begins a global upgrade and the previous run was interrupted (EBUSY on
 * node.exe, EPERM during cleanup, etc.). They block subsequent upgrades and
 * can hold a partial copy of an older version that confuses `npm view`.
 *
 * Returns the list of absolute paths to remove. Pure so it's unit-tested with
 * an injected fs.
 */
export function findOrphanStagings(
  globalRoot: string,
  exists: (p: string) => boolean = fs.existsSync,
  readdir: (p: string) => string[] = (p) => fs.readdirSync(p)
): string[] {
  const orphans: string[] = [];
  if (!exists(globalRoot)) return orphans;
  let entries: string[];
  try {
    entries = readdir(globalRoot);
  } catch {
    return orphans;
  }
  for (const entry of entries) {
    if (!entry.startsWith('.codegraph-vba-')) continue;
    // npm registry + paths use forward slashes uniformly; force path.posix so
    // Windows `path.join` doesn't substitute backslashes (which would mismatch
    // the npm cache's own internal paths).
    orphans.push(path.posix.join(globalRoot, entry));
  }
  return orphans;
}

/**
 * Remove orphan npm staging dirs from the global node_modules tree. Best-effort
 * — failures are logged but never fatal to the upgrade (a leftover orphan is
 * annoying; blocking the upgrade over it is worse). Exposed for tests.
 *
 * Returns the count successfully removed.
 */
export function cleanupOrphanStagings(
  deps: UpgradeDeps,
  opts: { rm?: (p: string) => boolean; exists?: (p: string) => boolean; readdir?: (p: string) => string[]; globalRoot?: string | null } = {}
): number {
  const rm = opts.rm ?? ((p: string) => { try { fs.rmSync(p, { recursive: true, force: true }); return true; } catch { return false; } });
  const exists = opts.exists ?? fs.existsSync;
  const readdir = opts.readdir ?? ((p: string) => fs.readdirSync(p));
  const globalRoot = opts.globalRoot !== undefined ? opts.globalRoot : resolveNpmGlobalRoot(deps);
  if (!globalRoot) return 0;
  const orphans = findOrphanStagings(globalRoot, exists, readdir);
  let removed = 0;
  for (const orphan of orphans) {
    if (rm(orphan)) removed++;
  }
  return removed;
}

/**
 * Read the installed package.json and return its `version` field, or null if
 * the file can't be read or parsed. Used by the post-install verification to
 * catch silent stale-install failure modes (npm reports success but the
 * on-disk version is older than the target). Exposed for tests.
 */
export function readInstalledPackageVersion(
  deps: UpgradeDeps,
  opts: { readFile?: (p: string) => string; exists?: (p: string) => boolean; globalRoot?: string | null } = {}
): string | null {
  const exists = opts.exists ?? fs.existsSync;
  const readFile = opts.readFile ?? ((p: string) => fs.readFileSync(p, 'utf-8')) as (p: string) => string;
  const root = opts.globalRoot ?? resolveNpmGlobalRoot(deps);
  if (!root) return null;
  const pkgPath = path.join(root, 'node_modules', NPM_PACKAGE, 'package.json');
  if (!exists(pkgPath)) return null;
  try {
    const parsed = JSON.parse(readFile(pkgPath));
    return typeof parsed?.version === 'string' ? parsed.version : null;
  } catch {
    return null;
  }
}

/**
 * Last-resort fallback: install from the dist tarball URL, which bypasses the
 * npm cache and staging dirs entirely. Used when the regular `npm install -g`
 * fails outright OR reports success but the installed package.json doesn't
 * match the target (stale cache, stuck staging, EPERM mid-install).
 *
 * `targetVersion` must be the RESOLVED version (e.g. `'0.9.9'`), NOT the
 * user-supplied spec — `latest` is a dist-tag, not a literal tarball name;
 * the URL needs the real semver.
 */
function upgradeNpmFromTarball(targetVersion: string, deps: UpgradeDeps): number {
  if (!targetVersion || targetVersion === 'latest') {
    deps.error('Cannot resolve a tarball URL without a resolved version; pin one: `codegraph-vba upgrade <version>`.');
    return 1;
  }
  const url = `https://registry.npmjs.org/${NPM_PACKAGE}/-/${NPM_PACKAGE.replace(/^@/, '')}-${targetVersion}.tgz`;
  const args = ['install', '-g', '--force', url];
  deps.log(c.dim(`Fallback install from tarball: npm ${args.join(' ')}`));
  const inv = npmInvocation(deps.platform, args);
  return deps.run(inv.cmd, inv.args, process.env);
}

// ---------------------------------------------------------------------------
// Production deps wiring (used by the CLI)
// ---------------------------------------------------------------------------

/**
 * True if `cmd` resolves to an executable on PATH. A pure-Node PATH scan — NOT
 * a spawned `command -v`/`which`: `command` is a shell builtin (no standalone
 * binary on Debian, though macOS ships one), and `which` isn't guaranteed
 * present on minimal images, so spawning either is unreliable. Scanning PATH
 * ourselves behaves identically on every platform.
 */
export function hasCommand(cmd: string): boolean {
  const isWin = process.platform === 'win32';
  const dirs = (process.env.PATH || process.env.Path || '').split(path.delimiter).filter(Boolean);
  const exts = isWin ? (process.env.PATHEXT || '.EXE;.CMD;.BAT;.COM').split(';') : [''];
  for (const dir of dirs) {
    for (const ext of exts) {
      const candidate = path.join(dir, cmd + ext);
      try {
        if (!fs.statSync(candidate).isFile()) continue;
        if (isWin) return true;
        fs.accessSync(candidate, fs.constants.X_OK);
        return true;
      } catch {
        /* not here / not executable — keep scanning */
      }
    }
  }
  return false;
}

export function defaultRun(cmd: string, args: string[], env?: NodeJS.ProcessEnv): number {
  const r = spawnSync(cmd, args, {
    stdio: 'inherit',
    env: env ?? process.env,
    shell: process.platform === 'win32',
    windowsHide: true,
  });
  if (r.error) return -1;
  return r.status ?? -1;
}

export function defaultCapture(cmd: string, args: string[]): { code: number; stdout: string } | null {
  // stdio is piped (the default with `encoding`), so nothing the probed
  // command prints reaches the user's terminal. The timeout keeps a wedged
  // probe from hanging the upgrade's last step.
  const r = spawnSync(cmd, args, { encoding: 'utf-8', windowsHide: true, timeout: 30_000 });
  if (r.error) return null;
  return { code: r.status ?? -1, stdout: r.stdout ?? '' };
}
