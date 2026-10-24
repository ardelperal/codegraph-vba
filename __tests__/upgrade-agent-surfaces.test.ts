/**
 * Issue #332: `upgrade` must not rewrite agent surfaces (instruction files,
 * MCP entries, hooks) by default, and `install` must leave instruction files
 * that gentle-ai manages alone.
 *
 * Every test runs against a fixture HOME in a temp dir. HOME, USERPROFILE,
 * APPDATA and XDG_CONFIG_HOME are overridden, so no real agent config is
 * ever read or written.
 */

import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import { ALL_TARGETS, getTarget } from '../src/installer/targets/registry';
import { refreshTargets } from '../src/installer';
import { runUpgrade, type UpgradeDeps } from '../src/upgrade';

const GENTLE_AI_BLOCK = [
  '<!-- gentle-ai:codegraph-guidance -->',
  '## CodeGraph',
  'Initialise CodeGraph immediately in every repository.',
  '<!-- /gentle-ai:codegraph-guidance -->',
].join('\n');

const ENV_KEYS = [
  'HOME',
  'USERPROFILE',
  'APPDATA',
  'XDG_CONFIG_HOME',
  'HERMES_HOME',
  'CODEGRAPH_NO_INSTALL_REFRESH',
  'CODEGRAPH_NO_PROMPT_HOOK',
  'CODEGRAPH_PROMPT_HOOK',
] as const;

/** Map of every file under `dir` (relative path → content). */
function snapshot(dir: string): Record<string, string> {
  const out: Record<string, string> = {};
  const walk = (d: string): void => {
    for (const entry of fs.readdirSync(d, { withFileTypes: true })) {
      const p = path.join(d, entry.name);
      if (entry.isDirectory()) walk(p);
      else out[path.relative(dir, p)] = fs.readFileSync(p, 'utf-8');
    }
  };
  walk(dir);
  return out;
}

function fakeDeps(over: Partial<UpgradeDeps> = {}): {
  deps: UpgradeDeps;
  calls: Array<{ cmd: string; args: string[] }>;
  logs: string[];
} {
  const calls: Array<{ cmd: string; args: string[] }> = [];
  const logs: string[] = [];
  const deps: UpgradeDeps = {
    currentVersion: '1.0.0',
    method: { kind: 'npm', scope: 'global' },
    resolveLatest: async () => '2.0.0',
    run: (cmd, args) => {
      calls.push({ cmd, args });
      return 0;
    },
    capture: () => null,
    hasCommand: () => true,
    log: (m) => logs.push(m),
    warn: (m) => logs.push(m),
    error: (m) => logs.push(m),
    platform: 'linux',
    installedPackageVersion: '2.0.0',
    orphanCleanupOverride: () => 0,
    ...over,
  };
  return { deps, calls, logs };
}

describe('Issue #332 — fixture home with a gentle-ai-managed CLAUDE.md', () => {
  let tmpHome: string;
  let tmpCwd: string;
  let origCwd: string;
  let savedEnv: Record<string, string | undefined>;

  beforeEach(() => {
    tmpHome = fs.mkdtempSync(path.join(os.tmpdir(), 'cg-332-home-'));
    tmpCwd = fs.mkdtempSync(path.join(os.tmpdir(), 'cg-332-cwd-'));
    origCwd = process.cwd();
    savedEnv = Object.fromEntries(ENV_KEYS.map((k) => [k, process.env[k]]));
    for (const k of ENV_KEYS) delete process.env[k];
    process.env.HOME = tmpHome;
    process.env.USERPROFILE = tmpHome;
    process.env.APPDATA = path.join(tmpHome, '.config');
    process.env.XDG_CONFIG_HOME = path.join(tmpHome, '.config');
    process.chdir(tmpCwd);

    // Claude Code is configured (MCP entry present) and gentle-ai owns its
    // CLAUDE.md. Codex is configured too, with a plain AGENTS.md that has no
    // gentle-ai marker: it must keep receiving the CodeGraph block.
    fs.writeFileSync(
      path.join(tmpHome, '.claude.json'),
      JSON.stringify({ mcpServers: { codegraph: { command: 'codegraph', args: ['serve', '--mcp'] } } }, null, 2),
    );
    fs.mkdirSync(path.join(tmpHome, '.claude'), { recursive: true });
    fs.writeFileSync(path.join(tmpHome, '.claude', 'CLAUDE.md'), `# Mine\n\n${GENTLE_AI_BLOCK}\n`);
    fs.mkdirSync(path.join(tmpHome, '.codex'), { recursive: true });
    fs.writeFileSync(
      path.join(tmpHome, '.codex', 'config.toml'),
      '[mcp_servers.codegraph]\ncommand = "codegraph"\nargs = ["serve", "--mcp"]\n',
    );
    fs.writeFileSync(path.join(tmpHome, '.codex', 'AGENTS.md'), '# Codex notes\n');
  });

  afterEach(() => {
    process.chdir(origCwd);
    for (const k of ENV_KEYS) {
      if (savedEnv[k] === undefined) delete process.env[k];
      else process.env[k] = savedEnv[k];
    }
    fs.rmSync(tmpHome, { recursive: true, force: true });
    fs.rmSync(tmpCwd, { recursive: true, force: true });
  });

  describe('upgrade without --refresh-agents', () => {
    it('leaves every agent file byte-identical and runs no refresh or hook write', async () => {
      const before = snapshot(tmpHome);
      const { deps, calls } = fakeDeps();

      const code = await runUpgrade({}, deps);

      expect(code).toBe(0);
      expect(snapshot(tmpHome)).toEqual(before);
      expect(calls.some((c) => c.args.includes('--refresh'))).toBe(false);
      expect(fs.existsSync(path.join(tmpHome, '.claude', 'settings.json'))).toBe(false);
    });

    it('summary names the agent surfaces it did not touch and the opt-in flag', async () => {
      const { deps, logs } = fakeDeps();

      await runUpgrade({}, deps);

      const summary = logs.join('\n');
      expect(summary).toMatch(/--refresh-agents/);
      expect(summary).toMatch(/instruction files/);
      expect(summary).toMatch(/MCP entries/);
      expect(summary).toMatch(/prompt hook/);
    });

    it('CODEGRAPH_NO_INSTALL_REFRESH and CODEGRAPH_NO_PROMPT_HOOK do not matter when not opted in', async () => {
      // Kill-switches only narrow the explicit opt-in; the default path already
      // skips every surface, so setting them must not change the result.
      process.env.CODEGRAPH_NO_INSTALL_REFRESH = '1';
      const before = snapshot(tmpHome);
      const { deps, calls } = fakeDeps();

      await runUpgrade({}, deps);

      expect(calls.some((c) => c.args.includes('--refresh'))).toBe(false);
      expect(snapshot(tmpHome)).toEqual(before);
    });
  });

  describe('upgrade with the explicit --refresh-agents opt-in', () => {
    it('still refreshes configured agents via `codegraph install --refresh`', async () => {
      const { deps, calls } = fakeDeps();

      const code = await runUpgrade({ refreshAgents: true }, deps);

      expect(code).toBe(0);
      expect(calls).toContainEqual({ cmd: 'codegraph', args: ['install', '--refresh'] });
    });

    it('still wires the Claude prompt hook for a configured Claude install', async () => {
      const { deps } = fakeDeps();

      await runUpgrade({ refreshAgents: true }, deps);

      expect(fs.existsSync(path.join(tmpHome, '.claude', 'settings.json'))).toBe(true);
    });

    it('CODEGRAPH_NO_INSTALL_REFRESH=1 still disables the refresh under the opt-in', async () => {
      process.env.CODEGRAPH_NO_INSTALL_REFRESH = '1';
      const { deps, calls } = fakeDeps();

      await runUpgrade({ refreshAgents: true }, deps);

      expect(calls.some((c) => c.args.includes('--refresh'))).toBe(false);
    });

    it('CODEGRAPH_NO_PROMPT_HOOK=1 still disables the prompt hook under the opt-in', async () => {
      process.env.CODEGRAPH_NO_PROMPT_HOOK = '1';
      const { deps } = fakeDeps();

      await runUpgrade({ refreshAgents: true }, deps);

      expect(fs.existsSync(path.join(tmpHome, '.claude', 'settings.json'))).toBe(false);
    });
  });

  describe('install and install --refresh with a gentle-ai-managed file', () => {
    it('refreshTargets adds no CODEGRAPH_START block to the managed file, reports it skipped, and still updates other files', () => {
      const reports = refreshTargets(ALL_TARGETS, 'global');

      const claudeText = fs.readFileSync(path.join(tmpHome, '.claude', 'CLAUDE.md'), 'utf-8');
      expect(claudeText).toBe(`# Mine\n\n${GENTLE_AI_BLOCK}\n`);
      expect(claudeText).not.toContain('CODEGRAPH_START');

      const claude = reports.find((r) => r.id === 'claude');
      expect(claude?.skippedPaths).toContain(path.join(tmpHome, '.claude', 'CLAUDE.md'));

      const codexAgents = fs.readFileSync(path.join(tmpHome, '.codex', 'AGENTS.md'), 'utf-8');
      expect(codexAgents).toContain('<!-- CODEGRAPH_START -->');
    });

    it('claude install (not refresh) reports CLAUDE.md as skipped and leaves it unchanged', () => {
      const claude = getTarget('claude');
      if (!claude) throw new Error('claude target missing');
      const before = fs.readFileSync(path.join(tmpHome, '.claude', 'CLAUDE.md'), 'utf-8');

      const result = claude.install('global', { autoAllow: false });

      expect(result.files).toContainEqual({
        path: path.join(tmpHome, '.claude', 'CLAUDE.md'),
        action: 'skipped',
      });
      expect(fs.readFileSync(path.join(tmpHome, '.claude', 'CLAUDE.md'), 'utf-8')).toBe(before);
    });

    it('upgrade (default and --refresh-agents), install and install --refresh never change skills directories', async () => {
      // Skills reach projects through the fleet propagator, never through
      // tool installers (issue #332 scope note). Pin mtimes to a past date so
      // any rewrite, even with identical bytes, shows up as a change.
      const past = new Date('2020-01-01T00:00:00Z');
      const sentinels = [
        path.join(tmpHome, '.claude', 'skills', 'x', 'SKILL.md'),
        path.join(tmpHome, '.agents', 'skills', 'x', 'SKILL.md'),
        path.join(tmpHome, '.codex', 'skills', 'x', 'SKILL.md'),
        path.join(tmpHome, '.config', 'opencode', 'skills', 'x', 'SKILL.md'),
      ];
      for (const f of sentinels) {
        fs.mkdirSync(path.dirname(f), { recursive: true });
        fs.writeFileSync(f, `# sentinel ${path.relative(tmpHome, f)}\n`);
        fs.utimesSync(f, past, past);
      }
      const state = () =>
        sentinels.map((f) => ({ f, bytes: fs.readFileSync(f, 'utf-8'), mtimeMs: fs.statSync(f).mtimeMs }));
      const before = state();

      await runUpgrade({}, fakeDeps().deps);
      await runUpgrade({ refreshAgents: true }, fakeDeps().deps);
      for (const target of ALL_TARGETS) {
        target.install('global', { autoAllow: true, promptHook: true });
      }
      refreshTargets(ALL_TARGETS, 'global');

      expect(state()).toEqual(before);
      for (const f of sentinels) expect(fs.existsSync(f)).toBe(true);
    });
  });
});
