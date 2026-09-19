/**
 * Bundle launcher names (issue #326).
 *
 * Agent MCP configs and Dysflow launch `codegraph-vba`, so a GitHub-only
 * install (no npm) must put that name on PATH next to `codegraph`. These tests
 * extract the REAL launcher block from `scripts/build-bundle.sh` and the REAL
 * link block from `install.sh` (between their marker comments) and run them
 * against temp fixtures — no build, no download.
 *
 * POSIX only: both blocks are shell, and the install.sh block uses symlinks.
 */

import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { spawnSync } from 'child_process';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

const ROOT = path.join(__dirname, '..');

function extractBlock(file: string, marker: string): string {
  const lines = fs.readFileSync(path.join(ROOT, file), 'utf8').replace(/\r\n/g, '\n').split('\n');
  const i = lines.findIndex((l) => l.trim() === `# >>> ${marker}`);
  const j = lines.findIndex((l) => l.trim() === `# <<< ${marker}`);
  if (i < 0 || j < 0 || j <= i) throw new Error(`${marker} markers not found in ${file}`);
  return lines.slice(i + 1, j).join('\n');
}

function shq(s: string): string {
  return `'${s.replace(/'/g, `'\\''`)}'`;
}

function run(shell: string, script: string): { code: number; stdout: string; stderr: string } {
  const r = spawnSync(shell, ['-c', script], { encoding: 'utf8' });
  return { code: r.status ?? -1, stdout: r.stdout ?? '', stderr: r.stderr ?? '' };
}

describe.skipIf(process.platform === 'win32')('bundle launchers (#326)', () => {
  let tmp: string;

  beforeEach(() => {
    tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'cg-launchers-'));
  });
  afterEach(() => {
    fs.rmSync(tmp, { recursive: true, force: true });
  });

  function stageBundle(osfam: string): string {
    const stage = path.join(tmp, 'stage');
    fs.mkdirSync(path.join(stage, 'bin'), { recursive: true });
    // Stand-in for the vendored Node: prints the args it was exec'd with.
    const fakeNode = path.join(tmp, 'fake-node');
    fs.writeFileSync(fakeNode, '#!/bin/sh\necho "node $*"\n', { mode: 0o755 });
    const script =
      `set -eu\nOSFAM=${shq(osfam)}\nSTAGE=${shq(stage)}\nNODE_BIN=${shq(fakeNode)}\n` +
      extractBlock('scripts/build-bundle.sh', 'CODEGRAPH_LAUNCHERS');
    const r = run('bash', script);
    expect(r.code, r.stderr).toBe(0);
    return stage;
  }

  it('POSIX bundles ship an executable codegraph-vba launcher that runs the CLI', () => {
    const stage = stageBundle('linux');
    const launcher = path.join(stage, 'bin', 'codegraph-vba');
    expect(fs.statSync(launcher).mode & 0o111).not.toBe(0);

    const r = spawnSync(launcher, ['--version'], { encoding: 'utf8' });
    expect(r.status).toBe(0);
    expect(r.stdout).toContain(`${path.join(stage, 'lib/dist/bin/codegraph.js')} --version`);
  });

  it('Windows bundles ship codegraph-vba.cmd identical to codegraph.cmd', () => {
    const stage = stageBundle('win32');
    const cmd = fs.readFileSync(path.join(stage, 'bin', 'codegraph.cmd'), 'utf8');
    expect(fs.readFileSync(path.join(stage, 'bin', 'codegraph-vba.cmd'), 'utf8')).toBe(cmd);
  });

  function linkLaunchers(dest: string): { binDir: string; stdout: string } {
    const binDir = path.join(tmp, 'bin');
    const script =
      `set -eu\nBIN_DIR=${shq(binDir)}\ndest=${shq(dest)}\n` +
      extractBlock('install.sh', 'CODEGRAPH_LINK_LAUNCHERS');
    const r = run('sh', script);
    expect(r.code, r.stderr).toBe(0);
    return { binDir, stdout: r.stdout };
  }

  it('install.sh links both launchers onto PATH', () => {
    const dest = stageBundle('linux');
    const { binDir } = linkLaunchers(dest);
    expect(fs.readlinkSync(path.join(binDir, 'codegraph'))).toBe(path.join(dest, 'bin', 'codegraph'));
    expect(fs.readlinkSync(path.join(binDir, 'codegraph-vba'))).toBe(path.join(dest, 'bin', 'codegraph-vba'));
  });

  it('install.sh leaves no dangling codegraph-vba link for an older bundle that lacks it', () => {
    const dest = path.join(tmp, 'old-bundle');
    fs.mkdirSync(path.join(dest, 'bin'), { recursive: true });
    fs.writeFileSync(path.join(dest, 'bin', 'codegraph'), '#!/bin/sh\n', { mode: 0o755 });
    const { binDir } = linkLaunchers(dest);
    expect(fs.existsSync(path.join(binDir, 'codegraph'))).toBe(true);
    expect(fs.existsSync(path.join(binDir, 'codegraph-vba'))).toBe(false);
  });
});
