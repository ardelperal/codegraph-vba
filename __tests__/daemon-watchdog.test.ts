import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { spawn } from 'child_process';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import {
  getRegistryDir,
  isProcessAlive,
  registerDaemon,
  deregisterDaemon,
  type DaemonRecord,
} from '../src/mcp/daemon-registry';
import { DaemonWatchdog } from '../src/mcp/daemon-watchdog';
import { decodeLockInfo, getDaemonPidPath } from '../src/mcp/daemon-paths';

/** A pid that's guaranteed dead: spawn a trivial process, let it exit, reap it. */
async function deadPid(): Promise<number> {
  const child = spawn(process.execPath, ['-e', 'process.exit(0)']);
  const pid = child.pid!;
  await new Promise<void>((r) => child.on('exit', () => r()));
  await new Promise((r) => setTimeout(r, 50)); // let the OS reap it
  return pid;
}

function rec(root: string, pid: number, startedAt = Date.now()): DaemonRecord {
  return { root, pid, version: '1.0.0', socketPath: `${root}/.codegraph/daemon.sock`, startedAt };
}

function fakeProject(): string {
  // Make a fake .codegraph-vba/ dir (note the suffix — codegraph-vba's
  // getCodeGraphDir() uses `.codegraph-vba`, NOT `.codegraph`) so
  // findNearestCodeGraphRoot() and getDaemonPidPath() both resolve correctly.
  // isInitialized() requires BOTH the dir AND a codegraph.db file, so create
  // an empty placeholder db so the project looks "initialized" to the daemon
  // logic.
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'cg-watchdog-'));
  fs.mkdirSync(path.join(root, '.codegraph-vba'), { recursive: true });
  fs.writeFileSync(path.join(root, '.codegraph-vba', 'codegraph.db'), '');
  return root;
}

describe('daemon-watchdog', () => {
  let tmpHome: string;
  let prevHome: string | undefined;
  let prevUserProfile: string | undefined;

  beforeEach(() => {
    tmpHome = fs.mkdtempSync(path.join(os.tmpdir(), 'cg-watchdog-home-'));
    prevHome = process.env.HOME;
    prevUserProfile = process.env.USERPROFILE;
    process.env.HOME = tmpHome;
    process.env.USERPROFILE = tmpHome;
    expect(getRegistryDir().startsWith(tmpHome)).toBe(true);
  });

  afterEach(() => {
    if (prevHome === undefined) delete process.env.HOME; else process.env.HOME = prevHome;
    if (prevUserProfile === undefined) delete process.env.USERPROFILE; else process.env.USERPROFILE = prevUserProfile;
    try { fs.rmSync(tmpHome, { recursive: true, force: true }); } catch { /* ignore */ }
  });

  describe('checkAndRespawn', () => {
    it('returns false when a live daemon already exists for the root', async () => {
      const root = fakeProject();
      // Write a live daemon's pidfile — the canonical pointer. Our own
      // process is alive, so isProcessAlive(process.pid) returns true and
      // checkAndRespawn() should skip the respawn.
      fs.writeFileSync(
        getDaemonPidPath(root),
        JSON.stringify({ pid: process.pid, version: '1.0.0', socketPath: '', startedAt: Date.now() }, null, 2)
      );

      const wd = new DaemonWatchdog();
      wd.watch(root);
      const spawned = await wd.checkAndRespawn(root);
      expect(spawned).toBe(false);
      fs.unlinkSync(getDaemonPidPath(root));
    });

    it('returns false (no-op) when the root has no .codegraph/ to bind a daemon under', async () => {
      const root = fs.mkdtempSync(path.join(os.tmpdir(), 'cg-no-cg-')); // no .codegraph/
      const wd = new DaemonWatchdog();
      wd.watch(root);
      const spawned = await wd.checkAndRespawn(root);
      expect(spawned).toBe(false);
    });

    it('returns true after tick() detects a dead daemon', async () => {
      const root = fakeProject();
      const dead = await deadPid();
      // Place a dead daemon's pidfile so hasLiveDaemon() returns false.
      fs.writeFileSync(
        getDaemonPidPath(root),
        JSON.stringify({ pid: dead, version: '1.0.0', socketPath: '', startedAt: 0 })
      );

      let spawnCalls = 0;
      const wd = new DaemonWatchdog({
        scriptPath: 'codegraph-vba',
        spawnFn: () => { spawnCalls++; return true; },
      });
      wd.watch(root);
      await wd.tick();
      expect(spawnCalls).toBe(1);
      fs.unlinkSync(getDaemonPidPath(root));
    });
  });

  describe('watchdog tick → respawn integration', () => {
    it('a dead daemon record is followed by another spawn on the next tick', async () => {
      const root = fakeProject();
      // Dead pidfile at the canonical path.
      fs.writeFileSync(
        getDaemonPidPath(root),
        JSON.stringify({ pid: await deadPid(), version: '1.0.0', socketPath: '', startedAt: 0 })
      );

      let spawnCalls = 0;
      const wd = new DaemonWatchdog({
        scriptPath: 'codegraph-vba',
        spawnFn: () => { spawnCalls++; return true; },
      });
      wd.watch(root);
      await wd.tick();
      await wd.tick(); // second tick: still dead (we don't simulate a real spawn replacing the pidfile), spawns again
      expect(spawnCalls).toBe(2);
      fs.unlinkSync(getDaemonPidPath(root));
    });

    it('start() begins the polling loop; stop() ends it', async () => {
      const root = fakeProject();
      fs.writeFileSync(
        getDaemonPidPath(root),
        JSON.stringify({ pid: await deadPid(), version: '1.0.0', socketPath: '', startedAt: 0 })
      );

      let spawnCalls = 0;
      const wd = new DaemonWatchdog({
        scriptPath: 'codegraph-vba',
        intervalMs: 5_000, // long enough that ONLY the explicit ticks fire
        spawnFn: () => { spawnCalls++; return true; },
      });
      wd.watch(root);
      wd.start();
      try {
        // `start()` fires one immediate tick (fire-and-forget). Give the
        // microtask queue a chance to drain before we sample.
        await new Promise((r) => setImmediate(r));
        expect(spawnCalls).toBe(1);
      } finally {
        wd.stop();
      }
      fs.unlinkSync(getDaemonPidPath(root));
    });

    it('stop() prevents further polls after the interval ticks', async () => {
      const root = fakeProject();
      fs.writeFileSync(
        getDaemonPidPath(root),
        JSON.stringify({ pid: await deadPid(), version: '1.0.0', socketPath: '', startedAt: 0 })
      );

      let spawnCalls = 0;
      const wd = new DaemonWatchdog({
        scriptPath: 'codegraph-vba',
        intervalMs: 20, // poll fast so we can sample after stop()
        spawnFn: () => { spawnCalls++; return true; },
      });
      wd.watch(root);
      wd.start();
      // Let the immediate tick + a couple of interval ticks happen.
      await new Promise((r) => setTimeout(r, 80));
      const beforeStop = spawnCalls;
      expect(beforeStop).toBeGreaterThanOrEqual(1);
      wd.stop();
      // After stop(), the interval is cleared — spawnCalls must not grow.
      await new Promise((r) => setTimeout(r, 80));
      expect(spawnCalls).toBe(beforeStop);
      fs.unlinkSync(getDaemonPidPath(root));
    });
  });

  describe('hasLiveDaemon', () => {
    it('returns false when no pid file exists', () => {
      const root = fakeProject();
      const wd = new DaemonWatchdog();
      expect(wd.hasLiveDaemon(root)).toBe(false);
    });

    it('returns true when the pid file points at our own process', () => {
      const root = fakeProject();
      const pidPath = getDaemonPidPath(root);
      const body = JSON.stringify(
        { pid: process.pid, version: '1.0.0', socketPath: '', startedAt: Date.now() },
        null,
        2
      );
      fs.writeFileSync(pidPath, body);
      const wd = new DaemonWatchdog();
      expect(wd.hasLiveDaemon(root)).toBe(true);
      fs.unlinkSync(pidPath);
    });

    it('returns false when the pid file points at a dead process', async () => {
      const root = fakeProject();
      const pidPath = getDaemonPidPath(root);
      fs.writeFileSync(
        pidPath,
        JSON.stringify({ pid: await deadPid(), version: '1.0.0', socketPath: '', startedAt: 0 })
      );
      const wd = new DaemonWatchdog();
      expect(wd.hasLiveDaemon(root)).toBe(false);
      fs.unlinkSync(pidPath);
    });
  });
});