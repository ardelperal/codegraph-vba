/**
 * Daemon liveness watchdog — issue #116.
 *
 * The MCP server (or any long-lived host) registers project roots it serves.
 * Every `intervalMs`, the watchdog checks each root's daemon (via the
 * {@link listDaemons} discovery index + liveness probe). If a daemon that
 * SHOULD exist is dead (kill -9, Stop-Process, crash, …), the watchdog
 * respawns it with the same spawnDetachedDaemon recipe the launcher uses
 * (`process.execPath` + `process.execArgv` + `scriptPath serve --mcp --path <root>`).
 *
 * Why a watchdog (vs. spawn-on-demand at MCP request time): the issue's repro
 * is `Stop-Process -Id <pid>` while the MCP keeps serving — file edits are
 * missed until the user manually runs `codegraph-vba sync`. The watchdog
 * closes that gap in <intervalMs>.
 *
 * Spawned detaches (own session/process group) so closing the MCP process
 * does not take the watchdog down with it.
 */

import { spawn } from 'child_process';
import * as fs from 'fs';
import * as path from 'path';
import { findNearestCodeGraphRoot, getCodeGraphDir } from '../directory';
import { decodeLockInfo, getDaemonPidPath, getDaemonSocketCandidates } from './daemon-paths';
import { isProcessAlive } from './daemon-registry';

const DAEMON_INTERNAL_ENV = 'CODEGRAPH_DAEMON_INTERNAL';
const HOST_PPID_ENV = 'CODEGRAPH_HOST_PPID';

const DEFAULT_INTERVAL_MS = 30_000;

/**
 * Default spawn implementation: spawn the daemon detached and unref so it
 * survives the launcher's exit. Extracted so tests can inject a stub.
 */
function defaultSpawnFn(
  nodePath: string,
  args: string[],
  opts: { detached: boolean; env: NodeJS.ProcessEnv; windowsHide?: boolean; stdio?: 'ignore' | [ 'ignore', number, number ] }
): boolean {
  const child = spawn(nodePath, args, {
    detached: opts.detached,
    stdio: opts.stdio,
    windowsHide: opts.windowsHide,
    env: opts.env,
  });
  child.unref();
  return true;
}

export interface DaemonWatchdogOptions {
  /** Poll interval (ms). Default 30s. */
  intervalMs?: number;
  /** Override the path to the codegraph binary used to spawn the daemon. */
  scriptPath?: string;
  /** Override node executable used to spawn the daemon. */
  nodePath?: string;
  /**
   * Override the spawn implementation. Used by tests; production callers
   * leave this unset to use the real `child_process.spawn`.
   */
  spawnFn?: (nodePath: string, args: string[], opts: { detached: boolean; env: NodeJS.ProcessEnv; windowsHide?: boolean; stdio?: 'ignore' | [ 'ignore', number, number ] }) => boolean;
}

/**
 * How long the watchdog waits for a freshly-spawned daemon to bind its socket
 * before declaring the spawn a failure (next poll retries). Same ~6s budget
 * the launcher uses for cold start.
 */
const SPAWN_BIND_MAX_RETRIES = 240;
const SPAWN_BIND_RETRY_DELAY_MS = 25;

/**
 * Per-project daemon liveness watchdog. Registers roots; polls every
 * `intervalMs`; respawns a daemon if the live one dies.
 */
export class DaemonWatchdog {
  private readonly roots = new Set<string>();
  private interval: NodeJS.Timeout | null = null;
  private readonly intervalMs: number;
  private readonly scriptPath: string;
  private readonly nodePath: string;
  private readonly spawnFn: (nodePath: string, args: string[], opts: { detached: boolean; env: NodeJS.ProcessEnv; windowsHide?: boolean; stdio?: 'ignore' | [ 'ignore', number, number ] }) => boolean;

  constructor(opts: DaemonWatchdogOptions = {}) {
    this.intervalMs = opts.intervalMs ?? DEFAULT_INTERVAL_MS;
    // Resolve the spawn target at construction time. `process.argv[1]` is the
    // script that the launcher used; reusing it keeps the detached daemon
    // launch identical to the launcher's spawn.
    this.scriptPath = opts.scriptPath ?? process.argv[1] ?? '';
    this.nodePath = opts.nodePath ?? process.execPath;
    this.spawnFn = opts.spawnFn ?? defaultSpawnFn;
  }

  /** Watch a project root: if its daemon dies, respawn it. Idempotent. */
  watch(root: string): void {
    this.roots.add(path.resolve(root));
  }

  /** Stop watching a root. Idempotent. */
  unwatch(root: string): void {
    this.roots.delete(path.resolve(root));
  }

  /** Number of roots currently being watched. */
  size(): number {
    return this.roots.size;
  }

  /**
   * Start the polling loop. No-op if already running. Detached from any caller;
   * safe to call from a request handler.
   */
  start(): void {
    if (this.interval) return;
    // Fire one check immediately so the user doesn't wait `intervalMs` for
    // the first round after restart.
    void this.tick();
    this.interval = setInterval(() => void this.tick(), this.intervalMs);
    // unref so the watchdog never keeps the event loop alive on its own.
    this.interval.unref?.();
  }

  /** Stop the polling loop. Idempotent. */
  stop(): void {
    if (this.interval) {
      clearInterval(this.interval);
      this.interval = null;
    }
  }

  /**
   * Run one round: for each watched root, check daemon liveness and respawn
   * if missing. Exposed for tests + manual triggers.
   */
  async tick(): Promise<void> {
    for (const root of this.roots) {
      await this.checkAndRespawn(root);
    }
  }

  /**
   * If no live daemon serves `root`, spawn one. Exposed for tests + manual
   * triggers. Resolves to `true` if a daemon was spawned, `false` if one
   * was already running or no `.codegraph/` is reachable from `root`.
   */
  async checkAndRespawn(root: string): Promise<boolean> {
    // We can't trust `listDaemons` to surface a dead daemon for us — it filters
    // out dead entries from its return value (even with `prune: false`, see
    // daemon-registry.ts). So we check the pidfile directly: it IS the
    // authoritative pointer (the registry is a discovery index, the lockfile
    // is the source of truth).
    if (this.hasLiveDaemon(root)) return false;
    // Only spawn if the project has a `.codegraph/` — no point spawning a
    // daemon for an uninitialized project.
    if (!findNearestCodeGraphRoot(root)) return false;
    return this.spawn(root);
  }

  /** Spawn a fresh daemon for `root`. Detached; safe to call from any context. */
  spawn(root: string): boolean {
    if (!this.scriptPath) return false;
    const logPath = path.join(getCodeGraphDir(root), 'daemon.log');
    let logFd: number | null = null;
    let stdio: 'ignore' | [ 'ignore', number, number ] = 'ignore';
    try {
      logFd = fs.openSync(logPath, 'a');
      stdio = [ 'ignore', logFd, logFd ];
    } catch {
      stdio = 'ignore';
    }
    try {
      // Don't leak the watchdog's host pid into the daemon (would trip its PPID
      // watchdog on this very process). Scrub it on spawn.
      const env: NodeJS.ProcessEnv = { ...process.env, [DAEMON_INTERNAL_ENV]: '1' };
      delete env[HOST_PPID_ENV];
      this.spawnFn(
        this.nodePath,
        [ ...process.execArgv, this.scriptPath, 'serve', '--mcp', '--path', root ],
        {
          detached: true,
          windowsHide: true,
          env,
          stdio,
        }
      );
      return true;
    } catch {
      return false;
    } finally {
      if (logFd !== null) {
        try { fs.closeSync(logFd); } catch { /* ignore */ }
      }
    }
  }

  /**
   * Wait for a daemon's socket to appear at one of the candidate paths under
   * `root`. Used by callers that want a synchronous spawn-then-use flow; not
   * required by the watchdog itself (which only needs liveness, not binding).
   */
  async waitForSocket(root: string, maxRetries = SPAWN_BIND_MAX_RETRIES): Promise<boolean> {
    const candidates = getDaemonSocketCandidates(root);
    for (let i = 0; i < maxRetries; i++) {
      for (const candidate of candidates) {
        try {
          // Any successful stat means the socket/pipe exists. The launcher
          // does a real connect on top of this; we only need existence.
          fs.statSync(candidate);
          return true;
        } catch {
          /* not yet */
        }
      }
      await new Promise<void>((r) => setTimeout(r, SPAWN_BIND_RETRY_DELAY_MS));
    }
    return false;
  }

  /** True if a daemon's pid file at `root` points at a live process. */
  hasLiveDaemon(root: string): boolean {
    let raw: string;
    try {
      raw = fs.readFileSync(getDaemonPidPath(root), 'utf8');
    } catch {
      return false;
    }
    const info = decodeLockInfo(raw);
    if (!info) return false;
    return isProcessAlive(info.pid);
  }
}