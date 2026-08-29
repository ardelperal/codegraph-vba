/**
 * Global daemon registry + stop/list control — the discovery layer behind
 * `codegraph list` and `codegraph stop [--all]`.
 *
 * Every per-project daemon already writes an authoritative lockfile at
 * `<root>/.codegraph/daemon.pid`. That's enough to stop ONE daemon you can name,
 * but there's no central place to find them ALL — which `list` and `stop --all`
 * need. So each daemon also drops a tiny record under `~/.codegraph/daemons/` on
 * start and removes it on graceful shutdown.
 *
 * The registry is a DISCOVERY index, never a source of truth: the live pid is.
 * A SIGKILL'd daemon can't remove its own record, so readers prune any record
 * whose pid is dead (`isProcessAlive`). Every write/read is best-effort — a
 * registry hiccup must never break the daemon or a command; worst case `list`
 * momentarily misses or over-lists one, which the next liveness prune corrects.
 *
 * Cross-platform by construction: only files + `process.kill(pid, signal)`,
 * which behave consistently on macOS/Linux (real signals) and Windows (mapped to
 * TerminateProcess). Validated live on all three.
 */
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import * as crypto from 'crypto';
import * as net from 'net';
import { getDaemonLifecyclePath, getDaemonPidPath, getDaemonReleaseLeasePath, getDaemonReleaseRecoveryPath, getDaemonSocketCandidates, decodeLockInfo } from './daemon-paths';

export interface DaemonRecord {
  /** Realpath'd project root the daemon serves. */
  root: string;
  pid: number;
  version: string;
  socketPath: string;
  /** Epoch ms when the daemon bound its socket. */
  startedAt: number;
}

/**
 * `~/.codegraph/daemons` — GLOBAL, keyed off the home install dir. (The
 * `CODEGRAPH_DIR` env var only renames the per-project index dir, not this.)
 */
export function getRegistryDir(): string {
  return path.join(os.homedir(), '.codegraph', 'daemons');
}

function recordPath(root: string): string {
  const hash = crypto.createHash('sha256').update(path.resolve(root)).digest('hex').slice(0, 16);
  return path.join(getRegistryDir(), `${hash}.json`);
}

/**
 * Is `pid` a live process? `kill(pid, 0)` sends no signal — it just probes:
 * ESRCH ⇒ dead, EPERM ⇒ alive but not ours (still alive). Same liveness check
 * the PPID watchdog (#277) and daemon lock arbitration use.
 */
export function isProcessAlive(pid: number): boolean {
  if (!Number.isInteger(pid) || pid <= 0) return false;
  try {
    process.kill(pid, 0);
    return true;
  } catch (err) {
    return (err as NodeJS.ErrnoException).code === 'EPERM';
  }
}

/** Best-effort: record this daemon so `list`/`stop --all` can find it. */
export function registerDaemon(rec: DaemonRecord): void {
  try {
    fs.mkdirSync(getRegistryDir(), { recursive: true });
    fs.writeFileSync(recordPath(rec.root), JSON.stringify(rec, null, 2) + '\n', { mode: 0o600 });
  } catch {
    /* best-effort — list's liveness prune tolerates a missing record */
  }
}

/** Best-effort: drop this daemon's record on graceful shutdown. */
export function deregisterDaemon(root: string): void {
  try {
    fs.unlinkSync(recordPath(root));
  } catch {
    /* already gone */
  }
}

/**
 * All registered daemons whose process is still alive, newest first. Dead/garbage
 * records are deleted as a side effect (self-healing) unless `prune` is false.
 */
export function listDaemons(opts: { prune?: boolean } = {}): DaemonRecord[] {
  const prune = opts.prune ?? true;
  const dir = getRegistryDir();
  let files: string[];
  try {
    files = fs.readdirSync(dir).filter((f) => f.endsWith('.json'));
  } catch {
    return []; // no registry dir yet
  }

  const live: DaemonRecord[] = [];
  for (const file of files) {
    const full = path.join(dir, file);
    let rec: DaemonRecord | null = null;
    try {
      rec = JSON.parse(fs.readFileSync(full, 'utf8')) as DaemonRecord;
    } catch {
      rec = null;
    }
    const valid = rec && typeof rec.pid === 'number' && typeof rec.root === 'string';
    if (valid && isProcessAlive(rec!.pid)) {
      live.push(rec!);
    } else if (prune) {
      try { fs.unlinkSync(full); } catch { /* ignore */ }
    }
  }
  return live.sort((a, b) => b.startedAt - a.startedAt);
}

function sameGeneration(actual: Partial<ExpectedDaemon> | null, expected: ExpectedDaemon): boolean {
  return Boolean(actual) && actual!.pid === expected.pid && actual!.startedAt === expected.startedAt &&
    actual!.socketPath === expected.socketPath;
}

/** Delete only artifacts that still belong to the stopped daemon generation. */
function cleanupDaemonArtifacts(root: string, expected: ExpectedDaemon, beforeDelete?: () => void): void {
  let lock: ReturnType<typeof decodeLockInfo> = null;
  let registry: DaemonRecord | null = null;
  try { lock = decodeLockInfo(fs.readFileSync(getDaemonPidPath(root), 'utf8')); } catch { /* gone */ }
  try { registry = JSON.parse(fs.readFileSync(recordPath(root), 'utf8')) as DaemonRecord; } catch { /* gone */ }

  const lockMatches = sameGeneration(lock, expected);
  const registryMatches = sameGeneration(registry, expected);
  const successorExists = Boolean(lock && !lockMatches) || Boolean(registry && !registryMatches);
  beforeDelete?.();

  if (lockMatches) {
    try { fs.unlinkSync(getDaemonPidPath(root)); } catch { /* raced */ }
  }
  if (registryMatches) {
    try { fs.unlinkSync(recordPath(root)); } catch { /* raced */ }
  }
  // A successor can reuse the stable socket path. Never unlink it when either
  // authoritative ownership record has already advanced to a new generation.
  if (!successorExists && process.platform !== 'win32' && expected.socketPath) {
    try { fs.unlinkSync(expected.socketPath); } catch { /* gone */ }
  }
}

const sleep = (ms: number): Promise<void> => new Promise((r) => setTimeout(r, ms));

async function waitForDeath(pid: number, timeoutMs: number): Promise<boolean> {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (!isProcessAlive(pid)) return true;
    await sleep(100);
  }
  return !isProcessAlive(pid);
}

export interface StopResult {
  root: string;
  pid: number | null;
  outcome: 'released' | 'not-running' | 'no-daemon' | 'identity-mismatch' | 'unreachable' | 'termination-failed';
  failure?: string;
}

export interface ExpectedDaemon {
  root: string;
  pid: number;
  version: string;
  socketPath: string;
  startedAt: number;
}

export interface DaemonReleaseDeps {
  isAlive?: (pid: number) => boolean;
  waitForDeath?: (pid: number, timeoutMs: number) => Promise<boolean>;
  requestControl?: (expected: ExpectedDaemon) => Promise<'releasing' | 'identity-mismatch' | 'unreachable'>;
  beforeArtifactDelete?: () => void;
  leaseLinkSync?: typeof fs.linkSync;
  afterLifecycleAcquired?: () => void;
  afterReleaseLeasePublished?: (expected: ExpectedDaemon) => void;
}

export const DAEMON_RELEASE_LEASE_TTL_MS = 30_000;
const DAEMON_RELEASE_HEARTBEAT_MS = 5_000;
const DAEMON_LIFECYCLE_TTL_MS = 30_000;

export interface DaemonLifecycleLock { token: string; ownerPid: number; createdAt: number; expiresAt: number }

function lifecycleArtifacts(root: string): string[] {
  const fixed = getDaemonLifecyclePath(root);
  const dir = path.dirname(fixed);
  const prefix = path.basename(fixed);
  try { return fs.readdirSync(dir).filter((name) => name === prefix || name.startsWith(`${prefix}.claim.`)).map((name) => path.join(dir, name)); }
  catch { return []; }
}

function cleanupExpiredLifecycleArtifacts(root: string, now: number): void {
  const fixed = getDaemonLifecyclePath(root);
  for (const artifact of lifecycleArtifacts(root)) {
    let expiresAt: number;
    try {
      const value = JSON.parse(fs.readFileSync(artifact, 'utf8')) as Partial<DaemonLifecycleLock>;
      expiresAt = typeof value.expiresAt === 'number' ? value.expiresAt : fs.statSync(artifact).mtimeMs + DAEMON_LIFECYCLE_TTL_MS;
    } catch {
      try { expiresAt = fs.statSync(artifact).mtimeMs + DAEMON_LIFECYCLE_TTL_MS; } catch { continue; }
    }
    if (expiresAt > now) continue;
    if (artifact === fixed) {
      const claim = `${fixed}.claim.${crypto.randomUUID()}`;
      try { fs.renameSync(fixed, claim); } catch { continue; }
      try { fs.unlinkSync(claim); } catch { /* gone */ }
    } else {
      try { fs.unlinkSync(artifact); } catch { /* gone */ }
    }
  }
}

/** Atomically acquire the shared startup/release publication arbiter. */
export function tryAcquireDaemonLifecycleLock(
  root: string,
  options: { now?: number; linkSync?: typeof fs.linkSync } = {},
): DaemonLifecycleLock | null {
  const now = options.now ?? Date.now();
  cleanupExpiredLifecycleArtifacts(root, now);
  if (lifecycleArtifacts(root).length !== 0) return null;
  const fixed = getDaemonLifecyclePath(root);
  const lock: DaemonLifecycleLock = { token: crypto.randomUUID(), ownerPid: process.pid, createdAt: now, expiresAt: now + DAEMON_LIFECYCLE_TTL_MS };
  const tmp = `${fixed}.${lock.token}.tmp`;
  try {
    fs.writeFileSync(tmp, JSON.stringify(lock), { mode: 0o600 });
    try { (options.linkSync ?? fs.linkSync)(tmp, fixed); } catch (err) {
      const code = (err as NodeJS.ErrnoException).code;
      if (code === 'EEXIST') return null;
      if (code !== 'EPERM' && code !== 'ENOTSUP' && code !== 'EXDEV') throw err;
      let fd: number;
      try { fd = fs.openSync(fixed, 'wx', 0o600); } catch (openErr) {
        if ((openErr as NodeJS.ErrnoException).code === 'EEXIST') return null;
        throw openErr;
      }
      try { fs.writeSync(fd, JSON.stringify(lock)); fs.fsyncSync(fd); } finally { fs.closeSync(fd); }
    }
  } finally {
    try { fs.unlinkSync(tmp); } catch { /* gone */ }
  }
  return lock;
}

export function releaseDaemonLifecycleLock(root: string, lock: DaemonLifecycleLock): void {
  const fixed = getDaemonLifecyclePath(root);
  try {
    const current = JSON.parse(fs.readFileSync(fixed, 'utf8')) as Partial<DaemonLifecycleLock>;
    if (current.token === lock.token) fs.unlinkSync(fixed);
  } catch { /* gone */ }
}

export interface DaemonReleaseLease {
  token: string;
  ownerPid: number;
  createdAt: number;
  heartbeatAt: number;
  expiresAt: number;
  generation: Pick<ExpectedDaemon, 'pid' | 'startedAt' | 'socketPath'>;
}

function decodeReleaseLease(raw: string): DaemonReleaseLease | null {
  try {
    const value = JSON.parse(raw) as Partial<DaemonReleaseLease>;
    if (typeof value.token !== 'string' || typeof value.ownerPid !== 'number' ||
        typeof value.createdAt !== 'number' || typeof value.heartbeatAt !== 'number' ||
        typeof value.expiresAt !== 'number' || !value.generation ||
        typeof value.generation.pid !== 'number' || typeof value.generation.startedAt !== 'number' ||
        typeof value.generation.socketPath !== 'string') return null;
    return value as DaemonReleaseLease;
  } catch { return null; }
}

interface RecoveryMarker { token: string; createdAt: number; expiresAt: number }

function recoveryArtifacts(root: string): string[] {
  const markerPath = getDaemonReleaseRecoveryPath(root);
  const dir = path.dirname(markerPath);
  const prefix = path.basename(markerPath);
  try { return fs.readdirSync(dir).filter((name) => name === prefix || name.startsWith(`${prefix}.claim.`)).map((name) => path.join(dir, name)); }
  catch { return []; }
}

/** Recover marker/claim files left by a crashed stale-recovery owner. */
function cleanupExpiredRecoveryArtifacts(root: string, now: number): void {
  const markerPath = getDaemonReleaseRecoveryPath(root);
  for (const artifact of recoveryArtifacts(root)) {
    let expiresAt = 0;
    try {
      const raw = JSON.parse(fs.readFileSync(artifact, 'utf8')) as Partial<RecoveryMarker & DaemonReleaseLease>;
      expiresAt = typeof raw.expiresAt === 'number' ? raw.expiresAt : fs.statSync(artifact).mtimeMs + DAEMON_RELEASE_LEASE_TTL_MS;
    } catch {
      try { expiresAt = fs.statSync(artifact).mtimeMs + DAEMON_RELEASE_LEASE_TTL_MS; } catch { continue; }
    }
    if (expiresAt > now) continue;
    if (artifact === markerPath) {
      const claim = `${markerPath}.claim.marker-${crypto.randomUUID()}`;
      try { fs.renameSync(markerPath, claim); } catch { continue; }
      try { fs.unlinkSync(claim); } catch { /* gone */ }
    } else {
      // Claims are immutable token-specific generations; deleting the expired
      // claim cannot target a newly-created fixed marker or lease.
      try { fs.unlinkSync(artifact); } catch { /* gone */ }
    }
  }
}

function acquireRecoveryMarker(root: string, now: number): RecoveryMarker | null {
  cleanupExpiredRecoveryArtifacts(root, now);
  if (recoveryArtifacts(root).length !== 0) return null;
  const markerPath = getDaemonReleaseRecoveryPath(root);
  const marker: RecoveryMarker = { token: crypto.randomUUID(), createdAt: now, expiresAt: now + DAEMON_RELEASE_LEASE_TTL_MS };
  try { fs.writeFileSync(markerPath, JSON.stringify(marker), { flag: 'wx', mode: 0o600 }); }
  catch { return null; }
  return marker;
}

function releaseRecoveryMarker(root: string, marker: RecoveryMarker): void {
  const markerPath = getDaemonReleaseRecoveryPath(root);
  try {
    const current = JSON.parse(fs.readFileSync(markerPath, 'utf8')) as Partial<RecoveryMarker>;
    if (current.token === marker.token) fs.unlinkSync(markerPath);
  } catch { /* gone or replaced */ }
}

/** Recover one expired immutable lease generation under an exclusive marker. */
function recoverStaleReleaseLease(root: string, now: number, afterRead?: () => void): boolean {
  const leasePath = getDaemonReleaseLeasePath(root);
  let current: DaemonReleaseLease | null = null;
  let modifiedAt = now;
  try {
    current = decodeReleaseLease(fs.readFileSync(leasePath, 'utf8'));
    modifiedAt = fs.statSync(leasePath).mtimeMs;
  } catch { return true; }
  // Fallback exclusive-create can be observed while its complete body is being
  // written. Fail closed for one full TTL; a crashed partial write then expires.
  const expired = current ? current.expiresAt <= now : modifiedAt + DAEMON_RELEASE_LEASE_TTL_MS <= now;
  if (!expired) return false;
  afterRead?.();

  const marker = acquireRecoveryMarker(root, now);
  if (!marker) return false;
  const claimPath = `${getDaemonReleaseRecoveryPath(root)}.claim.lease-${current?.token ?? crypto.randomUUID()}`;
  try {
    try { fs.renameSync(leasePath, claimPath); } catch { return true; }
    let claimed: DaemonReleaseLease | null = null;
    let claimExpiry = now + DAEMON_RELEASE_LEASE_TTL_MS;
    try {
      claimed = decodeReleaseLease(fs.readFileSync(claimPath, 'utf8'));
      claimExpiry = claimed?.expiresAt ?? fs.statSync(claimPath).mtimeMs + DAEMON_RELEASE_LEASE_TTL_MS;
    } catch { /* fail closed below */ }

    if (claimExpiry > now) {
      // Heartbeat won after our preliminary expired read. Keep its renewed
      // generation authoritative; the marker blocks startup during restoration.
      if (fs.existsSync(leasePath)) {
        const fixed = decodeReleaseLease(fs.readFileSync(leasePath, 'utf8'));
        if (fixed && fixed.token === claimed?.token && fixed.expiresAt >= claimExpiry) {
          try { fs.unlinkSync(claimPath); } catch { /* gone */ }
          return false;
        }
        return false;
      }
      fs.renameSync(claimPath, leasePath);
      return false;
    }
    try { fs.unlinkSync(claimPath); } catch { /* gone */ }
    return true;
  } finally {
    releaseRecoveryMarker(root, marker);
  }
}

/** True while intentional release exclusively owns this root's lifecycle. */
export function hasActiveDaemonReleaseLease(
  root: string,
  now: number = Date.now(),
  hooks: { afterRecoveryRead?: () => void } = {},
): boolean {
  const leasePath = getDaemonReleaseLeasePath(root);
  cleanupExpiredRecoveryArtifacts(root, now);
  if (recoveryArtifacts(root).length !== 0) return true;
  if (!fs.existsSync(leasePath)) return false;
  return !recoverStaleReleaseLease(root, now, hooks.afterRecoveryRead);
}

/** Atomically claim release/cleanup ownership for one daemon generation. */
export function tryAcquireDaemonReleaseLease(
  root: string,
  expected: ExpectedDaemon,
  options: { now?: number; linkSync?: typeof fs.linkSync } = {},
): DaemonReleaseLease | null {
  const now = options.now ?? Date.now();
  if (hasActiveDaemonReleaseLease(root, now)) return null;
  const leasePath = getDaemonReleaseLeasePath(root);
  const lease: DaemonReleaseLease = {
    token: crypto.randomUUID(),
    ownerPid: process.pid,
    createdAt: now,
    heartbeatAt: now,
    expiresAt: now + DAEMON_RELEASE_LEASE_TTL_MS,
    generation: { pid: expected.pid, startedAt: expected.startedAt, socketPath: expected.socketPath },
  };
  // Publish a fully-written immutable record in one atomic exclusive link, the
  // same primitive used by daemon.pid. No observer can see an empty lease.
  const tmpPath = `${leasePath}.${lease.token}.tmp`;
  try {
    fs.writeFileSync(tmpPath, JSON.stringify(lease, null, 2) + '\n', { mode: 0o600 });
    try { (options.linkSync ?? fs.linkSync)(tmpPath, leasePath); } catch (err) {
      if ((err as NodeJS.ErrnoException).code === 'EEXIST') return null;
      // Mirror daemon.pid's portable fallback for filesystems without hard
      // links. O_EXCL keeps acquisition exclusive; readers fail closed while
      // the complete body is written through the owning descriptor.
      let fd: number;
      try { fd = fs.openSync(leasePath, 'wx', 0o600); } catch (openErr) {
        if ((openErr as NodeJS.ErrnoException).code === 'EEXIST') return null;
        throw openErr;
      }
      try {
        fs.writeSync(fd, JSON.stringify(lease, null, 2) + '\n');
        fs.fsyncSync(fd);
      } finally {
        fs.closeSync(fd);
      }
    }
  } finally {
    try { fs.unlinkSync(tmpPath); } catch { /* gone */ }
  }
  return lease;
}

/** Publish a complete renewed generation atomically under the recovery marker. */
export function refreshDaemonReleaseLease(root: string, lease: DaemonReleaseLease, now: number = Date.now()): boolean {
  const leasePath = getDaemonReleaseLeasePath(root);
  const marker = acquireRecoveryMarker(root, now);
  if (!marker) return false;
  const tmpPath = `${leasePath}.${lease.token}.heartbeat.tmp`;
  try {
    const current = decodeReleaseLease(fs.readFileSync(leasePath, 'utf8'));
    if (current?.token !== lease.token) return false;
    lease.heartbeatAt = now;
    lease.expiresAt = now + DAEMON_RELEASE_LEASE_TTL_MS;
    fs.writeFileSync(tmpPath, JSON.stringify(lease, null, 2) + '\n', { mode: 0o600 });
    fs.renameSync(tmpPath, leasePath);
    return true;
  } finally {
    try { fs.unlinkSync(tmpPath); } catch { /* published or gone */ }
    releaseRecoveryMarker(root, marker);
  }
}

function releaseDaemonReleaseLease(root: string, lease: DaemonReleaseLease): void {
  const leasePath = getDaemonReleaseLeasePath(root);
  let current: DaemonReleaseLease | null = null;
  try { current = decodeReleaseLease(fs.readFileSync(leasePath, 'utf8')); } catch { return; }
  if (current?.token !== lease.token) return;
  try { fs.unlinkSync(leasePath); } catch { /* gone */ }
}

/** Resolve an explicit project root to the same canonical form daemons use. */
export function canonicalDaemonRoot(root: string): string {
  return fs.realpathSync(path.resolve(root));
}

/** Stable map/set key for aliases of the same canonical root. */
export function canonicalDaemonRootKey(root: string): string {
  const canonical = canonicalDaemonRoot(root);
  return process.platform === 'win32' ? canonical.toLowerCase() : canonical;
}

/** Project-scoped, idempotent release contract shared by CLI and MCP. */
export async function releaseDaemonAt(root: string, deps: DaemonReleaseDeps = {}): Promise<StopResult> {
  return stopDaemonAt(canonicalDaemonRoot(root), deps);
}

/**
 * Release the daemon serving `root` through its authenticated socket handshake,
 * wait for confirmed process death, then sweep ownership artifacts. Never sends
 * a signal based only on pidfile/registry liveness. `root` must be realpath'd.
 */
export async function stopDaemonAt(root: string, deps: DaemonReleaseDeps = {}): Promise<StopResult> {
  let lifecycle: DaemonLifecycleLock | null = null;
  for (let attempt = 0; attempt < 100 && !lifecycle; attempt++) {
    lifecycle = tryAcquireDaemonLifecycleLock(root);
    if (!lifecycle) await sleep(10);
  }
  if (!lifecycle) {
    return { root, pid: null, outcome: 'unreachable', failure: 'Project lifecycle arbitration is busy.' };
  }
  let expected: ExpectedDaemon | null = null;
  let lease: DaemonReleaseLease | null = null;
  try {
    deps.afterLifecycleAcquired?.();
    try {
      const info = decodeLockInfo(fs.readFileSync(getDaemonPidPath(root), 'utf8'));
      if (info) expected = { root, ...info };
    } catch { /* no lockfile */ }
    if (!expected) {
      const rec = listDaemons({ prune: false }).find((r) => path.resolve(r.root) === path.resolve(root));
      if (rec) expected = { ...rec, root };
    }
    if (!expected) return { root, pid: null, outcome: 'no-daemon' };
    lease = tryAcquireDaemonReleaseLease(root, expected, { linkSync: deps.leaseLinkSync });
    if (!lease) return { root, pid: expected.pid, outcome: 'unreachable', failure: 'A project release is already in progress.' };
    deps.afterReleaseLeasePublished?.(expected);
  } finally {
    releaseDaemonLifecycleLock(root, lifecycle);
  }

  const alive = deps.isAlive ?? isProcessAlive;
  const heartbeat = setInterval(() => refreshDaemonReleaseLease(root, lease), DAEMON_RELEASE_HEARTBEAT_MS);
  heartbeat.unref?.();
  try {
    if (!alive(expected.pid)) {
      cleanupDaemonArtifacts(root, expected, deps.beforeArtifactDelete);
      return { root, pid: expected.pid, outcome: 'not-running' };
    }

    const control = await (deps.requestControl ?? requestDaemonRelease)(expected);
    if (control !== 'releasing') {
      return {
        root,
        pid: expected.pid,
        outcome: control,
        failure: control === 'identity-mismatch'
          ? 'The reachable process did not prove ownership of this canonical root and daemon generation.'
          : 'The daemon control socket could not be reached; no process was signalled.',
      };
    }
    const died = await (deps.waitForDeath ?? waitForDeath)(expected.pid, 5000);
    if (!died) {
      return {
        root,
        pid: expected.pid,
        outcome: 'termination-failed',
        failure: 'The daemon accepted release but did not terminate before the timeout; ownership artifacts were preserved.',
      };
    }
    cleanupDaemonArtifacts(root, expected, deps.beforeArtifactDelete);
    return { root, pid: expected.pid, outcome: 'released' };
  } finally {
    clearInterval(heartbeat);
    releaseDaemonReleaseLease(root, lease);
  }
}

export async function requestDaemonRelease(
  expected: ExpectedDaemon,
): Promise<'releasing' | 'identity-mismatch' | 'unreachable'> {
  const candidates = [expected.socketPath, ...getDaemonSocketCandidates(expected.root)]
    .filter((value, index, all) => Boolean(value) && all.indexOf(value) === index);
  for (const socketPath of candidates) {
    const result = await requestReleaseAtSocket(socketPath, expected);
    if (result !== 'unreachable') return result;
  }
  return 'unreachable';
}

function requestReleaseAtSocket(
  socketPath: string,
  expected: ExpectedDaemon,
): Promise<'releasing' | 'identity-mismatch' | 'unreachable'> {
  return new Promise((resolve) => {
    const socket = net.createConnection(socketPath);
    socket.setEncoding('utf8');
    let buffer = '';
    let phase: 'hello' | 'response' = 'hello';
    let settled = false;
    const finish = (result: 'releasing' | 'identity-mismatch' | 'unreachable') => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      socket.destroy();
      resolve(result);
    };
    const timer = setTimeout(() => finish('unreachable'), 1500);
    timer.unref?.();
    socket.on('error', () => finish('unreachable'));
    socket.on('close', () => finish('unreachable'));
    socket.on('data', (chunk: string) => {
      buffer += chunk;
      let newline: number;
      while ((newline = buffer.indexOf('\n')) !== -1) {
        const line = buffer.slice(0, newline);
        buffer = buffer.slice(newline + 1);
        let value: Record<string, unknown>;
        try { value = JSON.parse(line) as Record<string, unknown>; } catch { finish('identity-mismatch'); return; }
        if (phase === 'hello') {
          const verified =
            value.protocol === 1 && value.pid === expected.pid &&
            value.codegraph === expected.version && value.root === expected.root &&
            value.startedAt === expected.startedAt && value.socketPath === expected.socketPath;
          if (!verified) { finish('identity-mismatch'); return; }
          phase = 'response';
          socket.write(JSON.stringify({
            codegraph_control: 1,
            action: 'release',
            root: expected.root,
            pid: expected.pid,
            startedAt: expected.startedAt,
          }) + '\n');
        } else {
          finish(value.codegraph_control === 1 && value.outcome === 'releasing' ? 'releasing' : 'identity-mismatch');
          return;
        }
      }
    });
  });
}

/** Stop every registered, live daemon. */
export async function stopAllDaemons(): Promise<StopResult[]> {
  const results: StopResult[] = [];
  for (const rec of listDaemons()) {
    results.push(await stopDaemonAt(rec.root));
  }
  return results;
}
