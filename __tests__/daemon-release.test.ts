import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { Command } from 'commander';
import * as fs from 'fs';
import * as net from 'net';
import * as os from 'os';
import * as path from 'path';
import { registerDaemonStopCommand, runDaemonStop } from '../src/bin/daemon-release';
import { DaemonWatchdog } from '../src/mcp/daemon-watchdog';
import { decodeLockInfo, getDaemonPidPath, getDaemonReleaseLeasePath, getDaemonReleaseRecoveryPath, getDaemonSocketCandidates } from '../src/mcp/daemon-paths';
import { clearStaleDaemonLock, Daemon, tryAcquireDaemonLock } from '../src/mcp/daemon';
import {
  releaseDaemonAt,
  requestDaemonRelease,
  deregisterDaemon,
  DAEMON_RELEASE_LEASE_TTL_MS,
  hasActiveDaemonReleaseLease,
  listDaemons,
  refreshDaemonReleaseLease,
  registerDaemon,
  tryAcquireDaemonReleaseLease,
  type ExpectedDaemon,
  type StopResult,
} from '../src/mcp/daemon-registry';
import { ProxyReleaseCoordinator, ProxyRequestBarrier } from '../src/mcp/proxy';
import { getStaticTools, ToolHandler } from '../src/mcp/tools';

function project(parent: string, name: string): string {
  const root = path.join(parent, name);
  fs.mkdirSync(path.join(root, '.codegraph-vba'), { recursive: true });
  fs.writeFileSync(path.join(root, '.codegraph-vba', 'codegraph.db'), '');
  return fs.realpathSync(root);
}

function metadata(root: string, socketPath = ''): ExpectedDaemon {
  return { root, pid: 424242, version: 'test', socketPath, startedAt: 123456 };
}

function writeMetadata(expected: ExpectedDaemon): void {
  fs.writeFileSync(getDaemonPidPath(expected.root), JSON.stringify(expected));
}

function socketPathFor(temp: string): string {
  return process.platform === 'win32'
    ? `\\\\.\\pipe\\cg-release-${process.pid}-${Date.now()}-${Math.random()}`
    : path.join(temp, `release-${Date.now()}-${Math.random()}.sock`);
}

async function protocolDaemon(expected: ExpectedDaemon): Promise<net.Server> {
  const server = net.createServer((socket) => {
    socket.write(JSON.stringify({ codegraph: expected.version, pid: expected.pid, socketPath: expected.socketPath, root: expected.root, startedAt: expected.startedAt, protocol: 1 }) + '\n');
    let input = '';
    socket.setEncoding('utf8');
    socket.on('data', (chunk: string) => {
      input += chunk;
      if (!input.includes('\n')) return;
      const control = JSON.parse(input.slice(0, input.indexOf('\n')));
      const accepted = control.codegraph_control === 1 && control.action === 'release' && control.root === expected.root && control.pid === expected.pid && control.startedAt === expected.startedAt;
      socket.end(JSON.stringify({ codegraph_control: 1, outcome: accepted ? 'releasing' : 'identity-mismatch' }) + '\n');
    });
  });
  await new Promise<void>((resolve, reject) => server.listen(expected.socketPath, resolve).once('error', reject));
  return server;
}

describe('project-scoped daemon release', () => {
  let temp: string;
  let previousTools: string | undefined;

  beforeEach(() => {
    temp = fs.mkdtempSync(path.join(os.tmpdir(), 'cg-release-'));
    previousTools = process.env.CODEGRAPH_MCP_TOOLS;
    delete process.env.CODEGRAPH_MCP_TOOLS;
  });

  it('blocks successor publication when release wins lifecycle arbitration', async () => {
    const root = project(temp, 'release-wins');
    const expected = metadata(root);
    writeMetadata(expected);
    let successor: ReturnType<typeof tryAcquireDaemonLock> | undefined;

    const result = await releaseDaemonAt(root, {
      isAlive: () => false,
      afterReleaseLeasePublished: () => { successor = tryAcquireDaemonLock(root); },
    });

    expect(successor?.kind).toBe('taken');
    expect(result.outcome).toBe('not-running');
  });

  it('makes release target the exact generation published when startup wins arbitration', async () => {
    const root = project(temp, 'startup-wins');
    let targeted: ExpectedDaemon | undefined;
    let releasePromise: Promise<StopResult> | undefined;

    const startup = tryAcquireDaemonLock(root, {
      afterLifecycleAcquired: () => {
        releasePromise = releaseDaemonAt(root, {
          isAlive: () => true,
          requestControl: async (expected) => { targeted = expected; return 'releasing'; },
          waitForDeath: async () => true,
        });
      },
    });
    expect(startup.kind).toBe('acquired');
    const result = await releasePromise!;

    expect(result.outcome).toBe('released');
    expect(targeted).toMatchObject(startup.kind === 'acquired' ? startup.info : {});
  });

  it('releases a normally bound daemon using one generation across pidfile, registry, hello, and control identity', async () => {
    const root = project(temp, 'normal-generation');
    const acquired = tryAcquireDaemonLock(root);
    expect(acquired.kind).toBe('acquired');
    if (acquired.kind !== 'acquired') return;
    const daemon = new Daemon(root, { generation: acquired.info, idleTimeoutMs: 0, maxIdleMs: 0, exit: () => {} });
    const started = await daemon.start();

    expect(started.lock).toEqual(acquired.info);
    expect(decodeLockInfo(fs.readFileSync(getDaemonPidPath(root), 'utf8'))).toEqual(acquired.info);
    expect(listDaemons({ prune: false }).find((record) => record.root === root)).toMatchObject(acquired.info);

    const released = await releaseDaemonAt(root, { isAlive: () => true, waitForDeath: async () => true });
    expect(released).toMatchObject({ outcome: 'released', pid: acquired.info.pid });
  });

  it.runIf(process.platform !== 'win32')('preserves the acquired generation when the socket relocates', async () => {
    const root = project(temp, 'relocated-generation');
    const acquired = tryAcquireDaemonLock(root);
    expect(acquired.kind).toBe('acquired');
    if (acquired.kind !== 'acquired') return;
    const candidates = getDaemonSocketCandidates(root);
    expect(candidates.length).toBeGreaterThan(1);
    const daemon = new Daemon(root, {
      generation: acquired.info,
      idleTimeoutMs: 0,
      maxIdleMs: 0,
      exit: () => {},
      listenSocket: (socketPath, onConnection) => {
        if (socketPath === candidates[0]) {
          const error = new Error('unsupported socket location') as NodeJS.ErrnoException;
          error.code = 'ENOTSUP';
          return Promise.reject(error);
        }
        return new Promise<net.Server>((resolve, reject) => {
          const server = net.createServer(onConnection);
          server.once('error', reject);
          server.listen(socketPath, () => resolve(server));
        });
      },
    });
    const started = await daemon.start();
    const relocated = { ...acquired.info, socketPath: candidates[1] };

    expect(started.lock).toEqual(relocated);
    expect(decodeLockInfo(fs.readFileSync(getDaemonPidPath(root), 'utf8'))).toEqual(relocated);
    expect(listDaemons({ prune: false }).find((record) => record.root === root)).toMatchObject(relocated);

    const released = await releaseDaemonAt(root, { isAlive: () => true, waitForDeath: async () => true });
    expect(released).toMatchObject({ outcome: 'released', pid: acquired.info.pid });
  });

  afterEach(() => {
    vi.useRealTimers();
    if (previousTools === undefined) delete process.env.CODEGRAPH_MCP_TOOLS;
    else process.env.CODEGRAPH_MCP_TOOLS = previousTools;
    fs.rmSync(temp, { recursive: true, force: true });
  });

  it('dispatches the CLI handler with the explicit root without reading dist', async () => {
    const root = project(temp, 'cli');
    let received = '';
    const output = await runDaemonStop(root, async (pathArg) => {
      received = pathArg;
      return { root, pid: null, outcome: 'no-daemon' };
    });
    expect(received).toBe(root);
    expect(JSON.parse(output)).toEqual({ root, pid: null, outcome: 'no-daemon' });
  });

  it('parses daemon stop --path through the same source registration used by the CLI', async () => {
    const root = project(temp, 'cli-parser');
    const program = new Command().exitOverride();
    const daemon = program.command('daemon');
    let received = '';
    let output = '';
    registerDaemonStopCommand(
      daemon,
      async (pathArg) => { received = pathArg; return JSON.stringify({ root: pathArg, outcome: 'no-daemon' }); },
      (value) => { output = value; },
    );
    await program.parseAsync(['node', 'codegraph', 'daemon', 'stop', '--path', root]);
    expect(received).toBe(root);
    expect(JSON.parse(output)).toEqual({ root, outcome: 'no-daemon' });
  });

  it('authenticates daemon root and generation over the existing socket handshake', async () => {
    const root = project(temp, 'protocol');
    const expected = metadata(root, socketPathFor(temp));
    const server = await protocolDaemon(expected);
    try {
      expect(await requestDaemonRelease(expected)).toBe('releasing');
      expect(await requestDaemonRelease({ ...expected, startedAt: expected.startedAt + 1 })).toBe('identity-mismatch');
    } finally {
      await new Promise<void>((resolve) => server.close(() => resolve()));
    }
  });

  it('never signals a live PID when socket identity cannot be proven', async () => {
    const root = project(temp, 'stale');
    const expected = metadata(root);
    writeMetadata(expected);
    const result = await releaseDaemonAt(root, {
      isAlive: () => true,
      requestControl: async () => 'identity-mismatch',
    });
    expect(result.outcome).toBe('identity-mismatch');
    expect(fs.existsSync(getDaemonPidPath(root))).toBe(true);
  });

  it('preserves Windows ownership artifacts when accepted termination is delayed past timeout', async () => {
    const root = project(temp, 'windows-failure');
    const expected = metadata(root);
    writeMetadata(expected);
    const result = await releaseDaemonAt(root, {
      isAlive: () => true,
      requestControl: async () => 'releasing',
      waitForDeath: async () => false,
    });
    expect(result).toMatchObject({ root, pid: expected.pid, outcome: 'termination-failed' });
    expect(fs.existsSync(getDaemonPidPath(root))).toBe(true);
    expect(fs.existsSync(getDaemonReleaseLeasePath(root))).toBe(false);
  });

  it('cleans only the selected root after daemon death is positively confirmed', async () => {
    const first = metadata(project(temp, 'first'));
    const second = metadata(project(temp, 'second'));
    writeMetadata(first);
    writeMetadata(second);
    const result = await releaseDaemonAt(first.root, {
      isAlive: () => true,
      requestControl: async () => 'releasing',
      waitForDeath: async () => true,
    });
    expect(result.outcome).toBe('released');
    expect(fs.existsSync(getDaemonPidPath(first.root))).toBe(false);
    expect(fs.existsSync(getDaemonPidPath(second.root))).toBe(true);
  });

  it('preserves successor ownership artifacts when a new daemon takes over before cleanup', async () => {
    const root = project(temp, 'takeover');
    const socketPath = path.join(temp, 'takeover.sock');
    const stopped = { ...metadata(root, socketPath), pid: 111, startedAt: 100 };
    const successor = { ...stopped, pid: process.pid, startedAt: 200 };
    writeMetadata(stopped);
    if (process.platform !== 'win32') fs.writeFileSync(socketPath, 'successor');
    try {
      const result = await releaseDaemonAt(root, {
        isAlive: () => true,
        requestControl: async () => 'releasing',
        waitForDeath: async () => {
          writeMetadata(successor);
          registerDaemon(successor);
          return true;
        },
      });
      expect(result.outcome).toBe('released');
      expect(JSON.parse(fs.readFileSync(getDaemonPidPath(root), 'utf8'))).toMatchObject({ pid: successor.pid, startedAt: successor.startedAt });
      if (process.platform !== 'win32') expect(fs.readFileSync(socketPath, 'utf8')).toBe('successor');
    } finally {
      deregisterDaemon(root);
    }
  });

  it('excludes successor acquisition between generation validation and artifact deletion', async () => {
    const root = project(temp, 'lease-adversary');
    const stopped = { ...metadata(root, path.join(temp, 'lease-adversary.sock')), pid: 2147483646, startedAt: 100 };
    writeMetadata(stopped);
    let attempted: ReturnType<typeof tryAcquireDaemonLock> | null = null;

    const result = await releaseDaemonAt(root, {
      isAlive: () => true,
      requestControl: async () => 'releasing',
      waitForDeath: async () => true,
      beforeArtifactDelete: () => {
        expect(fs.existsSync(getDaemonReleaseLeasePath(root))).toBe(true);
        expect(clearStaleDaemonLock(getDaemonPidPath(root), stopped.pid)).toBe(true);
        attempted = tryAcquireDaemonLock(root);
      },
    });

    expect(result.outcome).toBe('released');
    expect(attempted?.kind).toBe('taken');
    expect(fs.existsSync(getDaemonPidPath(root))).toBe(false);
    expect(fs.existsSync(getDaemonReleaseLeasePath(root))).toBe(false);

    const successor = tryAcquireDaemonLock(root);
    expect(successor.kind).toBe('acquired');
    fs.unlinkSync(getDaemonPidPath(root));
  });

  it('recovers an expired lease even when its owner PID was reused by a live process', () => {
    const root = project(temp, 'stale-release-lease');
    const now = Date.now();
    fs.writeFileSync(getDaemonReleaseLeasePath(root), JSON.stringify({
      token: 'dead-owner',
      ownerPid: process.pid,
      createdAt: 1,
      heartbeatAt: 1,
      expiresAt: now - 1,
      generation: { pid: 111, startedAt: 100, socketPath: path.join(temp, 'stale.sock') },
    }));

    expect(hasActiveDaemonReleaseLease(root, now)).toBe(false);
    const acquired = tryAcquireDaemonLock(root);
    expect(acquired.kind).toBe('acquired');
    expect(fs.existsSync(getDaemonReleaseLeasePath(root))).toBe(false);
    fs.unlinkSync(getDaemonPidPath(root));
  });

  it('keeps a non-expired lease active regardless of owner PID liveness', () => {
    const root = project(temp, 'active-release-lease');
    const now = Date.now();
    fs.writeFileSync(getDaemonReleaseLeasePath(root), JSON.stringify({
      token: 'active-owner',
      ownerPid: 2147483647,
      createdAt: now,
      heartbeatAt: now,
      expiresAt: now + DAEMON_RELEASE_LEASE_TTL_MS,
      generation: { pid: 111, startedAt: 100, socketPath: path.join(temp, 'active.sock') },
    }));

    expect(hasActiveDaemonReleaseLease(root, now)).toBe(true);
    expect(tryAcquireDaemonLock(root).kind).toBe('taken');
  });

  it('refreshes the active release lease before its expiry', async () => {
    vi.useFakeTimers();
    const root = project(temp, 'heartbeat-release-lease');
    const expected = { ...metadata(root), pid: 424239 };
    writeMetadata(expected);
    const started = Date.now();
    vi.setSystemTime(started);
    let confirmDeath!: (died: boolean) => void;
    const release = releaseDaemonAt(root, {
      isAlive: () => true,
      requestControl: async () => 'releasing',
      waitForDeath: async () => new Promise<boolean>((resolve) => { confirmDeath = resolve; }),
    });
    await vi.advanceTimersByTimeAsync(0);
    const initial = JSON.parse(fs.readFileSync(getDaemonReleaseLeasePath(root), 'utf8'));
    await vi.advanceTimersByTimeAsync(5_000);
    const refreshed = JSON.parse(fs.readFileSync(getDaemonReleaseLeasePath(root), 'utf8'));
    expect(refreshed.heartbeatAt).toBeGreaterThan(initial.heartbeatAt);
    expect(refreshed.expiresAt).toBeGreaterThan(initial.expiresAt);
    confirmDeath(true);
    expect((await release).outcome).toBe('released');
  });

  it('preserves a heartbeat renewal that races an expired-lease recovery read', () => {
    const root = project(temp, 'heartbeat-recovery-race');
    const expected = { ...metadata(root), pid: 424238 };
    const now = Date.now();
    const lease = tryAcquireDaemonReleaseLease(root, expected, {
      now: now - DAEMON_RELEASE_LEASE_TTL_MS - 1,
    });
    expect(lease).not.toBeNull();

    const active = hasActiveDaemonReleaseLease(root, now, {
      afterRecoveryRead: () => {
        expect(refreshDaemonReleaseLease(root, lease!, now)).toBe(true);
      },
    });

    expect(active).toBe(true);
    const renewed = JSON.parse(fs.readFileSync(getDaemonReleaseLeasePath(root), 'utf8'));
    expect(renewed.token).toBe(lease!.token);
    expect(renewed.expiresAt).toBe(now + DAEMON_RELEASE_LEASE_TTL_MS);
    expect(fs.existsSync(getDaemonReleaseRecoveryPath(root))).toBe(false);
    expect(tryAcquireDaemonLock(root).kind).toBe('taken');
    fs.unlinkSync(getDaemonReleaseLeasePath(root));
  });

  it('recovers a crash-stale recovery marker after bounded expiry', () => {
    const root = project(temp, 'stale-recovery-marker');
    const markerPath = getDaemonReleaseRecoveryPath(root);
    const now = Date.now();
    fs.writeFileSync(markerPath, JSON.stringify({
      token: 'active-recovery',
      createdAt: now,
      expiresAt: now + DAEMON_RELEASE_LEASE_TTL_MS,
    }));
    expect(hasActiveDaemonReleaseLease(root, now)).toBe(true);
    expect(tryAcquireDaemonLock(root).kind).toBe('taken');

    fs.writeFileSync(markerPath, JSON.stringify({
      token: 'crashed-recovery',
      createdAt: now - DAEMON_RELEASE_LEASE_TTL_MS - 1,
      expiresAt: now - 1,
    }));

    expect(hasActiveDaemonReleaseLease(root, now)).toBe(false);
    expect(fs.existsSync(markerPath)).toBe(false);
    const acquired = tryAcquireDaemonLock(root);
    expect(acquired.kind).toBe('acquired');
    fs.unlinkSync(getDaemonPidPath(root));
  });

  it('fails closed on a partial lease until its bounded expiry', () => {
    const root = project(temp, 'partial-release-lease');
    const leasePath = getDaemonReleaseLeasePath(root);
    const now = Date.now();
    fs.writeFileSync(leasePath, '{');
    expect(hasActiveDaemonReleaseLease(root, now)).toBe(true);
    expect(tryAcquireDaemonLock(root).kind).toBe('taken');
    const expired = new Date(now - DAEMON_RELEASE_LEASE_TTL_MS - 1_000);
    fs.utimesSync(leasePath, expired, expired);
    expect(hasActiveDaemonReleaseLease(root, now)).toBe(false);
  });

  it.each(['EPERM', 'ENOTSUP'])('falls back to exclusive create when hard links fail with %s', async (code) => {
    const root = project(temp, `lease-fallback-${code}`);
    const expected = { ...metadata(root, path.join(temp, `${code}.sock`)), pid: 424240 };
    writeMetadata(expected);
    const linkFailure = () => { throw Object.assign(new Error(code), { code }); };
    let competingLease: ReturnType<typeof tryAcquireDaemonReleaseLease> | undefined;
    let startup: ReturnType<typeof tryAcquireDaemonLock> | undefined;

    const result = await releaseDaemonAt(root, {
      isAlive: () => true,
      leaseLinkSync: linkFailure,
      requestControl: async () => {
        competingLease = tryAcquireDaemonReleaseLease(root, expected, { linkSync: linkFailure });
        startup = tryAcquireDaemonLock(root);
        return 'releasing';
      },
      waitForDeath: async () => true,
    });

    expect(competingLease).toBeNull();
    expect(startup?.kind).toBe('taken');
    expect(result.outcome).toBe('released');
    expect(fs.existsSync(getDaemonPidPath(root))).toBe(false);
    expect(fs.existsSync(getDaemonReleaseLeasePath(root))).toBe(false);
  });

  it('closes an overlapping watchdog tick/release race before spawn', async () => {
    const root = project(temp, 'watchdog');
    let openGate!: () => void;
    const gate = new Promise<void>((resolve) => { openGate = resolve; });
    let spawns = 0;
    const released: StopResult = { root, pid: null, outcome: 'no-daemon' };
    const watchdog = new DaemonWatchdog({
      scriptPath: 'codegraph',
      spawnFn: () => { spawns++; return true; },
      beforeSpawn: () => gate,
      releaseFn: async () => released,
    });
    watchdog.watch(root);
    const tick = watchdog.checkAndRespawn(root);
    const release = watchdog.release(root);
    openGate();
    await Promise.all([tick, release]);
    expect(watchdog.size()).toBe(0);
    expect(spawns).toBe(0);
  });

  it('uses one canonical watchdog key across a junction alias during tick/release', async () => {
    const root = project(temp, 'canonical-root');
    const alias = path.join(temp, 'canonical-alias');
    fs.symlinkSync(root, alias, 'junction');
    let openGate!: () => void;
    const gate = new Promise<void>((resolve) => { openGate = resolve; });
    let spawns = 0;
    const watchdog = new DaemonWatchdog({
      scriptPath: 'codegraph',
      spawnFn: () => { spawns++; return true; },
      beforeSpawn: () => gate,
      releaseFn: async (releasedRoot) => ({ root: releasedRoot, pid: null, outcome: 'no-daemon' }),
    });
    watchdog.watch(alias);
    const tick = watchdog.checkAndRespawn(alias);
    const release = watchdog.release(root);
    openGate();
    await Promise.all([tick, release]);
    expect(watchdog.size()).toBe(0);
    expect(spawns).toBe(0);
  });

  it('keeps a canonical release tombstone across watch aliases until explicit resume', async () => {
    const root = project(temp, 'watch-resume-root');
    const alias = path.join(temp, 'watch-resume-alias');
    fs.symlinkSync(root, alias, 'junction');
    let spawns = 0;
    const watchdog = new DaemonWatchdog({
      scriptPath: 'codegraph',
      spawnFn: () => { spawns++; return true; },
      releaseFn: async (releasedRoot) => ({ root: releasedRoot, pid: null, outcome: 'no-daemon' }),
    });
    watchdog.watch(root);
    await watchdog.release(root);
    watchdog.watch(alias);
    expect(watchdog.size()).toBe(0);
    expect(await watchdog.checkAndRespawn(alias)).toBe(false);
    expect(spawns).toBe(0);
    watchdog.resume(alias);
    expect(watchdog.size()).toBe(1);
    expect(await watchdog.checkAndRespawn(root)).toBe(true);
    expect(spawns).toBe(1);
  });

  it('exposes opt-in destructive MCP metadata and typed idempotency', async () => {
    const root = project(temp, 'mcp');
    expect((await new ToolHandler(null).execute('codegraph_release', { path: root })).isError).toBe(true);
    process.env.CODEGRAPH_MCP_TOOLS = 'explore,release';
    const definition = getStaticTools().find((tool) => tool.name === 'codegraph_release');
    expect(definition?.annotations).toMatchObject({ destructiveHint: true, idempotentHint: true });
    const result = await new ToolHandler(null).execute('codegraph_release', { path: root });
    expect(JSON.parse(result.content[0]!.text)).toEqual({ root, pid: null, outcome: 'no-daemon' });
  });

  it('makes proxy release a barrier, deduplicates concurrent release, and rejects later calls', async () => {
    const root = project(temp, 'proxy');
    const coordinator = new ProxyReleaseCoordinator(root);
    const writes: Record<string, unknown>[] = [];
    let resolveRelease!: (value: StopResult) => void;
    let calls = 0;
    const release = () => { calls++; return new Promise<StopResult>((resolve) => { resolveRelease = resolve; }); };
    const message = (id: number, name: string) => ({ jsonrpc: '2.0', id, method: 'tools/call', params: { name, arguments: { path: root } } });

    expect(coordinator.handle(message(1, 'codegraph_release'), true, release, (value) => writes.push(value as Record<string, unknown>))).toBe(true);
    expect(coordinator.getState()).toBe('releasing');
    expect(coordinator.handle(message(2, 'codegraph_release'), true, release, (value) => writes.push(value as Record<string, unknown>))).toBe(true);
    expect(coordinator.handle(message(3, 'codegraph_explore'), true, release, (value) => writes.push(value as Record<string, unknown>))).toBe(true);
    await new Promise((resolve) => setImmediate(resolve));
    expect(calls).toBe(1);
    resolveRelease({ root, pid: 7, outcome: 'released' });
    await new Promise((resolve) => setImmediate(resolve));
    expect(coordinator.getState()).toBe('released');
    const releaseResponses = writes.filter((value) => {
      const result = value.result as { content?: Array<{ text?: string }> } | undefined;
      const text = result?.content?.[0]?.text;
      if (!text || !text.startsWith('{')) return false;
      return JSON.parse(text).outcome === 'released';
    });
    expect(releaseResponses).toHaveLength(2);
    expect(JSON.stringify(writes)).toContain('no longer accepts project calls');
    expect(coordinator.handle(message(4, 'codegraph_explore'), true, release, (value) => writes.push(value as Record<string, unknown>))).toBe(true);
  });

  it('does not intercept proxy release when the tool is not opted in', () => {
    const root = project(temp, 'proxy-opt-in');
    const coordinator = new ProxyReleaseCoordinator(root);
    const handled = coordinator.handle(
      { jsonrpc: '2.0', id: 1, method: 'tools/call', params: { name: 'codegraph_release', arguments: { path: root } } },
      false,
      async () => ({ root, pid: null, outcome: 'no-daemon' }),
      () => { throw new Error('must not write'); },
    );
    expect(handled).toBe(false);
    expect(coordinator.getState()).toBe('active');
  });

  it('drains an unresolved earlier request before release and rejects later requests synchronously', async () => {
    const root = project(temp, 'proxy-drain');
    let finishEarlier!: () => void;
    const earlier = new Promise<void>((resolve) => { finishEarlier = resolve; });
    const coordinator = new ProxyReleaseCoordinator(root, () => earlier);
    const writes: Record<string, unknown>[] = [];
    let releaseCalls = 0;
    const release = async (): Promise<StopResult> => {
      releaseCalls++;
      return { root, pid: 9, outcome: 'released' };
    };
    const releaseMessage = { jsonrpc: '2.0', id: 1, method: 'tools/call', params: { name: 'codegraph_release', arguments: { path: root } } };
    const laterMessage = { jsonrpc: '2.0', id: 2, method: 'tools/call', params: { name: 'codegraph_explore', arguments: { query: 'x' } } };

    expect(coordinator.handle(releaseMessage, true, release, (value) => writes.push(value as Record<string, unknown>))).toBe(true);
    expect(coordinator.getState()).toBe('releasing');
    expect(coordinator.handle(laterMessage, true, release, (value) => writes.push(value as Record<string, unknown>))).toBe(true);
    expect(releaseCalls).toBe(0);
    expect(JSON.stringify(writes)).toContain('no longer accepts project calls');
    finishEarlier();
    await new Promise((resolve) => setImmediate(resolve));
    expect(releaseCalls).toBe(1);
    expect(coordinator.getState()).toBe('released');
  });

  it('drains a real pending-before-connect request after it transitions to inflight', async () => {
    const root = project(temp, 'proxy-pending');
    const barrier = new ProxyRequestBarrier();
    const coordinator = new ProxyReleaseCoordinator(root, () => barrier.drain());
    const earlier = JSON.stringify({ jsonrpc: '2.0', id: 41, method: 'tools/call', params: { name: 'codegraph_explore' } });
    barrier.buffer(earlier);
    const sent: string[] = [];
    const writes: Record<string, unknown>[] = [];
    let releaseCalls = 0;
    const release = async (): Promise<StopResult> => {
      releaseCalls++;
      return { root, pid: 9, outcome: 'released' };
    };
    const releaseMessage = { jsonrpc: '2.0', id: 42, method: 'tools/call', params: { name: 'codegraph_release', arguments: { path: root } } };
    const laterMessage = { jsonrpc: '2.0', id: 43, method: 'tools/call', params: { name: 'codegraph_explore', arguments: { query: 'later' } } };

    expect(coordinator.handle(releaseMessage, true, release, (value) => writes.push(value as Record<string, unknown>))).toBe(true);
    expect(coordinator.handle(laterMessage, true, release, (value) => writes.push(value as Record<string, unknown>))).toBe(true);
    await new Promise((resolve) => setImmediate(resolve));
    expect(releaseCalls).toBe(0);
    expect(barrier.pendingCount()).toBe(1);
    barrier.flush((line) => sent.push(line));
    expect(sent).toEqual([earlier]);
    expect(barrier.pendingCount()).toBe(0);
    expect(barrier.inflightCount()).toBe(1);
    await new Promise((resolve) => setImmediate(resolve));
    expect(releaseCalls).toBe(0);
    barrier.settle(41);
    await new Promise((resolve) => setImmediate(resolve));
    expect(releaseCalls).toBe(1);
    expect(JSON.stringify(writes)).toContain('no longer accepts project calls');
  });

  it('keys concurrent owned and foreign releases independently', async () => {
    const owned = project(temp, 'proxy-owned');
    const foreign = project(temp, 'proxy-foreign');
    const coordinator = new ProxyReleaseCoordinator(owned);
    const resolvers = new Map<string, (result: StopResult) => void>();
    const calls: string[] = [];
    const release = (root: string) => {
      calls.push(root);
      return new Promise<StopResult>((resolve) => resolvers.set(root, resolve));
    };
    const write = () => { /* response shape is covered by the owned release test */ };
    const message = (id: number, root: string) => ({ jsonrpc: '2.0', id, method: 'tools/call', params: { name: 'codegraph_release', arguments: { path: root } } });

    expect(coordinator.handle(message(1, owned), true, release, write)).toBe(true);
    expect(coordinator.handle(message(2, foreign), true, release, write)).toBe(true);
    await new Promise((resolve) => setImmediate(resolve));
    expect(new Set(calls)).toEqual(new Set([owned, foreign]));
    expect(coordinator.getState(owned)).toBe('releasing');
    expect(coordinator.getState(foreign)).toBe('releasing');
    resolvers.get(foreign)!({ root: foreign, pid: 2, outcome: 'released' });
    await new Promise((resolve) => setImmediate(resolve));
    expect(coordinator.getState(foreign)).toBe('released');
    expect(coordinator.getState(owned)).toBe('releasing');
    resolvers.get(owned)!({ root: owned, pid: 1, outcome: 'released' });
    await new Promise((resolve) => setImmediate(resolve));
    expect(coordinator.getState(owned)).toBe('released');
  });
});
