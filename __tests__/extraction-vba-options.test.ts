/**
 * Issue #243 — `VbaExtractionOptions`, the one object that replaces seven
 * positional VBA-extraction parameters across five files.
 *
 * This is a PURE REFACTOR, so every test here is an equivalence test: the new
 * object form must produce the SAME `ExtractionResult` as the old positional
 * form, and the knobs must keep gating exactly as they did.
 *
 * Acceptance criteria pinned here (from the issue body):
 *   - object form and legacy positional form produce deep-equal
 *     `ExtractionResult` (minus the wall-clock `durationMs`)
 *   - an option set on the object survives a REAL worker round-trip through
 *     `parse-pool` — i.e. it crosses `structuredClone` intact
 *   - `maxRaiseFanout` still gates as before (regression guard for #152)
 *
 * Real files, real extractor, a real `worker_threads.Worker`. No mocking.
 */
import { describe, it, expect } from 'vitest';
import { Worker } from 'node:worker_threads';
import { VbaExtractor, DEFAULT_MAX_RAISE_FANOUT } from '../src/extraction/vba-extractor';
import { extractFromSource } from '../src/extraction/tree-sitter';
import { ParseWorkerPool, type ParsePoolWorker } from '../src/extraction/parse-pool';
import type { VbaExtractionOptions } from '../src/extraction/vba/options';
import type { ExtractionResult, Language } from '../src/types';

const FILE = 'src/classes/OptionsProbe.cls';

/**
 * A module that exercises BOTH knobs at once:
 *   - `#If DEV` makes the result depend on `targets`
 *   - `RaiseEvent Changed` makes it depend on `maxRaiseFanout`
 */
const VBA_SRC = [
  'Attribute VB_Name = "OptionsProbe"',
  'Public Event Changed()',
  '',
  '#If DEV Then',
  'Public Sub DevOnly()',
  '    Helper',
  'End Sub',
  '#End If',
  '',
  'Public Sub DoWork()',
  '    RaiseEvent Changed',
  '    Helper',
  'End Sub',
  '',
  'Public Sub Helper()',
  'End Sub',
  '',
].join('\n');

/** Build a `.cls` with one event raised `n` times, one raise site per Sub. */
function buildClassWithRaisedEvent(eventName: string, n: number): string {
  const lines = ['Attribute VB_Name = "FanoutProbe"', `Public Event ${eventName}()`];
  for (let i = 0; i < n; i++) {
    lines.push(`Public Sub RaiseSite_${i}()`, `    RaiseEvent ${eventName}`, 'End Sub');
  }
  return lines.join('\n');
}

/**
 * `ExtractionResult.durationMs` and `Node.updatedAt` are wall-clock readings,
 * so they differ between two runs of identical input. They are the ONLY fields
 * an equivalence assertion may ignore. Everything else — every node field,
 * every edge, every unresolved reference, every error — must match.
 */
function comparable(r: ExtractionResult): unknown {
  const { durationMs: _durationMs, ...rest } = r;
  return {
    ...rest,
    nodes: rest.nodes.map(({ updatedAt: _updatedAt, ...node }) => node),
  };
}

const names = (r: ExtractionResult): string[] => r.nodes.map((n) => n.name).sort();
const raiseEdges = (r: ExtractionResult) => r.edges.filter((e) => e.kind === 'raises-event');

describe('Issue #243 — VbaExtractor options object vs the legacy positionals', () => {
  it('object form and legacy positional form produce deep-equal results', () => {
    const objectForm = new VbaExtractor(FILE, VBA_SRC, {
      targets: { DEV: true },
      maxRaiseFanout: 7,
    }).extract();
    // The `@deprecated` overload: targets 3rd, maxRaiseFanout 4th.
    const legacyForm = new VbaExtractor(FILE, VBA_SRC, { DEV: true }, 7).extract();

    expect(comparable(objectForm)).toEqual(comparable(legacyForm));
    // Guard the guard: if `targets` were silently dropped, both runs would be
    // equal AND wrong. `DevOnly` only survives preprocessing when DEV is true.
    expect(names(objectForm)).toContain('DevOnly');
  });

  it('targets threads through the object exactly as the positional did (DEV off)', () => {
    const objectForm = new VbaExtractor(FILE, VBA_SRC, { targets: { DEV: false } }).extract();
    const legacyForm = new VbaExtractor(FILE, VBA_SRC, { DEV: false }).extract();

    expect(comparable(objectForm)).toEqual(comparable(legacyForm));
    expect(names(objectForm)).not.toContain('DevOnly');
  });

  it('no options, empty options, and explicit-undefined positionals are all the same run', () => {
    const bare = new VbaExtractor(FILE, VBA_SRC).extract();
    const emptyObject = new VbaExtractor(FILE, VBA_SRC, {}).extract();
    const explicitUndefined = new VbaExtractor(FILE, VBA_SRC, undefined, undefined).extract();

    expect(comparable(emptyObject)).toEqual(comparable(bare));
    expect(comparable(explicitUndefined)).toEqual(comparable(bare));
    // The zero-config default leaves the DEV branch out.
    expect(names(bare)).not.toContain('DevOnly');
  });

  it('an absent maxRaiseFanout still falls back to DEFAULT_MAX_RAISE_FANOUT', () => {
    const src = buildClassWithRaisedEvent('Changed', DEFAULT_MAX_RAISE_FANOUT + 10);
    const viaObject = new VbaExtractor(FILE, src, { targets: undefined }).extract();
    const viaLegacy = new VbaExtractor(FILE, src, undefined, undefined).extract();

    expect(comparable(viaObject)).toEqual(comparable(viaLegacy));
    // Default 50 < 60 raise sites → the gate fires on both.
    expect(raiseEdges(viaObject)).toHaveLength(0);
  });
});

describe('Issue #243 — maxRaiseFanout still gates as before (regression guard for #152)', () => {
  it('over the threshold drops every raises-event edge and flags the node', () => {
    const src = buildClassWithRaisedEvent('Changed', 12);
    const r = new VbaExtractor(FILE, src, { maxRaiseFanout: 10 }).extract();

    const event = r.nodes.find((n) => n.kind === 'event' && n.name === 'Changed');
    expect(event?.metadata?.highFanout).toBe(true);
    expect(event?.metadata?.raiseCount).toBe(12);
    expect(raiseEdges(r)).toHaveLength(0);
  });

  it('at the threshold every edge survives (the gate is a strict `>`)', () => {
    const src = buildClassWithRaisedEvent('Changed', 10);
    const r = new VbaExtractor(FILE, src, { maxRaiseFanout: 10 }).extract();

    const event = r.nodes.find((n) => n.kind === 'event' && n.name === 'Changed');
    expect(event?.metadata?.highFanout).toBeUndefined();
    expect(raiseEdges(r)).toHaveLength(10);
  });

  it('the object form and the positional form gate identically', () => {
    const src = buildClassWithRaisedEvent('Changed', 12);
    const viaObject = new VbaExtractor(FILE, src, { maxRaiseFanout: 10 }).extract();
    const viaLegacy = new VbaExtractor(FILE, src, undefined, 10).extract();

    expect(comparable(viaObject)).toEqual(comparable(viaLegacy));
  });
});

describe('Issue #243 — extractFromSource merges the legacy positionals into vbaOptions', () => {
  it('trailing vbaOptions matches the legacy 5th/6th positionals', () => {
    const viaObject = extractFromSource(FILE, VBA_SRC, 'vba', undefined, undefined, undefined, true, {
      targets: { DEV: true },
      maxRaiseFanout: 7,
    });
    const viaLegacy = extractFromSource(FILE, VBA_SRC, 'vba', undefined, { DEV: true }, 7, true);

    expect(comparable(viaObject)).toEqual(comparable(viaLegacy));
    expect(names(viaObject)).toContain('DevOnly');
  });

  it('the object wins per field when both forms are supplied', () => {
    const src = buildClassWithRaisedEvent('Changed', 12);
    // Positional says "cap at 1000" (no gate); the object says "cap at 10".
    const merged = extractFromSource(FILE, src, 'vba', undefined, undefined, 1000, true, {
      maxRaiseFanout: 10,
    });
    const objectOnly = extractFromSource(FILE, src, 'vba', undefined, undefined, undefined, true, {
      maxRaiseFanout: 10,
    });
    const positionalOnly = extractFromSource(FILE, src, 'vba', undefined, undefined, 1000, true);

    expect(comparable(merged)).toEqual(comparable(objectOnly));
    expect(raiseEdges(merged)).toHaveLength(0);
    // Sanity: the positional value alone really would NOT have gated.
    expect(raiseEdges(positionalOnly)).toHaveLength(12);
  });

  it('a field absent from both forms falls through to the extractor default', () => {
    const viaNothing = extractFromSource(FILE, VBA_SRC, 'vba');
    const viaEmptyObject = extractFromSource(FILE, VBA_SRC, 'vba', undefined, undefined, undefined, true, {});

    expect(comparable(viaEmptyObject)).toEqual(comparable(viaNothing));
  });
});

describe('Issue #243 — VbaExtractionOptions is plain, structured-cloneable data', () => {
  it('survives structuredClone unchanged (no functions, RegExp, Map or class instances)', () => {
    const options: VbaExtractionOptions = {
      targets: { DEV: true, WIN64: false },
      maxRaiseFanout: 17,
      sqlWrappers: ['EjecutarSQL', 'AbrirRecordset'],
    };

    // This is the exact serialization `worker.postMessage` performs. A
    // function or a RegExp in the object would throw DataCloneError or lose
    // its identity here — which is why the compiled matcher lives on the
    // extractor context, not in this object.
    expect(structuredClone(options)).toEqual(options);
  });
});

/**
 * The REAL worker round-trip. A `worker_threads.Worker` speaks the same
 * {load-grammars → grammars-loaded} / {parse → parse-result} protocol the
 * production `parse-worker` does, and echoes the `vbaOptions` it received back
 * across the boundary. Both hops are genuine `structuredClone`s, so this fails
 * the moment a non-cloneable value is added to `VbaExtractionOptions`.
 *
 * It deliberately does NOT load tree-sitter grammars or a built `dist/`: the
 * claim under test is that the OPTIONS survive the boundary, not that the
 * extractor runs inside a worker (the extraction suite covers that).
 */
const ECHO_WORKER = `
const { parentPort } = require('node:worker_threads');
parentPort.on('message', (msg) => {
  if (msg.type === 'load-grammars') {
    parentPort.postMessage({ type: 'grammars-loaded' });
    return;
  }
  if (msg.type !== 'parse') return;
  parentPort.postMessage({
    type: 'parse-result',
    id: msg.id,
    result: {
      nodes: [],
      edges: [],
      unresolvedReferences: [],
      // The echo channel: the options exactly as they arrived in the worker.
      errors: [{ message: JSON.stringify(msg.vbaOptions ?? null) }],
      durationMs: 0,
    },
  });
});
`;

describe('Issue #243 — vbaOptions survives a real parse-pool worker round-trip', () => {
  it('arrives in the worker with every field intact', async () => {
    const pool = new ParseWorkerPool({
      languages: ['vba' as Language],
      size: 1,
      createWorker: () =>
        new Worker(ECHO_WORKER, { eval: true }) as unknown as ParsePoolWorker,
    });

    const vbaOptions: VbaExtractionOptions = {
      targets: { DEV: true, WIN64: false },
      maxRaiseFanout: 17,
      sqlWrappers: ['EjecutarSQL', 'AbrirRecordset'],
    };

    try {
      const result = await pool.requestParse({
        filePath: FILE,
        content: VBA_SRC,
        language: 'vba' as Language,
        vbaOptions,
      });

      const echoed = JSON.parse(result.errors[0]!.message) as VbaExtractionOptions;
      expect(echoed).toEqual(vbaOptions);
    } finally {
      await pool.destroy();
    }
  });

  it('an omitted vbaOptions arrives as undefined, not as a mangled object', async () => {
    const pool = new ParseWorkerPool({
      languages: ['vba' as Language],
      size: 1,
      createWorker: () =>
        new Worker(ECHO_WORKER, { eval: true }) as unknown as ParsePoolWorker,
    });

    try {
      const result = await pool.requestParse({
        filePath: FILE,
        content: VBA_SRC,
        language: 'vba' as Language,
      });

      expect(JSON.parse(result.errors[0]!.message)).toBeNull();
    } finally {
      await pool.destroy();
    }
  });
});
