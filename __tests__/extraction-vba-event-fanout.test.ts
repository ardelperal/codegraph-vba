/**
 * Issue #152 — per-file fanout gate for `RaiseEvent <EventName>` edges.
 *
 * Acceptance criteria (from the issue body):
 *   - [x] Form with `AfterUpdate` raised from 60 sites → 0 `raises-event`
 *         edges to that event node, and the node has
 *         `metadata.highFanout: true, metadata.raiseCount: 60`.
 *   - [x] Events with < 50 raise sites pass through unchanged.
 *   - [x] `codegraph_explore` on the high-fanout event still returns the
 *         event node, declaration site, and handler(s) — only the
 *         `raise_sites` list is condensed.
 *   - [x] `MAX_RAISE_FANOUT` is configurable via `codegraph.json` under
 *         the `vba` section.
 *
 * The default threshold is 50 (matches the issue spec) and is exposed as
 * `DEFAULT_MAX_RAISE_FANOUT` from `src/extraction/vba-extractor.ts`. The
 * test file builds minimal `.cls` / `.form.txt` fixtures and runs the
 * VbaExtractor directly — no SQLite, no MCP, no daemon.
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as os from 'node:os';
import { VbaExtractor, DEFAULT_MAX_RAISE_FANOUT } from '../src/extraction/vba-extractor';
import { loadVbaConfig, clearProjectConfigCache } from '../src/project-config';

function extract(filePath: string, source: string, maxRaiseFanout?: number) {
  return new VbaExtractor(filePath, source, undefined, maxRaiseFanout).extract();
}

/** Build a `.cls` with one event raised `n` times across N small Sub procs. */
function buildClassWithRaisedEvent(eventName: string, n: number): string {
  const lines: string[] = ['Attribute VB_Name = "BigForm"'];
  lines.push(`Public Event ${eventName}()`);
  for (let i = 0; i < n; i++) {
    lines.push(`Public Sub RaiseSite_${i}()`);
    lines.push(`    RaiseEvent ${eventName}`);
    lines.push('End Sub');
  }
  return lines.join('\n');
}

describe('Issue #152 — per-file RaiseEvent fanout gate (default threshold)', () => {
  it('DEFAULT_MAX_RAISE_FANOUT is 50 (per the issue spec)', () => {
    expect(DEFAULT_MAX_RAISE_FANOUT).toBe(50);
  });

  it('event with < 50 raise sites passes through unchanged', () => {
    // 30 sites is well under the default 50. Every edge must survive
    // the gate and the event node must NOT carry `metadata.highFanout`.
    const src = buildClassWithRaisedEvent('Changed', 30);
    const r = extract('src/classes/BigForm.cls', src);

    const event = r.nodes.find((n) => n.kind === 'event' && n.name === 'Changed');
    expect(event).toBeDefined();
    expect(event?.metadata?.highFanout).toBeUndefined();
    expect(event?.metadata?.raiseCount).toBeUndefined();

    const raiseEdges = r.edges.filter((e) => e.kind === 'raises-event');
    expect(raiseEdges).toHaveLength(30);
    for (const e of raiseEdges) {
      expect(e.target).toBe(event?.id);
      expect(e.metadata?.eventName).toBe('Changed');
    }
  });

  it('event with 50 raise sites passes through unchanged (threshold is strict `>`)', () => {
    // 50 equals the default; the gate fires when count STRICTLY EXCEEDS
    // the threshold, so 50 must NOT be flagged. The issue body phrases
    // it as "exceeds 50" → strict greater-than.
    const src = buildClassWithRaisedEvent('Changed', DEFAULT_MAX_RAISE_FANOUT);
    const r = extract('src/classes/BigForm.cls', src);

    const event = r.nodes.find((n) => n.kind === 'event' && n.name === 'Changed');
    expect(event).toBeDefined();
    expect(event?.metadata?.highFanout).toBeUndefined();
    const raiseEdges = r.edges.filter((e) => e.kind === 'raises-event');
    expect(raiseEdges).toHaveLength(DEFAULT_MAX_RAISE_FANOUT);
  });

  it('event with 60 raise sites is gated: 0 edges, metadata.highFanout: true, metadata.raiseCount: 60', () => {
    // The headline acceptance criterion: a form with `AfterUpdate` raised
    // from 60 sites → 0 `raises-event` edges to that node, the node
    // itself is stamped with the high-fanout marker + the original count.
    const src = buildClassWithRaisedEvent('AfterUpdate', 60);
    const r = extract('src/classes/BigForm.cls', src);

    const event = r.nodes.find((n) => n.kind === 'event' && n.name === 'AfterUpdate');
    expect(event).toBeDefined();
    // The event node itself MUST stay — the gate only drops edges.
    expect(event?.kind).toBe('event');
    expect(event?.metadata?.highFanout).toBe(true);
    expect(event?.metadata?.raiseCount).toBe(60);

    const raiseEdges = r.edges.filter((e) => e.kind === 'raises-event');
    expect(raiseEdges).toHaveLength(0);
  });

  it('mixed in the same file: under-threshold event keeps its edges, over-threshold event is gated', () => {
    // `Focused` raised 5 times (kept) and `AfterUpdate` raised 60 times
    // (gated) in the same .cls. The two events must NOT bleed into each
    // other — the gate is per-event.
    const lines: string[] = ['Attribute VB_Name = "MixedForm"'];
    lines.push('Public Event Focused()');
    lines.push('Public Event AfterUpdate()');
    for (let i = 0; i < 5; i++) {
      lines.push(`Public Sub FocusedSite_${i}()`);
      lines.push('    RaiseEvent Focused');
      lines.push('End Sub');
    }
    for (let i = 0; i < 60; i++) {
      lines.push(`Public Sub AfterUpdateSite_${i}()`);
      lines.push('    RaiseEvent AfterUpdate');
      lines.push('End Sub');
    }
    const r = extract('src/classes/MixedForm.cls', lines.join('\n'));

    const focused = r.nodes.find((n) => n.kind === 'event' && n.name === 'Focused');
    const afterUpdate = r.nodes.find((n) => n.kind === 'event' && n.name === 'AfterUpdate');
    expect(focused).toBeDefined();
    expect(afterUpdate).toBeDefined();

    // Focused (5 sites, under 50) — unchanged.
    expect(focused?.metadata?.highFanout).toBeUndefined();
    const focusedEdges = r.edges.filter(
      (e) => e.kind === 'raises-event' && e.target === focused?.id,
    );
    expect(focusedEdges).toHaveLength(5);

    // AfterUpdate (60 sites, over 50) — gated.
    expect(afterUpdate?.metadata?.highFanout).toBe(true);
    expect(afterUpdate?.metadata?.raiseCount).toBe(60);
    const afterUpdateEdges = r.edges.filter(
      (e) => e.kind === 'raises-event' && e.target === afterUpdate?.id,
    );
    expect(afterUpdateEdges).toHaveLength(0);
  });

  it('gated event node stays reachable: contains edge from module/class is preserved', () => {
    // `codegraph_explore` on the high-fanout event must still return the
    // event node + its declaration site. The gate must NOT orphan the
    // node — the module/class → event `contains` edge is a separate
    // edge kind and survives.
    const src = buildClassWithRaisedEvent('AfterUpdate', 60);
    const r = extract('src/classes/BigForm.cls', src);

    const event = r.nodes.find((n) => n.kind === 'event' && n.name === 'AfterUpdate');
    const moduleOrClass = r.nodes.find((n) => n.kind === 'class' && n.name === 'BigForm');
    expect(event).toBeDefined();
    expect(moduleOrClass).toBeDefined();

    const containsEdge = r.edges.find(
      (e) =>
        e.kind === 'contains' &&
        e.source === moduleOrClass?.id &&
        e.target === event?.id,
    );
    expect(containsEdge).toBeDefined();
  });
});

describe('Issue #152 — explicit threshold overrides the default', () => {
  it('passing maxRaiseFanout=10 gates a 15-site event and keeps a 5-site event', () => {
    // Low threshold (10) to make the test cheap — same shape as the 60-site
    // headline case but with smaller fixtures.
    const lines: string[] = ['Attribute VB_Name = "TunableForm"'];
    lines.push('Public Event Rare()');
    lines.push('Public Event Noisy()');
    for (let i = 0; i < 5; i++) {
      lines.push(`Public Sub RareSite_${i}()`);
      lines.push('    RaiseEvent Rare');
      lines.push('End Sub');
    }
    for (let i = 0; i < 15; i++) {
      lines.push(`Public Sub NoisySite_${i}()`);
      lines.push('    RaiseEvent Noisy');
      lines.push('End Sub');
    }
    const r = extract('src/classes/TunableForm.cls', lines.join('\n'), 10);

    const rare = r.nodes.find((n) => n.kind === 'event' && n.name === 'Rare');
    const noisy = r.nodes.find((n) => n.kind === 'event' && n.name === 'Noisy');

    // Rare (5 sites) — under 10, unchanged.
    expect(rare?.metadata?.highFanout).toBeUndefined();
    const rareEdges = r.edges.filter(
      (e) => e.kind === 'raises-event' && e.target === rare?.id,
    );
    expect(rareEdges).toHaveLength(5);

    // Noisy (15 sites) — over 10, gated.
    expect(noisy?.metadata?.highFanout).toBe(true);
    expect(noisy?.metadata?.raiseCount).toBe(15);
    const noisyEdges = r.edges.filter(
      (e) => e.kind === 'raises-event' && e.target === noisy?.id,
    );
    expect(noisyEdges).toHaveLength(0);
  });
});

describe('Issue #152 — codegraph.json vba.maxRaiseFanout is honored end-to-end', () => {
  let dir: string;

  beforeEach(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'cg-vba-fanout-'));
    clearProjectConfigCache();
  });

  afterEach(() => {
    clearProjectConfigCache();
    try {
      fs.rmSync(dir, { recursive: true, force: true });
    } catch {
      // ignore cleanup EPERM on Windows
    }
  });

  it('returns undefined when codegraph.json has no vba block (the default)', () => {
    expect(loadVbaConfig(dir).maxRaiseFanout).toBeUndefined();
  });

  it('loads a well-formed vba.maxRaiseFanout value', () => {
    fs.writeFileSync(
      path.join(dir, 'codegraph.json'),
      JSON.stringify({ vba: { maxRaiseFanout: 25 } }),
    );
    expect(loadVbaConfig(dir).maxRaiseFanout).toBe(25);
  });

  it('warns and ignores a non-numeric / negative vba.maxRaiseFanout value', () => {
    fs.writeFileSync(
      path.join(dir, 'codegraph.json'),
      JSON.stringify({ vba: { maxRaiseFanout: 'fifty' } }),
    );
    expect(loadVbaConfig(dir).maxRaiseFanout).toBeUndefined();

    clearProjectConfigCache();
    fs.writeFileSync(
      path.join(dir, 'codegraph.json'),
      JSON.stringify({ vba: { maxRaiseFanout: -5 } }),
    );
    expect(loadVbaConfig(dir).maxRaiseFanout).toBeUndefined();
  });

  it('coexists with vba.targets in the same codegraph.json block', () => {
    // The two knobs share the `vba` block but serve different purposes —
    // `targets` is for conditional-compilation, `maxRaiseFanout` is for
    // the per-file `RaiseEvent` gate. Both must load independently.
    fs.writeFileSync(
      path.join(dir, 'codegraph.json'),
      JSON.stringify({
        vba: {
          targets: { Win64: true, Mac: false },
          maxRaiseFanout: 75,
        },
      }),
    );
    const cfg = loadVbaConfig(dir);
    expect(cfg.targets).toEqual({ Win64: true, Mac: false });
    expect(cfg.maxRaiseFanout).toBe(75);
  });
});
