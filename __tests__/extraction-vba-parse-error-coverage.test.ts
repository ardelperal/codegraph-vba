/**
 * Issue #213 — both `severity: 'error'` catch-all paths in the VBA subsystem
 * were untested, AND a latent defect: when a throw escaped the walk, edges
 * that had been pushed with `source: ''` (held in
 * `pendingModuleOrClassSource` for re-attribution once the module/class node
 * existed) were left dangling — `finalize()` and the drain never ran because
 * the catch was outside the try and the throw bypassed both. The caller
 * received a partially-corrupt graph with no signal to reject it wholesale.
 *
 * Three behavioural contracts are pinned here:
 *
 *   1. **Catch records a `parse_error` and does NOT re-throw** — a malformed
 *      `.bas`/`.cls`/`.form.txt`/`.report.txt` must never crash a whole
 *      index run. The error envelope is `severity: 'error'`,
 *      `code: 'parse_error'`, and `filePath` matches the input.
 *
 *   2. **`nodes` and `edges` return a valid (possibly partial) result**
 *      rather than throwing out of `extract()`. `extract()` must always
 *      resolve with a structurally valid `ExtractionResult` so the caller
 *      can index the partial graph and surface the error.
 *
 *   3. **No edge in the returned `edges` carries `source: ''`** — the
 *      partial result must be a coherent graph fragment, never a dangling
 *      edge pointing nowhere. The empty-source placeholder exists only to
 *      be re-attributed BEFORE the result leaves `extract()`; a throw must
 *      not leak that internal placeholder.
 *
 * Test seam: the catch paths are reached by stubbing a collaborator the
 * walk calls (the cleanest such seam is `applyRaiseFanoutGate` in the
 * code extractor — it runs AFTER the walks and AFTER `finalize()`, so the
 * walks have already pushed `contains` / `references` edges with
 * `source: ''` into `pendingModuleOrClassSource` by the time the gate
 * throws). The form-extractor catch is reached by stubbing `walkBlocks`
 * directly on a single instance (private methods are reachable via
 * `(instance as unknown as Record<string, unknown>)` because TS `private`
 * is compile-time only).
 */
import { afterEach, describe, it, expect, vi } from 'vitest';
import { VbaExtractor } from '../src/extraction/vba-extractor';
import { VbaFormExtractor } from '../src/extraction/vba-form-extractor';

afterEach(() => {
  vi.restoreAllMocks();
});

describe('VbaExtractor — parse_error catch-all (Issue #213)', () => {
  it('records severity=error, code=parse_error, filePath and does not throw', () => {
    // Public Sub emits a function node AND a `contains` edge held in
    // pendingModuleOrClassSource with source:'' (see
    // `pushContainsFromModule`). The walk completes successfully — by
    // the time the gate runs, that placeholder edge is in the queue
    // waiting to be re-attributed.
    const src = [
      'Attribute VB_Name = "BoomMod"',
      'Public Sub DoStuff()',
      'End Sub',
    ].join('\n');
    const filePath = 'src/modules/BoomMod.bas';

    // Build the extractor, then stub the post-walk gate on its context.
    // The gate runs AFTER the walks and AFTER `finalize()` — by the
    // time it throws, the walks have already produced their `contains`
    // edges in `pendingModuleOrClassSource`.
    const ex = new VbaExtractor(filePath, src);
    const ctx = (
      ex as unknown as {
        ctx: { applyRaiseFanoutGate: (n: number) => void };
      }
    ).ctx;
    vi.spyOn(ctx, 'applyRaiseFanoutGate').mockImplementation(() => {
      throw new Error('forced throw from gate');
    });

    let result: ReturnType<VbaExtractor['extract']> | undefined;
    expect(() => {
      result = ex.extract();
    }).not.toThrow();

    expect(result).toBeDefined();
    expect(result!.errors).toHaveLength(1);
    expect(result!.errors[0]?.severity).toBe('error');
    expect(result!.errors[0]?.code).toBe('parse_error');
    expect(result!.errors[0]?.filePath).toBe(filePath);
    expect(result!.errors[0]?.message).toMatch(/forced throw from gate/);
  });

  it('never returns an edge with source: empty after a throw', () => {
    const src = [
      'Attribute VB_Name = "BoomMod2"',
      'Public Sub DoStuff()',
      'End Sub',
    ].join('\n');
    const filePath = 'src/modules/BoomMod2.bas';

    const ex = new VbaExtractor(filePath, src);
    const ctx = (
      ex as unknown as {
        ctx: { applyRaiseFanoutGate: (n: number) => void };
      }
    ).ctx;
    vi.spyOn(ctx, 'applyRaiseFanoutGate').mockImplementation(() => {
      throw new Error('forced throw for empty-source invariant');
    });

    const result = ex.extract();

    // Walk produced edges — at least the contains edge for DoStuff must
    // exist in pendingModuleOrClassSource. The catch must have either
    // dropped the dangling edges or drained them before returning.
    const dangling = result.edges.filter((e) => e.source === '');
    expect(
      dangling,
      `expected zero dangling edges after throw, got ${dangling.length}: ${JSON.stringify(dangling)}`,
    ).toEqual([]);

    // pendingModuleOrClassSource is the internal queue and should be
    // drained before return (the `moduleOrClassNode` was never created
    // because the gate threw before that step).
    expect(result.errors).toHaveLength(1);
    expect(result.errors[0]?.code).toBe('parse_error');
  });
});

describe('VbaFormExtractor — parse_error catch-all (Issue #213)', () => {
  it('records severity=error, code=parse_error, filePath and does not throw', () => {
    // A form with a real Begin/End block and a TextBox — enough to push
    // a few edges/nodes into `this` before the stubbed throw.
    const filePath = 'src/forms/Form_Boom.form.txt';
    const src = [
      'Begin Form',
      '    Caption = "Boom"',
      '    Begin TextBox',
      '        Name = "txtBoom"',
      '    End',
      'End',
    ].join('\n');

    const ex = new VbaFormExtractor(filePath, src);
    // Force walkBlocks to throw after the file node, layout node, and
    // unresolved-reference have been pushed (they live in extract()'s
    // try block, before the walkBlocks call).
    vi.spyOn(
      ex as unknown as { walkBlocks: (src: string, id: string) => void },
      'walkBlocks',
    ).mockImplementation(() => {
      throw new Error('forced throw from walkBlocks');
    });

    let result: ReturnType<VbaFormExtractor['extract']> | undefined;
    expect(() => {
      result = ex.extract();
    }).not.toThrow();

    expect(result).toBeDefined();
    expect(result!.errors).toHaveLength(1);
    expect(result!.errors[0]?.severity).toBe('error');
    expect(result!.errors[0]?.code).toBe('parse_error');
    expect(result!.errors[0]?.filePath).toBe(filePath);
    expect(result!.errors[0]?.message).toMatch(/forced throw from walkBlocks/);

    // The file node + form-layout node + form-binding unresolved
    // reference are pushed BEFORE walkBlocks runs in extract(), so they
    // survive the throw. The walk's emits (control nodes, RecordSource
    // edges) do NOT.
    const fileNode = result!.nodes.find((n) => n.kind === 'file');
    const layoutNode = result!.nodes.find(
      (n) => n.kind === 'form-layout' || n.kind === 'report-layout',
    );
    expect(fileNode).toBeDefined();
    expect(layoutNode).toBeDefined();
    expect(
      result!.unresolvedReferences.some(
        (u) => u.metadata?.synthesizedBy === 'vba-form-binding',
      ),
    ).toBe(true);
    // No control nodes from the (thrown) walk.
    expect(
      result!.nodes.some((n) => n.kind === 'form-instance-control'),
    ).toBe(false);
  });
});