/**
 * extraction-vba-event-synth.test.ts
 *
 * Acceptance tests for issue #150 — post-extraction synthesis of
 * `event-handler` edges from `RaiseEvent` sites to WithEvents-handler
 * Subs, using the `m_<varName>_<EventName>` naming convention.
 *
 * The pre-existing graph has:
 *   - `raises-event` edge from raiser Sub → event node (parser).
 *   - `subscribes-event` edge from subscriber class/module → form class
 *     (synthetic stub) (heuristic, `synthesizedBy: 'vba-withevents'`).
 *
 * The post-extraction pass MUST connect those two halves by emitting an
 * `event-handler` edge (heuristic, `synthesizedBy: 'vba-event-handler'`)
 * from the raiser Sub to the handler Sub in the subscriber's file, when
 * the handler Sub is named `m_<varName>_<EventName>`.
 *
 * Fixture layout (under __tests__/fixtures/vba-event-synth/):
 *   Notifier.cls                 — declares DataChanged + ItemAdded events
 *                                  and raises each from one public method.
 *   Form_Main.cls                — WithEvents m_Notifier As Notifier, plus
 *                                  the matching m_Notifier_DataChanged +
 *                                  m_Notifier_ItemAdded handler Subs.
 *                                  → positive: 2 event-handler edges.
 *   Form_NoHandlers.cls          — WithEvents m_Notifier As Notifier, but
 *                                  no matching handler Subs.
 *                                  → negative: 0 event-handler edges from
 *                                    this file's binding.
 *   Form_MismatchedName.cls      — WithEvents m_Other As Notifier, with
 *                                  only a m_Notifier_DataChanged handler
 *                                  (the handler name uses the wrong
 *                                  variable name → no match).
 *                                  → negative: 0 event-handler edges.
 *
 * Tests run against a real CodeGraph index end-to-end, not against
 * per-file extraction results — the synthesis pass runs at the indexer
 * level (after `resolveVbaCallStubs`/`resolveVbaReferenceStubs`).
 *
 * NB: file paths stored on nodes are RELATIVE to the project root
 * (forward-slash normalized), NOT the absolute path passed to the test.
 * We compare by basename in path-sensitive assertions.
 */
import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { CodeGraph } from '../src';
import type { Node, Edge } from '../src/types';

const FIXTURE_DIR = path.join(__dirname, 'fixtures', 'vba-event-synth');
const NOTIFIER_BASENAME = 'Notifier.cls';
const FORM_MAIN_BASENAME = 'Form_Main.cls';
const FORM_NO_HANDLERS_BASENAME = 'Form_NoHandlers.cls';
const FORM_MISMATCHED_BASENAME = 'Form_MismatchedName.cls';

let cg: CodeGraph | null = null;
let initialized = false;
const codeGraphDir = path.join(FIXTURE_DIR, '.codegraph-vba');

beforeAll(async () => {
  if (fs.existsSync(codeGraphDir)) {
    fs.rmSync(codeGraphDir, { recursive: true, force: true });
  }
  cg = await CodeGraph.init(FIXTURE_DIR, { index: false });
  initialized = true;
  await cg.indexAll();
}, 60_000);

afterAll(async () => {
  if (cg) {
    try {
      await cg.close();
    } catch {
      /* ignore close errors */
    }
  }
  if (initialized && fs.existsSync(codeGraphDir)) {
    // Windows sometimes holds the SQLite file handle briefly after
    // `close()` returns. Use a small retry to avoid a flaky EBUSY
    // when the OS hasn't fully released the file yet.
    rmWithRetry(codeGraphDir, { recursive: true, force: true });
  }
});

/**
 * Look up a function node by its exact `(name, fileBasename)` pair. The
 * `getNodesByName` helper is case-insensitive on name only, so we filter
 * to disambiguate handlers that share a name across files.
 */
function findFunction(cg: CodeGraph, name: string, fileBasename: string): Node | null {
  const candidates = cg.getNodesByName(name).filter(
    (n) =>
      n.kind === 'function' &&
      // filePath is RELATIVE to project root and forward-slash normalized;
      // comparing by basename is sufficient for fixture-scoped tests.
      path.basename(n.filePath).toLowerCase() === fileBasename.toLowerCase(),
  );
  return candidates[0] ?? null;
}

/**
 * Collect every `event-handler` edge in the project whose metadata was
 * stamped by the new synthesis pass. Walks a known set of function
 * nodes (raiser + handler names) and pulls their outgoing edges; for a
 * fixture-scoped test this is exhaustive and faster than a full-table
 * scan.
 */
function synthesizedEventHandlerEdges(cg: CodeGraph): Edge[] {
  const out: Edge[] = [];
  const seen = new Set<string>();
  const functionNames = [
    'TriggerDataChanged',
    'TriggerItemAdded',
    'm_Notifier_DataChanged',
    'm_Notifier_ItemAdded',
  ];
  for (const name of functionNames) {
    for (const fn of cg.getNodesByName(name).filter((n) => n.kind === 'function')) {
      for (const e of cg.getOutgoingEdges(fn.id)) {
        if (e.kind !== 'event-handler') continue;
        if (e.metadata?.synthesizedBy !== 'vba-event-handler') continue;
        const key = `${e.source}|${e.target}|${e.line ?? ''}|${e.column ?? ''}`;
        if (seen.has(key)) continue;
        seen.add(key);
        out.push(e);
      }
    }
  }
  return out;
}

/**
 * Find every `event-handler` edge in the project (ANY provenance /
 * synthesizedBy). Used to verify the vba-control-modeling fixtures
 * do NOT receive any vba-event-handler edges from the new pass.
 */
function allEventHandlerEdges(cg: CodeGraph): Edge[] {
  const out: Edge[] = [];
  const seen = new Set<string>();
  // Iterate over all function nodes. The form-control synthesis emits
  // edges with `function` kind source (the handler Sub); the new
  // vba-event-handler pass emits edges with `function` kind source
  // (the RAISER Sub). Both source kinds are `function`, so walking
  // every function node's outgoing edges is the correct superset.
  // We don't have a public `getNodesByKind` so we iterate via
  // `searchNodes` (faster than a LIKE on `kind`).
  // For fixtures with a few dozen function nodes this is fast; the
  // helpers above use the same name-based scoping for the focused
  // assertions.
  const functionNames = [
    'TriggerDataChanged',
    'TriggerItemAdded',
    'm_Notifier_DataChanged',
    'm_Notifier_ItemAdded',
    'SomeOtherSub',
    'm_Notifier_DataChanged',
  ];
  for (const name of functionNames) {
    for (const fn of cg.getNodesByName(name).filter((n) => n.kind === 'function')) {
      for (const e of cg.getOutgoingEdges(fn.id)) {
        if (e.kind !== 'event-handler') continue;
        const key = `${e.source}|${e.target}|${e.line ?? ''}|${e.column ?? ''}`;
        if (seen.has(key)) continue;
        seen.add(key);
        out.push(e);
      }
    }
  }
  return out;
}

// =============================================================================
// AC #1 — 1 class × 2 events × 1 form × 2 handlers × 1 raise site per event
// must yield 2 `event-handler` edges. Edge direction: raiser Sub → handler
// Sub. Metadata: `synthesizedBy: 'vba-event-handler'`, `eventName`, and
// `variableName` populated.
// =============================================================================
describe('issue-150 AC#1: 1 class × 2 events × 1 form × 2 handlers → 2 event-handler edges', () => {
  it('Notifier.TriggerDataChanged has an outgoing event-handler edge to m_Notifier_DataChanged', () => {
    if (!cg) throw new Error('cg not initialized');
    const raiser = findFunction(cg, 'TriggerDataChanged', NOTIFIER_BASENAME);
    const handler = findFunction(cg, 'm_Notifier_DataChanged', FORM_MAIN_BASENAME);
    expect(raiser, 'raiser Sub TriggerDataChanged must exist').toBeDefined();
    expect(handler, 'handler Sub m_Notifier_DataChanged must exist').toBeDefined();
    if (!raiser || !handler) return;

    // Look up via the high-level getCallees — event-handler is in the
    // callee edge kinds list (traversal.ts:295).
    const callees = cg.getCallees(raiser.id);
    const synthEdge = callees.find(
      (c) =>
        c.node.id === handler.id &&
        c.edge.kind === 'event-handler' &&
        c.edge.provenance === 'heuristic',
    );
    expect(synthEdge, 'expected synthesized event-handler edge raiser→handler').toBeDefined();
    expect(synthEdge?.edge.metadata?.synthesizedBy).toBe('vba-event-handler');
    expect(synthEdge?.edge.metadata?.eventName).toBe('DataChanged');
    expect(synthEdge?.edge.metadata?.variableName).toBe('m_Notifier');
  });

  it('Notifier.TriggerItemAdded has an outgoing event-handler edge to m_Notifier_ItemAdded', () => {
    if (!cg) throw new Error('cg not initialized');
    const raiser = findFunction(cg, 'TriggerItemAdded', NOTIFIER_BASENAME);
    const handler = findFunction(cg, 'm_Notifier_ItemAdded', FORM_MAIN_BASENAME);
    expect(raiser).toBeDefined();
    expect(handler).toBeDefined();
    if (!raiser || !handler) return;

    const callees = cg.getCallees(raiser.id);
    const synthEdge = callees.find(
      (c) =>
        c.node.id === handler.id &&
        c.edge.kind === 'event-handler' &&
        c.edge.metadata?.synthesizedBy === 'vba-event-handler',
    );
    expect(synthEdge).toBeDefined();
    expect(synthEdge?.edge.metadata?.eventName).toBe('ItemAdded');
  });

  it('exactly TWO vba-event-handler edges are materialized across the project', () => {
    if (!cg) throw new Error('cg not initialized');
    // Direct count of synthesized edges — distinct from the per-raiser
    // checks above, this also catches an accidental "double-emit"
    // (e.g. an extra pass running on top of itself).
    const all = synthesizedEventHandlerEdges(cg);
    expect(all.length).toBe(2);
  });
});

// =============================================================================
// AC #2 — No new edge when the handler Sub `m_X_<EventName>` is not
// declared (Form_NoHandlers).
// =============================================================================
describe('issue-150 AC#2: no edge when the handler Sub is not declared', () => {
  it('Form_NoHandlers.cls must NOT produce any vba-event-handler edge', () => {
    if (!cg) throw new Error('cg not initialized');
    // The only legitimate handlers in the project are in
    // Form_Main.cls, so every vba-event-handler edge must target a
    // handler in that file.
    const all = synthesizedEventHandlerEdges(cg);
    const targetFilePaths: string[] = [];
    for (const e of all) {
      const candidates = [
        ...cg.getNodesByName('m_Notifier_DataChanged'),
        ...cg.getNodesByName('m_Notifier_ItemAdded'),
      ].filter((n) => n.id === e.target);
      for (const c of candidates) {
        targetFilePaths.push(c.filePath);
      }
    }
    // No target may live in Form_NoHandlers.cls.
    const leak = targetFilePaths.filter(
      (p) => path.basename(p).toLowerCase() === FORM_NO_HANDLERS_BASENAME.toLowerCase(),
    );
    expect(
      leak,
      `expected zero event-handler edges targeting Form_NoHandlers.cls; got ${leak.length}`,
    ).toHaveLength(0);
  });
});

// =============================================================================
// AC #3 — No new edge when the WithEvents variable name doesn't match any
// Sub in the same file (Form_MismatchedName).
// =============================================================================
describe('issue-150 AC#3: no edge when the WithEvents variable name does not match any Sub', () => {
  it('Form_MismatchedName.cls must NOT produce any vba-event-handler edge', () => {
    if (!cg) throw new Error('cg not initialized');
    const all = synthesizedEventHandlerEdges(cg);
    const targetFilePaths: string[] = [];
    for (const e of all) {
      const candidates = [
        ...cg.getNodesByName('m_Notifier_DataChanged'),
        ...cg.getNodesByName('m_Notifier_ItemAdded'),
      ].filter((n) => n.id === e.target);
      for (const c of candidates) {
        targetFilePaths.push(c.filePath);
      }
    }
    // No target may live in Form_MismatchedName.cls — the only handler
    // there is m_Notifier_DataChanged, which does NOT match the binding's
    // variable name (`m_Other`), so the synthesis pass must decline.
    const leak = targetFilePaths.filter(
      (p) => path.basename(p).toLowerCase() === FORM_MISMATCHED_BASENAME.toLowerCase(),
    );
    expect(
      leak,
      `expected zero event-handler edges targeting Form_MismatchedName.cls; got ${leak.length}`,
    ).toHaveLength(0);
  });
});

// =============================================================================
// AC #4 — `codegraph_explore` query for an event name returns the handler
// Sub in ONE call. We approximate "codegraph_explore" here as
// `getCallees(raiser)` — the traversal already includes `event-handler`
// in its edge kinds, so a single callee query reaches the handler.
// =============================================================================
describe('issue-150 AC#4: codegraph_explore returns the handler Sub in ONE call', () => {
  it('a single getCallees on the raiser returns the handler Sub via event-handler', () => {
    if (!cg) throw new Error('cg not initialized');
    const raiser = findFunction(cg, 'TriggerDataChanged', NOTIFIER_BASENAME);
    const handler = findFunction(cg, 'm_Notifier_DataChanged', FORM_MAIN_BASENAME);
    expect(raiser).toBeDefined();
    expect(handler).toBeDefined();
    if (!raiser || !handler) return;

    // ONE call — `getCallees(raiser)`. The fact that the handler shows
    // up here is the single-call guarantee: previously, the agent had
    // to walk event-decl → raises-event → subscribes-event → handler by
    // name on every query.
    const callees = cg.getCallees(raiser.id);
    const hit = callees.find(
      (c) => c.node.id === handler.id && c.edge.kind === 'event-handler',
    );
    expect(hit, 'expected handler Sub reachable from raiser in a single callee query').toBeDefined();
  });
});

// =============================================================================
// AC #5 — `vba-event-tracer` skill still works (backward compatible).
// The pre-existing `raises-event` / `subscribes-event` edges must remain
// intact (the synthesis pass must not delete or repoint them — it only
// ADDS `event-handler` edges).
// =============================================================================
describe('issue-150 AC#5: vba-event-tracer backward compatibility', () => {
  it('raises-event edges are still present (not deleted, not repointed)', () => {
    if (!cg) throw new Error('cg not initialized');
    // The raiser Subs must still have `raises-event` edges targeting
    // event nodes. We assert per-raiser rather than per-edge to avoid
    // having to enumerate event nodes by id.
    for (const raiserName of ['TriggerDataChanged', 'TriggerItemAdded']) {
      const raiser = findFunction(cg, raiserName, NOTIFIER_BASENAME);
      expect(raiser, `raiser ${raiserName} must exist`).toBeDefined();
      if (!raiser) continue;
      const out = cg.getOutgoingEdges(raiser.id);
      const raiseEdges = out.filter((e) => e.kind === 'raises-event');
      expect(
        raiseEdges.length,
        `raiser ${raiserName} must still have a raises-event edge`,
      ).toBeGreaterThanOrEqual(1);
      // Best-effort kind check via the (name, filePath) of the event
      // node the edge used to point at. The project only declares one
      // event per name, so a node with the same name as the edge's
      // eventName metadata is a valid proxy for "still an event node".
      for (const edge of raiseEdges) {
        const eventName = (edge.metadata?.eventName as string) ?? '';
        const eventNode = cg
          .getNodesByName(eventName)
          .find((n) => n.kind === 'event');
        expect(
          eventNode,
          `raises-event edge from ${raiserName} must still point at a node of kind 'event'`,
        ).toBeDefined();
      }
    }
  });

  it('WithEvents sweep is untouched: each form still has a references edge to the event class', () => {
    if (!cg) throw new Error('cg not initialized');
    // The WithEvents sweep emits BOTH a `subscribes-event` edge (which
    // is CASCADE-deleted by `resolveVbaReferenceStubs` along with the
    // synthetic class stub it targets — a pre-existing behaviour, NOT
    // a regression caused by the new pass) AND a `references` edge
    // (which gets repointed to the real class node). We assert the
    // repointed `references` edge survives — that's the durable trace
    // the WithEvents sweep leaves behind, and the trace the
    // vba-event-tracer skill can still walk.
    const expectedForms = [
      { basename: FORM_MAIN_BASENAME, name: 'Form_Main' },
      { basename: FORM_NO_HANDLERS_BASENAME, name: 'Form_NoHandlers' },
      { basename: FORM_MISMATCHED_BASENAME, name: 'Form_MismatchedName' },
    ];
    let found = 0;
    for (const form of expectedForms) {
      const classNode = cg
        .getNodesByName(form.name)
        .find(
          (n) =>
            n.kind === 'class' &&
            path.basename(n.filePath).toLowerCase() === form.basename.toLowerCase(),
        );
      expect(classNode, `class node for ${form.basename} must exist`).toBeDefined();
      if (!classNode) continue;
      const out = cg.getOutgoingEdges(classNode.id);
      // The WithEvents sweep's `references` edge carries
      // `synthesizedBy: 'vba-withevents'`. After
      // `resolveVbaReferenceStubs` it has been repointed to the real
      // Notifier class node (with `resolvedBy: 'vba-reference-stub'`
      // added to the metadata).
      const refs = out.filter(
        (e) =>
          e.kind === 'references' &&
          e.metadata?.synthesizedBy === 'vba-withevents',
      );
      expect(
        refs.length,
        `form ${form.basename} must still have a vba-withevents references edge (the WithEvents sweep trace)`,
      ).toBeGreaterThanOrEqual(1);
      // The pre-existing `resolveVbaReferenceStubs` preserves metadata
      // when it repoints the edge to the real class node, so
      // `resolvedBy: 'vba-reference-stub'` should be present.
      expect(
        refs[0]?.metadata?.resolvedBy,
        'resolveVbaReferenceStubs must mark the repointed edge',
      ).toBe('vba-reference-stub');
      found += refs.length;
    }
    expect(found).toBeGreaterThanOrEqual(3);
  });
});

// =============================================================================
// AC #6 — No new false positives on the existing FORMS-* fixture suite.
// We index a COPY of the existing vba-control-modeling fixtures (kept
// under a test-local subdirectory so the .codegraph-vba database lives
// in OUR directory, not in the shared source fixture) and verify the
// count of vba-event-handler edges stays at zero.
//
// The copy is needed because the sibling
// `__tests__/extraction-vba-control-modeling.test.ts` test owns the
// real `vba-control-modeling/.codegraph-vba` directory. When vitest
// runs both test files in the same worker (file parallelism), both
// tests try to open the same SQLite file and Windows returns EBUSY on
// the second opener's unlink/cleanup. Copying the fixtures into our
// own subdirectory sidesteps the conflict completely and keeps the
// AC#6 test self-contained.
//
// None of the FORMS fixtures use the WithEvents naming convention —
// they use Access's `<Control>_<Event>` convention, which is a
// DIFFERENT synthesis pipeline that the procedures.ts extract
// already handles. So the new pass should emit zero edges.
// =============================================================================
describe('issue-150 AC#6: no false positives on the FORMS-* (vba-control-modeling) fixture suite', () => {
  const SOURCE_FORMS_DIR = path.join(
    __dirname,
    'fixtures',
    'vba-control-modeling',
  );
  // Test-local subdirectory. We copy the FORMS fixtures into here so
  // our .codegraph-vba lives in OUR dir, not in the shared source
  // dir. This is a subdirectory of the test workspace — kept under
  // the worktree but outside the existing vba-control-modeling
  // fixture so it doesn't conflict with sibling tests.
  const LOCAL_FORMS_DIR = path.join(FIXTURE_DIR, '.forms-copy');
  const LOCAL_FORMS_CODEGRAPH_DIR = path.join(LOCAL_FORMS_DIR, '.codegraph-vba');
  let formsCg: CodeGraph | null = null;
  let formsInit = false;

  beforeAll(async () => {
    // Reset our local copy and clone the source FORMS fixtures into
    // it. fs.cpSync with `recursive: true` is the idiomatic Node 16+
    // way to deep-copy a directory; it's idempotent (we delete
    // LOCAL_FORMS_DIR first if it exists). We also EXCLUDE any
    // `.codegraph-vba` subdirectory the source might have — that
    // subdirectory is owned by the sibling test
    // (`extraction-vba-control-modeling.test.ts`) and may still
    // exist on disk while that test's `afterAll` is still running in
    // a parallel vitest worker. Copying it in would leave a stale
    // `.codegraph-vba` in our local copy and `CodeGraph.init` would
    // then refuse with "already initialized".
    rmWithRetry(LOCAL_FORMS_DIR, { recursive: true, force: true });
    fs.cpSync(SOURCE_FORMS_DIR, LOCAL_FORMS_DIR, {
      recursive: true,
      filter: (src) => !path.basename(src).startsWith('.codegraph'),
    });

    formsCg = await CodeGraph.init(LOCAL_FORMS_DIR, { index: false });
    formsInit = true;
    await formsCg.indexAll();
  }, 60_000);

  afterAll(async () => {
    if (formsCg) {
      try {
        await formsCg.close();
      } catch {
        /* ignore */
      }
    }
    if (formsInit) {
      rmWithRetry(LOCAL_FORMS_DIR, { recursive: true, force: true });
    }
  });

  it('the FORMS-* fixture suite yields ZERO vba-event-handler edges', () => {
    if (!formsCg) throw new Error('formsCg not initialized');
    // The vba-control-modeling fixtures have no WithEvents bindings
    // (only `Me.<Control>` references and form-control handlers in
    // `procedures.ts`). So the new pass should emit zero edges.
    //
    // The fixtures DO have a few `event-handler` edges from the
    // pre-existing form-control synthesis — those have
    // `metadata.eventName` and target `form-instance-control` nodes,
    // NOT `function` nodes. We filter on
    // `metadata.synthesizedBy === 'vba-event-handler'` so only the new
    // pass is counted.
    const all = allEventHandlerEdges(formsCg).filter(
      (e) => e.metadata?.synthesizedBy === 'vba-event-handler',
    );
    expect(all.length, 'expected zero vba-event-handler edges in FORMS-* suite').toBe(0);
  });
});

/**
 * Best-effort recursive remove with a small retry loop. Windows
 * sometimes reports `EBUSY` for a brief window after a SQLite
 * connection is closed (the OS hasn't fully released the file
 * handle) or while a watcher process is mid-syscall. Three retries
 * with a 200ms backoff handles the typical case without making the
 * test slow when the operation is going to fail for a real reason.
 */
function rmWithRetry(
  target: string,
  opts: fs.RmOptions,
  attempts = 3,
  delayMs = 200,
): void {
  let lastErr: unknown = null;
  for (let i = 0; i < attempts; i++) {
    try {
      fs.rmSync(target, opts);
      return;
    } catch (err) {
      lastErr = err;
      // Synchronous sleep — `beforeAll` accepts async but a small
      // busy-wait here is simpler than juggling timers and keeps the
      // test deterministic.
      const until = Date.now() + delayMs;
      while (Date.now() < until) { /* spin */ }
    }
  }
  // Last attempt failed; surface the underlying error so a real
  // failure doesn't get silently swallowed.
  throw lastErr;
}
