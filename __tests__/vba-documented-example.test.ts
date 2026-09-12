/**
 * Issue #300 — the Access example in the README has to be real.
 *
 * This executes the documented sequence step for step, in the same order, and
 * asserts the identities, kinds and relationships it shows. It is a behavioral
 * test, not a prose snapshot: nothing here greps the README for a heading, and
 * nothing passes because a sentence exists. If the graph contract moves, the
 * documented example fails here before a user finds out by following it.
 *
 * Documented at: README.md → "VBA / Access + Dysflow integration" →
 * "Worked example: from a control to the tables it touches".
 */
import { afterEach, describe, expect, it } from 'vitest';
import {
  indexConsumerSemanticsCorpus,
  IndexedCorpus,
} from './helpers/vba-consumer-semantics';

let corpus: IndexedCorpus | null = null;

afterEach(async () => {
  if (corpus) {
    await corpus.dispose();
    corpus = null;
  }
});

describe('issue #300 — documented Access guidance', () => {
  it('documented Access traversal example matches indexed fixture', async () => {
    corpus = await indexConsumerSemanticsCorpus();
    const cg = corpus.cg;

    // Step 1 — resolve the control WITH its layout.
    const btnSave = cg
      .searchNodes('btnSave', { kinds: ['form-instance-control'], languages: ['vba'] })
      .map(({ node }) => node)
      .find((node) => node.filePath.endsWith('Form_Orders.form.txt'));
    expect(btnSave).toBeDefined();
    expect(btnSave!.kind).toBe('form-instance-control');

    // Step 2 — the binding is stored handler -> control.
    const binding = cg
      .getIncomingEdges(btnSave!.id)
      .filter((edge) => edge.kind === 'event-handler');
    expect(binding).toHaveLength(1);
    expect(binding[0]!.metadata?.eventName).toBe('Click');
    const handler = cg.searchNodes('btnSave_Click', {
      kinds: ['function'],
      languages: ['vba'],
    })
      .map(({ node }) => node)
      .find((node) => node.filePath.endsWith('Form_Orders.cls'));
    expect(binding[0]!.source).toBe(handler!.id);
    // And nothing points the other way: there is no reverse or duplicated edge.
    expect(
      cg.getOutgoingEdges(btnSave!.id).filter((edge) => edge.kind === 'event-handler'),
    ).toEqual([]);

    // Step 3 — what runs, and what it reaches.
    const behavior = cg.getBehaviorEvidence({ nodeId: btnSave!.id });
    expect(behavior.evidence[0]!.handler).toBe('btnSave_Click');
    expect(behavior.evidence[0]!.callPath).toEqual(['btnSave_Click', 'SaveOrderTotals']);
    expect(behavior.evidence[0]!.tables).toEqual(['tblOrderLines', 'tblProducts']);

    // Step 4 — the tables are reached through the saved query.
    expect(
      behavior.context.data.find((d) => d.name === 'tblOrderLines')!.throughQuery,
    ).toBe('qryOrderTotals');
  });

  it('documented layout scoping distinguishes same-named controls', async () => {
    corpus = await indexConsumerSemanticsCorpus();

    expect(
      corpus.cg.getBehaviorEvidence({ name: 'btnSave', layout: 'Form_Invoices' })
        .evidence[0]!.callPath,
    ).toEqual(['btnSave_Click', 'SaveInvoiceTotals']);

    const unscoped = corpus.cg.getBehaviorEvidence({ name: 'btnSave' });
    expect(unscoped.evidence).toEqual([]);
    expect(unscoped.context.ambiguous[0]!.matches.map((m) => m.layout).sort()).toEqual([
      'Form_Invoices',
      'Form_Orders',
    ]);
  });

  it('documented layout node kinds are what the extractor emits', async () => {
    corpus = await indexConsumerSemanticsCorpus();
    const kindsIn = (file: string) =>
      new Set(
        (
          corpus!.db
            .prepare('SELECT DISTINCT kind FROM nodes WHERE file_path = ?')
            .all(file) as Array<{ kind: string }>
        ).map((row) => row.kind),
      );

    const formKinds = kindsIn('Form_Orders.form.txt');
    expect(formKinds).toContain('form-layout');
    expect(formKinds).toContain('form-instance-control');
    expect(formKinds).not.toContain('module');
    const reportKinds = kindsIn('Report_Sales.report.txt');
    expect(reportKinds).toContain('report-layout');
    expect(reportKinds).toContain('form-instance-control');

    // The documented invariant: a layout file emits no procedures.
    for (const file of ['Form_Orders.form.txt', 'Report_Sales.report.txt']) {
      expect([...kindsIn(file)].filter((k) => ['function', 'method'].includes(k))).toEqual([]);
    }
  });

  it('documented lifecycle and expression bindings carry their stated markers', async () => {
    corpus = await indexConsumerSemanticsCorpus();

    // Lifecycle: handler -> layout, tagged scope 'form'.
    const layout = corpus.cg.searchNodes('Form_Orders', {
      kinds: ['form-layout'],
      languages: ['vba'],
    })[0]!.node;
    const lifecycle = corpus.cg
      .getIncomingEdges(layout.id)
      .filter((edge) => edge.kind === 'event-handler');
    expect(lifecycle).toHaveLength(1);
    expect(lifecycle[0]!.metadata).toMatchObject({ eventName: 'Load', scope: 'form' });

    // Expression-wired: same direction, tagged by the synthesizer.
    const btnAudit = corpus.cg
      .searchNodes('btnAudit', { kinds: ['form-instance-control'], languages: ['vba'] })[0]!
      .node;
    const expression = corpus.cg
      .getIncomingEdges(btnAudit.id)
      .filter((edge) => edge.kind === 'event-handler');
    expect(expression).toHaveLength(1);
    expect(expression[0]!.metadata).toMatchObject({
      eventName: 'Click',
      synthesizedBy: 'vba-expression-handler',
    });
  });
});
