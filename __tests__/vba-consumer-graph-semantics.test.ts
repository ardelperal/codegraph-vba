/**
 * Issue #298 — the two Access consumer helpers must read the graph the
 * production extractor actually emits.
 *
 * Every case below indexes the shared corpus through real extraction and
 * resolution (see `helpers/vba-consumer-semantics.ts`). Nothing inserts an
 * edge by hand, so a traversal that only works against a hypothetical graph
 * fails here instead of passing on its own fixture.
 */
import { afterEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { DatabaseConnection } from '../src/db';
import { QueryBuilder } from '../src/db/queries';
import { traverseGraph, TraversalNode } from '../src/utils/backtrace-helpers';
import { runImpactAnalysis } from '../src/utils/sql-impact-helpers';
import {
  indexConsumerSemanticsCorpus,
  IndexedCorpus,
} from './helpers/vba-consumer-semantics';

let corpus: IndexedCorpus | null = null;
const scratch: Array<() => void> = [];

afterEach(async () => {
  if (corpus) {
    await corpus.dispose();
    corpus = null;
  }
  while (scratch.length) scratch.pop()!();
});

function names(nodes: TraversalNode[]): string[] {
  return nodes.map((child) => child.name);
}

function flatten(node: TraversalNode | null): string[] {
  if (!node) return [];
  return [node.name, ...node.children.flatMap(flatten)];
}

function controlId(
  loaded: IndexedCorpus,
  control: string,
  layoutFile: string,
): string {
  const hits = loaded.cg
    .searchNodes(control, { kinds: ['form-instance-control'], languages: ['vba'] })
    .filter(({ node }) => path.basename(node.filePath) === layoutFile);
  expect(hits.length, `${control} in ${layoutFile}`).toBe(1);
  return hits[0]!.node.id;
}

describe('issue #298 — canonical Access graph consumption', () => {
  it('indexed control reaches canonical handler and callees', async () => {
    corpus = await indexConsumerSemanticsCorpus();
    const btnSave = controlId(corpus, 'btnSave', 'Form_Orders.form.txt');

    const result = traverseGraph(corpus.db, btnSave);

    expect(result.warnings).toEqual([]);
    expect(result.cycle_detected).toBe(false);
    expect(result.tree?.kind).toBe('form-instance-control');
    expect(names(result.tree!.children)).toEqual(['btnSave_Click']);
    expect(names(result.tree!.children[0]!.children)).toEqual(['SaveOrderTotals']);
  });

  it('form-level lifecycle handlers are reachable from the layout node', async () => {
    corpus = await indexConsumerSemanticsCorpus();
    const layout = corpus.cg.searchNodes('Form_Orders', {
      kinds: ['form-layout'],
      languages: ['vba'],
    })[0]!.node;

    const result = traverseGraph(corpus.db, layout.id);

    expect(flatten(result.tree)).toContain('Form_Load');
    expect(flatten(result.tree)).toContain('RefreshOrders');
  });

  it('reports expose their section handlers the same way forms do', async () => {
    corpus = await indexConsumerSemanticsCorpus();
    const detalle = controlId(corpus, 'Detalle', 'Report_Sales.report.txt');

    const result = traverseGraph(corpus.db, detalle);

    expect(names(result.tree!.children)).toEqual(['Detalle_Format']);
    expect(names(result.tree!.children[0]!.children)).toEqual(['FormatSalesRow']);
  });

  it('same-named controls stay layout-scoped', async () => {
    corpus = await indexConsumerSemanticsCorpus();
    const orders = traverseGraph(
      corpus.db,
      controlId(corpus, 'btnSave', 'Form_Orders.form.txt'),
    );
    const invoices = traverseGraph(
      corpus.db,
      controlId(corpus, 'btnSave', 'Form_Invoices.form.txt'),
    );

    expect(flatten(orders.tree)).toContain('SaveOrderTotals');
    expect(flatten(orders.tree)).not.toContain('SaveInvoiceTotals');
    expect(flatten(invoices.tree)).toContain('SaveInvoiceTotals');
    expect(flatten(invoices.tree)).not.toContain('SaveOrderTotals');
  });

  it('does not present containment or data references as call steps', async () => {
    corpus = await indexConsumerSemanticsCorpus();
    const saveOrderTotals = corpus.cg.searchNodes('SaveOrderTotals', {
      kinds: ['function'],
      languages: ['vba'],
    })[0]!.node;

    const result = traverseGraph(corpus.db, saveOrderTotals.id);

    // The procedure references qryOrderTotals and is contained by modOrders;
    // neither is a call step.
    expect(flatten(result.tree)).toEqual(['SaveOrderTotals']);
  });

  it('SQL caller reaches owning layout without direct UI binding', async () => {
    corpus = await indexConsumerSemanticsCorpus();

    const result = runImpactAnalysis(corpus.db, corpus.dir, 'qryOrderTotals');

    expect(result.form_bindings).toEqual([]);
    expect(result.downstream_impact.vba_callers).toContain('modOrders.bas');
    expect(result.downstream_impact.forms).toEqual(['Form_Orders']);
  });

  it('report-owned SQL impact names the report layout', async () => {
    corpus = await indexConsumerSemanticsCorpus();

    const result = runImpactAnalysis(corpus.db, corpus.dir, 'qrySalesTotals');

    expect(result.form_bindings).toEqual([]);
    expect(result.downstream_impact.forms).toEqual(['Report_Sales']);
  });

  it('semantic traversal terminates cycles', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'vba-consumer-cycle-'));
    scratch.push(() => fs.rmSync(dir, { recursive: true, force: true }));
    const conn = DatabaseConnection.initialize(path.join(dir, 'cycle.db'));
    scratch.push(() => conn.close());
    const q = new QueryBuilder(conn.getDb());
    const node = (id: string, name: string, kind: string) => ({
      id,
      kind,
      name,
      qualifiedName: name,
      filePath: 'modCycle.bas',
      language: 'vba',
      startLine: 1,
      endLine: 1,
      startColumn: 0,
      endColumn: 0,
      updatedAt: Date.now(),
    });
    q.insertNodes([node('a', 'Ping', 'function'), node('b', 'Pong', 'function')]);
    q.insertEdges([
      { source: 'a', target: 'b', kind: 'calls' },
      { source: 'b', target: 'a', kind: 'calls' },
    ]);

    const result = traverseGraph(conn.getDb(), 'a');

    expect(result.cycle_detected).toBe(true);
    expect(flatten(result.tree)).toEqual(['Ping', 'Pong', 'Ping']);
  });

  it('reports a missing root explicitly instead of failing silently', async () => {
    corpus = await indexConsumerSemanticsCorpus();

    const result = traverseGraph(corpus.db, 'no-such-node');

    expect(result.tree).toBeNull();
    expect(result.warnings).toEqual(['START_NODE_NOT_FOUND']);
  });
});
