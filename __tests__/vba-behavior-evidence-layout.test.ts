import { buildBehaviorEvidence } from '../src/graph/behavior-evidence';
import { QueryBuilder } from '../src/db/queries';
import { Node } from '../src/types';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { indexConsumerSemanticsCorpus, IndexedCorpus } from './helpers/vba-consumer-semantics';

let corpus: IndexedCorpus;
beforeAll(async () => { corpus = await indexConsumerSemanticsCorpus(); });
afterAll(async () => { await corpus?.dispose(); });

describe('issue #310 — handler layout selectors', () => {
  it.each([
    ['btnSave_Click', 'Form_Orders', 'Form_Orders.form.txt'],
    ['Report_Open', 'Report_Sales', 'Report_Sales.report.txt'],
    ['Detalle_Format', 'Report_Sales', 'Report_Sales.report.txt'],
    ['AuditNow', 'Form_Admin', 'Form_Admin.form.txt'],
  ])('resolves %s by indexed layout name and file', (name, layout, file) => {
    const byName = corpus.cg.getBehaviorEvidence({ name, layout });
    const byFile = corpus.cg.getBehaviorEvidence({ name, layout: file.toUpperCase() });
    expect(byName.target).toMatchObject({ name, layout });
    expect(byFile.target).toEqual(byName.target);
    expect(byFile.evidence).toEqual(byName.evidence);
    expect(byFile.evidence.length).toBeGreaterThan(0);
    expect(corpus.cg.getBehaviorEvidence({ name, layout: 'Missing.form.txt' }).target).toBeNull();
  });

  it('keeps unscoped repeated handlers ambiguous and scopes either owner', () => {
    const ambiguous = corpus.cg.getBehaviorEvidence({ name: 'btnSave_Click' });
    expect(ambiguous.context.notes).toContain('AMBIGUOUS_TARGET');
    expect(ambiguous.target).toBeNull();
    expect(ambiguous.context.ambiguous[0].matches.map(m => m.layout).sort())
      .toEqual(['Form_Invoices', 'Form_Orders']);
    expect(corpus.cg.getBehaviorEvidence({ name: 'btnSave_Click', layout: 'Form_Invoices.form.txt' }).target)
      .toMatchObject({ name: 'btnSave_Click', layout: 'Form_Invoices' });
  });
});

// A shared expression has several legitimate bindings, not one guessed owner.
it('scopes shared expression wiring and preserves same-layout ambiguity', () => {
  const node = (id: string, name: string, kind: Node['kind'], filePath: string): Node => ({
    id, name, kind, filePath, qualifiedName: id, language: 'vba',
    startLine: 1, endLine: 2, startColumn: 0, endColumn: 0,
  });
  const fn = node('fn', 'Shared', 'function', 'Utilities.bas');
  const first = node('a', 'First', 'form-layout', 'First.form.txt');
  const second = node('b', 'Second', 'report-layout', 'Second.report.txt');
  const nodes = [fn, first, second];
  const queries = {
    getNodesByLowerName: (name: string) => nodes.filter(n => n.name.toLowerCase() === name),
    getNodeById: (id: string) => nodes.find(n => n.id === id) ?? null,
    getIncomingEdges: () => [],
    getOutgoingEdges: (id: string, kinds: string[]) => ['fn', 'other'].includes(id) && kinds.includes('event-handler')
      ? [first, second].map((layout, i) => ({ source: id, target: layout.id,
        metadata: { synthesizedBy: 'vba-expression-handler', eventName: i ? 'Open' : 'Load' } })) : [],
    getUnresolvedReferencesForFile: () => [],
  } as unknown as QueryBuilder;
  expect(buildBehaviorEvidence(queries, { name: 'Shared' }).target?.layout).toBeNull();
  const scoped = buildBehaviorEvidence(queries, { name: 'Shared', layout: 'Second.report.txt' });
  expect(scoped.target?.layout).toBe('Second');
  expect(scoped.context.handlers[0]).toMatchObject({ scope: 'expression', event: 'Open' });
  nodes.push({ ...fn, id: 'other', filePath: 'Other.bas' });
  expect(buildBehaviorEvidence(queries, { name: 'Shared', layout: 'Second.report.txt' })
    .context.notes).toContain('AMBIGUOUS_TARGET');
});
