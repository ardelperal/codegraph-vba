import { afterEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { CodeGraph } from '../src';
import { VbaFormExtractor } from '../src/extraction/vba-form-extractor';
import { normalizeAccessObjectName } from '../src/resolution/vba-access-object-name';

const open: Array<{ cg: CodeGraph; dir: string }> = [];

afterEach(async () => {
  while (open.length) {
    const { cg, dir } = open.pop()!;
    await cg.close();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

describe('Access object name normalization', () => {
  it.each([
    ['Child', 'child'],
    ['Form.Child', 'child'],
    ['form_child', 'child'],
    ['REPORT.Child', 'child'],
    ['Report_Child', 'child'],
  ])('normalizes %s case-insensitively', (input, expected) => {
    expect(normalizeAccessObjectName(input)).toBe(expected);
  });
});

describe('issue #136 SourceObject extraction', () => {
  it('captures form/report embeddings and keeps missing targets unresolved without stubs', () => {
    const source = `Begin Form
  Begin Subform
    Name ="subBare"
    SourceObject ="MissingChild"
  End
  Begin Subform
    Name ="subForm"
    SourceObject ="Form.Child"
  End
  Begin Subform
    Name ="subReport"
    SourceObject ="Report.Monthly"
  End
End`;
    const r = new VbaFormExtractor('Form_Parent.form.txt', source).extract();
    const controls = r.nodes.filter((n) => n.kind === 'form-instance-control');
    expect(controls.map((n) => [n.name, n.metadata?.sourceObject])).toEqual([
      ['subBare', 'MissingChild'],
      ['subForm', 'Form.Child'],
      ['subReport', 'Report.Monthly'],
    ]);
    expect(r.unresolvedReferences).toEqual(expect.arrayContaining([
      expect.objectContaining({ referenceName: 'MissingChild', metadata: { synthesizedBy: 'vba-source-object', embeds: true, accessObjectKind: 'form' } }),
      expect.objectContaining({ referenceName: 'Child', metadata: { synthesizedBy: 'vba-source-object', embeds: true, accessObjectKind: 'form' } }),
      expect.objectContaining({ referenceName: 'Monthly', metadata: { synthesizedBy: 'vba-source-object', embeds: true, accessObjectKind: 'report' } }),
    ]));
    expect(r.nodes.some((n) => n.name === 'MissingChild')).toBe(false);
  });

  it('routes Table/Query SourceObject values through data-binding placeholders, never embeddings', () => {
    const source = `Begin Form
  Begin Subform
    Name ="tableData"
    SourceObject ="Table.Orders"
  End
  Begin Subform
    Name ="queryData"
    SourceObject ="Query.ActiveOrders"
  End
End`;
    const r = new VbaFormExtractor('Form_Parent.form.txt', source).extract();
    expect(r.unresolvedReferences.filter((u) => u.metadata?.synthesizedBy === 'vba-source-object')).toHaveLength(0);
    const edges = r.edges.filter((e) => e.metadata?.synthesizedBy === 'vba-source-object');
    expect(edges).toHaveLength(2);
    expect(edges.every((e) => e.metadata?.access === 'read')).toBe(true);
    expect(r.nodes.filter((n) => n.kind === 'class').map((n) => n.name).sort()).toEqual(['ActiveOrders', 'Orders']);
  });
});

async function indexFixture(reverse: boolean): Promise<CodeGraph> {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), `vba-source-object-${reverse ? 'z' : 'a'}-`));
  const fixture = path.join(__dirname, 'fixtures', 'vba-source-object');
  const files = fs.readdirSync(fixture).sort((a, b) => reverse ? b.localeCompare(a) : a.localeCompare(b));
  for (const [i, file] of files.entries()) {
    const prefix = reverse ? String(files.length - i) : String(i);
    const destination = path.join(dir, prefix, file);
    fs.mkdirSync(path.dirname(destination), { recursive: true });
    fs.copyFileSync(path.join(fixture, file), destination);
  }
  const cg = await CodeGraph.init(dir, { index: false });
  open.push({ cg, dir });
  await cg.indexAll();
  return cg;
}

describe('SourceObject graph resolution', () => {
  it.each([false, true])('links the parent control to Form_Child regardless of file order (reverse=%s)', async (reverse) => {
    const cg = await indexFixture(reverse);
    const child = cg.searchNodes('Form_Child', { kinds: ['form-layout'], languages: ['vba'] })[0]?.node;
    const control = cg.searchNodes('subChild', { kinds: ['form-instance-control'], languages: ['vba'] })[0]?.node;
    expect(child).toBeDefined();
    expect(control).toBeDefined();
    const incoming = cg.getIncomingEdges(child!.id).filter((e) => e.kind === 'references');
    expect(incoming).toEqual(expect.arrayContaining([
      expect.objectContaining({ source: control!.id, provenance: 'heuristic', metadata: expect.objectContaining({ synthesizedBy: 'vba-source-object', embeds: true }) }),
    ]));
  });
});
