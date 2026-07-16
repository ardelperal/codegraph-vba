import { afterEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { CodeGraph } from '../src';
import { VbaFormExtractor } from '../src/extraction/vba-form-extractor';

const open: Array<{ cg: CodeGraph; dir: string }> = [];

afterEach(async () => {
  while (open.length) {
    const { cg, dir } = open.pop()!;
    await cg.close();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

describe('issue #137 expression event extraction', () => {
  it('emits scoped refs for call expressions and ignores macros, event procedures, and non-calls', () => {
    const result = new VbaFormExtractor('Form_Orders.form.txt', `Begin Form
  OnCurrent ="=RefreshState()"
  Begin CommandButton
    Name ="cmdRun"
    OnClick ="=DoThing([Id],""a"")"
    OnDblClick ="[Event Procedure]"
    OnMouseMove ="ToolbarMacro"
    OnEnter ="=Forms!Orders!Id"
    OnExit ="=DoThing("
    OnGotFocus ="=DoThing() trailing"
  End
End`).extract();

    const control = result.nodes.find((node) => node.kind === 'form-instance-control');
    const layout = result.nodes.find((node) => node.kind === 'form-layout');
    expect(result.unresolvedReferences).toEqual(expect.arrayContaining([
      expect.objectContaining({
        fromNodeId: control?.id,
        referenceName: 'DoThing',
        referenceKind: 'event-handler',
        metadata: { eventName: 'Click', synthesizedBy: 'vba-expression-handler' },
      }),
      expect.objectContaining({
        fromNodeId: layout?.id,
        referenceName: 'RefreshState',
        referenceKind: 'event-handler',
        metadata: { eventName: 'Current', synthesizedBy: 'vba-expression-handler' },
      }),
    ]));
    expect(result.unresolvedReferences.filter((ref) => ref.metadata?.synthesizedBy === 'vba-expression-handler')).toHaveLength(2);
    expect(result.nodes.filter((node) => node.kind === 'function')).toHaveLength(0);
  });
});

async function indexProject(includeFunction = true): Promise<CodeGraph> {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'vba-expression-events-'));
  fs.writeFileSync(path.join(dir, 'Form_Orders.form.txt'), `Begin Form
  OnCurrent ="=RefreshState()"
  Begin CommandButton
    Name ="cmdRun"
    OnClick ="=DoThing([Id],""a"")"
  End
End`);
  if (includeFunction) {
    fs.writeFileSync(path.join(dir, 'Modulo.bas'), `Attribute VB_Name = "Modulo"
Public Function DoThing(ByVal Id As Variant, ByVal label As String) As Boolean
  DoThing = True
End Function

Public Function RefreshState() As Boolean
  RefreshState = True
End Function`);
  }
  const cg = await CodeGraph.init(dir, { index: false });
  open.push({ cg, dir });
  await cg.indexAll();
  return cg;
}

describe('issue #137 expression event graph resolution', () => {
  it('links resolved functions to control and form targets with event metadata', async () => {
    const cg = await indexProject();
    const control = cg.searchNodes('cmdRun', { kinds: ['form-instance-control'], languages: ['vba'] })[0]!.node;
    const layout = cg.searchNodes('Form_Orders', { kinds: ['form-layout'], languages: ['vba'] })[0]!.node;
    const doThing = cg.searchNodes('DoThing', { kinds: ['function'], languages: ['vba'] })[0]!.node;
    const refresh = cg.searchNodes('RefreshState', { kinds: ['function'], languages: ['vba'] })[0]!.node;

    expect(cg.getIncomingEdges(control.id)).toContainEqual(expect.objectContaining({
      source: doThing.id,
      kind: 'event-handler',
      provenance: 'heuristic',
      metadata: expect.objectContaining({ eventName: 'Click', synthesizedBy: 'vba-expression-handler' }),
    }));
    expect(cg.getIncomingEdges(layout.id)).toContainEqual(expect.objectContaining({
      source: refresh.id,
      kind: 'event-handler',
      metadata: expect.objectContaining({ eventName: 'Current', synthesizedBy: 'vba-expression-handler' }),
    }));
    expect(cg.getCallers(control.id).map(({ node }) => node.id)).toContain(doThing.id);
  });

  it('keeps missing functions unresolved without fabricating a node', async () => {
    const cg = await indexProject(false);
    expect(cg.searchNodes('DoThing', { languages: ['vba'] })).toHaveLength(0);
    const control = cg.searchNodes('cmdRun', { kinds: ['form-instance-control'], languages: ['vba'] })[0]!.node;
    expect(cg.getIncomingEdges(control.id).filter((edge) => edge.kind === 'event-handler')).toHaveLength(0);
    const queries = (cg as unknown as {
      queries: { getRetryableFailedReferences(names: string[]): Array<{ referenceName: string }> };
    }).queries;
    expect(queries.getRetryableFailedReferences(['DoThing'])).toContainEqual(
      expect.objectContaining({ referenceName: 'DoThing' }),
    );
  });

  it('follows the existing best-match convention when function names are ambiguous', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'vba-expression-events-ambiguous-'));
    fs.writeFileSync(path.join(dir, 'Form_Orders.form.txt'), `Begin Form
  Begin CommandButton
    Name ="cmdRun"
    OnClick ="=DoThing()"
  End
End`);
    for (const moduleName of ['First', 'Second']) {
      fs.writeFileSync(path.join(dir, `${moduleName}.bas`), `Attribute VB_Name = "${moduleName}"
Public Function DoThing() As Boolean
  DoThing = True
End Function`);
    }
    const cg = await CodeGraph.init(dir, { index: false });
    open.push({ cg, dir });
    await cg.indexAll();
    const control = cg.searchNodes('cmdRun', { kinds: ['form-instance-control'], languages: ['vba'] })[0]!.node;
    const edges = cg.getIncomingEdges(control.id).filter((edge) => edge.kind === 'event-handler');
    expect(edges).toHaveLength(1);
    expect(cg.searchNodes('DoThing', { kinds: ['function'], languages: ['vba'] }).map(({ node }) => node.id)).toContain(edges[0]!.source);
  });

  it('recreates the handler edge when its function is modified, deleted, and restored', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'vba-expression-events-sync-'));
    const modulePath = path.join(dir, 'Modulo.bas');
    fs.writeFileSync(path.join(dir, 'Form_Orders.form.txt'), `Begin Form
  Begin CommandButton
    Name ="cmdRun"
    OnClick ="=DoThing()"
  End
End`);
    fs.writeFileSync(path.join(dir, 'Form_Orders.cls'), `Attribute VB_Name = "Form_Orders"`);
    const writeFunction = (comment = '') => fs.writeFileSync(modulePath, `Attribute VB_Name = "Modulo"
Public Function DoThing() As Boolean
  ${comment}
  DoThing = True
End Function`);
    writeFunction();
    const cg = await CodeGraph.init(dir, { index: false });
    open.push({ cg, dir });
    await cg.indexAll();
    const control = cg.searchNodes('cmdRun', { kinds: ['form-instance-control'], languages: ['vba'] })[0]!.node;
    const handlerEdges = () => cg.getIncomingEdges(control.id).filter((edge) => edge.kind === 'event-handler');
    expect(handlerEdges()).toHaveLength(1);
    writeFunction("' changed");
    await cg.sync();
    expect(handlerEdges()).toHaveLength(1);
    fs.rmSync(modulePath);
    await cg.sync();
    expect(handlerEdges()).toHaveLength(0);
    writeFunction("' restored");
    await cg.sync();
    expect(handlerEdges()).toHaveLength(1);
  });
});
