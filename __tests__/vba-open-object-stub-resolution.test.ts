import { afterEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { CodeGraph } from '../src';

const opened: Array<{ cg: CodeGraph; dir: string }> = [];

afterEach(async () => {
  while (opened.length) {
    const { cg, dir } = opened.pop()!;
    await cg.close();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

const navSource = `Attribute VB_Name = "Nav"
Option Explicit
Private Const FORM_NAME_CONST As String = "FormNC"
Public Sub AbrirNC()
  DoCmd.OpenForm FORM_NAME_CONST
End Sub
Public Sub AbrirInforme()
  DoCmd.OpenReport "monthly"
End Sub`;

const formSource = `Attribute VB_Name = "Form_FormNC"
Begin Form
  Name ="FormNC"
End`;

const formClassSource = `VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "Form_FormNC"
Option Explicit
Private Sub Form_Load()
End Sub`;

const reportSource = `Attribute VB_Name = "Report_Monthly"
Begin Report
  Name ="Monthly"
End`;

async function indexFixture(layoutFirst: boolean, includeTargets = true): Promise<CodeGraph> {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'vba-open-object-'));
  const write = (order: number, name: string, content: string) => {
    const target = path.join(dir, String(order), name);
    fs.mkdirSync(path.dirname(target), { recursive: true });
    fs.writeFileSync(target, content);
  };
  write(layoutFirst ? 9 : 0, 'Nav.bas', navSource);
  if (includeTargets) {
    write(layoutFirst ? 0 : 9, 'Form_FormNC.form.txt', formSource);
    write(layoutFirst ? 1 : 8, 'Form_FormNC.cls', formClassSource);
    write(layoutFirst ? 2 : 7, 'Report_Monthly.report.txt', reportSource);
  }
  const cg = await CodeGraph.init(dir, { index: false });
  opened.push({ cg, dir });
  await cg.indexAll();
  return cg;
}

describe('DoCmd opened-object stub resolution', () => {
  it.each([false, true])('bridges real forms and reports regardless of index order (layoutFirst=%s)', async (layoutFirst) => {
    const cg = await indexFixture(layoutFirst);
    const caller = cg.searchNodes('AbrirNC', { kinds: ['function'], languages: ['vba'] })[0]!.node;
    const form = cg.searchNodes('Form_FormNC', { kinds: ['form-layout'], languages: ['vba'] }).map((r) => r.node);
    expect(form).toHaveLength(1);
    expect(form[0]!.metadata?.stub).not.toBe(true);

    const opensForm = cg.getOutgoingEdges(caller.id).filter((e) => e.kind === 'opens-form');
    expect(opensForm).toEqual([
      expect.objectContaining({
        target: form[0]!.id,
        provenance: 'heuristic',
        metadata: expect.objectContaining({ synthesizedBy: 'vba-opens-form', targetFormName: 'FormNC' }),
      }),
    ]);

    const formClass = cg.searchNodes('Form_FormNC', { kinds: ['class'], languages: ['vba'] })[0]!.node;
    const formLoad = cg.searchNodes('Form_Load', { kinds: ['function'], languages: ['vba'] })[0]!.node;
    expect(cg.getOutgoingEdges(form[0]!.id)).toEqual(expect.arrayContaining([
      expect.objectContaining({ kind: 'references', target: formClass.id }),
    ]));
    expect(cg.getOutgoingEdges(formClass.id)).toEqual(expect.arrayContaining([
      expect.objectContaining({ kind: 'contains', target: formLoad.id }),
    ]));

    const reportCaller = cg.searchNodes('AbrirInforme', { kinds: ['function'], languages: ['vba'] })[0]!.node;
    const report = cg.searchNodes('Report_Monthly', { kinds: ['report-layout'], languages: ['vba'] }).map((r) => r.node);
    expect(report).toHaveLength(1);
    expect(cg.getOutgoingEdges(reportCaller.id)).toEqual(expect.arrayContaining([
      expect.objectContaining({
        kind: 'opens-report',
        target: report[0]!.id,
        provenance: 'heuristic',
        metadata: expect.objectContaining({ synthesizedBy: 'vba-opens-report', targetReportName: 'monthly' }),
      }),
    ]));
  });

  it('preserves the current placeholder behavior when targets are not indexed', async () => {
    const cg = await indexFixture(false, false);
    const form = cg.searchNodes('FormNC', { kinds: ['form-layout'], languages: ['vba'] })[0]!.node;
    const report = cg.searchNodes('monthly', { kinds: ['report-layout'], languages: ['vba'] })[0]!.node;
    expect(form.metadata?.stub).toBe(true);
    expect(report.metadata?.stub).toBe(true);
    expect(cg.getIncomingEdges(form.id)).toEqual(expect.arrayContaining([expect.objectContaining({ kind: 'opens-form' })]));
    expect(cg.getIncomingEdges(report.id)).toEqual(expect.arrayContaining([expect.objectContaining({ kind: 'opens-report' })]));
  });

  it('does not resurrect placeholders or duplicate nodes on a consecutive full index', async () => {
    const cg = await indexFixture(false);
    const nodeCountBefore = cg.getStats().nodeCount;
    const before = [
      cg.searchNodes('Form_FormNC', { kinds: ['form-layout'], languages: ['vba'] }).length,
      cg.searchNodes('Report_Monthly', { kinds: ['report-layout'], languages: ['vba'] }).length,
    ];
    await cg.indexAll();
    expect(cg.getStats().nodeCount).toBe(nodeCountBefore);
    const after = [
      cg.searchNodes('Form_FormNC', { kinds: ['form-layout'], languages: ['vba'] }).length,
      cg.searchNodes('Report_Monthly', { kinds: ['report-layout'], languages: ['vba'] }).length,
    ];
    expect(after).toEqual(before);
    expect(after).toEqual([1, 1]);
  });

  it('preserves navigation across target-only modify, delete, and re-add syncs', async () => {
    const cg = await indexFixture(false);
    const dir = opened[opened.length - 1]!.dir;
    const formPath = path.join(dir, '9', 'Form_FormNC.form.txt');
    const reportPath = path.join(dir, '7', 'Report_Monthly.report.txt');

    fs.appendFileSync(formPath, '\n');
    fs.appendFileSync(reportPath, '\n');
    await cg.sync();
    expect(cg.searchNodes('Form_FormNC', { kinds: ['form-layout'] })).toHaveLength(1);
    expect(cg.searchNodes('Report_Monthly', { kinds: ['report-layout'] })).toHaveLength(1);

    fs.rmSync(formPath);
    fs.rmSync(reportPath);
    await cg.sync();
    const formStub = cg.searchNodes('FormNC', { kinds: ['form-layout'] })[0]!.node;
    const reportStub = cg.searchNodes('monthly', { kinds: ['report-layout'] })[0]!.node;
    expect(formStub.metadata?.stub).toBe(true);
    expect(reportStub.metadata?.stub).toBe(true);
    expect(cg.getIncomingEdges(formStub.id)).toEqual(expect.arrayContaining([expect.objectContaining({ kind: 'opens-form' })]));
    expect(cg.getIncomingEdges(reportStub.id)).toEqual(expect.arrayContaining([expect.objectContaining({ kind: 'opens-report' })]));

    fs.writeFileSync(formPath, formSource);
    fs.writeFileSync(reportPath, reportSource);
    await cg.sync();
    const form = cg.searchNodes('Form_FormNC', { kinds: ['form-layout'] }).map((r) => r.node);
    const report = cg.searchNodes('Report_Monthly', { kinds: ['report-layout'] }).map((r) => r.node);
    expect(form).toHaveLength(1);
    expect(report).toHaveLength(1);
    expect(form[0]!.metadata?.stub).not.toBe(true);
    expect(report[0]!.metadata?.stub).not.toBe(true);
    expect(cg.getIncomingEdges(form[0]!.id)).toEqual(expect.arrayContaining([expect.objectContaining({ kind: 'opens-form' })]));
    expect(cg.getIncomingEdges(report[0]!.id)).toEqual(expect.arrayContaining([expect.objectContaining({ kind: 'opens-report' })]));
  });
});
