/**
 * `scripts/vba-coverage-probe.mjs` — the committed coverage probe behind
 * every acceptance criterion in `docs/vba-node-discovery-plan.md`.
 *
 * The probe is a measuring instrument, so this suite pins EXACT counts over
 * a real three-file fixture set written to a temp dir (one `.bas`, one
 * `.cls`, one `.form.txt`). A drift in any number is either a real
 * extraction change — which the roadmap wants to see — or a probe bug.
 * Either way it must not pass silently.
 *
 * The single most important assertion here is the stub discriminator: a
 * synthesized call target and a declared procedure are BOTH `kind:
 * 'function'` and BOTH `visibility: 'public'`. Only `metadata.stub === true`
 * separates them. `counts the stub discriminator, not just kind === function`
 * is the regression guard — without it the reference corpus reads 9,972
 * procedures instead of 3,840.
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { VbaExtractor } from '../src/extraction/vba-extractor';
import { VbaFormExtractor } from '../src/extraction/vba-form-extractor';
import { SqlQueryExtractor } from '../src/extraction/sql-query-extractor';
// @ts-expect-error — plain ESM script, no type declarations by design.
import { runProbe, formatReport, fileExtensionKey, dispatchFor } from '../scripts/vba-coverage-probe.mjs';

const extractors = { VbaExtractor, VbaFormExtractor, SqlQueryExtractor };

/**
 * `modProbe.bas` — 3 declared procedures. `SaveRecord` reaches a same-file
 * call (`CalcTotal`), a synthesized stub (`VBA.DoEvents`), an `opens-form`
 * edge (`DoCmd.OpenForm`), a `vba-sql-table` edge (`tblExpediente`) and a
 * second stub (`CurrentDb.Execute`). `Idle` deliberately reaches nothing.
 */
const BAS_SOURCE = [
  'Attribute VB_Name = "modProbe"',
  'Option Explicit',
  '',
  'Public Sub SaveRecord()',
  '    Dim total As Long',
  '    total = CalcTotal()',
  '    VBA.DoEvents',
  '    DoCmd.OpenForm "FormMain"',
  '    CurrentDb.Execute "UPDATE tblExpediente SET Estado = 1"',
  'End Sub',
  '',
  'Private Function CalcTotal() As Long',
  '    CalcTotal = 1',
  'End Function',
  '',
  'Public Sub Idle()',
  'End Sub',
  '',
].join('\n');

/**
 * `Riesgo.cls` — 1 Sub + 1 Property Get (both land as `function` nodes),
 * plus two more stubs (`Collection.Add`, a second `VBA.DoEvents`). The
 * second `VBA.DoEvents` is what proves stub counting is per-file: node ids
 * are file-scoped, so the same stub name in two files is two nodes.
 */
const CLS_SOURCE = [
  'Attribute VB_Name = "Riesgo"',
  'Option Explicit',
  '',
  'Public Sub Registrar()',
  '    Dim items As New Collection',
  '    items.Add 1',
  '    VBA.DoEvents',
  'End Sub',
  '',
  'Public Property Get Nombre() As String',
  '    Nombre = "x"',
  'End Property',
  '',
].join('\n');

/**
 * `FormMain.form.txt` — a Dysflow SaveAsText export with 2 controls and a
 * `RecordSource`. Emits the `vba-form-binding` unresolved reference to the
 * sibling module name, which `formsReferenced` unions with the `opens-form`
 * target from the `.bas` (both resolve to `FormMain`, so the count is 1).
 */
const FORM_SOURCE = [
  'Begin Form',
  '    RecordSource = "tblExpediente"',
  '    Begin TextBox',
  '        Name = "txtCodigo"',
  '    End',
  '    Begin CommandButton',
  '        Name = "btnGuardar"',
  '    End',
  'End',
  '',
].join('\n');

describe('vba-coverage-probe', () => {
  let tmpDir: string;

  beforeEach(() => {
    tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'vba-coverage-probe-'));
    fs.mkdirSync(path.join(tmpDir, 'src', 'modules'), { recursive: true });
    fs.mkdirSync(path.join(tmpDir, 'src', 'classes'), { recursive: true });
    fs.mkdirSync(path.join(tmpDir, 'src', 'forms'), { recursive: true });
    fs.writeFileSync(path.join(tmpDir, 'src', 'modules', 'modProbe.bas'), BAS_SOURCE);
    fs.writeFileSync(path.join(tmpDir, 'src', 'classes', 'Riesgo.cls'), CLS_SOURCE);
    fs.writeFileSync(path.join(tmpDir, 'src', 'forms', 'FormMain.form.txt'), FORM_SOURCE);
  });

  afterEach(() => {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  });

  it('dispatches each fixture by extension', async () => {
    const report = await runProbe([tmpDir], { extractors });

    expect(report.files).toEqual({
      total: 3,
      dispatched: 3,
      skipped: 0,
      byExtension: { '.bas': 1, '.cls': 1, '.form.txt': 1 },
      skippedByExtension: {},
    });
  });

  it('counts the stub discriminator, not just kind === function', async () => {
    const report = await runProbe([tmpDir], { extractors });

    // 5 declared: SaveRecord, CalcTotal, Idle (.bas) + Registrar, Nombre (.cls).
    expect(report.declaredProcedures).toBe(5);
    // 4 stubs: VBA.DoEvents x2 (one node per file), CurrentDb.Execute,
    // Collection.Add.
    expect(report.stubProcedures).toBe(4);
    // The whole point: the two buckets are disjoint and both are
    // `kind: 'function'`. Counting kind alone would report 9.
    expect(report.declaredProcedures + report.stubProcedures).toBe(9);
    expect(report.nodesByKind.function).toBe(9);
    expect(report.declaredProcedures).not.toBe(report.nodesByKind.function);
  });

  it('reports exact node counts by kind', async () => {
    const report = await runProbe([tmpDir], { extractors });

    expect(report.nodesByKind).toEqual({
      function: 9,
      class: 4,
      file: 3,
      'form-instance-control': 2,
      'form-layout': 2,
      property: 2,
      module: 1,
    });
  });

  it('reports exact edge counts by kind and by synthesizer', async () => {
    const report = await runProbe([tmpDir], { extractors });

    expect(report.edgesByKind).toEqual({
      contains: 7,
      calls: 5,
      references: 3,
      'opens-form': 1,
    });
    expect(report.edgesBySynthesizer).toEqual({
      '(none)': 8,
      'vba-name-resolution': 5,
      'vba-opens-form': 1,
      'vba-record-source': 1,
      'vba-sql-table': 1,
    });
  });

  it('reports exact unresolved references by kind', async () => {
    const report = await runProbe([tmpDir], { extractors });

    // The `.form.txt` -> sibling module binding is the only unresolved
    // reference this fixture set produces.
    expect(report.unresolvedByKind).toEqual({ references: 1 });
  });

  it('ranks stub targets by count, ties broken by name', async () => {
    const report = await runProbe([tmpDir], { extractors });

    expect(report.stubTargetsTop).toEqual([
      { name: 'VBA.DoEvents', count: 2 },
      { name: 'Collection.Add', count: 1 },
      { name: 'CurrentDb.Execute', count: 1 },
    ]);
  });

  it('reports distinct SQL tables and forms reached', async () => {
    const report = await runProbe([tmpDir], { extractors });

    expect(report.sqlTablesReferenced).toEqual({
      count: 1,
      names: ['tblExpediente'],
    });
    // `opens-form` targets `FormMain` and the form binding names `FormMain`
    // too — the union is 1, not 2.
    expect(report.formsReferenced).toEqual({ count: 1, names: ['FormMain'] });
  });

  it('counts declared procedures with zero outgoing edges and references', async () => {
    const report = await runProbe([tmpDir], { extractors });

    // CalcTotal, Idle and Nombre reach nothing. SaveRecord and Registrar do.
    expect(report.proceduresWithNoOutgoing).toBe(3);
  });

  it('reports no extraction errors for well-formed fixtures', async () => {
    const report = await runProbe([tmpDir], { extractors });

    expect(report.errors).toEqual({});
  });

  it('counts a file no extractor claims as skipped, not dispatched', async () => {
    fs.writeFileSync(path.join(tmpDir, 'src', 'notes.md'), '# not VBA\n');
    const report = await runProbe([tmpDir], { extractors });

    expect(report.files.total).toBe(4);
    expect(report.files.dispatched).toBe(3);
    expect(report.files.skipped).toBe(1);
    expect(report.files.skippedByExtension).toEqual({ '.md': 1 });
    // The skipped file must not perturb any measurement.
    expect(report.declaredProcedures).toBe(5);
    expect(report.stubProcedures).toBe(4);
  });

  it('dispatches a .sql file to SqlQueryExtractor', async () => {
    fs.mkdirSync(path.join(tmpDir, 'src', 'queries'), { recursive: true });
    fs.writeFileSync(
      path.join(tmpDir, 'src', 'queries', 'qryActivos.sql'),
      'SELECT Id FROM tblExpediente WHERE Estado = 1;\n',
    );
    const report = await runProbe([tmpDir], { extractors });

    expect(report.files.byExtension['.sql']).toBe(1);
    expect(report.nodesByKind.query).toBe(1);
    expect(report.edgesBySynthesizer['sql-query-table']).toBe(1);
  });

  it('is deterministic across runs over the same tree', async () => {
    const first = await runProbe([tmpDir], { extractors });
    const second = await runProbe([tmpDir], { extractors });

    expect(second).toEqual(first);
  });

  it('renders a markdown report carrying the headline numbers', async () => {
    const report = await runProbe([tmpDir], { extractors });
    const md = formatReport(report);

    expect(md).toContain('# VBA coverage probe');
    expect(md).toContain('| declared procedures (no `metadata.stub`) | 5 |');
    expect(md).toContain('| stub function nodes (`metadata.stub === true`) | 4 |');
    expect(md).toContain('| `function` nodes total | 9 |');
    expect(md).toContain('| `VBA.DoEvents` | 2 |');
  });
});

describe('vba-coverage-probe — dispatch helpers', () => {
  it('keeps the two-segment form/report extension intact', () => {
    expect(fileExtensionKey('src/forms/FormMain.form.txt')).toBe('.form.txt');
    expect(fileExtensionKey('src/reports/Report_X.report.txt')).toBe('.report.txt');
    expect(fileExtensionKey('src/modules/modX.bas')).toBe('.bas');
    expect(fileExtensionKey('src/queries/qryX.sql')).toBe('.sql');
    expect(fileExtensionKey('LICENSE')).toBe('(none)');
  });

  it('routes each extension to the extractor the indexer would use', () => {
    expect(dispatchFor('src/forms/FormMain.form.txt')).toBe('form');
    expect(dispatchFor('src/reports/Report_X.report.txt')).toBe('form');
    expect(dispatchFor('src/modules/modX.bas')).toBe('code');
    expect(dispatchFor('src/classes/X.cls')).toBe('code');
    expect(dispatchFor('src/legacy/X.frm')).toBe('code');
    expect(dispatchFor('src/legacy/X.dsr')).toBe('code');
    expect(dispatchFor('src/queries/qryX.sql')).toBe('sql');
    expect(dispatchFor('src/queries/queries.json')).toBeNull();
    expect(dispatchFor('README.md')).toBeNull();
  });
});
