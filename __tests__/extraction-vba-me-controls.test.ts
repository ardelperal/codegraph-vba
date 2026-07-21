import { afterEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { CodeGraph } from '../src';
import { VbaExtractor } from '../src/extraction/vba-extractor';
import { UnresolvedReference } from '../src/types';

const open: Array<{ cg: CodeGraph; dir: string }> = [];

afterEach(async () => {
  while (open.length) {
    const { cg, dir } = open.pop()!;
    await cg.close();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

async function indexFiles(files: Record<string, string>): Promise<CodeGraph> {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'vba-me-controls-'));
  for (const [relative, source] of Object.entries(files)) {
    const target = path.join(dir, relative);
    fs.mkdirSync(path.dirname(target), { recursive: true });
    fs.writeFileSync(target, source);
  }
  const cg = await CodeGraph.init(dir, { index: false });
  open.push({ cg, dir });
  await cg.indexAll();
  return cg;
}

describe('issue #140 Me.<Control> references', () => {
  it('emits scoped, deduplicated references from form procedures', () => {
    const result = new VbaExtractor('forms/Form_Prueba.cls', `Attribute VB_Name = "Form_Prueba"
Private Sub Uno()
  Me.txtNombre = "x"
  If Me.chkOk Then Me.txtNombre.SetFocus
  Me.txtNombre = Me.txtNombre
  Me.Requery
  Me.Inexistente = 1
End Sub
Private Sub Dos()
  Debug.Print Me.txtNombre
End Sub`).extract();

    const refs = result.unresolvedReferences.filter(
      (ref) => ref.metadata?.synthesizedBy === 'vba-me-control',
    );
    expect(refs.map((ref) => [ref.referenceName, ref.line, ref.referenceKind])).toEqual([
      ['txtNombre', 3, 'references'],
      ['chkOk', 4, 'references'],
      ['Requery', 6, 'property-get'],
      ['Inexistente', 7, 'references'],
      ['txtNombre', 10, 'references'],
    ]);
    expect(refs.every((ref) => ref.metadata?.siblingPath === 'forms/Form_Prueba.form.txt')).toBe(true);
    expect(refs.find((ref) => ref.referenceName === 'txtNombre')?.metadata?.access).toBe('write');
    expect(refs.find((ref) => ref.referenceName === 'Requery')?.metadata?.builtIn).toBe(true);
  });

  it('does not run the dot sweep for ordinary class modules', () => {
    const result = new VbaExtractor('Servicio.cls', `Attribute VB_Name = "Servicio"
Public Sub Ejecutar()
  Me.Algo = 1
End Sub`).extract();
    expect(result.unresolvedReferences.filter(
      (ref) => ref.metadata?.synthesizedBy === 'vba-me-control',
    )).toHaveLength(0);
  });

  it('uses the report sibling layout for report code-behind', () => {
    const result = new VbaExtractor('reports/Report_Resumen.cls', `Attribute VB_Name = "Report_Resumen"
Private Sub Detail_Print()
  Me.txtTotal = 1
End Sub`).extract();
    expect(result.unresolvedReferences).toEqual(expect.arrayContaining([
      expect.objectContaining({
        referenceName: 'txtTotal',
        referenceKind: 'references',
        metadata: expect.objectContaining({
          synthesizedBy: 'vba-me-control',
          siblingPath: 'reports/Report_Resumen.report.txt',
        }),
      }),
    ]));
  });

  it('binds case-insensitively only to controls in the sibling layout without adding nodes', async () => {
    const cls = `Attribute VB_Name = "Form_Prueba"
Private Sub Uno()
  Me.txtNombre = "x"
  If Me.CHKOK Then Me.txtNombre.SetFocus
  Me.Requery
  Me.Inexistente = 1
End Sub
Private Sub Dos()
  Debug.Print Me.txtNombre
End Sub`;
    const cg = await indexFiles({
      'forms/Form_Prueba.cls': cls,
      'forms/Form_Prueba.form.txt': `Begin Form
  Begin TextBox
    Name ="txtNombre"
  End
  Begin CheckBox
    Name ="chkOk"
  End
End`,
      'other/Form_Otra.form.txt': `Begin Form
  Begin TextBox
    Name ="Inexistente"
  End
End`,
    });

    const controls = cg.searchNodes('', { kinds: ['form-instance-control'], languages: ['vba'] });
    const txt = controls.find((r) => r.node.name === 'txtNombre')!.node;
    const chk = controls.find((r) => r.node.name === 'chkOk')!.node;
    const nodeCount = cg.getStats().nodeCount;
    const txtIncoming = cg.getIncomingEdges(txt.id).filter((edge) => edge.kind === 'references');
    const chkIncoming = cg.getIncomingEdges(chk.id).filter((edge) => edge.kind === 'references');

    expect(txtIncoming).toHaveLength(2);
    expect(chkIncoming).toHaveLength(1);
    expect([...txtIncoming, ...chkIncoming].every(
      (edge) => edge.metadata?.synthesizedBy === 'vba-me-control',
    )).toBe(true);
    expect(cg.getStats().nodeCount).toBe(nodeCount);
    expect(cg.searchNodes('Requery').some((r) => r.node.name === 'Requery')).toBe(false);
    expect(cg.searchNodes('Inexistente').filter((r) => r.node.filePath.endsWith('.cls'))).toHaveLength(0);
  });

  it('connects the statically visible real-fixture MotivoBorrado usages', async () => {
    const fixtureDir = path.join(__dirname, 'fixtures', 'vba', 'src', 'forms');
    const cg = await indexFiles({
      'forms/Form_FormNCAuditoriaMotivoEliminado.cls': fs.readFileSync(
        path.join(fixtureDir, 'Form_FormNCAuditoriaMotivoEliminado.cls'),
        'utf8',
      ),
      'forms/Form_FormNCAuditoriaMotivoEliminado.form.txt': fs.readFileSync(
        path.join(fixtureDir, 'Form_FormNCAuditoriaMotivoEliminado.form.txt'),
        'utf8',
      ),
    });
    const control = cg.searchNodes('MotivoBorrado', {
      kinds: ['form-instance-control'], languages: ['vba'],
    }).find((result) => result.node.filePath.endsWith('.form.txt'))!.node;
    const callers = cg.getIncomingEdges(control.id)
      .filter((edge) => edge.kind === 'references' && edge.metadata?.synthesizedBy === 'vba-me-control')
      .map((edge) => cg.getNode(edge.source)?.name)
      .filter(Boolean)
      .sort();
    expect(callers).toEqual(['EstableceColorBordes', 'EstablecerDatos', 'HaHabidoCambios']);
  });
});

describe('issue #211 Me.<X> / Me!<X> read/write classification', () => {
  function meRefs(source: string): UnresolvedReference[] {
    const result = new VbaExtractor('forms/Form_Prueba.cls', source).extract();
    return result.unresolvedReferences.filter(
      (ref) => ref.metadata?.synthesizedBy === 'vba-me-control',
    );
  }

  function findByName(refs: UnresolvedReference[], name: string, line: number): UnresolvedReference {
    const ref = refs.find((r) => r.referenceName === name && r.line === line);
    if (!ref) {
      throw new Error(`expected ${name}@${line}, got ${JSON.stringify(refs.map((r) => [r.referenceName, r.line, r.referenceKind]))}`);
    }
    return ref;
  }

  it('Me.Name in a comparison is classified as a read (builtIn branch: property-get)', () => {
    const src = `Attribute VB_Name = "Form_Prueba"
Private Sub Uno()
  If Me.Name = "frmX" Then Debug.Print "x"
End Sub`;
    const refs = meRefs(src);
    const ref = findByName(refs, 'Name', 3);
    expect(ref.referenceKind).toBe('property-get');
    expect(ref.metadata?.builtIn).toBe(true);
  });

  it('Me!importe in a comparison is classified as a read (bang branch: bang-get)', () => {
    const src = `Attribute VB_Name = "Form_Prueba"
Private Sub Uno()
  If Me!importe = 0 Then Debug.Print "x"
End Sub`;
    const refs = meRefs(src);
    const ref = findByName(refs, 'importe', 3);
    expect(ref.referenceKind).toBe('bang-get');
  });

  it('Me.txtNombre as a statement-form assignment is still classified as a write (control branch)', () => {
    const src = `Attribute VB_Name = "Form_Prueba"
Private Sub Uno()
  Me.txtNombre = "x"
End Sub`;
    const refs = meRefs(src);
    const ref = findByName(refs, 'txtNombre', 3);
    expect(ref.referenceKind).toBe('references');
    expect(ref.metadata?.access).toBe('write');
  });

  it('Me!importe as a statement-form assignment is still classified as a write (bang branch: bang-set)', () => {
    const src = `Attribute VB_Name = "Form_Prueba"
Private Sub Uno()
  Me!importe = 0
End Sub`;
    const refs = meRefs(src);
    const ref = findByName(refs, 'importe', 3);
    expect(ref.referenceKind).toBe('bang-set');
  });

  it('Me.txtNombre in a comparison is classified as a read (control branch)', () => {
    const src = `Attribute VB_Name = "Form_Prueba"
Private Sub Uno()
  If Me.txtNombre = "x" Then Debug.Print "y"
End Sub`;
    const refs = meRefs(src);
    const ref = findByName(refs, 'txtNombre', 3);
    expect(ref.referenceKind).toBe('references');
    expect(ref.metadata?.access).toBe('read');
  });

  it('mixed line: read then assignment on the same line is read (builtIn branch sanity)', () => {
    // A pathological line that has both an `=` in the after and a prior
    // before-content. The unified predicate must reject this as a read.
    const src = `Attribute VB_Name = "Form_Prueba"
Private Sub Uno()
  If (Me.Name = "frmX") Then Debug.Print "x"
End Sub`;
    const refs = meRefs(src);
    const ref = findByName(refs, 'Name', 3);
    expect(ref.referenceKind).toBe('property-get');
  });
});
