/**
 * VbaFormExtractor — tests for the `.form.txt` / `.report.txt` UI extractor.
 *
 * Coverage map (spec scenario id → test name):
 *   REQ-FORM-1 Form module named from VB_Name     → "Form module named from VB_Name"
 *   REQ-FORM-1 Form module named from filename    → "Form module named from filename when VB_Name absent"
 *   REQ-FORM-2 Single textbox control             → "single TextBox control emits property with controlType"
 *   REQ-FORM-2 Multiple controls produce props    → "TextBox + CommandButton emit two properties"
 *   REQ-FORM-3 Reports behave like forms          → "Report behaves identically to form"
 *   REQ-FORM-4 No code nodes from form UI         → "literal Sub keyword in form source produces no function nodes"
 *                                                  (also covers the empty form case)
 */
import { describe, it, expect } from 'vitest';
import { VbaFormExtractor } from '../src/extraction/vba-form-extractor';

function extract(filePath: string, source: string) {
  return new VbaFormExtractor(filePath, source).extract();
}

describe('VbaFormExtractor — form module with class binding (REQ-FORM-1)', () => {
  it('Form module named from VB_Name', () => {
    const src = `Attribute VB_Name = "Form_Main"
Begin
    Begin TextBox
        Name = "txtField"
    End
End`;
    const r = extract('src/forms/Form_Main.form.txt', src);
    // B2 (hueco 4): the file-level node is now `kind: 'form-layout'`,
    // not `module`. `.bas` standard modules keep emitting `module`.
    const formNode = r.nodes.find((n) => n.kind === 'form-layout');
    expect(formNode).toBeDefined();
    expect(formNode?.name).toBe('Form_Main');
    // The sibling-class binding is emitted as an UnresolvedReference (so
    // REQ-FORM-4 — no class nodes from form UI — is honored). It carries
    // synthesizedBy: vba-form-binding.
    const binding = r.unresolvedReferences.find(
      (u) => u.metadata?.synthesizedBy === 'vba-form-binding',
    );
    expect(binding).toBeDefined();
    expect(binding?.referenceName).toBe('Form_Main');
    expect(binding?.fromNodeId).toBe(formNode?.id);
  });

  it('Form module named from filename when VB_Name absent', () => {
    const src = `Begin
    Begin TextBox
        Name = "txtField"
    End
End`;
    const r = extract('src/forms/Form_Main.form.txt', src);
    const formNode = r.nodes.find((n) => n.kind === 'form-layout');
    expect(formNode).toBeDefined();
    expect(formNode?.name).toBe('Form_Main');
  });
});

describe('VbaFormExtractor — controls emit property nodes (REQ-FORM-2)', () => {
  it('single TextBox control emits property with controlType', () => {
    const src = `Attribute VB_Name = "Form_Main"
Begin
    Begin TextBox
        Name = "txtField"
    End
End`;
    const r = extract('src/forms/Form_Main.form.txt', src);
    const props = r.nodes.filter((n) => n.kind === 'property');
    expect(props).toHaveLength(1);
    expect(props[0]?.metadata?.controlType).toBe('TextBox');
  });

  it('TextBox + CommandButton emit two properties', () => {
    const src = `Attribute VB_Name = "Form_Main"
Begin
    Begin TextBox
        Name = "txtField"
    End
    Begin CommandButton
        Name = "btnOK"
    End
End`;
    const r = extract('src/forms/Form_Main.form.txt', src);
    const props = r.nodes.filter((n) => n.kind === 'property');
    expect(props).toHaveLength(2);
    const types = props.map((p) => p.metadata?.controlType).sort();
    expect(types).toEqual(['CommandButton', 'TextBox']);
  });
});

describe('VbaFormExtractor — reports behave like forms (REQ-FORM-3)', () => {
  it('Report behaves identically to form', () => {
    const src = `Attribute VB_Name = "Report_Orders"
Begin
    Begin TextBox
        Name = "txtOrderId"
    End
End`;
    const r = extract('src/reports/Report_Orders.report.txt', src);
    // B2 (hueco 4): the file-level node is now `kind: 'form-layout'`,
    // not `module`. The `.report.txt` path shares the same rename.
    const formNode = r.nodes.find((n) => n.kind === 'form-layout');
    expect(formNode).toBeDefined();
    expect(formNode?.name).toBe('Report_Orders');
    const binding = r.unresolvedReferences.find(
      (u) => u.metadata?.synthesizedBy === 'vba-form-binding',
    );
    expect(binding).toBeDefined();
    expect(binding?.referenceName).toBe('Report_Orders');
    const props = r.nodes.filter((n) => n.kind === 'property');
    expect(props).toHaveLength(1);
    expect(props[0]?.metadata?.controlType).toBe('TextBox');
  });
});

describe('VbaFormExtractor — no code nodes from form UI (REQ-FORM-4)', () => {
  it('literal Sub keyword in form source produces no function nodes', () => {
    const src = `Attribute VB_Name = "Form_Main"
' Sub Form_Load()
Begin
    Begin TextBox
        Name = "txtField"
    End
End`;
    const r = extract('src/forms/Form_Main.form.txt', src);
    // B2 (hueco 4): a `.form.txt` MUST NOT emit `module` nodes at all
    // (the form-level node is now `form-layout`). The `module` branch
    // of the filter is therefore vacuously true and only `function`
    // and `class` need to be checked.
    const codeNodes = r.nodes.filter(
      (n) =>
        n.kind === 'function' ||
        n.kind === 'class' ||
        n.kind === 'module',
    );
    expect(codeNodes).toHaveLength(0);
    // Class nodes also forbidden by REQ-FORM-4.
    const classes = r.nodes.filter((n) => n.kind === 'class');
    expect(classes).toHaveLength(0);
  });

  it('empty form file still emits form-layout + sibling-binding reference', () => {
    const r = extract('src/forms/Form_Main.form.txt', '');
    // B2 (hueco 4): the file-level node is `kind: 'form-layout'` now.
    const formNode = r.nodes.find((n) => n.kind === 'form-layout');
    expect(formNode).toBeDefined();
    expect(formNode?.name).toBe('Form_Main');
    const props = r.nodes.filter((n) => n.kind === 'property');
    expect(props).toHaveLength(0);
    const binding = r.unresolvedReferences.find(
      (u) => u.metadata?.synthesizedBy === 'vba-form-binding',
    );
    expect(binding).toBeDefined();
    expect(binding?.referenceName).toBe('Form_Main');
  });
});

/**
 * Fix 3: `Begin Form` and `Begin Section` are the form root and Access
 * section containers (Header/Detail/Footer), NOT user controls. They MUST
 * NOT appear as `property` nodes. The old `NON_CONTROL_TYPES` set was empty,
 * so every `Begin` line, including `Begin Form` and `Begin Section`, created
 * a property node.
 */
describe('VbaFormExtractor — Form and Section are not emitted as controls (Fix 3)', () => {
  it('Begin Form is filtered out and does not produce a property node', () => {
    const src = `Begin Form
    Begin TextBox
        Name = "txtField"
    End
End`;
    const r = extract('src/forms/Form_Main.form.txt', src);
    const props = r.nodes.filter((n) => n.kind === 'property');
    // Only the TextBox should be a property; Form is the root container.
    expect(props).toHaveLength(1);
    expect(props[0]?.metadata?.controlType).toBe('TextBox');
    // Explicitly: no property node with controlType 'Form'
    const formProps = props.filter((p) => p.metadata?.controlType === 'Form');
    expect(formProps).toHaveLength(0);
  });

  it('Begin Section is filtered out and does not produce a property node', () => {
    const src = `Begin Form
    Begin Section
    End
    Begin TextBox
        Name = "txtField"
    End
End`;
    const r = extract('src/forms/Form_Main.form.txt', src);
    const props = r.nodes.filter((n) => n.kind === 'property');
    // Only TextBox emits a property; Form and Section are containers.
    expect(props).toHaveLength(1);
    expect(props[0]?.metadata?.controlType).toBe('TextBox');
    const sectionProps = props.filter((p) => p.metadata?.controlType === 'Section');
    expect(sectionProps).toHaveLength(0);
  });

  it('Rectangle control is NOT filtered (it is a real Access control)', () => {
    const src = `Begin Form
    Begin Rectangle
        Name = "rectBorder"
    End
End`;
    const r = extract('src/forms/Form_Main.form.txt', src);
    const rectProps = r.nodes.filter(
      (n) => n.kind === 'property' && n.metadata?.controlType === 'Rectangle',
    );
    expect(rectProps).toHaveLength(1);
  });

  it('Image control is NOT filtered (it is a real Access control)', () => {
    const src = `Begin Form
    Begin Image
        Name = "imgLogo"
    End
End`;
    const r = extract('src/forms/Form_Main.form.txt', src);
    const imgProps = r.nodes.filter(
      (n) => n.kind === 'property' && n.metadata?.controlType === 'Image',
    );
    expect(imgProps).toHaveLength(1);
  });
});

/**
 * Issue #49 — RecordSource/RowSource bindings → `references` edges to
 * placeholder class nodes (one per table/query). The form-level
 * RecordSource attributes to the form-layout node with tag
 * `'vba-record-source'`; per-control RowSources attribute to the
 * enclosing `form-instance-control` node with tag `'vba-row-source'`.
 * Value-list controls (`RowSourceType = "Value List"`) are skipped.
 */
describe('VbaFormExtractor — RecordSource / RowSource bindings emit references edges (Issue #49)', () => {
  it('Issue #49: form-level RecordSource = "TbExpedientes" emits one references edge to a class placeholder', () => {
    const src = `Attribute VB_Name = "Form_Expedientes"
Begin Form
    RecordSource = "TbExpedientes"
    Begin TextBox
        Name = "txtId"
    End
End`;
    const r = extract('src/forms/Form_Expedientes.form.txt', src);
    const formNode = r.nodes.find((n) => n.kind === 'form-layout');
    expect(formNode).toBeDefined();
    // Exactly one references edge from form-layout with tag vba-record-source.
    const recordEdges = r.edges.filter(
      (e) =>
        e.kind === 'references' &&
        e.metadata?.synthesizedBy === 'vba-record-source',
    );
    expect(recordEdges).toHaveLength(1);
    expect(recordEdges[0]?.source).toBe(formNode?.id);
    // The edge target is a synthetic `class` placeholder named TbExpedientes.
    const target = r.nodes.find((n) => n.id === recordEdges[0]?.target);
    expect(target).toBeDefined();
    expect(target?.kind).toBe('class');
    expect(target?.name).toBe('TbExpedientes');
    expect(target?.qualifiedName).toBe('TbExpedientes');
  });

  it('form-level RecordSource binding carries access=read', () => {
    const src = `Attribute VB_Name = "Form_Expedientes"
Begin Form
    RecordSource = "TbExpedientes"
End`;
    const r = extract('src/forms/Form_Expedientes.form.txt', src);
    const recordEdge = r.edges.find(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-record-source',
    );
    expect(recordEdge?.metadata?.access).toBe('read');
  });

  it('control-level RowSource binding carries access=read', () => {
    const src = `Attribute VB_Name = "Form_Pedidos"
Begin Form
    Begin ComboBox
        Name = "cmbProvincias"
        RowSource = "SELECT Id, Nombre FROM TbProvincias"
    End
End`;
    const r = extract('src/forms/Form_Pedidos.form.txt', src);
    const rowEdge = r.edges.find(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-row-source',
    );
    expect(rowEdge?.metadata?.access).toBe('read');
  });

  it('Issue #49: form-level RecordSource with SELECT/FROM SQL emits one references edge per table', () => {
    const src = `Attribute VB_Name = "Form_Usuarios"
Begin Form
    RecordSource = "SELECT Id, Nombre FROM TbUsuarios ORDER BY Nombre"
    Begin TextBox
        Name = "txtNombre"
    End
End`;
    const r = extract('src/forms/Form_Usuarios.form.txt', src);
    const formNode = r.nodes.find((n) => n.kind === 'form-layout');
    expect(formNode).toBeDefined();
    const recordEdges = r.edges.filter(
      (e) =>
        e.kind === 'references' &&
        e.metadata?.synthesizedBy === 'vba-record-source',
    );
    // Exactly one edge to TbUsuarios (the SELECT/FROM/ORDER BY clauses
    // do not produce additional references — only the FROM clause's
    // table name does).
    expect(recordEdges).toHaveLength(1);
    expect(recordEdges[0]?.source).toBe(formNode?.id);
    const target = r.nodes.find((n) => n.id === recordEdges[0]?.target);
    expect(target?.name).toBe('TbUsuarios');
  });

  it('Issue #49: ComboBox RowSource SQL emits one references edge from the form-instance-control node', () => {
    const src = `Attribute VB_Name = "Form_Pedidos"
Begin Form
    Begin ComboBox
        Name = "cmbProvincias"
        RowSource = "SELECT Id, Nombre FROM TbProvincias"
    End
End`;
    const r = extract('src/forms/Form_Pedidos.form.txt', src);
    const controlNode = r.nodes.find(
      (n) => n.kind === 'form-instance-control' && n.name === 'cmbProvincias',
    );
    expect(controlNode).toBeDefined();
    const rowEdges = r.edges.filter(
      (e) =>
        e.kind === 'references' &&
        e.metadata?.synthesizedBy === 'vba-row-source',
    );
    expect(rowEdges).toHaveLength(1);
    // Edge source is the form-instance-control node, NOT the form-layout.
    expect(rowEdges[0]?.source).toBe(controlNode?.id);
    const target = r.nodes.find((n) => n.id === rowEdges[0]?.target);
    expect(target?.kind).toBe('class');
    expect(target?.name).toBe('TbProvincias');
  });

  it('Issue #49: ComboBox with RowSourceType = "Value List" emits ZERO references edges', () => {
    const src = `Attribute VB_Name = "Form_Pedidos"
Begin Form
    Begin ComboBox
        Name = "cmbEstados"
        RowSourceType = "Value List"
        RowSource = "uno;dos;tres"
    End
End`;
    const r = extract('src/forms/Form_Pedidos.form.txt', src);
    // No references edges from this combo — value-list is skipped.
    const valueListEdges = r.edges.filter(
      (e) =>
        e.kind === 'references' &&
        e.metadata?.synthesizedBy === 'vba-row-source',
    );
    expect(valueListEdges).toHaveLength(0);
    // And no class placeholder node for the literal list values.
    const synthClasses = r.nodes.filter(
      (n) => n.kind === 'class' && n.name === 'uno;dos;tres',
    );
    expect(synthClasses).toHaveLength(0);
  });

  it('Issue #49: mix of SQL RowSource and Value List RowSource — only SQL controls emit edges', () => {
    const src = `Attribute VB_Name = "Form_Pedidos"
Begin Form
    Begin ComboBox
        Name = "cmbProvincias"
        RowSource = "SELECT Id FROM TbProvincias"
    End
    Begin ComboBox
        Name = "cmbEstados"
        RowSourceType = "Value List"
        RowSource = "uno;dos;tres"
    End
    Begin ComboBox
        Name = "cmbCiudades"
        RowSource = "SELECT Id FROM TbCiudades ORDER BY Id"
    End
End`;
    const r = extract('src/forms/Form_Pedidos.form.txt', src);
    // Exactly 2 vba-row-source edges: TbProvincias + TbCiudades.
    // The Value List combo (cmbEstados) is skipped.
    const rowEdges = r.edges.filter(
      (e) =>
        e.kind === 'references' &&
        e.metadata?.synthesizedBy === 'vba-row-source',
    );
    expect(rowEdges).toHaveLength(2);
    const targets = rowEdges
      .map((e) => r.nodes.find((n) => n.id === e.target))
      .map((n) => n?.name)
      .sort();
    expect(targets).toEqual(['TbCiudades', 'TbProvincias']);
    // Each edge sources from its own combo control, not from the form-layout.
    const formNode = r.nodes.find((n) => n.kind === 'form-layout');
    for (const e of rowEdges) {
      expect(e.source).not.toBe(formNode?.id);
    }
  });

  it('Issue #49: REQ-FORM-4 invariant — no function/sub/module/event/declare/type nodes from form files with bindings', () => {
    const src = `Attribute VB_Name = "Form_Expedientes"
Begin Form
    RecordSource = "TbExpedientes"
    Begin ComboBox
        Name = "cmbEstados"
        RowSource = "SELECT Id FROM TbEstados"
    End
End`;
    const r = extract('src/forms/Form_Expedientes.form.txt', src);
    // The form's own node stays `form-layout` — never `class` (the
    // form's class binding is an UnresolvedReference, NOT a class node).
    const formFileClass = r.nodes.filter(
      (n) =>
        n.kind === 'class' &&
        n.name === 'Form_Expedientes' &&
        n.filePath.endsWith('.form.txt'),
    );
    expect(formFileClass).toHaveLength(0);
    // No executable-code kinds.
    const codeKinds = [
      'function',
      'module',
      'event',
      'declare',
      'type',
    ] as const;
    for (const k of codeKinds) {
      const codeNodes = r.nodes.filter((n) => n.kind === k);
      expect(codeNodes, `no ${k} nodes from form files`).toHaveLength(0);
    }
    // The synthetic `class` placeholder nodes ARE emitted for the
    // referenced tables — that's the new behavior, not a violation of
    // REQ-FORM-4 (which forbids the form's OWN class binding as a node).
    const tableClasses = r.nodes.filter(
      (n) => n.kind === 'class' && n.filePath.endsWith('.form.txt'),
    );
    expect(tableClasses.length).toBeGreaterThanOrEqual(2);
    const tableNames = tableClasses.map((n) => n.name).sort();
    expect(tableNames).toEqual(['TbEstados', 'TbExpedientes']);
  });

  it('Issue #49: RowSource at form-level (outside any control Begin block) falls back to form-layout as source', () => {
    const src = `Attribute VB_Name = "Form_Expedientes"
Begin Form
    RowSource = "SELECT Id FROM TbFoo"
End`;
    const r = extract('src/forms/Form_Expedientes.form.txt', src);
    const formNode = r.nodes.find((n) => n.kind === 'form-layout');
    expect(formNode).toBeDefined();
    const rowEdges = r.edges.filter(
      (e) =>
        e.kind === 'references' &&
        e.metadata?.synthesizedBy === 'vba-row-source',
    );
    expect(rowEdges).toHaveLength(1);
    // No form-instance-control scope → fall back to form-layout node.
    expect(rowEdges[0]?.source).toBe(formNode?.id);
    const target = r.nodes.find((n) => n.id === rowEdges[0]?.target);
    expect(target?.name).toBe('TbFoo');
  });

  it('Issue #49: same table referenced from RecordSource and RowSource collapses to ONE class node but TWO edges', () => {
    const src = `Attribute VB_Name = "Form_Expedientes"
Begin Form
    RecordSource = "TbExpedientes"
    Begin ComboBox
        Name = "cmbExpedientes"
        RowSource = "SELECT Id FROM TbExpedientes"
    End
End`;
    const r = extract('src/forms/Form_Expedientes.form.txt', src);
    const classNodes = r.nodes.filter(
      (n) =>
        n.kind === 'class' &&
        n.name === 'TbExpedientes' &&
        n.filePath.endsWith('.form.txt'),
    );
    // One placeholder class node for TbExpedientes regardless of how
    // many call sites reference it.
    expect(classNodes).toHaveLength(1);
    // But TWO edges: one vba-record-source from form-layout, one
    // vba-row-source from cmbExpedientes.
    const refs = r.edges.filter(
      (e) =>
        e.kind === 'references' &&
        e.target === classNodes[0]?.id,
    );
    expect(refs).toHaveLength(2);
    const tags = refs.map((e) => e.metadata?.synthesizedBy).sort();
    expect(tags).toEqual(['vba-record-source', 'vba-row-source']);
  });
});