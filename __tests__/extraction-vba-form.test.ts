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
    const moduleNode = r.nodes.find((n) => n.kind === 'module');
    expect(moduleNode).toBeDefined();
    expect(moduleNode?.name).toBe('Form_Main');
    // The sibling-class binding is emitted as an UnresolvedReference (so
    // REQ-FORM-4 — no class nodes from form UI — is honored). It carries
    // synthesizedBy: vba-form-binding.
    const binding = r.unresolvedReferences.find(
      (u) => u.metadata?.synthesizedBy === 'vba-form-binding',
    );
    expect(binding).toBeDefined();
    expect(binding?.referenceName).toBe('Form_Main');
    expect(binding?.fromNodeId).toBe(moduleNode?.id);
  });

  it('Form module named from filename when VB_Name absent', () => {
    const src = `Begin
    Begin TextBox
        Name = "txtField"
    End
End`;
    const r = extract('src/forms/Form_Main.form.txt', src);
    const moduleNode = r.nodes.find((n) => n.kind === 'module');
    expect(moduleNode).toBeDefined();
    expect(moduleNode?.name).toBe('Form_Main');
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
    const moduleNode = r.nodes.find((n) => n.kind === 'module');
    expect(moduleNode).toBeDefined();
    expect(moduleNode?.name).toBe('Report_Orders');
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
    const codeNodes = r.nodes.filter(
      (n) =>
        n.kind === 'function' ||
        n.kind === 'class' ||
        (n.kind === 'module' && n.name !== 'Form_Main'),
    );
    expect(codeNodes).toHaveLength(0);
    // Class nodes also forbidden by REQ-FORM-4.
    const classes = r.nodes.filter((n) => n.kind === 'class');
    expect(classes).toHaveLength(0);
  });

  it('empty form file still emits module + sibling-binding reference', () => {
    const r = extract('src/forms/Form_Main.form.txt', '');
    const moduleNode = r.nodes.find((n) => n.kind === 'module');
    expect(moduleNode).toBeDefined();
    expect(moduleNode?.name).toBe('Form_Main');
    const props = r.nodes.filter((n) => n.kind === 'property');
    expect(props).toHaveLength(0);
    const binding = r.unresolvedReferences.find(
      (u) => u.metadata?.synthesizedBy === 'vba-form-binding',
    );
    expect(binding).toBeDefined();
    expect(binding?.referenceName).toBe('Form_Main');
  });
});