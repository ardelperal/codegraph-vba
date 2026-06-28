/**
 * VbaExtractor — tests for the `.bas` / `.cls` / `.frm` / `.dsr` regex
 * extractor. Each `it()` corresponds to one or two spec scenarios from
 * `openspec/changes/vba-extractor/specs/vba-code-extraction/spec.md`.
 *
 * Coverage map (scenario id → test name):
 *   REQ-CODE-1  Public Sub in .bas                 → "Public Sub emits function with visibility"
 *   REQ-CODE-1  Private Function in .bas          → "Private Function emits function with visibility"
 *   REQ-CODE-1  Property declaration in .bas      → "Property Get emits function node"
 *   REQ-CODE-2  Method in .cls                     → "Public Function in .cls emits class+function+contains"
 *   REQ-CODE-3  Public Sub New sets marker        → "Public Sub New sets class initializer marker"
 *   REQ-CODE-3  Private Sub New sets marker       → "Private Sub New sets class initializer marker"
 *   REQ-CODE-3  Missing Sub New leaves unset      → "missing Sub New leaves hasClassInitializer unset"
 *   REQ-CODE-4  Same-file call emits plain calls  → "same-file call emits plain calls edge"
 *   REQ-CODE-4  Cross-module qualified call       → "cross-module qualified call carries synthesizedBy"
 *   REQ-CODE-4  Unresolvable call is silent        → "unresolvable call emits no edge and does not throw"
 *   REQ-CODE-5  Implements IFoo emits edge        → "Implements IFoo emits implements edge"
 *   REQ-CODE-6  Qualified Dim references outer    → "qualified Dim As references outer type"
 *   REQ-CODE-6  Unqualified Dim does not emit     → "unqualified Dim does not emit edge"
 *   REQ-CODE-7  WithEvents emits synthesized ref  → "WithEvents emits synthesized reference"
 *   REQ-CODE-8  FROM clause resolves table        → "DoCmd.RunSQL with FROM clause resolves table"
 *   REQ-CODE-8  UPDATE statement resolves table   → "CurrentDb.Execute UPDATE resolves table"
 *   REQ-CODE-8  INTO clause resolves table        → "DoCmd.RunSQL INSERT INTO resolves table"
 *   REQ-CODE-8  SQL inside VBA comment not match  → "SQL inside a VBA comment does not match"
 *   REQ-CODE-9  .form.txt input rejected          → "VbaExtractor on a .form.txt file emits zero code nodes"
 *   REQ-CODE-10 Option directives are inert       → "Option Explicit alone emits nothing"
 *   REQ-CODE-11 VB_Name attribute is used         → "Attribute VB_Name sets module name"
 *   REQ-CODE-11 Filename is used when VB_Name abs → "missing VB_Name falls back to file basename"
 */
import { describe, it, expect } from 'vitest';
import { VbaExtractor } from '../src/extraction/vba-extractor';

function extract(filePath: string, source: string) {
  return new VbaExtractor(filePath, source).extract();
}

describe('VbaExtractor — procedure declarations in .bas (REQ-CODE-1)', () => {
  it('Public Sub emits function with visibility', () => {
    const src = `Public Sub SaveRecord()
End Sub`;
    const r = extract('src/modules/modRepo.bas', src);
    const moduleNode = r.nodes.find((n) => n.kind === 'module');
    expect(moduleNode).toBeDefined();
    expect(moduleNode?.name).toBe('modRepo');
    const func = r.nodes.find((n) => n.kind === 'function' && n.name === 'SaveRecord');
    expect(func).toBeDefined();
    expect(func?.visibility).toBe('public');
    expect(func?.language).toBe('vba');
  });

  it('Private Function emits function with visibility', () => {
    const src = `Private Function CalcTotal() As Long
  CalcTotal = 1
End Function`;
    const r = extract('src/modules/modCalc.bas', src);
    const func = r.nodes.find((n) => n.kind === 'function' && n.name === 'CalcTotal');
    expect(func).toBeDefined();
    expect(func?.visibility).toBe('private');
  });

  it('Property Get emits function node', () => {
    const src = `Property Get Name() As String
  Name = "x"
End Property`;
    const r = extract('src/modules/modName.bas', src);
    const func = r.nodes.find((n) => n.kind === 'function' && n.name === 'Name');
    expect(func).toBeDefined();
    // No visibility keyword → 'public' default.
    expect(func?.visibility).toBe('public');
  });
});

describe('VbaExtractor — methods in .cls (REQ-CODE-2)', () => {
  it('Public Function in .cls emits class+function+contains edge', () => {
    const src = `Public Function Calc() As Long
  Calc = 1
End Function`;
    const r = extract('src/classes/CalcEngine.cls', src);
    const cls = r.nodes.find((n) => n.kind === 'class');
    const func = r.nodes.find((n) => n.kind === 'function' && n.name === 'Calc');
    expect(cls).toBeDefined();
    expect(cls?.name).toBe('CalcEngine');
    expect(func).toBeDefined();
    expect(func?.visibility).toBe('public');
    // contains edge from class to function
    const edge = r.edges.find((e) => e.kind === 'contains' && e.source === cls?.id && e.target === func?.id);
    expect(edge).toBeDefined();
  });
});

describe('VbaExtractor — Sub New class initializer marker (REQ-CODE-3)', () => {
  it('Public Sub New sets class initializer marker', () => {
    const src = `Public Sub New()
End Sub`;
    const r = extract('src/classes/Customer.cls', src);
    const cls = r.nodes.find((n) => n.kind === 'class');
    expect(cls).toBeDefined();
    const md = cls?.metadata as Record<string, unknown> | undefined;
    expect(md?.hasClassInitializer).toBe(true);
    expect(md?.initializerName).toBe('New');
  });

  it('Private Sub New sets class initializer marker', () => {
    const src = `Private Sub New()
End Sub`;
    const r = extract('src/classes/Internal.cls', src);
    const cls = r.nodes.find((n) => n.kind === 'class');
    const md = cls?.metadata as Record<string, unknown> | undefined;
    expect(md?.hasClassInitializer).toBe(true);
  });

  it('missing Sub New leaves hasClassInitializer unset', () => {
    const src = `Public Function DoWork() As Long
  DoWork = 1
End Function`;
    const r = extract('src/classes/Worker.cls', src);
    const cls = r.nodes.find((n) => n.kind === 'class');
    const md = cls?.metadata as Record<string, unknown> | undefined;
    expect(md?.hasClassInitializer).toBeFalsy();
    expect(md?.initializerName).toBeFalsy();
  });
});

describe('VbaExtractor — call sites (REQ-CODE-4)', () => {
  it('same-file call emits plain calls edge', () => {
    const src = `Sub Outer()
  Call Inner()
End Sub

Sub Inner()
End Sub`;
    const r = extract('src/modules/SameFile.bas', src);
    const outer = r.nodes.find((n) => n.kind === 'function' && n.name === 'Outer');
    const inner = r.nodes.find((n) => n.kind === 'function' && n.name === 'Inner');
    const edge = r.edges.find(
      (e) => e.kind === 'calls' && e.source === outer?.id && e.target === inner?.id,
    );
    expect(edge).toBeDefined();
    expect(edge?.provenance).toBeUndefined();
    expect(edge?.metadata).toBeUndefined();
  });

  it('cross-module qualified call carries synthesizedBy', () => {
    const src = `Sub RunIt()
  Call modHelpers.CalcTotal()
End Sub`;
    const r = extract('src/modules/Caller.bas', src);
    const edge = r.edges.find(
      (e) => e.kind === 'calls' && e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(edge).toBeDefined();
    expect(edge?.provenance).toBe('heuristic');
    // Target name should include the qualified receiver.
    const target = r.nodes.find((n) => n.id === edge?.target);
    expect(target?.name).toContain('modHelpers');
  });

  it('unresolvable call emits no edge and does not throw', () => {
    // A bare qualified reference (no parens) is not a call expression;
    // the CALL_RE requires `\s*\(` so this source emits zero edges.
    const src = `Sub RunIt()
  UnknownExternal.Whatever
End Sub`;
    const r = extract('src/modules/Solo.bas', src);
    // No edge should target anything starting with "UnknownExternal".
    const edgesToUnknown = r.edges.filter((e) => {
      const tgt = r.nodes.find((n) => n.id === e.target);
      return tgt?.name?.startsWith('UnknownExternal');
    });
    expect(edgesToUnknown).toHaveLength(0);
    expect(r.errors).toHaveLength(0);
  });
});

describe('VbaExtractor — Implements (REQ-CODE-5)', () => {
  it('Implements IFoo emits implements edge', () => {
    const src = `Implements IFoo

Public Sub IFoo_Do() Implements IFoo: End Sub`;
    const r = extract('src/classes/Bar.cls', src);
    const cls = r.nodes.find((n) => n.kind === 'class');
    const edge = r.edges.find(
      (e) => e.kind === 'implements' && e.source === cls?.id,
    );
    expect(edge).toBeDefined();
    const target = r.nodes.find((n) => n.id === edge?.target);
    expect(target?.name).toBe('IFoo');
  });
});

describe('VbaExtractor — qualified Dim (REQ-CODE-6)', () => {
  it('qualified Dim As references outer type', () => {
    const src = `Dim m_Calc As CalcEngine.Helper`;
    const r = extract('src/modules/UseCalc.bas', src);
    const edge = r.edges.find(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(edge).toBeDefined();
    const target = r.nodes.find((n) => n.id === edge?.target);
    expect(target?.name).toBe('CalcEngine');
  });

  it('unqualified Dim does not emit edge', () => {
    const src = `Dim m_Count As Long`;
    const r = extract('src/modules/Counter.bas', src);
    const refEdges = r.edges.filter((e) => e.kind === 'references');
    expect(refEdges).toHaveLength(0);
  });
});

describe('VbaExtractor — WithEvents (REQ-CODE-7)', () => {
  it('WithEvents emits synthesized reference', () => {
    const src = `WithEvents m_Form As Form_Main`;
    const r = extract('src/classes/FormListener.cls', src);
    const edge = r.edges.find(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-withevents',
    );
    expect(edge).toBeDefined();
    const target = r.nodes.find((n) => n.id === edge?.target);
    expect(target?.name).toBe('Form_Main');
  });
});

describe('VbaExtractor — SQL in strings (REQ-CODE-8)', () => {
  it('DoCmd.RunSQL with FROM clause resolves table', () => {
    const src = `Sub Q()
  DoCmd.RunSQL "SELECT * FROM tblCustomers"
End Sub`;
    const r = extract('src/modules/Q.bas', src);
    const edge = r.edges.find(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-sql-table',
    );
    expect(edge).toBeDefined();
    const target = r.nodes.find((n) => n.id === edge?.target);
    expect(target?.name).toBe('tblCustomers');
  });

  it('CurrentDb.Execute UPDATE resolves table', () => {
    const src = `Sub U()
  CurrentDb.Execute "UPDATE tblOrders SET Status = 1"
End Sub`;
    const r = extract('src/modules/U.bas', src);
    const edge = r.edges.find(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-sql-table',
    );
    expect(edge).toBeDefined();
    const target = r.nodes.find((n) => n.id === edge?.target);
    expect(target?.name).toBe('tblOrders');
  });

  it('DoCmd.RunSQL INSERT INTO resolves table', () => {
    const src = `Sub A()
  DoCmd.RunSQL "INSERT INTO tblAudit (Id) VALUES (1)"
End Sub`;
    const r = extract('src/modules/A.bas', src);
    const edge = r.edges.find(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-sql-table',
    );
    expect(edge).toBeDefined();
    const target = r.nodes.find((n) => n.id === edge?.target);
    expect(target?.name).toBe('tblAudit');
  });

  it('SQL inside a VBA comment does not match', () => {
    const src = `' DoCmd.RunSQL "SELECT * FROM tblFake"
Public Sub DoWork()
End Sub`;
    const r = extract('src/modules/Commented.bas', src);
    const target = r.nodes.find((n) => n.kind === 'file' || n.name === 'tblFake');
    // No reference target should be named "tblFake".
    const fakeNodes = r.nodes.filter((n) => n.name === 'tblFake');
    expect(fakeNodes).toHaveLength(0);
    const fakeEdges = r.edges.filter((e) => {
      const tgt = r.nodes.find((n) => n.id === e.target);
      return tgt?.name === 'tblFake';
    });
    expect(fakeEdges).toHaveLength(0);
  });
});

describe('VbaExtractor — .form.txt rejection (REQ-CODE-9)', () => {
  it('emits zero function/class/module nodes when given a .form.txt input', () => {
    const src = `Sub Form_Load()
End Sub`;
    const r = extract('src/forms/Form_X.form.txt', src);
    const codeNodes = r.nodes.filter(
      (n) => n.kind === 'function' || n.kind === 'class' || n.kind === 'module',
    );
    expect(codeNodes).toHaveLength(0);
  });
});

describe('VbaExtractor — Option directives are inert (REQ-CODE-10)', () => {
  it('Option Explicit alone emits nothing beyond the file node', () => {
    const src = `Option Explicit
Option Compare Database`;
    const r = extract('src/modules/Empty.bas', src);
    // No module/class/function nodes emitted for an option-directive-only file.
    const symbols = r.nodes.filter((n) =>
      n.kind === 'function' || n.kind === 'class' || n.kind === 'module',
    );
    expect(symbols).toHaveLength(0);
  });
});

describe('VbaExtractor — VB_Name attribute (REQ-CODE-11)', () => {
  it('Attribute VB_Name sets module name', () => {
    const src = `Attribute VB_Name = "modHelpers"
Public Sub DoThing()
End Sub`;
    const r = extract('src/modules/something.bas', src);
    const moduleNode = r.nodes.find((n) => n.kind === 'module');
    expect(moduleNode).toBeDefined();
    expect(moduleNode?.name).toBe('modHelpers');
  });

  it('missing VB_Name falls back to file basename', () => {
    const src = `Public Sub DoThing()
End Sub`;
    const r = extract('src/modules/modHelpers.bas', src);
    const moduleNode = r.nodes.find((n) => n.kind === 'module');
    expect(moduleNode).toBeDefined();
    expect(moduleNode?.name).toBe('modHelpers');
  });
});

/**
 * C1 invariant — node `startLine` MUST align with the original source line.
 *
 * Audit finding C1 (June 2026): preprocessing collapsed lines
 * (`joinLineContinuations` removed newlines, `stripVbaComments` dropped
 * Option/Rem lines with `continue`), so every emitted node's `startLine`
 * drifted further from its real position as the file grew. Live probe
 * against `ACAuditoriaOperaciones.cls` showed functions pointing at
 * blank lines or unrelated declarations.
 *
 * Fix: the pre-processing helpers now preserve line count (empty-string
 * placeholders + newline retention). These tests assert the end-to-end
 * invariant: every emitted node's `startLine` lands on the same line in
 * the ORIGINAL source as its `startLine` value points to in the
 * transformed source.
 */
describe('VbaExtractor — startLine aligns with original source (C1 invariant)', () => {
  it('a function preceded by Option directives has startLine on its declaration', () => {
    // Pattern every real Dysflow-exported .bas/.cls opens with.
    const src = [
      'Attribute VB_Name = "modHelpers"',                       // 1
      '',                                                      // 2
      'Option Compare Database',                               // 3
      'Option Explicit',                                       // 4
      '',                                                      // 5
      "' Public API",                                          // 6
      'Public Function Helper() As Long',                      // 7
      '    Helper = 42',                                       // 8
      'End Function',                                          // 9
    ].join('\n');
    const r = extract('src/modules/modHelpers.bas', src);
    const func = r.nodes.find((n) => n.kind === 'function' && n.name === 'Helper');
    expect(func).toBeDefined();
    // `Public Function Helper() As Long` is on line 7 of the original.
    // startLine must point to that line — not to a blank or unrelated
    // line caused by preprocessing drift.
    expect(func?.startLine).toBe(7);
  });

  it('a function preceded by Rem comment block has startLine on its declaration', () => {
    const src = [
      'Attribute VB_Name = "modHelpers"',                       // 1
      '',                                                      // 2
      'Rem ============================================',     // 3
      'Rem Módulo: helpers',                                   // 4
      'Rem ============================================',     // 5
      '',                                                      // 6
      'Public Function Helper() As Long',                      // 7
      '    Helper = 42',                                       // 8
      'End Function',                                          // 9
    ].join('\n');
    const r = extract('src/modules/modHelpers.bas', src);
    const func = r.nodes.find((n) => n.kind === 'function' && n.name === 'Helper');
    expect(func).toBeDefined();
    expect(func?.startLine).toBe(7);
  });

  it('a function using a line continuation has startLine on its declaration', () => {
    const src = [
      'Attribute VB_Name = "modHelpers"',                       // 1
      '',                                                      // 2
      'Option Compare Database',                               // 3
      'Option Explicit',                                       // 4
      '',                                                      // 5
      'Public Function Helper() As Long',                      // 6
      '    Helper = 1 _',                                      // 7
      '        + 2 _',                                         // 8
      '        + 3',                                           // 9
      'End Function',                                          // 10
    ].join('\n');
    const r = extract('src/modules/modHelpers.bas', src);
    const func = r.nodes.find((n) => n.kind === 'function' && n.name === 'Helper');
    expect(func).toBeDefined();
    // `Public Function Helper() As Long` is on line 6 of the original,
    // even though the body uses line continuations on lines 7-9.
    expect(func?.startLine).toBe(6);
  });
});

/**
 * W5 invariant — non-ASCII identifiers (Spanish VBA: Módulo, Cálculo,
 * Señal) MUST be matched. Audit finding W5 (June 2026): the original
 * regex used `[A-Za-z_]\w*` (ASCII-only), so unicode identifiers were
 * silently truncated. Fix: Unicode-aware classes `\p{L}[\p{L}\p{N}_]*`
 * with the `/u` flag on every regex that matches an identifier.
 */
describe('VbaExtractor — Unicode identifier handling (W5 invariant)', () => {
  it('extracts a function with a Spanish identifier', () => {
    const src = `Public Function Cálculo() As Long
    Cálculo = 42
End Function`;
    const r = extract('src/modules/modCalculo.bas', src);
    const func = r.nodes.find((n) => n.kind === 'function' && n.name === 'Cálculo');
    expect(func).toBeDefined();
    expect(func?.name).toBe('Cálculo');
  });

  it('extracts a Sub with an accented identifier', () => {
    const src = `Public Sub Módulo_Iniciar()
End Sub`;
    const r = extract('src/modules/modInit.bas', src);
    const func = r.nodes.find((n) => n.kind === 'function' && n.name === 'Módulo_Iniciar');
    expect(func).toBeDefined();
  });

  it('extracts a Sub with a ñ character (Módulo1.bas fixture)', () => {
    const src = `Public Sub Módulo1()
End Sub`;
    const r = extract('src/modules/Módulo1.bas', src);
    const func = r.nodes.find((n) => n.kind === 'function' && n.name === 'Módulo1');
    expect(func).toBeDefined();
  });

  it('extracts an Implements with an accented interface name', () => {
    const src = `VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MiClase"
Attribute VB_GlobalNameSpace = False
Implements INotificación`;
    const r = extract('src/classes/MiClase.cls', src);
    // The class emits from VB_Name; the Implements target is referenced.
    const implementsEdges = r.edges.filter((e) => e.kind === 'implements');
    expect(implementsEdges.length).toBeGreaterThan(0);
    // The referenced target name should be INotificación, not truncated.
    const targetNode = r.nodes.find((n) => implementsEdges.some((e) => e.target === n.id));
    expect(targetNode?.name).toBe('INotificación');
  });

  it('matches a Spanish-named SQL table inside a string', () => {
    const src = `Sub Q()
    DoCmd.RunSQL "SELECT * FROM tblÓrdenes"
End Sub`;
    const r = extract('src/modules/Q.bas', src);
    const edge = r.edges.find(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-sql-table',
    );
    expect(edge).toBeDefined();
    const target = r.nodes.find((n) => n.id === edge?.target);
    expect(target?.name).toBe('tblÓrdenes');
  });
});