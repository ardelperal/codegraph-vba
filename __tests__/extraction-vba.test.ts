/**
 * VbaExtractor — tests for the `.bas` / `.cls` / `.frm` / `.dsr` regex
 * extractor. Each `it()` corresponds to one or two spec scenarios from
 * `openspec/specs/vba-code-extraction/spec.md`.
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

  it('unresolvable qualified call (no parens) emits NO edge when receiver is not a local declared variable (Fix 2)', () => {
    // After Fix 2 (gated qualified stmt calls): a receiver that is not
    // declared as a file-local Dim/Private/Public variable is SILENT —
    // aligns with REQ-CODE-4 "Unresolvable call is silent".
    // The previous Fix 7 behavior (always emit) is now restricted to
    // receivers declared as file-local variables typed as project classes.
    const src = `Sub RunIt()
  UnknownExternal.Whatever
End Sub`;
    const r = extract('src/modules/Solo.bas', src);
    // UnknownExternal is not declared in the file → no edge emitted.
    const edgesToUnknown = r.edges.filter((e) => {
      const tgt = r.nodes.find((n) => n.id === e.target);
      return tgt?.name?.startsWith('UnknownExternal');
    });
    expect(edgesToUnknown).toHaveLength(0);
    // No extraction errors.
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

  it('getdb().OpenRecordset with SQL assigned to a variable resolves the table', () => {
    const src = `Sub Q()
  Dim m_SQL As String
  m_SQL = "SELECT * FROM tblCustomers " & _
          "WHERE Id = " & customerId
  Set rs = getdb().OpenRecordset(m_SQL)
End Sub`;
    const r = extract('src/modules/Q.bas', src);
    const edge = r.edges.find(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-sql-table',
    );
    expect(edge).toBeDefined();
    const target = r.nodes.find((n) => n.id === edge?.target);
    expect(target?.name).toBe('tblCustomers');
  });

  it('getdb().Execute with SQL assigned to a variable resolves the table', () => {
    const src = `Sub U()
  Dim m_SQL As String
  m_SQL = "UPDATE tblOrders " & _
          "SET Status = 1"
  getdb().Execute m_SQL
End Sub`;
    const r = extract('src/modules/U.bas', src);
    const edge = r.edges.find(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-sql-table',
    );
    expect(edge).toBeDefined();
    const target = r.nodes.find((n) => n.id === edge?.target);
    expect(target?.name).toBe('tblOrders');
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

  it('real Dysflow class fixtures emit SQL table edges for getdb() variable SQL', () => {
    const fs = require('node:fs') as typeof import('node:fs');
    const path = require('node:path') as typeof import('node:path');
    const fixtures = [
      path.resolve(__dirname, '..', '__tests__', 'fixtures', 'vba', 'src', 'classes', 'ACAuditoriaOperaciones.cls'),
      path.resolve(__dirname, '..', '__tests__', 'fixtures', 'vba', 'src', 'classes', 'ARAuditoria.cls'),
    ];
    const sqlEdges = fixtures.flatMap((fixture) => {
      if (!fs.existsSync(fixture)) return [];
      const r = extract(fixture, fs.readFileSync(fixture, 'utf8'));
      return r.edges
        .filter((e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-sql-table')
        .map((e) => ({ edge: e, nodes: r.nodes }));
    });
    expect(sqlEdges.length).toBeGreaterThanOrEqual(5);
    const targetNames = new Set(
      sqlEdges
        .map(({ edge, nodes }) => nodes.find((n) => n.id === edge.target)?.name)
        .filter(Boolean),
    );
    expect(targetNames).toContain('TbNCAuditoriaAccionCorrectivas');
    expect(targetNames).toContain('TbNCAuditoriaAccionesRealizadas');
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

/**
 * W2 invariant — `detectVbName` MUST walk past Access class metadata
 * headers (VERSION / BEGIN / MultiUse / END / Attribute …) to find the
 * VB_Name attribute. Audit W2 (June 2026): the previous implementation
 * returned null at the first non-Attribute line, so real Access .cls
 * files always fell through to the basename fallback.
 */
describe('VbaExtractor — detectVbName skips class metadata header (W2 invariant)', () => {
  it('a .cls with the standard Access header uses VB_Name as the class name', () => {
    const src = `VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ACAuditoriaOperaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IAuditable`;
    const r = extract('src/classes/ACAuditoriaOperaciones.cls', src);
    const cls = r.nodes.find((n) => n.kind === 'class');
    expect(cls).toBeDefined();
    // Class name MUST come from VB_Name (not the file basename).
    expect(cls?.name).toBe('ACAuditoriaOperaciones');
  });

  it('a .bas with VB_Name on the first non-empty line still works (regression)', () => {
    const src = `Attribute VB_Name = "modHelpers"
Option Explicit
Public Function DoThing()
End Function`;
    const r = extract('src/modules/something.bas', src);
    const mod = r.nodes.find((n) => n.kind === 'module');
    expect(mod?.name).toBe('modHelpers');
  });
});

/**
 * W4 invariant — calls on Access runtime objects (Me, DoCmd, Forms,
 * Application, etc.) MUST NOT synthesize a `function` node for the
 * receiver. Audit W4 (June 2026): one real-world .cls produced ~20 junk
 * `function` nodes (`rcdDatos.Fields`, `getdb().OpenRecordset`, …)
 * polluting search/explore output. The fix is a runtime-receiver
 * blacklist applied BEFORE synthesis.
 *
 * Note: `DoCmd.RunSQL` (and friends) still get SQL edge tracking via
 * `SQL_WRAPPERS` (REQ-CODE-8) — that path is independent of this one.
 */
describe('VbaExtractor — runtime receivers do not synthesize nodes (W4 invariant)', () => {
  it('Me.Refresh() does not emit a synthetic function node', () => {
    const src = `Public Sub X()
    Me.Refresh
End Sub`;
    const r = extract('src/forms/Form_X.cls', src);
    const synthFns = r.nodes.filter(
      (n) => n.kind === 'function' && (n.name === 'Me.Refresh' || n.name.includes('Me.')),
    );
    expect(synthFns).toHaveLength(0);
  });

  it('DoCmd.OpenForm does not emit a synthetic function node', () => {
    const src = `Public Sub X()
    DoCmd.OpenForm "MyForm"
End Sub`;
    const r = extract('src/modules/X.bas', src);
    const synthFns = r.nodes.filter(
      (n) => n.kind === 'function' && n.name.includes('DoCmd.'),
    );
    expect(synthFns).toHaveLength(0);
  });

  it('Forms!MyForm.Open does not emit a synthetic function node', () => {
    const src = `Public Sub X()
    Forms!MyForm.Visible = True
End Sub`;
    const r = extract('src/modules/X.bas', src);
    const synthFns = r.nodes.filter(
      (n) => n.kind === 'function' && n.name.includes('Forms.'),
    );
    expect(synthFns).toHaveLength(0);
  });

  it('DoCmd.RunSQL still emits a vba-sql-table edge (regression — W4 must not break REQ-CODE-8)', () => {
    const src = `Public Sub X()
    DoCmd.RunSQL "DELETE FROM tblOld"
End Sub`;
    const r = extract('src/modules/X.bas', src);
    const sqlEdges = r.edges.filter((e) => e.metadata?.synthesizedBy === 'vba-sql-table');
    expect(sqlEdges.length).toBeGreaterThan(0);
    // But NO synthetic DoCmd.RunSQL function node should be emitted.
    const synthFns = r.nodes.filter(
      (n) => n.kind === 'function' && n.name === 'DoCmd.RunSQL',
    );
    expect(synthFns).toHaveLength(0);
  });

  it('Application.StatusBar still does not emit a synthetic function node', () => {
    const src = `Public Sub X()
    Application.StatusBar = "Working..."
End Sub`;
    const r = extract('src/modules/X.bas', src);
    const synthFns = r.nodes.filter(
      (n) => n.kind === 'function' && n.name.includes('Application.'),
    );
    expect(synthFns).toHaveLength(0);
  });

  it('DAO recordset Fields access does not emit a synthetic function node', () => {
    const src = `Public Sub X()
    rcdDatos.Fields("ID") = 1
End Sub`;
    const r = extract('src/modules/X.bas', src);
    const synthFns = r.nodes.filter(
      (n) => n.kind === 'function' && n.name === 'rcdDatos.Fields',
    );
    expect(synthFns).toHaveLength(0);
  });
});

/**
 * W6 invariant — Property Get/Let/Set with the same name MUST all be
 * tracked. Audit W6 (June 2026): the previous Map<string, ProcInfo>
 * keyed by bare name kept only the last accessor, breaking same-file
 * call resolution. The fix is a multimap (Map<string, ProcInfo[]>).
 */
describe('VbaExtractor — Property Get/Let/Set coexistence (W6 invariant)', () => {
  it('all three accessors for the same property name are emitted', () => {
    const src = `VERSION 1.0 CLASS
BEGIN
END
Attribute VB_Name = "Thing"

Public Property Get Name() As String
    Name = "x"
End Property

Public Property Let Name(v As String)
End Property

Public Property Set Name(v As Object)
End Property`;
    const r = extract('src/classes/Thing.cls', src);
    // The class emits 3 function nodes, one per accessor.
    const nameFns = r.nodes.filter((n) => n.kind === 'function' && n.name === 'Name');
    expect(nameFns.length).toBe(3);
  });
});

/**
 * S3 invariant — `Dim x As SomeType` (unqualified, no dot) MUST emit a
 * `references` edge to `SomeType` when `SomeType` is not a primitive.
 * Audit S3 (June 2026): the previous implementation required a `.` in
 * the type (DIM_QUAL_RE), so `Dim AC As ACAuditoria` emitted no edge
 * and class→class flows were invisible — defeating the tool's value
 * for the dominant VBA dependency form.
 */
describe('VbaExtractor — unqualified Dim emits a class reference (S3 invariant)', () => {
  it('Dim x As UserClass emits a vba-name-resolution reference to UserClass', () => {
    const src = `Public AC As ACAuditoria`;
    const r = extract('src/classes/Container.cls', src);
    const edge = r.edges.find(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(edge).toBeDefined();
    const target = r.nodes.find((n) => n.id === edge?.target);
    expect(target?.name).toBe('ACAuditoria');
  });

  it('Dim x As Long (primitive) does NOT emit a reference edge', () => {
    const src = `Public Count As Long`;
    const r = extract('src/modules/Mod.bas', src);
    const edges = r.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(edges).toHaveLength(0);
  });

  it('Dim x As String (primitive) does NOT emit a reference edge', () => {
    const src = `Public Name As String`;
    const r = extract('src/modules/Mod.bas', src);
    const edges = r.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(edges).toHaveLength(0);
  });

  it('Dim x As Variant (primitive) does NOT emit a reference edge', () => {
    const src = `Public Data As Variant`;
    const r = extract('src/modules/Mod.bas', src);
    const edges = r.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(edges).toHaveLength(0);
  });

  it('Dim x As SomeType with Spanish identifier emits a reference', () => {
    const src = `Public Notif As INotificación`;
    const r = extract('src/modules/Mod.bas', src);
    const edge = r.edges.find(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(edge).toBeDefined();
    const target = r.nodes.find((n) => n.id === edge?.target);
    expect(target?.name).toBe('INotificación');
  });
});

/**
 * S4 invariant — `Implements IFoo` MUST use the `parser` provenance,
 * NOT `heuristic`. Audit S4 (June 2026): the previous implementation
 * tagged the static, source-declared edge as `heuristic` — but
 * `heuristic` is reserved for guessed/inferred edges, which Implements
 * is not. The fix adds a new `parser` provenance value (generalizes
 * `tree-sitter` for non-tree-sitter extractors) and uses it here.
 */
describe('VbaExtractor — Implements edge uses parser provenance (S4 invariant)', () => {
  it('Implements edge has provenance === "parser", not "heuristic"', () => {
    const src = `VERSION 1.0 CLASS
BEGIN
END
Attribute VB_Name = "MiClase"
Implements IAuditable`;
    const r = extract('src/classes/MiClase.cls', src);
    const implEdge = r.edges.find((e) => e.kind === 'implements');
    expect(implEdge).toBeDefined();
    expect(implEdge?.provenance).toBe('parser');
    expect(implEdge?.provenance).not.toBe('heuristic');
  });
});

/**
 * S5 invariant — single-line colon-separated procedure declarations
 * (e.g. `Public Sub X(): End Sub`) MUST end the procedure for the
 * purpose of proc-stack tracking. Audit S5 (June 2026): the previous
 * procedureEndRe was anchored at line start, so the proc stack never
 * popped for these declarations and subsequent lines were treated as
 * still inside the procedure.
 */
describe('VbaExtractor — colon-separated single-line procedures (S5 invariant)', () => {
  it('a single-line Public Sub X(): End Sub is recognized as ending at the colon', () => {
    const src = `Public Sub One(): End Sub
Public Sub Two()
    Debug.Print "inside two"
End Sub`;
    const r = extract('src/modules/Mod.bas', src);
    // Both procedures emit function nodes.
    const procs = r.nodes.filter(
      (n) => n.kind === 'function' && (n.name === 'One' || n.name === 'Two'),
    );
    expect(procs).toHaveLength(2);
    // The Debug.Print inside `Two` should emit NO calls edge to a
    // (nonexistent) "One" function — which would happen if the proc
    // stack incorrectly kept `One` pushed because the End Sub was
    // never recognized.
    const calls = r.edges.filter((e) => e.kind === 'calls');
    expect(calls).toHaveLength(0);
  });

  it('a single-line Function F(): End Function is recognized', () => {
    const src = `Public Function Add(): End Function`;
    const r = extract('src/modules/Mod.bas', src);
    const fn = r.nodes.find((n) => n.kind === 'function' && n.name === 'Add');
    expect(fn).toBeDefined();
  });
});

/**
 * C2 invariant — every emitted `function` node MUST have `endLine > startLine`
 * (or equal for the single-line colon-separated case) so `codegraph_explore`
 * returns the full body, not just the signature line.
 *
 * Audit C2 (June 2026): the previous implementation set `endLine = lineNum`
 * (same as `startLine`), so explore returned only the signature and the
 * agent fell back to Read for every procedure.
 */
describe('VbaExtractor — function nodes carry the full body span (C2 invariant)', () => {
  it('a multi-line Public Function has endLine on the End Function line', () => {
    const src = [
      'Public Function Calc() As Long',                // 1 — startLine = 1
      '    Dim x As Long',                            // 2
      '    x = 1',                                    // 3
      '    Calc = x * 2',                             // 4
      'End Function',                                 // 5 — endLine = 5
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const fn = r.nodes.find((n) => n.kind === 'function' && n.name === 'Calc');
    expect(fn).toBeDefined();
    expect(fn?.startLine).toBe(1);
    expect(fn?.endLine).toBe(5);
    expect(fn?.endLine).toBeGreaterThan(fn?.startLine ?? 0);
  });

  it('a multi-line Sub has endLine on the End Sub line (span covers body)', () => {
    const src = [
      "Public Sub DoWork()",                          // 1
      "    Debug.Print \"a\"",                        // 2
      "    Debug.Print \"b\"",                        // 3
      "    Debug.Print \"c\"",                        // 4
      "End Sub",                                       // 5
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const fn = r.nodes.find((n) => n.kind === 'function' && n.name === 'DoWork');
    expect(fn).toBeDefined();
    expect(fn?.startLine).toBe(1);
    expect(fn?.endLine).toBe(5);
  });

  it('a single-line colon-separated Sub has endLine == startLine (no body to span)', () => {
    const src = `Public Sub One(): End Sub`;
    const r = extract('src/modules/Mod.bas', src);
    const fn = r.nodes.find((n) => n.kind === 'function' && n.name === 'One');
    expect(fn).toBeDefined();
    expect(fn?.startLine).toBe(1);
    expect(fn?.endLine).toBe(1);
  });

  it('a real fixture procedure has endLine > startLine (recall probe invariant)', () => {
    // Index the REAL Dysflow fixture — every REAL procedure there must
    // have a body span, not just a signature line. (Synthetic function
    // nodes from qualified-call synthesis are excluded — they represent
    // call-site expressions, not declarations; that's M1's concern.)
    const fs = require('node:fs') as typeof import('node:fs');
    const path = require('node:path') as typeof import('node:path');
    const fix = path.resolve(
      __dirname,
      '..',
      '__tests__',
      'fixtures',
      'vba',
      'src',
      'classes',
      'ACAuditoriaOperaciones.cls',
    );
    if (!fs.existsSync(fix)) return; // skip if fixture moved
    const r = extract(fix, fs.readFileSync(fix, 'utf8'));
    // Real procedure nodes have a bare name (no `.`) — synthesized
    // qualified-call nodes do.
    const realProcs = r.nodes.filter(
      (n) => n.kind === 'function' && !n.name.includes('.'),
    );
    expect(realProcs.length).toBeGreaterThan(5);
    const noBody = realProcs.filter((n) => n.endLine === n.startLine);
    expect(noBody).toHaveLength(0);
  });
});

/**
 * H1 invariant — statement-form Sub calls (no parens, no `Call` keyword)
 * MUST emit same-file `calls` edges. Audit H1 (June 2026): the CALL_RE
 * only matched the parens form, so the dominant VBA idiom
 * (`EstablecerDatos m_Error` at statement position) was invisible. On
 * the real form fixture, this dropped recall from "should be high" to
 * near-zero for plain (non-qualified) Sub calls.
 */
describe('VbaExtractor — statement-form Sub calls (H1 invariant)', () => {
  it('a bare statement call without parens emits a same-file calls edge', () => {
    const src = [
      'Public Sub Outer()',
      '    Inner 1, 2',
      'End Sub',
      '',
      'Public Sub Inner(a As Long, b As Long)',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const calls = r.edges.filter((e) => e.kind === 'calls');
    expect(calls.length).toBeGreaterThanOrEqual(1);
    const callEdge = calls.find((e) => {
      const t = r.nodes.find((n) => n.id === e.target);
      return t?.name === 'Inner';
    });
    expect(callEdge).toBeDefined();
    const source = r.nodes.find((n) => n.id === callEdge?.source);
    expect(source?.name).toBe('Outer');
  });

  it('Call Sub() (with Call keyword AND parens) is caught by the existing CALL_RE', () => {
    const src = [
      'Public Sub Outer()',
      '    Call Inner(1, 2)',
      'End Sub',
      'Public Sub Inner(a As Long, b As Long)',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const calls = r.edges.filter((e) => e.kind === 'calls');
    expect(calls.length).toBeGreaterThanOrEqual(1);
  });

  it('Call Sub (with Call keyword, no parens) emits a calls edge', () => {
    const src = [
      'Public Sub Outer()',
      '    Call Inner 1, 2',
      'End Sub',
      'Public Sub Inner(a As Long, b As Long)',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const calls = r.edges.filter((e) => e.kind === 'calls');
    expect(calls.length).toBeGreaterThanOrEqual(1);
  });

  it('a bare no-argument statement call emits a same-file calls edge', () => {
    const src = [
      'Public Sub Outer()',
      '    Inner',
      'End Sub',
      'Public Sub Inner()',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const callEdge = r.edges.find((e) => {
      if (e.kind !== 'calls') return false;
      const source = r.nodes.find((n) => n.id === e.source);
      const target = r.nodes.find((n) => n.id === e.target);
      return source?.name === 'Outer' && target?.name === 'Inner';
    });
    expect(callEdge).toBeDefined();
  });

  it('a statement call with parentheses inside argument expressions emits a calls edge', () => {
    const src = [
      'Public Sub Outer()',
      '    Inner Nz(x, 0), 2',
      'End Sub',
      'Public Sub Inner(a As Long, b As Long)',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const callEdge = r.edges.find((e) => {
      if (e.kind !== 'calls') return false;
      const source = r.nodes.find((n) => n.id === e.source);
      const target = r.nodes.find((n) => n.id === e.target);
      return source?.name === 'Outer' && target?.name === 'Inner';
    });
    expect(callEdge).toBeDefined();
  });

  it('a normal parens-form call is not double-counted by statement-form detection', () => {
    const src = [
      'Public Sub Outer()',
      '    Inner(1, 2)',
      'End Sub',
      'Public Sub Inner(a As Long, b As Long)',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const innerCalls = r.edges.filter((e) => {
      if (e.kind !== 'calls') return false;
      const target = r.nodes.find((n) => n.id === e.target);
      return target?.name === 'Inner';
    });
    expect(innerCalls).toHaveLength(1);
  });

  it('an assignment (X = ...) does NOT trigger a statement call', () => {
    const src = [
      'Public Sub Outer()',
      '    x = 1',
      'End Sub',
      'Public Sub Inner(a As Long, b As Long)',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const calls = r.edges.filter((e) => e.kind === 'calls');
    expect(calls).toHaveLength(0);
  });

  it('a declaration (Dim x As Foo) does NOT trigger a statement call', () => {
    const src = [
      'Public Sub Outer()',
      '    Dim x As Long',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const calls = r.edges.filter((e) => e.kind === 'calls');
    expect(calls).toHaveLength(0);
  });

  it('a real Dysflow form fixture has same-file calls edges from the statement-form pattern', () => {
    // The form fixture has lines like `EstablecerDatos m_Error` and
    // `CorreoAlAdministrador m_Error` — bare statement calls. Verify
    // these get captured.
    const fs = require('node:fs') as typeof import('node:fs');
    const path = require('node:path') as typeof import('node:path');
    const fix = path.resolve(
      __dirname,
      '..',
      '__tests__',
      'fixtures',
      'vba',
      'src',
      'forms',
      'Form_FormNCAuditoriaMotivoEliminado.cls',
    );
    if (!fs.existsSync(fix)) return;
    const r = extract(fix, fs.readFileSync(fix, 'utf8'));
    const statementCalls = r.edges.filter((e) => {
      if (e.kind !== 'calls') return false;
      const source = r.nodes.find((n) => n.id === e.source);
      const t = r.nodes.find((n) => n.id === e.target);
      return (
        (e.line === 15 && source?.name === 'Form_Load' && t?.name === 'EstablecerDatos') ||
        (e.line === 70 && source?.name === 'EstablecerDatos' && t?.name === 'EstableceColorBordes') ||
        (e.line === 109 && source?.name === 'MotivoBorrado_AfterUpdate' && t?.name === 'EstableceColorBordes') ||
        (e.line === 180 && source?.name === 'ComandoGrabar_Click' && t?.name === 'RellenarDatosObjeto') ||
        (e.line === 195 && source?.name === 'ComandoGrabar_Click' && t?.name === 'EstablecerDatos')
      );
    });
    expect(statementCalls).toHaveLength(5);
  });
});

// ---------------------------------------------------------------------------
// Fix 1: Property Get/Let/Set endLine + caller attribution
// ---------------------------------------------------------------------------

/**
 * Fix 1: when Property Get and Property Let/Set share the same name, the
 * proc-stack in `sweepCallsAndSql` was always pushing `bucket[0]` (the Get's
 * ProcInfo) regardless of which accessor's declaration line we were on. This
 * caused the wrong `endLine` to be set (the second `End Property` updated the
 * Get node instead of the Let/Set node) and calls inside the Let/Set body to
 * be attributed to the Get node.
 */
describe('VbaExtractor — Property Get/Let/Set endLine attribution (Fix 1)', () => {
  it('Property Get and Property Set with the same name have non-overlapping endLine spans', () => {
    const src = [
      'Property Get Documentos() As String',  // 1
      '  Documentos = m_str',                 // 2
      'End Property',                          // 3
      '',                                      // 4
      'Property Set Documentos(v As String)', // 5
      '  m_str = v',                           // 6
      'End Property',                          // 7
    ].join('\n');
    const r = extract('src/classes/X.cls', src);
    const getNode = r.nodes.find(
      (n) => n.kind === 'function' && n.name === 'Documentos' && n.startLine === 1,
    );
    const setNode = r.nodes.find(
      (n) => n.kind === 'function' && n.name === 'Documentos' && n.startLine === 5,
    );
    expect(getNode).toBeDefined();
    expect(setNode).toBeDefined();
    // Each accessor must end on its own `End Property` line.
    expect(getNode?.endLine).toBe(3);
    expect(setNode?.endLine).toBe(7);
    // Spans must not overlap.
    expect(getNode?.endLine ?? 0).toBeLessThan(setNode?.startLine ?? 999);
  });

  it('Property Get/Let/Set spans are correct for all three accessors', () => {
    const src = [
      'Property Get Val() As Long',           // 1
      '  Val = m_val',                         // 2
      'End Property',                           // 3
      '',                                       // 4
      'Property Let Val(v As Long)',            // 5
      '  m_val = v',                            // 6
      'End Property',                           // 7
      '',                                       // 8
      'Property Set Val(v As Object)',          // 9
      '  Set m_val = v',                        // 10
      'End Property',                           // 11
    ].join('\n');
    const r = extract('src/classes/X.cls', src);
    const getN = r.nodes.find((n) => n.kind === 'function' && n.name === 'Val' && n.startLine === 1);
    const letN = r.nodes.find((n) => n.kind === 'function' && n.name === 'Val' && n.startLine === 5);
    const setN = r.nodes.find((n) => n.kind === 'function' && n.name === 'Val' && n.startLine === 9);
    expect(getN?.endLine).toBe(3);
    expect(letN?.endLine).toBe(7);
    expect(setN?.endLine).toBe(11);
  });

  it('calls inside Property Let body attribute to the Let node, not the Get node', () => {
    const src = [
      'Property Get Val() As Long',            // 1
      '  Val = m_val',                          // 2
      'End Property',                            // 3
      '',                                        // 4
      'Property Let Val(v As Long)',             // 5
      '  Helper',                                // 6  bare statement call
      'End Property',                            // 7
      '',                                        // 8
      'Private Sub Helper()',                    // 9
      'End Sub',                                 // 10
    ].join('\n');
    const r = extract('src/classes/X.cls', src);
    const letNode = r.nodes.find(
      (n) => n.kind === 'function' && n.name === 'Val' && n.startLine === 5,
    );
    const helperNode = r.nodes.find((n) => n.kind === 'function' && n.name === 'Helper');
    expect(letNode).toBeDefined();
    expect(helperNode).toBeDefined();
    // The calls edge source MUST be the Let node, not the Get node.
    const callEdge = r.edges.find(
      (e) => e.kind === 'calls' && e.source === letNode?.id && e.target === helperNode?.id,
    );
    expect(callEdge).toBeDefined();
    // And no calls edge should have source === Get node.
    const getNode = r.nodes.find(
      (n) => n.kind === 'function' && n.name === 'Val' && n.startLine === 1,
    );
    const wrongCallEdge = r.edges.find(
      (e) => e.kind === 'calls' && e.source === getNode?.id && e.target === helperNode?.id,
    );
    expect(wrongCallEdge).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// Fix 2: Duplicate `references` edge for qualified `Dim x As Foo.Bar`
// ---------------------------------------------------------------------------

/**
 * Fix 2: `Dim x As Foo.Bar` was matched by BOTH `DIM_QUAL_RE` (which extracts
 * the outer type `Foo`) AND `DIM_UNQUAL_RE` (which also extracts `Foo` via the
 * same capture), causing two identical `references` edges with the same target.
 */
describe('VbaExtractor — no duplicate references for qualified Dim (Fix 2)', () => {
  it('Dim x As Foo.Bar emits exactly ONE references edge to the outer type (Fix 2)', () => {
    const src = `Dim m_Obj As DAO.Database`;
    const r = extract('src/classes/X.cls', src);
    const refEdges = r.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    // Must be exactly 1, not 2.
    expect(refEdges).toHaveLength(1);
    const target = r.nodes.find((n) => n.id === refEdges[0]?.target);
    expect(target?.name).toBe('DAO');
  });

  it('Dim x As Foo.Bar emits ONE node for the outer type (no duplicate synthetic nodes)', () => {
    const src = `Dim m_Obj As Scripting.Dictionary`;
    const r = extract('src/modules/Mod.bas', src);
    const daoNodes = r.nodes.filter((n) => n.name === 'Scripting');
    expect(daoNodes).toHaveLength(1);
  });
});

// ---------------------------------------------------------------------------
// Fix 4: PRIMITIVE_TYPES guard is case-insensitive
// ---------------------------------------------------------------------------

/**
 * Fix 4: `PRIMITIVE_TYPES.has(typeName)` was comparing against PascalCase
 * entries (`'Long'`, `'String'`, …) but VBA is case-insensitive, so
 * `Dim x As long` / `Dim x As STRING` slipped through and created phantom
 * `references` nodes named `long` / `STRING`.
 */
describe('VbaExtractor — primitive type guard is case-insensitive (Fix 4)', () => {
  it('Dim x As long (all-lowercase) does NOT emit a reference edge', () => {
    const src = `Dim x As long`;
    const r = extract('src/modules/Mod.bas', src);
    const refs = r.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(refs).toHaveLength(0);
  });

  it('Dim x As STRING (all-uppercase) does NOT emit a reference edge', () => {
    const src = `Dim x As STRING`;
    const r = extract('src/modules/Mod.bas', src);
    const refs = r.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(refs).toHaveLength(0);
  });

  it('Dim x As Long (PascalCase — existing) still does NOT emit a reference edge', () => {
    const src = `Dim x As Long`;
    const r = extract('src/modules/Mod.bas', src);
    const refs = r.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(refs).toHaveLength(0);
  });

  it('Dim x As lONGpTR (mixed-case LongPtr) does NOT emit a reference edge', () => {
    const src = `Dim x As lONGpTR`;
    const r = extract('src/modules/Mod.bas', src);
    const refs = r.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(refs).toHaveLength(0);
  });
});

// ---------------------------------------------------------------------------
// Fix 5: SQL-table synthetic nodes — one node per unique table name
// ---------------------------------------------------------------------------

/**
 * Fix 5: `emitReference` was keying the synthetic node id on
 * `generateNodeId(filePath, 'class', name, lineNum)` where `lineNum` varies
 * per reference site. This meant the `synthClassNodeIds` de-dup never fired
 * for different lines and the same table referenced from N procedures created
 * N separate nodes.
 */
describe('VbaExtractor — one synthetic node per SQL table name (Fix 5)', () => {
  it('the same SQL table referenced in two different procedures emits exactly ONE node', () => {
    const src = [
      'Sub A()',
      '  DoCmd.RunSQL "SELECT * FROM TbFoo"',
      'End Sub',
      '',
      'Sub B()',
      '  DoCmd.RunSQL "UPDATE TbFoo SET x = 1"',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const fooNodes = r.nodes.filter((n) => n.name === 'TbFoo');
    // Must be exactly ONE node, not two.
    expect(fooNodes).toHaveLength(1);
  });

  it('two DIFFERENT SQL tables each produce their own node', () => {
    const src = [
      'Sub A()',
      '  DoCmd.RunSQL "SELECT * FROM TbFoo"',
      'End Sub',
      '',
      'Sub B()',
      '  DoCmd.RunSQL "SELECT * FROM TbBar"',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const fooNodes = r.nodes.filter((n) => n.name === 'TbFoo');
    const barNodes = r.nodes.filter((n) => n.name === 'TbBar');
    expect(fooNodes).toHaveLength(1);
    expect(barNodes).toHaveLength(1);
  });

  it('the same Dim type referenced on multiple lines produces exactly ONE node', () => {
    // Same fix applies to Dim references, not just SQL.
    const src = [
      'Dim a As ACAuditoria',
      'Dim b As ACAuditoria',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const acNodes = r.nodes.filter((n) => n.name === 'ACAuditoria');
    expect(acNodes).toHaveLength(1);
  });
});

// ---------------------------------------------------------------------------
// Fix 7: Qualified statement-form calls (no parens) emit heuristic edges
// ---------------------------------------------------------------------------

/**
 * Fix 7: `m_Obj.Method arg1, arg2` (no parens) emitted zero `calls` edges
 * because `CALL_RE` required `\s*\(`. This is the dominant cross-object call
 * shape in real Dysflow-exported fixtures. Now a statement-form qualified call
 * is detected by `detectQualifiedStatementCall` and emits a heuristic edge.
 */
describe('VbaExtractor — qualified statement-form calls emit heuristic edges (Fix 7)', () => {
  it('Receiver.Method arg (no parens) emits a heuristic calls edge when receiver is a local project-class var', () => {
    // Fix 2 (Issue #2): receiver must be declared as a file-local variable
    // typed as a simple (non-qualified, non-primitive) class.
    const src = [
      'Sub Outer()',
      '  Dim m_NCOp As NCOperaciones',
      '  m_NCOp.Registrar m_ARAlInicio, p_Error',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const hEdges = r.edges.filter(
      (e) => e.kind === 'calls' && e.provenance === 'heuristic' &&
             e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(hEdges.length).toBeGreaterThanOrEqual(1);
    const target = r.nodes.find((n) => n.id === hEdges[0]?.target);
    // #12a (intentional, not a regression — see proposal.md Affected Areas):
    // the stub's name now uses the RESOLVED CLASS TYPE (`NCOperaciones`),
    // not the raw variable name (`m_NCOp`), so it matches the real `.cls`
    // method's `${className}.${proc}` qualifiedName shape for the #12b
    // resolver's exact-match lookup.
    expect(target?.name).toBe('NCOperaciones.Registrar');
    expect(target?.metadata?.stub).toBe(true);
  });

  it('Receiver.Method (no args, no parens) also emits a heuristic calls edge when receiver is declared', () => {
    // Fix 2 (Issue #2): receiver must be in the local var type map.
    const src = [
      'Sub Outer()',
      '  Dim m_Obj As SomeClass',
      '  m_Obj.Init',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const hEdges = r.edges.filter(
      (e) => e.kind === 'calls' && e.provenance === 'heuristic',
    );
    expect(hEdges.length).toBeGreaterThanOrEqual(1);
    const target = r.nodes.find((n) => n.id === hEdges[0]?.target);
    // #12a (intentional, not a regression): resolved-type name, not the
    // raw variable name.
    expect(target?.name).toBe('SomeClass.Init');
    expect(target?.metadata?.stub).toBe(true);
  });

  it('Receiver.Method(args) (paren form) is not double-counted by the statement form detector', () => {
    const src = [
      'Sub Outer()',
      '  m_NCOp.Registrar(arg1, arg2)',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    // CALL_RE handles the paren form; detectQualifiedStatementCall must skip it.
    const hEdges = r.edges.filter(
      (e) => e.kind === 'calls' && e.provenance === 'heuristic',
    );
    expect(hEdges).toHaveLength(1);
  });

  it('blacklisted receivers do not emit qualified statement calls', () => {
    const src = [
      'Sub Outer()',
      '  DoCmd.OpenForm "MyForm"',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    // DoCmd is in RUNTIME_RECEIVER_BLACKLIST — no heuristic edge.
    const hEdgesDoCmd = r.edges.filter((e) => {
      const tgt = r.nodes.find((n) => n.id === e.target);
      return e.kind === 'calls' && tgt?.name?.startsWith('DoCmd');
    });
    expect(hEdgesDoCmd).toHaveLength(0);
  });

  it('property assignment (Receiver.Prop = value) does NOT emit a calls edge', () => {
    const src = [
      'Sub Outer()',
      '  m_Obj.Status = 1',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const callEdges = r.edges.filter((e) => e.kind === 'calls');
    expect(callEdges).toHaveLength(0);
  });

  it('qualified statement call does NOT emit an edge when receiver is typed as a qualified/runtime type (Fix 2)', () => {
    // DAO.Recordset is a qualified type → rcdDatos.AddNew must be silent.
    const src = [
      'Sub Outer()',
      '  Dim rcdDatos As DAO.Recordset',
      '  rcdDatos.AddNew',
      '  rcdDatos.Update',
      '  rcdDatos.Close',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const hEdges = r.edges.filter((e) => e.kind === 'calls' && e.provenance === 'heuristic');
    expect(hEdges).toHaveLength(0);
  });

  it('legitimate cross-object call on project-class var still emits edge after Fix 2 gating', () => {
    // Replicates the real fixture pattern: m_AROp is typed as ARAuditoriaOperaciones
    // (a project class, not qualified, not primitive) → edge must be emitted.
    const src = [
      'Sub Eliminar()',
      '  Dim m_AROp As ARAuditoriaOperaciones',
      '  Set m_AROp = New ARAuditoriaOperaciones',
      '  m_AROp.Eliminar p_Error',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const hEdges = r.edges.filter(
      (e) => e.kind === 'calls' && e.provenance === 'heuristic' &&
             e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(hEdges.length).toBeGreaterThanOrEqual(1);
    const target = r.nodes.find((n) => n.id === hEdges[0]?.target);
    // #12a (intentional, not a regression): resolved-type name, not the
    // raw variable name.
    expect(target?.name).toBe('ARAuditoriaOperaciones.Eliminar');
    expect(target?.metadata?.stub).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// #12a: qualified call-stub nodes/edges are tagged `metadata.stub === true`
// so a post-extraction resolver pass (#12b) can find and repoint them.
// ---------------------------------------------------------------------------

describe('VbaExtractor — call-stub metadata tagging (#12a)', () => {
  it('qualified paren-form call (scanCallSites): stub node + edge carry the stub-tagging contract', () => {
    const src = [
      'Sub Outer()',
      '  modHelpers.CalcTotal(1, 2)',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const hEdge = r.edges.find(
      (e) => e.kind === 'calls' && e.provenance === 'heuristic' &&
             e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(hEdge).toBeDefined();
    expect(hEdge?.metadata?.stub).toBe(true);
    // `modHelpers` isn't a declared local var — receiverType falls back to
    // the raw receiver text (no class-type resolution to apply).
    expect(hEdge?.metadata?.receiverType).toBe('modHelpers');
    expect(hEdge?.metadata?.member).toBe('CalcTotal');
    const target = r.nodes.find((n) => n.id === hEdge?.target);
    expect(target?.metadata?.stub).toBe(true);
    expect(target?.name).toBe('modHelpers.CalcTotal');
  });

  it('class-typed qualified statement-form call: stub node + edge carry the stub-tagging contract, name uses the RESOLVED type', () => {
    const src = [
      'Sub Outer()',
      '  Dim m_NCOp As NCOperaciones',
      '  m_NCOp.Registrar m_ARAlInicio, p_Error',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const hEdge = r.edges.find(
      (e) => e.kind === 'calls' && e.provenance === 'heuristic' &&
             e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(hEdge).toBeDefined();
    expect(hEdge?.metadata?.stub).toBe(true);
    // Proposal (#12): class-typed receivers resolve via localVarTypeMap —
    // the stub's name/qualifiedName is the RESOLVED CLASS name (matching
    // the real `.cls` method's `${className}.${proc}` qualifiedName shape),
    // not the raw variable name `m_NCOp`.
    expect(hEdge?.metadata?.receiverType).toBe('NCOperaciones');
    expect(hEdge?.metadata?.member).toBe('Registrar');
    const target = r.nodes.find((n) => n.id === hEdge?.target);
    expect(target?.metadata?.stub).toBe(true);
    expect(target?.name).toBe('NCOperaciones.Registrar');
  });
});

// ---------------------------------------------------------------------------
// Issue #1: Dim x As New <Type> — must reference the real type, not `New`
// ---------------------------------------------------------------------------

describe('VbaExtractor — Dim As New references the actual class (Issue #1)', () => {
  it('Dim x As New SomeClass emits a reference to SomeClass, not New', () => {
    const src = `Dim conn As New ACAuditoria`;
    const r = extract('src/classes/X.cls', src);
    const refs = r.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(refs).toHaveLength(1);
    const target = r.nodes.find((n) => n.id === refs[0]?.target);
    expect(target?.name).toBe('ACAuditoria');
    // Must NOT reference 'New' as a type.
    const newNode = r.nodes.find((n) => n.name === 'New');
    expect(newNode).toBeUndefined();
  });

  it('Dim rs As New DAO.Recordset emits a reference to DAO (outer), not New', () => {
    const src = `Dim rs As New DAO.Recordset`;
    const r = extract('src/modules/Mod.bas', src);
    const refs = r.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(refs).toHaveLength(1);
    const target = r.nodes.find((n) => n.id === refs[0]?.target);
    expect(target?.name).toBe('DAO');
    const newNode = r.nodes.find((n) => n.name === 'New');
    expect(newNode).toBeUndefined();
  });

  it('Dim c As New Collection emits reference to Collection (not New)', () => {
    // Collection is not a VBA primitive; it's a built-in class. The type
    // name is captured correctly without the New keyword.
    const src = `Dim c As New Collection`;
    const r = extract('src/modules/Mod.bas', src);
    // Collection is not in PRIMITIVE_TYPES, so a reference is emitted.
    const collNodes = r.nodes.filter((n) => n.name === 'Collection');
    expect(collNodes.length).toBeGreaterThanOrEqual(1);
    const newNode = r.nodes.find((n) => n.name === 'New');
    expect(newNode).toBeUndefined();
  });

  it('Set x = New SomeClass (Set statement) does NOT emit a Dim-style references edge', () => {
    // Set assignment is not a Dim declaration; the extractor should not
    // emit a references edge from it.
    const src = [
      'Sub X()',
      '  Set m_AROp = New ARAuditoriaOperaciones',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const refEdges = r.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    // No Dim declaration on that line → no references edge.
    expect(refEdges).toHaveLength(0);
  });
});

// ---------------------------------------------------------------------------
// Issue #3: Multi-variable Dim — all types must be referenced
// ---------------------------------------------------------------------------

describe('VbaExtractor — multi-variable Dim emits all type references (Issue #3)', () => {
  it('Dim a As Foo, b As Bar emits references to both Foo and Bar', () => {
    const src = `Dim a As Foo, b As Bar`;
    const r = extract('src/modules/Mod.bas', src);
    const refs = r.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(refs).toHaveLength(2);
    const names = refs.map((e) => r.nodes.find((n) => n.id === e.target)?.name).sort();
    expect(names).toEqual(['Bar', 'Foo']);
  });

  it('Dim a As Long, b As Bar skips Long (primitive) and emits only Bar', () => {
    const src = `Dim a As Long, b As Bar`;
    const r = extract('src/modules/Mod.bas', src);
    const refs = r.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(refs).toHaveLength(1);
    const target = r.nodes.find((n) => n.id === refs[0]?.target);
    expect(target?.name).toBe('Bar');
  });

  it('Dim a As Foo, b As New Bar emits references to both Foo and Bar (Fix 1+3)', () => {
    const src = `Dim a As Foo, b As New Bar`;
    const r = extract('src/modules/Mod.bas', src);
    const refs = r.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    const names = refs.map((e) => r.nodes.find((n) => n.id === e.target)?.name).sort();
    expect(names).toContain('Foo');
    expect(names).toContain('Bar');
    // Must not reference 'New' as a type name.
    const newNode = r.nodes.find((n) => n.name === 'New');
    expect(newNode).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// Issue #2: Call scanning must be string-aware
// ---------------------------------------------------------------------------

describe('VbaExtractor — call patterns inside string literals are ignored (Issue #2)', () => {
  it('modHelper.BuildQuery(123) inside a string literal does not emit a calls edge', () => {
    const src = [
      'Sub X()',
      '  m_SQL = "EXEC modHelper.BuildQuery(123)"',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const callEdges = r.edges.filter((e) => e.kind === 'calls');
    expect(callEdges).toHaveLength(0);
  });

  it('myModule.MyFunc(id) inside a DoCmd.RunSQL string does not emit a calls edge', () => {
    // DoCmd is blacklisted at the receiver level; the string content must
    // also be masked so `myModule.MyFunc` is not matched by CALL_RE.
    const src = [
      'Sub X()',
      '  DoCmd.RunSQL "SELECT myModule.MyFunc(id) FROM tbl"',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const callEdges = r.edges.filter((e) => {
      const tgt = r.nodes.find((n) => n.id === e.target);
      return e.kind === 'calls' && tgt?.name?.startsWith('myModule');
    });
    expect(callEdges).toHaveLength(0);
  });
});

// ---------------------------------------------------------------------------
// Issue #4: getdb().Execute / getdb().OpenRecordset with inline SQL literal
// ---------------------------------------------------------------------------

describe('VbaExtractor — getdb() inline SQL literal resolved (Issue #4)', () => {
  it('getdb().Execute with inline SQL literal resolves the table', () => {
    const src = [
      'Sub Del()',
      '  getdb().Execute "DELETE FROM tblFoo"',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const edge = r.edges.find(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-sql-table',
    );
    expect(edge).toBeDefined();
    const target = r.nodes.find((n) => n.id === edge?.target);
    expect(target?.name).toBe('tblFoo');
  });

  it('getdb().OpenRecordset with inline SQL literal resolves the table', () => {
    const src = [
      'Sub Q()',
      '  Set rs = getdb().OpenRecordset "SELECT * FROM tblBar"',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const edge = r.edges.find(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-sql-table',
    );
    expect(edge).toBeDefined();
    const target = r.nodes.find((n) => n.id === edge?.target);
    expect(target?.name).toBe('tblBar');
  });
});

// ---------------------------------------------------------------------------
// Issue #5: SQL + Rem-with-string false positive
// ---------------------------------------------------------------------------

describe('VbaExtractor — Rem comment after SQL does not produce false table refs (Issue #5)', () => {
  it('DoCmd.RunSQL "real" Rem "fake" yields only tblReal, not tblFake', () => {
    const src = [
      'Sub X()',
      '  DoCmd.RunSQL "SELECT * FROM tblReal" Rem "SELECT * FROM tblFake"',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const sqlEdges = r.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-sql-table',
    );
    const tableNames = sqlEdges.map((e) => r.nodes.find((n) => n.id === e.target)?.name);
    expect(tableNames).toContain('tblReal');
    expect(tableNames).not.toContain('tblFake');
  });

  it('trailing bare Rem at EOL (no trailing space) strips the remainder', () => {
    // REM_MIDLINE must handle `\s+Rem$` (end-of-line) not just `\s+Rem\s`.
    const src = [
      'Sub X()',
      '  DoCmd.RunSQL "SELECT * FROM tblReal" Rem',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const sqlEdges = r.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-sql-table',
    );
    expect(sqlEdges).toHaveLength(1);
    const target = r.nodes.find((n) => n.id === sqlEdges[0]?.target);
    expect(target?.name).toBe('tblReal');
  });
});

// ---------------------------------------------------------------------------
// Issue #13: trackSqlVariableAssignment must ACCUMULATE across
// self-referential concatenation (`sql = sql & "..."`), not overwrite.
// ---------------------------------------------------------------------------

describe('VbaExtractor — SQL variable accumulation across self-referential concatenation (#13)', () => {
  it('two-fragment self-referential concat retains the FROM table from the first fragment', () => {
    const src = [
      'Sub Q()',
      '  Dim sql As String',
      '  sql = "SELECT * FROM tblA"',
      '  sql = sql & " WHERE x=1"',
      '  db.Execute sql',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Q.bas', src);
    const sqlEdges = r.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-sql-table',
    );
    const tableNames = sqlEdges.map((e) => r.nodes.find((n) => n.id === e.target)?.name);
    // Previously: overwrite dropped the first fragment's `FROM tblA` entirely.
    expect(tableNames).toContain('tblA');
  });

  it('three-plus fragment accumulation retains tables from every fragment', () => {
    const src = [
      'Sub Q()',
      '  Dim sql As String',
      '  sql = "SELECT * FROM tblA"',
      '  sql = sql & " INTO tblB"',
      '  sql = sql & " WHERE x=1"',
      '  db.Execute sql',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Q.bas', src);
    const sqlEdges = r.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-sql-table',
    );
    const tableNames = sqlEdges.map((e) => r.nodes.find((n) => n.id === e.target)?.name);
    expect(tableNames).toContain('tblA');
    expect(tableNames).toContain('tblB');
  });

  it('fresh (non-self-referential) reassignment after use resets tracking', () => {
    const src = [
      'Sub Q()',                            // line 1
      '  Dim sql As String',                // line 2
      '  sql = "SELECT * FROM tblA"',       // line 3
      '  db.Execute sql',                   // line 4 — first call, tblA
      '  sql = "UPDATE tblC SET x=1"',      // line 5 — fresh reassignment
      '  db.Execute sql',                   // line 6 — second call, tblC only
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Q.bas', src);
    const sqlEdges = r.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-sql-table',
    );
    const secondCallEdges = sqlEdges.filter((e) => e.line === 6);
    const secondCallTables = secondCallEdges.map(
      (e) => r.nodes.find((n) => n.id === e.target)?.name,
    );
    expect(secondCallTables).toContain('tblC');
    expect(secondCallTables).not.toContain('tblA');
  });

  it('case-insensitive self-reference (Sql = sql & ...) still accumulates', () => {
    const src = [
      'Sub Q()',
      '  Dim Sql As String',
      '  Sql = "SELECT * FROM tblA"',
      '  Sql = sql & " WHERE x=1"',
      '  db.Execute Sql',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Q.bas', src);
    const sqlEdges = r.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-sql-table',
    );
    const tableNames = sqlEdges.map((e) => r.nodes.find((n) => n.id === e.target)?.name);
    expect(tableNames).toContain('tblA');
  });
});


describe('VbaExtractor � API declarations and VBA conditional compilation', () => {
  it('extracts Public Declare PtrSafe Sub as a single-line declare node with metadata', () => {
    const src = [
      'Option Explicit',
      'Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)',
      'Sub UseSleep()',
      '  Sleep 1000',
      'End Sub',
    ].join('\n');

    const r = extract('src/modules/modApi.bas', src);
    const sleep = r.nodes.find((n) => n.kind === 'declare' && n.name === 'Sleep');
    expect(sleep).toBeDefined();
    expect(sleep?.visibility).toBe('public');
    expect(sleep?.startLine).toBe(2);
    expect(sleep?.endLine).toBe(2);
    expect(sleep?.metadata).toEqual(expect.objectContaining({
      dll: 'kernel32',
      declareKind: 'sub',
      ptrSafe: true,
    }));

    const caller = r.nodes.find((n) => n.kind === 'function' && n.name === 'UseSleep');
    const call = r.edges.find((e) => e.kind === 'calls' && e.source === caller?.id && e.target === sleep?.id);
    expect(call).toBeDefined();
  });

  it('conditional compilation hides inactive duplicate Declare branch', () => {
    const src = [
      '#If VBA7 Then',
      'Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long',
      '#Else',
      'Public Declare Function GetTickCount Lib "kernel32" () As Long',
      '#End If',
    ].join('\n');

    const r = extract('src/modules/modApi.bas', src);
    const declarations = r.nodes.filter((n) => n.kind === 'declare' && n.name === 'GetTickCount');
    expect(declarations).toHaveLength(1);
    expect(declarations[0]?.startLine).toBe(2);
    expect(declarations[0]?.metadata?.ptrSafe).toBe(true);
  });
});

describe('VbaExtractor � custom db variables and OpenForm constants', () => {
  it('extracts inline SQL executed through custom variables ending in db', () => {
    const src = [
      'Sub Q(p_db As Object)',
      '  p_db.Execute "SELECT * FROM Employees"',
      '  m_Db.OpenRecordset "SELECT * FROM Orders"',
      'End Sub',
    ].join('\n');

    const r = extract('src/modules/modSql.bas', src);
    const tables = r.edges
      .filter((e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-sql-table')
      .map((e) => r.nodes.find((n) => n.id === e.target)?.name)
      .sort();

    expect(tables).toEqual(['Employees', 'Orders']);
  });

  it('extracts SQL variable execution through a custom db variable but not db_test', () => {
    const src = [
      'Sub Q(p_db As Object, db_test As Object)',
      '  sql = "SELECT * FROM Included"',
      '  p_db.Execute sql',
      '  db_test.Execute "SELECT * FROM Excluded"',
      'End Sub',
    ].join('\n');

    const r = extract('src/modules/modSql.bas', src);
    const tables = r.edges
      .filter((e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-sql-table')
      .map((e) => r.nodes.find((n) => n.id === e.target)?.name);

    expect(tables).toContain('Included');
    expect(tables).not.toContain('Excluded');
  });

  it('resolves local string constants in DoCmd.OpenForm and falls back for unknown constants', () => {
    const src = [
      'Const FORM_EMPLOYEES = "frmEmployees", FORM_ORDERS As String = "frmOrders"',
      'Sub OpenKnown()',
      '  DoCmd.OpenForm FORM_EMPLOYEES',
      '  DoCmd.OpenForm FORM_ORDERS',
      '  DoCmd.OpenForm FORM_UNKNOWN',
      'End Sub',
    ].join('\n');

    const r = extract('src/modules/modForms.bas', src);
    const targets = r.edges
      .filter((e) => e.kind === 'opens-form')
      .map((e) => r.nodes.find((n) => n.id === e.target)?.name)
      .sort();

    expect(targets).toEqual(['FORM_UNKNOWN', 'frmEmployees', 'frmOrders']);
  });
});

// ---------------------------------------------------------------------------
// Issue #47: `Global` module-level and `Static` procedure-local declarations
// must emit the same `references` edge and `localVarTypeMap` registration as
// their `Dim` / `Private` / `Public` siblings today. The previous regex
// `DIM_DECL_PREFIX_RE` only matched `Dim|Private|Public`; today a real-world
// module using the `Global` keyword for a module-level typed instance
// (`Global g_Client As Class_Client`) ends up invisible to the extractor and
// downstream stub/edge resolution. Fix: extend the prefix alternation to
// include `Global` (and `Static` for the procedure-local case). The primitive
// gate (PRIMITIVE_TYPES) must still suppress `Global gsNombre As String`
// cleanly — primitives never emit a references edge and never enter
// `localVarTypeMap`.
// ---------------------------------------------------------------------------

describe('VbaExtractor — `Global` module-level declarations (Issue #47)', () => {
  it('Global g_Client As Class_Client emits a vba-name-resolution references edge to Class_Client', () => {
    const src = `Global g_Client As Class_Client`;
    const r = extract('src/modules/modClient.bas', src);
    const refs = r.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(refs).toHaveLength(1);
    const target = r.nodes.find((n) => n.id === refs[0]?.target);
    expect(target?.name).toBe('Class_Client');
  });

  it('Global g_Client As Class_Client registers g_Client in localVarTypeMap (so stub/edge resolution resolves the receiver as the class type)', () => {
    // Indirect proof of `localVarTypeMap.set('gclient', { outer: 'Class_Client', ... })`.
    // The `receiverType` of a class-typed qualified statement-form call is
    // populated from the map; if the entry is missing the metadata would
    // degrade to the raw variable name. See the existing
    // "class-typed qualified statement-form call" test (#12a) for the contract.
    const src = [
      'Global g_Client As Class_Client',
      'Sub Outer()',
      '  g_Client.DoWork',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/modClient.bas', src);
    const hEdge = r.edges.find(
      (e) => e.kind === 'calls' && e.provenance === 'heuristic' &&
             e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(hEdge).toBeDefined();
    expect(hEdge?.metadata?.receiverType).toBe('Class_Client');
    expect(hEdge?.metadata?.member).toBe('DoWork');
  });

  it('Global gsNombre As String (primitive) emits no reference edge and does not enter localVarTypeMap', () => {
    const src = [
      'Global gsNombre As String',
      'Sub Outer()',
      '  gsNombre = "x"',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/modClient.bas', src);
    const refs = r.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(refs).toHaveLength(0);
    // Indirect proof that `gsnombre` is NOT in `localVarTypeMap`: a subsequent
    // qualified statement-form call on `gsNombre` cannot have a class-shaped
    // `receiverType` since `gsNombre` was never typed as a class. There must
    // be no calls edge attributed to `gsnombre` as a class receiver.
    const stmtEdges = r.edges.filter(
      (e) => e.kind === 'calls' && e.provenance === 'heuristic' &&
             e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(stmtEdges).toHaveLength(0);
  });

  it('Global Const X = 1 stays OUT of the Dim sweep (Const has its own sweepEnumsAndConsts path)', () => {
    // Regression guard: the negative lookahead
    // `(?!(?:Function|Sub|Property|Const|WithEvents)\b)` in DIM_DECL_PREFIX_RE
    // keeps `Global Const` routed to `sweepEnumsAndConsts`. Adding `Global`
    // to the alternation MUST NOT change that contract — `Const` in any
    // combination still goes through the Const sweep.
    const src = `Global Const MY_GLOBAL = 1`;
    const r = extract('src/modules/modConsts.bas', src);
    // The Dim sweep would have emitted a class-style `vba-name-resolution`
    // edge with `references`. Const sweep emits a `constant` node instead.
    const dimRefs = r.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(dimRefs).toHaveLength(0);
    const constNode = r.nodes.find((n) => n.kind === 'constant' && n.name === 'MY_GLOBAL');
    expect(constNode).toBeDefined();
  });
});

describe('VbaExtractor — `Static` procedure-local declarations (Issue #47)', () => {
  it('Static m_Cache As MiClase inside a procedure emits a vba-name-resolution references edge to MiClase', () => {
    const src = [
      'Sub Outer()',
      '  Static m_Cache As MiClase',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/modCache.bas', src);
    const refs = r.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(refs).toHaveLength(1);
    const target = r.nodes.find((n) => n.id === refs[0]?.target);
    expect(target?.name).toBe('MiClase');
  });

  it('Static m_Cache As MiClase registers m_Cache in localVarTypeMap', () => {
    // Indirect proof: a class-typed qualified statement-form call on
    // `m_Cache` resolves the receiver as the declared class type.
    const src = [
      'Sub Outer()',
      '  Static m_Cache As MiClase',
      '  m_Cache.Get',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/modCache.bas', src);
    const hEdge = r.edges.find(
      (e) => e.kind === 'calls' && e.provenance === 'heuristic' &&
             e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(hEdge).toBeDefined();
    expect(hEdge?.metadata?.receiverType).toBe('MiClase');
    expect(hEdge?.metadata?.member).toBe('Get');
  });

  it('Static s_Label As String (primitive) inside a procedure emits no reference edge', () => {
    const src = [
      'Sub Outer()',
      '  Static s_Label As String',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/modCache.bas', src);
    const refs = r.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(refs).toHaveLength(0);
  });
});

describe('VbaExtractor — regression: `Dim`/`Private`/`Public` shape unchanged (Issue #47)', () => {
  // Regression guard for the existing REQ-CODE-6 / S3 / Issue #1/3 paths.
  // These tests intentionally mirror the original wording so that the new
  // `Global`/`Static` alternation does NOT shift `Dim` semantics in any way.
  it('Dim x As Foo still emits a vba-name-resolution references edge to Foo (unchanged shape)', () => {
    const src = `Dim x As Foo`;
    const r = extract('src/modules/Mod.bas', src);
    const edge = r.edges.find(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(edge).toBeDefined();
    const target = r.nodes.find((n) => n.id === edge?.target);
    expect(target?.name).toBe('Foo');
  });

  it('Private p_X As SomeType still emits a vba-name-resolution references edge (unchanged shape)', () => {
    const src = `Private p_X As SomeType`;
    const r = extract('src/modules/Mod.bas', src);
    const edge = r.edges.find(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(edge).toBeDefined();
    const target = r.nodes.find((n) => n.id === edge?.target);
    expect(target?.name).toBe('SomeType');
  });

  it('Public AC As ACAuditoria still emits a vba-name-resolution references edge (unchanged shape)', () => {
    const src = `Public AC As ACAuditoria`;
    const r = extract('src/modules/Mod.bas', src);
    const edge = r.edges.find(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(edge).toBeDefined();
    const target = r.nodes.find((n) => n.id === edge?.target);
    expect(target?.name).toBe('ACAuditoria');
  });
});

describe('VbaExtractor — Declare statements emit metadata.isDeclare (spec compliance)', () => {
  // The spec at openspec/changes/archive/vba-api-declarations/specs/
  // vba-extraction-enhancements/spec.md (Requirement: DLL API Declarations
  // Extraction, item 3) mandates that every declare node carries
  // `metadata.isDeclare === true`. The other metadata fields (dll,
  // declareKind, ptrSafe, optional aliasName) keep their existing values;
  // this block adds the missing flag and a regression guard for the
  // existing field shapes.

  it('bare Declare Sub (no PtrSafe, no Alias) carries metadata.isDeclare === true', () => {
    const src = [
      'Option Explicit',
      'Declare Sub Foo Lib "kernel32" ()',
    ].join('\n');

    const r = extract('src/modules/modApi.bas', src);
    const foo = r.nodes.find((n) => n.kind === 'declare' && n.name === 'Foo');
    expect(foo).toBeDefined();
    expect(foo?.metadata?.isDeclare).toBe(true);
  });

  it('Declare PtrSafe Function with Alias carries metadata.isDeclare === true and round-trips alias', () => {
    const src = [
      'Option Explicit',
      'Declare PtrSafe Function Bar Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal uType As Long) As Long',
    ].join('\n');

    const r = extract('src/modules/modApi.bas', src);
    const bar = r.nodes.find((n) => n.kind === 'declare' && n.name === 'Bar');
    expect(bar).toBeDefined();
    expect(bar?.metadata?.isDeclare).toBe(true);
    expect(bar?.metadata?.aliasName).toBe('MessageBoxA');
  });

  it('regression: existing Declare metadata fields (dll, declareKind, ptrSafe, aliasName) keep their shape', () => {
    // Regression guard. The fix in src/extraction/vba-extractor.ts is
    // purely additive — it MUST NOT alter the shape of metadata.dll,
    // metadata.declareKind, metadata.ptrSafe, or metadata.aliasName.
    const src = [
      'Option Explicit',
      'Declare PtrSafe Function GetTickCount Lib "kernel32" Alias "GetTickCount64" () As Long',
    ].join('\n');

    const r = extract('src/modules/modApi.bas', src);
    const node = r.nodes.find((n) => n.kind === 'declare' && n.name === 'GetTickCount');
    expect(node).toBeDefined();
    expect(node?.metadata).toEqual(expect.objectContaining({
      isDeclare: true,
      dll: 'kernel32',
      declareKind: 'function',
      ptrSafe: true,
      aliasName: 'GetTickCount64',
    }));
  });
});

