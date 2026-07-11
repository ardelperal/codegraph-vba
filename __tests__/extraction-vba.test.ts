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
import { generateNodeId } from '../src/extraction/tree-sitter-helpers';

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

  it('qualified paren-form call on primitive local variable is silent', () => {
    const src = `Sub RunIt()
  Dim nCount As Long
  nCount.ToString()
End Sub`;
    const r = extract('src/modules/Solo.bas', src);
    const edgesToPrimitive = r.edges.filter((e) => {
      const target = r.nodes.find((n) => n.id === e.target);
      return e.kind === 'calls' && target?.name === 'nCount.ToString';
    });
    expect(edgesToPrimitive).toHaveLength(0);
    const primitiveStub = r.nodes.find((n) => n.kind === 'function' && n.name === 'nCount.ToString');
    expect(primitiveStub).toBeUndefined();
    expect(r.errors).toHaveLength(0);
  });

  it('qualified paren-form call on external local variable is silent', () => {
    const src = `Sub RunIt()
  Dim rcdDatos As DAO.Recordset
  rcdDatos.AddNew()
End Sub`;
    const r = extract('src/modules/Solo.bas', src);
    const edgesToExternal = r.edges.filter((e) => {
      const target = r.nodes.find((n) => n.id === e.target);
      return e.kind === 'calls' && target?.name === 'rcdDatos.AddNew';
    });
    expect(edgesToExternal).toHaveLength(0);
    const externalStub = r.nodes.find((n) => n.kind === 'function' && n.name === 'rcdDatos.AddNew');
    expect(externalStub).toBeUndefined();
    expect(r.errors).toHaveLength(0);
  });

  it('qualified statement-form call on undeclared module receiver emits heuristic edge', () => {
    const src = `Sub RunIt()
  modUtils.Foo arg
End Sub`;
    const r = extract('src/modules/Solo.bas', src);
    const edge = r.edges.find(
      (e) => e.kind === 'calls' && e.provenance === 'heuristic' &&
             e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(edge).toBeDefined();
    expect(edge?.metadata?.receiverType).toBe('modUtils');
    expect(edge?.metadata?.member).toBe('Foo');
    const target = r.nodes.find((n) => n.id === edge?.target);
    expect(target?.name).toBe('modUtils.Foo');
    expect(r.errors).toHaveLength(0);
  });

  it('qualified statement-form call on primitive (Variant) local variable is silent', () => {
    // Issue #40: the unified `shouldProcessQualifiedCall` gate must
    // suppress the heuristic stub for declared primitive-typed locals in
    // BOTH call shapes. The statement form is the shape that was silently
    // dropping cross-module calls before #40; after the fix it must also
    // suppress primitive stubs symmetrically with the paren form.
    const src = `Sub RunIt()
  Dim x As Variant
  x.Foo arg
End Sub`;
    const r = extract('src/modules/Solo.bas', src);
    const edgesToPrimitive = r.edges.filter((e) => {
      const target = r.nodes.find((n) => n.id === e.target);
      return e.kind === 'calls' && target?.name === 'x.Foo';
    });
    expect(edgesToPrimitive).toHaveLength(0);
    const primitiveStub = r.nodes.find((n) => n.kind === 'function' && n.name === 'x.Foo');
    expect(primitiveStub).toBeUndefined();
    expect(r.errors).toHaveLength(0);
  });

  it('qualified statement-form call on external (DAO) local variable is silent', () => {
    // Issue #40: the unified gate must suppress the heuristic stub for
    // declared external-typed locals (e.g. `DAO.Database`) in the
    // statement form, mirroring the paren-form behaviour pinned at
    // line 175. Without this, `db.OpenRecordset sql` would emit a
    // `db.OpenRecordset` function stub that does not exist anywhere in
    // the project.
    const src = `Sub RunIt()
  Dim db As DAO.Database
  db.OpenRecordset sql
End Sub`;
    const r = extract('src/modules/Solo.bas', src);
    const edgesToExternal = r.edges.filter((e) => {
      const target = r.nodes.find((n) => n.id === e.target);
      return e.kind === 'calls' && target?.name === 'db.OpenRecordset';
    });
    expect(edgesToExternal).toHaveLength(0);
    const externalStub = r.nodes.find((n) => n.kind === 'function' && n.name === 'db.OpenRecordset');
    expect(externalStub).toBeUndefined();
    expect(r.errors).toHaveLength(0);
  });

  // Issue #40 — deterministic acceptance criteria (a)–(c).
  //
  // The original "fix" by commit 9b1787a only wired the unified
  // `shouldProcessQualifiedCall` gate into the PAREN-form call site,
  // leaving the statement-form path gating on the old
  // `isLocalProjectClassVar` (which returns `false` for undeclared
  // receivers). The cherry-pick onto main has now applied the same gate
  // to the statement form (see the unified-gate comment at line ~1940
  // in `vba-extractor.ts`). These three tests pin the three acceptance
  // criteria from the issue body so that any future regression is
  // caught.

  it('AC (a): statement-form `db.OpenRecordset sql` with `Dim db As DAO.Database` does NOT emit a `db.OpenRecordset` stub', () => {
    const src = `Sub RunIt()
  Dim db As DAO.Database
  Set rs = db.OpenRecordset(sql)
End Sub`;
    const r = extract('src/modules/Solo.bas', src);
    const stub = r.nodes.find((n) => n.kind === 'function' && n.name === 'db.OpenRecordset');
    expect(stub).toBeUndefined();
    const callsToStub = r.edges.filter((e) => {
      const t = r.nodes.find((n) => n.id === e.target);
      return e.kind === 'calls' && t?.name === 'db.OpenRecordset';
    });
    expect(callsToStub).toHaveLength(0);
    expect(r.errors).toHaveLength(0);
  });

  it('AC (b): both paren-form and statement-form `modUtils.Foo` emit a heuristic calls edge when sibling module `modUtils` exists', () => {
    // Index a sibling module so the resolver can repoint the heuristic
    // stub. We test the extractor in isolation (no resolver), so the
    // acceptance is "edge + stub emitted" — the resolver behaviour is
    // covered by `resolution` tests separately.
    //
    // Two call shapes on different lines each emit one heuristic edge
    // to a stub named `modUtils.Foo`. The shared `callDedupe` set keeps
    // the same (caller, qualified, line) tuple from emitting twice;
    // here the two calls live on DIFFERENT lines, so both edges land.
    const src = `Sub RunIt()
  modUtils.Foo arg
  modUtils.Foo(arg)
End Sub`;
    const r = extract('src/modules/Caller.bas', src);
    const hEdges = r.edges.filter(
      (e) => e.kind === 'calls' && e.provenance === 'heuristic' &&
             e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    // Both call shapes should each produce one heuristic edge.
    expect(hEdges.length).toBeGreaterThanOrEqual(2);
    const modUtilsEdges = hEdges.filter((e) => {
      const t = r.nodes.find((n) => n.id === e.target);
      return e.metadata?.receiverType === 'modUtils' &&
             e.metadata?.member === 'Foo' &&
             t?.name === 'modUtils.Foo';
    });
    expect(modUtilsEdges.length).toBe(2);
    expect(r.errors).toHaveLength(0);
  });

  it('AC (c): statement-form `m_NCOp.Registrar args` with `Dim m_NCOp As NCOperaciones` emits a resolved edge to `NCOperaciones.Registrar`', () => {
    const src = `Sub RunIt()
  Dim m_NCOp As NCOperaciones
  m_NCOp.Registrar args
End Sub`;
    const r = extract('src/modules/Caller.bas', src);
    const edges = r.edges.filter(
      (e) => e.kind === 'calls' && e.metadata?.receiverType === 'NCOperaciones' &&
             e.metadata?.member === 'Registrar',
    );
    expect(edges.length).toBeGreaterThanOrEqual(1);
    const target = r.nodes.find((n) => n.id === edges[0]?.target);
    // Resolved — the receiver type is the class name (NCOperaciones),
    // not the local variable name (m_NCOp).
    expect(target?.name).toBe('NCOperaciones.Registrar');
    expect(target?.metadata?.stub).toBe(true);
    // The un-resolved stub must NOT exist (that was the original P1
    // pollution bug).
    const wrongStub = r.nodes.find((n) => n.kind === 'function' && n.name === 'm_NCOp.Registrar');
    expect(wrongStub).toBeUndefined();
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

  it('db.Execute with schema-qualified FROM (dbo.tblCustomers) emits one composite reference', () => {
    // REGRESSION GUARD for the SQL_TABLE_RE schema-prefix extension:
    // previously the regex stopped at `dbo` (period is not in `\p{L}[\p{L}\p{N}_]*`)
    // and silently dropped `tblCustomers`. The fix extends the capture to allow an
    // optional bracketed/unbracketed schema prefix followed by `.`, so the whole
    // `dbo.tblCustomers` comes through as a single composite table reference.
    const src = `Sub Q()
  db.Execute "SELECT * FROM dbo.tblCustomers"
End Sub`;
    const r = extract('src/modules/Q.bas', src);
    const edges = r.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-sql-table',
    );
    const tableTargets = edges.map((e) => r.nodes.find((n) => n.id === e.target)?.name);
    expect(tableTargets).toEqual(['dbo.tblCustomers']);
    // The schema prefix must NOT leak as a separate node.
    expect(r.nodes.some((n) => n.name === 'dbo')).toBe(false);
    expect(r.nodes.some((n) => n.name === 'tblCustomers')).toBe(false);
  });

  it('db.Execute with bracketed schema-qualified FROM ([My Schema].[My Table]) emits one composite reference', () => {
    // The unwrapped form is what the consumer code already emits for plain bracketed
    // names (`[Order Details]` → `Order Details`), so the schema-qualified form is
    // also unwrapped: `[My Schema].[My Table]` → `My Schema.My Table`. Documented
    // in the commit body and applied consistently to BOTH regexes.
    const src = `Sub Q()
  db.Execute "SELECT * FROM [My Schema].[My Table]"
End Sub`;
    const r = extract('src/modules/Q.bas', src);
    const edges = r.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-sql-table',
    );
    const tableTargets = edges.map((e) => r.nodes.find((n) => n.id === e.target)?.name);
    expect(tableTargets).toEqual(['My Schema.My Table']);
    // Neither the bracketed schema nor the bracketed table name should leak.
    expect(r.nodes.some((n) => n.name === '[My Schema]')).toBe(false);
    expect(r.nodes.some((n) => n.name === '[My Table]')).toBe(false);
    expect(r.nodes.some((n) => n.name === 'My Schema')).toBe(false);
    expect(r.nodes.some((n) => n.name === 'My Table')).toBe(false);
  });

  it('db.Execute with plain (un-qualified) FROM still emits just the table name (regression guard)', () => {
    // The new schema-prefix is OPTIONAL — `FROM tblCustomers` must produce exactly
    // one node named `tblCustomers`, byte-identical to the pre-fix behaviour.
    const src = `Sub Q()
  db.Execute "SELECT * FROM tblCustomers"
End Sub`;
    const r = extract('src/modules/Q.bas', src);
    const edges = r.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-sql-table',
    );
    const tableTargets = edges.map((e) => r.nodes.find((n) => n.id === e.target)?.name);
    expect(tableTargets).toEqual(['tblCustomers']);
    expect(r.nodes.some((n) => n.name === 'tblCustomers')).toBe(true);
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

  // REQ-CODE-8 (access direction): every `vba-sql-table` reference carries a
  // `metadata.access` of `'read'` or `'write'`, derived from the SQL verb, so
  // consumers can answer "who WRITES table X" vs "who READS table X" — mirrors
  // the read/write tagging TempVars references already carry.
  function sqlEdgeFor(r: ReturnType<typeof extract>, tableName: string) {
    return r.edges.find((e) => {
      if (e.kind !== 'references' || e.metadata?.synthesizedBy !== 'vba-sql-table') return false;
      return r.nodes.find((n) => n.id === e.target)?.name === tableName;
    });
  }

  it('SELECT FROM tags the table access=read', () => {
    const src = `Sub R()
  DoCmd.RunSQL "SELECT * FROM tblCustomers"
End Sub`;
    const r = extract('src/modules/R.bas', src);
    expect(sqlEdgeFor(r, 'tblCustomers')?.metadata?.access).toBe('read');
  });

  it('UPDATE tags the table access=write', () => {
    const src = `Sub U()
  CurrentDb.Execute "UPDATE tblOrders SET Status = 1"
End Sub`;
    const r = extract('src/modules/U.bas', src);
    expect(sqlEdgeFor(r, 'tblOrders')?.metadata?.access).toBe('write');
  });

  it('INSERT INTO tags the target table access=write', () => {
    const src = `Sub I()
  DoCmd.RunSQL "INSERT INTO tblAudit (Id) VALUES (1)"
End Sub`;
    const r = extract('src/modules/I.bas', src);
    expect(sqlEdgeFor(r, 'tblAudit')?.metadata?.access).toBe('write');
  });

  it('DELETE FROM tags the target table access=write (FROM after DELETE is a write, not a read)', () => {
    const src = `Sub D()
  DoCmd.RunSQL "DELETE FROM tblOld"
End Sub`;
    const r = extract('src/modules/D.bas', src);
    expect(sqlEdgeFor(r, 'tblOld')?.metadata?.access).toBe('write');
  });

  it('INSERT INTO ... SELECT FROM tags target=write and source=read distinctly', () => {
    const src = `Sub Copy()
  getdb().Execute "INSERT INTO tblArchive SELECT * FROM tblLive"
End Sub`;
    const r = extract('src/modules/Copy.bas', src);
    expect(sqlEdgeFor(r, 'tblArchive')?.metadata?.access).toBe('write');
    expect(sqlEdgeFor(r, 'tblLive')?.metadata?.access).toBe('read');
  });

  it('JOIN source table in a SELECT is access=read', () => {
    const src = `Sub J()
  getdb().OpenRecordset "SELECT * FROM tblA INNER JOIN tblB ON tblA.Id = tblB.Id"
End Sub`;
    const r = extract('src/modules/J.bas', src);
    expect(sqlEdgeFor(r, 'tblA')?.metadata?.access).toBe('read');
    expect(sqlEdgeFor(r, 'tblB')?.metadata?.access).toBe('read');
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

  it('does not break extraction when a comment ends with a line continuation _', () => {
    const src = [
      'Attribute VB_Name = "modHelpers"',
      '',
      'Public Function Helper() As Long',
      '    \' This is a comment ending in _',
      '    Call OtherFunc',
      'End Function',
      '',
      'Public Sub OtherFunc()',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/modHelpers.bas', src);
    const edge = r.edges.find((e) => e.kind === 'calls');
    expect(edge).toBeDefined();
  });

  it('extracts calls spanning multiple lines with a comment after the continuation character', () => {
    const src = [
      'Attribute VB_Name = "modHelpers"',
      '',
      'Public Function Helper() As Long',
      '    Call OtherFunc(1, _ \' comment here',
      '                   2)',
      'End Function',
      '',
      'Public Sub OtherFunc()',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/modHelpers.bas', src);
    const edge = r.edges.find((e) => e.kind === 'calls');
    expect(edge).toBeDefined();
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

// ---------------------------------------------------------------------------
// Issue #44: VBA bang (`!`) operator — `Me!Ctl`, `Forms!Form!Ctl`.
//
// Access VBA uses the bang operator as an alias for default-collection
// (Me controls) and default-member (Forms collection, Recordsets) access:
//
//   Me!txtNombre                   — control on the own form (= Me.txtNombre)
//   Forms!FormPrincipal!txtEstado  — cross-form control access
//   Forms("FormPrincipal")!txtEstado — same, paren form
//   Forms!FormX                    — form alone (no control segment)
//   rs!Campo                       — recordset field (STRETCH SCOPE: pinned
//                                    out of scope here — see regression test)
//
// Before this fix only `Me.<Control>` was scanned by `ME_CONTROL_RE`; the
// bang form was silently invisible. `Forms!…` is doubly invisible because
// `Forms` is in `RUNTIME_RECEIVER_BLACKLIST`, so the generic `CALL_RE`
// path also skips it. The fix extends `ME_CONTROL_RE` to accept the bang
// (`Me[.!]`) and adds a dedicated `FORMS_BANG_RE` scanner that emits an
// `UnresolvedReference` to the form tagged `metadata.synthesizedBy ===
// 'vba-forms-bang'`. The bang-form scanner runs independently of `CALL_RE`
// (mirroring the `DoCmd.OpenForm` precedent) so the runtime-receiver
// blacklist does not drop the reference.
//
// Reference forms covered:
//   `Forms!FormX!txtY.Value`     — bang form, control segment present
//   `Forms("FormX")!txtY`        — paren form, control segment present
//   `Forms!FormX`                — bang form, no control segment
//   `Forms![Mi Formulario]`      — bracketed form name (mirrors #54)
//
// NOT covered (stretch scope, pinned by test):
//   `rs!Campo`                   — recordset field access
// ---------------------------------------------------------------------------

describe('VbaExtractor — bang operator (Issue #44)', () => {
  it('Me!txtFoo emits a vba-me-control UnresolvedReference to txtFoo', () => {
    // Regression: Me!txtFoo must produce the EXACT same emission shape as
    // Me.txtFoo — one UnresolvedReference with referenceName 'txtFoo' and
    // metadata.synthesizedBy === 'vba-me-control'.
    const src = `Public Sub X()
    Me!txtFoo = "Hello"
End Sub`;
    const r = extract('src/forms/Form_X.cls', src);
    const refs = r.unresolvedReferences.filter(
      (u) => u.metadata?.synthesizedBy === 'vba-me-control',
    );
    expect(refs.length).toBe(1);
    expect(refs[0]?.referenceName).toBe('txtFoo');
    expect(refs[0]?.referenceKind).toBe('references');
  });

  it('Me!txtFoo and Me.txtFoo produce byte-identical UnresolvedReferences (regression parity)', () => {
    const srcDot = `Public Sub X()
    Me.txtFoo = "Hello"
End Sub`;
    const srcBang = `Public Sub X()
    Me!txtFoo = "Hello"
End Sub`;
    const rDot = extract('src/forms/Form_X.cls', srcDot);
    const rBang = extract('src/forms/Form_X.cls', srcBang);
    const dotRef = rDot.unresolvedReferences.filter(
      (u) => u.metadata?.synthesizedBy === 'vba-me-control',
    );
    const bangRef = rBang.unresolvedReferences.filter(
      (u) => u.metadata?.synthesizedBy === 'vba-me-control',
    );
    expect(bangRef.length).toBe(dotRef.length);
    expect(bangRef.length).toBe(1);
    expect(bangRef[0]?.referenceName).toBe(dotRef[0]?.referenceName);
    expect(bangRef[0]?.referenceKind).toBe(dotRef[0]?.referenceKind);
  });

  it('Forms!FormX!txtY.Value = 1 emits a vba-forms-bang UnresolvedReference to FormX', () => {
    const src = `Public Sub X()
    Forms!FormX!txtY.Value = 1
End Sub`;
    const r = extract('src/modules/X.bas', src);
    const refs = r.unresolvedReferences.filter(
      (u) => u.metadata?.synthesizedBy === 'vba-forms-bang',
    );
    expect(refs.length).toBeGreaterThanOrEqual(1);
    const toFormX = refs.find((u) => u.referenceName === 'FormX');
    expect(toFormX).toBeDefined();
    expect(toFormX?.referenceKind).toBe('references');
  });

  it('Forms!FormX (no control segment) emits a vba-forms-bang UnresolvedReference to FormX', () => {
    const src = `Public Sub X()
    Set f = Forms!FormX
End Sub`;
    const r = extract('src/modules/X.bas', src);
    const refs = r.unresolvedReferences.filter(
      (u) => u.metadata?.synthesizedBy === 'vba-forms-bang',
    );
    expect(refs.length).toBe(1);
    expect(refs[0]?.referenceName).toBe('FormX');
  });

  it('Forms("FormX")!txtY emits the same vba-forms-bang UnresolvedReference to FormX as Forms!FormX!txtY', () => {
    const src = `Public Sub X()
    Forms("FormX")!txtY.Value = 1
End Sub`;
    const r = extract('src/modules/X.bas', src);
    const refs = r.unresolvedReferences.filter(
      (u) => u.metadata?.synthesizedBy === 'vba-forms-bang',
    );
    expect(refs.length).toBeGreaterThanOrEqual(1);
    const toFormX = refs.find((u) => u.referenceName === 'FormX');
    expect(toFormX).toBeDefined();
    expect(toFormX?.referenceKind).toBe('references');
  });

  it('Forms!FormX.Foo (post-form property access, not a control) does NOT emit a vba-forms-bang reference', () => {
    // The form's `.Foo` is a property access on the form object (e.g.
    // `Forms!FormX.Recordsource`), not a bang control access. The
    // FORMS_BANG_RE must only match the bang shape; trailing `.Property`
    // must not be misinterpreted as another bang segment.
    const src = `Public Sub X()
    Forms!FormX.Foo = 1
End Sub`;
    const r = extract('src/modules/X.bas', src);
    const refs = r.unresolvedReferences.filter(
      (u) => u.metadata?.synthesizedBy === 'vba-forms-bang',
    );
    expect(refs).toHaveLength(0);
  });

  it('rs!Campo (recordset field access) is STRETCH SCOPE — emits zero vba-me-control / vba-forms-bang references', () => {
    // Out of scope for Issue #44: `rs!Campo` on a DAO/ADO recordset is
    // treated as a default-member field read. The current extractor emits
    // nothing for it (the generic call-site path is suppressed by the
    // runtime-receiver blacklist when `rs` is a recordset field ref), and
    // this test pins that behavior so a future change can be reviewed
    // explicitly against the bang-form scope decision.
    const src = `Public Sub X()
    rs!Campo = "x"
End Sub`;
    const r = extract('src/modules/X.bas', src);
    const meCtrlRefs = r.unresolvedReferences.filter(
      (u) => u.metadata?.synthesizedBy === 'vba-me-control',
    );
    const formsBangRefs = r.unresolvedReferences.filter(
      (u) => u.metadata?.synthesizedBy === 'vba-forms-bang',
    );
    expect(meCtrlRefs).toHaveLength(0);
    expect(formsBangRefs).toHaveLength(0);
  });

  it('Forms!FormX!txtY does NOT emit a synthetic function node (W4 invariant preserved)', () => {
    // Defense-in-depth guard: even though the dedicated FORMS_BANG_RE
    // scanner emits UnresolvedReferences, no synthetic `function` node
    // should be synthesized for the form/control receiver (would pollute
    // the graph per the W4 invariant — see Forms!MyForm.Open test above).
    const src = `Public Sub X()
    Forms!FormX!txtY.Value = 1
End Sub`;
    const r = extract('src/modules/X.bas', src);
    const synthFns = r.nodes.filter(
      (n) => n.kind === 'function' && /Forms|FormX|txtY/.test(n.name),
    );
    expect(synthFns).toHaveLength(0);
  });
});

// ---------------------------------------------------------------------------
// Issue #48: DoCmd.OpenReport & DoCmd.OpenQuery built-in modeling — sibling
// coverage to the OpenForm hueco-6 implementation. OpenReport mirrors the
// OpenForm stub+edge pattern (Report_<Name> qualifiedName, report-layout
// stub kind, opens-report edge kind, targetReportName metadata key,
// vba-opens-report synthesizedBy). OpenQuery is structurally different: it
// emits a single UnresolvedReference (not a stub + edge) so the resolver
// binds to the REAL `query` node emitted by SqlQueryExtractor for
// `queries/<Name>.sql` — same shape as vba-me-control / vba-forms-bang.
//
// Acceptance criteria covered (one test each):
//   1.  Literal target emits opens-report edge + report-layout stub
//   2.  Const-resolved target emits opens-report edge with resolved name
//   3.  Unresolved Const falls back to bare identifier (no edge skipped)
//   4.  Stub id is deterministic across re-index (id formula check)
//   5.  DoCmd.OpenReport does NOT emit a synthetic `function` node (W4)
//   6.  Two DoCmd.OpenReport "SameName" calls → exactly ONE stub (de-dup)
//   7.  Literal target emits one UnresolvedReference with synthesizedBy
//       `vba-opens-query`, no synthetic `function` / `query` node
//   8.  Const-resolved query emits UnresolvedReference with resolved name
//   9.  Unresolved Const falls back to bare identifier
//   10. DoCmd.OpenQuery does NOT emit a synthetic `function` node (W4)
// ---------------------------------------------------------------------------

describe('VbaExtractor — DoCmd.OpenReport built-in modeling (Issue #48)', () => {
  it('DoCmd.OpenReport "InformeMensual" (literal) emits an opens-report edge to a report-layout stub', () => {
    const src = `Public Sub PrintIt()
    DoCmd.OpenReport "InformeMensual"
End Sub`;
    const r = extract('src/modules/modReports.bas', src);

    const edge = r.edges.find((e) => e.kind === 'opens-report');
    expect(edge).toBeDefined();
    expect(edge?.provenance).toBe('heuristic');
    expect(edge?.metadata?.synthesizedBy).toBe('vba-opens-report');
    expect(edge?.metadata?.targetReportName).toBe('InformeMensual');
    // OpenForm's key MUST NOT leak into OpenReport metadata.
    expect(edge?.metadata?.targetFormName).toBeUndefined();

    const stub = r.nodes.find((n) => n.id === edge?.target);
    expect(stub).toBeDefined();
    expect(stub?.kind).toBe('report-layout');
    expect(stub?.name).toBe('InformeMensual');
    expect(stub?.qualifiedName).toBe('Report_InformeMensual');
    expect(stub?.metadata?.stub).toBe(true);
  });

  it('DoCmd.OpenReport REPORT_MENSUAL (Const-resolved) emits the resolved name "InformeMensual"', () => {
    // Mirrors the OpenForm Const-fallback test shape — known Const is
    // resolved to its literal value, unknown Const falls back to the bare
    // identifier.
    const src = [
      'Const REPORT_MENSUAL As String = "InformeMensual"',
      'Sub PrintKnown()',
      '  DoCmd.OpenReport REPORT_MENSUAL',
      '  DoCmd.OpenReport REPORT_UNKNOWN',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/modReports.bas', src);

    const reportEdges = r.edges.filter((e) => e.kind === 'opens-report');
    expect(reportEdges.length).toBe(2);

    const targets = reportEdges
      .map((e) => r.nodes.find((n) => n.id === e.target)?.name)
      .sort();
    expect(targets).toEqual(['InformeMensual', 'REPORT_UNKNOWN']);

    // Both edges must carry the resolved targetReportName in metadata.
    const toInformeMensual = reportEdges.find(
      (e) => e.metadata?.targetReportName === 'InformeMensual',
    );
    expect(toInformeMensual).toBeDefined();
    expect(toInformeMensual?.metadata?.synthesizedBy).toBe('vba-opens-report');

    // Unknown Const: edge still emitted, falls back to bare identifier.
    const toUnknown = reportEdges.find(
      (e) => e.metadata?.targetReportName === 'REPORT_UNKNOWN',
    );
    expect(toUnknown).toBeDefined();
    const unknownStub = r.nodes.find((n) => n.id === toUnknown?.target);
    expect(unknownStub?.qualifiedName).toBe('Report_REPORT_UNKNOWN');
  });

  it('stub id for DoCmd.OpenReport "InformeMensual" is deterministic across re-index (matches generateNodeId formula)', () => {
    // Deterministic-id invariant: the stub id is computed from a synthetic
    // file path (`synthetic:opensReportStub/<Name>.form.txt`) and the
    // kind/name/line tuple. Re-indexing the same source MUST produce the
    // SAME id (so per-file INSERT OR REPLACE collapses to a no-op and the
    // graph stays stable). The synthetic path uses `.form.txt` for both
    // OpenForm and OpenReport (see `emitOpensStubEdge` comment); the
    // dispatch table's `moduleNamePrefix` is what carries the
    // `Report_<Name>` vs `Form_<Name>` qualifiedName convention.
    const src = `Public Sub PrintIt()
    DoCmd.OpenReport "InformeMensual"
End Sub`;

    const expectedStubId = generateNodeId(
      'synthetic:opensReportStub/InformeMensual.report.txt',
      'report-layout',
      'InformeMensual',
      0,
    );

    const r1 = extract('src/modules/modReports.bas', src);
    const r2 = extract('src/modules/modReports.bas', src);
    const edge1 = r1.edges.find((e) => e.kind === 'opens-report');
    const edge2 = r2.edges.find((e) => e.kind === 'opens-report');
    expect(edge1?.target).toBe(expectedStubId);
    expect(edge2?.target).toBe(expectedStubId);
    expect(edge1?.target).toBe(edge2?.target);
  });

  it('DoCmd.OpenReport does not emit a synthetic function node (W4 invariant)', () => {
    // Mirrors the W4 guard for OpenForm (line ~647). `DoCmd` is in
    // RUNTIME_RECEIVER_BLACKLIST, so the generic CALL_RE path skips it;
    // the dedicated dispatch must NOT regress to emitting a synthetic
    // `DoCmd.OpenReport` `function` node (would pollute the graph per W4).
    const src = `Public Sub X()
    DoCmd.OpenReport "InformeMensual"
End Sub`;
    const r = extract('src/modules/X.bas', src);
    const synthFns = r.nodes.filter(
      (n) => n.kind === 'function' && n.name.includes('DoCmd.OpenReport'),
    );
    expect(synthFns).toHaveLength(0);
  });

  it('two DoCmd.OpenReport "SameName" calls in one file produce exactly ONE report-layout stub (de-dup invariant)', () => {
    // The opensStubIdsByKey cache (keyed by `${cacheKey}:${lowerName}`)
    // must collapse N call sites to a single stub. Verifies the Issue #48
    // refactor preserved the de-dup contract while moving from a name-keyed
    // cache to a (method, name)-keyed cache.
    const src = [
      'Sub PrintTwice()',
      '  DoCmd.OpenReport "SameName"',
      '  DoCmd.OpenReport "SameName"',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/modReports.bas', src);

    const stubs = r.nodes.filter(
      (n) => n.kind === 'report-layout' && n.name === 'SameName',
    );
    expect(stubs).toHaveLength(1);

    const edges = r.edges.filter((e) => e.kind === 'opens-report');
    expect(edges.length).toBe(2);
    // Both edges MUST point at the SAME stub id (the de-dup invariant).
    const uniqueTargets = new Set(edges.map((e) => e.target));
    expect(uniqueTargets.size).toBe(1);
  });
});

describe('VbaExtractor — DoCmd.OpenQuery built-in modeling (Issue #48)', () => {
  it('DoCmd.OpenQuery "Consulta1" (literal) emits ONE UnresolvedReference with synthesizedBy vba-opens-query', () => {
    // OpenQuery is structurally different from OpenForm/OpenReport: it
    // emits an UnresolvedReference (NOT a stub + edge) so the resolver
    // binds to the REAL `query` node emitted by SqlQueryExtractor for
    // `queries/<Name>.sql`. Same emission shape as vba-me-control /
    // vba-forms-bang.
    const src = `Public Sub OpenIt()
    DoCmd.OpenQuery "Consulta1"
End Sub`;
    const r = extract('src/modules/modQueries.bas', src);

    const refs = r.unresolvedReferences.filter(
      (u) => u.metadata?.synthesizedBy === 'vba-opens-query',
    );
    expect(refs.length).toBe(1);
    expect(refs[0]?.referenceName).toBe('Consulta1');
    expect(refs[0]?.referenceKind).toBe('references');

    // No stub nodes (OpenQuery doesn't synthesize one).
    const stubs = r.nodes.filter((n) => n.kind === 'report-layout' || n.kind === 'form-layout');
    expect(stubs.some((s) => s.name === 'Consulta1')).toBe(false);
    // No opens-form / opens-report edges either (the resolver binds the
    // UnresolvedReference when the matching .sql is indexed).
    const openEdges = r.edges.filter(
      (e) => e.kind === 'opens-form' || e.kind === 'opens-report',
    );
    expect(openEdges).toHaveLength(0);
  });

  it('DoCmd.OpenQuery CONSULTA_DEPURACION (Const-resolved) emits UnresolvedReference with resolved name', () => {
    const src = [
      'Const CONSULTA_DEPURACION As String = "qDepuracion"',
      'Sub OpenDepuration()',
      '  DoCmd.OpenQuery CONSULTA_DEPURACION',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/modQueries.bas', src);

    const refs = r.unresolvedReferences.filter(
      (u) => u.metadata?.synthesizedBy === 'vba-opens-query',
    );
    expect(refs.length).toBe(1);
    expect(refs[0]?.referenceName).toBe('qDepuracion');
    expect(refs[0]?.referenceKind).toBe('references');
  });

  it('DoCmd.OpenQuery CONSULTA_UNKNOWN (Const not defined) falls back to bare identifier', () => {
    const src = `Public Sub OpenIt()
    DoCmd.OpenQuery CONSULTA_UNKNOWN
End Sub`;
    const r = extract('src/modules/modQueries.bas', src);

    const refs = r.unresolvedReferences.filter(
      (u) => u.metadata?.synthesizedBy === 'vba-opens-query',
    );
    expect(refs.length).toBe(1);
    expect(refs[0]?.referenceName).toBe('CONSULTA_UNKNOWN');
  });

  it('DoCmd.OpenQuery does not emit a synthetic function node (W4 invariant)', () => {
    // Mirrors the W4 guard for OpenForm / OpenReport. The dedicated
    // OpenQuery scanner emits UnresolvedReferences, NOT synthetic
    // `function` / `query` nodes (the real `query` node already exists in
    // the index once `queries/<Name>.sql` is processed, and creating
    // stubs would compete with the binding).
    const src = `Public Sub X()
    DoCmd.OpenQuery "Consulta1"
End Sub`;
    const r = extract('src/modules/X.bas', src);
    const synthFns = r.nodes.filter(
      (n) => n.kind === 'function' && n.name.includes('DoCmd.OpenQuery'),
    );
    expect(synthFns).toHaveLength(0);
    // VbaExtractor must not synthesize a `query` node either — that kind
    // belongs to SqlQueryExtractor's output, not the VBA scanner.
    const synthQueries = r.nodes.filter((n) => n.kind === 'query');
    expect(synthQueries).toHaveLength(0);
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
    // Declared project-class receivers stay eligible and resolve to the class
    // name so the heuristic edge can match real `.cls` methods.
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
    // Declared project-class receivers remain eligible through the unified
    // qualified-call gate.
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
// Factory return-type inference: `Set x = Factory()` types x from the
// same-file function's declared return type so `x.Method` resolves to the
// factory's class instead of a dead-end `x.Method` stub.
// ---------------------------------------------------------------------------

describe('VbaExtractor — Set x = Factory() types x from the function return type', () => {
  it('undeclared x: Set x = CrearFoo() makes x.Method resolve to the factory class Foo', () => {
    const src = `Public Function CrearFoo() As Foo
End Function
Public Sub Usar()
  Set x = CrearFoo()
  x.Hacer 1
End Sub`;
    const r = extract('src/modules/Fabrica.bas', src);
    const edge = r.edges.find(
      (e) => e.kind === 'calls' && e.metadata?.member === 'Hacer',
    );
    expect(edge).toBeDefined();
    // Without the factory inference, receiverType would be the raw var `x`.
    expect(edge?.metadata?.receiverType).toBe('Foo');
    const target = r.nodes.find((n) => n.id === edge?.target);
    expect(target?.name).toBe('Foo.Hacer');
  });

  it('emits a vba-factory-return reference edge to the factory class', () => {
    const src = `Public Function CrearFoo() As Foo
End Function
Public Sub Usar()
  Set x = CrearFoo()
End Sub`;
    const r = extract('src/modules/Fabrica.bas', src);
    const ref = r.edges.find(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-factory-return',
    );
    expect(ref).toBeDefined();
    const target = r.nodes.find((n) => n.id === ref?.target);
    expect(target?.name).toBe('Foo');
  });

  it('factory return type overrides a generic Dim x As Object', () => {
    const src = `Public Function CrearFoo() As Foo
End Function
Public Sub Usar()
  Dim x As Object
  Set x = CrearFoo()
  x.Hacer 1
End Sub`;
    const r = extract('src/modules/Fabrica.bas', src);
    const edge = r.edges.find(
      (e) => e.kind === 'calls' && e.metadata?.member === 'Hacer',
    );
    expect(edge?.metadata?.receiverType).toBe('Foo');
  });

  it('an explicit Dim x As Bar (project class) wins over the factory return type', () => {
    const src = `Public Function CrearFoo() As Foo
End Function
Public Sub Usar()
  Dim x As Bar
  Set x = CrearFoo()
  x.Hacer 1
End Sub`;
    const r = extract('src/modules/Fabrica.bas', src);
    const edge = r.edges.find(
      (e) => e.kind === 'calls' && e.metadata?.member === 'Hacer',
    );
    expect(edge?.metadata?.receiverType).toBe('Bar');
  });

  it('a primitive-returning function does not type x (Set is invalid there anyway) — no over-reach', () => {
    // CrearId returns Long; a bare `y.Method` with y assigned from CrearId
    // must NOT resolve to `Long.Method`.
    const src = `Public Function CrearId() As Long
End Function
Public Sub Usar()
  Set y = CrearId()
  y.Hacer 1
End Sub`;
    const r = extract('src/modules/Fabrica.bas', src);
    const foundLong = r.edges.some(
      (e) => e.kind === 'calls' && e.metadata?.receiverType === 'Long',
    );
    expect(foundLong).toBe(false);
  });

  it('cross-file factory (function not in this file) leaves x untyped — no false resolution', () => {
    const src = `Public Sub Usar()
  Set x = CrearExterno()
  x.Hacer 1
End Sub`;
    const r = extract('src/modules/Consumidor.bas', src);
    // No same-file return type is known, so the receiver stays the raw `x`
    // (existing undeclared-receiver behavior) — never a bogus class.
    const edge = r.edges.find(
      (e) => e.kind === 'calls' && e.metadata?.member === 'Hacer',
    );
    expect(edge?.metadata?.receiverType).toBe('x');
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

// ---------------------------------------------------------------------------
// Antigravity audit Task 2: SQL_WRAPPERS must capture the FULL
// `&`-concatenation chain on a single physical line, not just the first
// literal. Today `db.Execute "FROM A" & " JOIN B"` only emits A; the wrapper
// regex's `((?:[^"]|"")*)` stops at the first closing quote. Cross-physical-
// line concatenation via `_` continuation is OUT OF SCOPE for v1 — see
// commit message for the deferred-work note.
// ---------------------------------------------------------------------------

describe('VbaExtractor — SQL wrapper captures multi-literal concat chains', () => {
  function sqlTableNames(r: ReturnType<typeof extract>): string[] {
    return r.edges
      .filter((e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-sql-table')
      .map((e) => r.nodes.find((n) => n.id === e.target)?.name)
      .filter((n): n is string => typeof n === 'string');
  }

  it('db.Execute with `&`-joined literals on the same line emits BOTH tables', () => {
    const src = [
      'Sub Q()',
      '  db.Execute "SELECT * FROM A" & " JOIN B ON b.id = a.id"',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Q.bas', src);
    const tableNames = sqlTableNames(r);
    // The bug: today only A is emitted because the wrapper regex stops at
    // the first closing quote. After the fix, both A and B must be emitted.
    expect(tableNames).toContain('A');
    expect(tableNames).toContain('B');
  });

  it('db.Execute single literal (regression guard) emits ONLY that table', () => {
    const src = [
      'Sub Q()',
      '  db.Execute "SELECT * FROM tblA"',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Q.bas', src);
    const tableNames = sqlTableNames(r);
    expect(tableNames).toEqual(['tblA']);
  });

  it('DoCmd.RunSQL single literal (regression guard) emits ONLY that table', () => {
    const src = [
      'Sub Q()',
      '  DoCmd.RunSQL "DELETE FROM tblOld"',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Q.bas', src);
    const tableNames = sqlTableNames(r);
    expect(tableNames).toEqual(['tblOld']);
  });

  it('DoCmd.RunSQL with `&`-joined literals emits BOTH tables', () => {
    const src = [
      'Sub Q()',
      '  DoCmd.RunSQL "DELETE FROM tblOld" & " WHERE id < 100"',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Q.bas', src);
    const tableNames = sqlTableNames(r);
    expect(tableNames).toContain('tblOld');
    // No extra table from the WHERE clause, but the wrapper must not have
    // been confused by the chain. We assert the exact set: only tblOld.
    expect(tableNames).toEqual(['tblOld']);
  });

  it('leading empty literal `"" & "FROM X"` emits ONLY X (no false positive)', () => {
    const src = [
      'Sub Q()',
      '  db.Execute "" & "FROM tblX"',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Q.bas', src);
    const tableNames = sqlTableNames(r);
    expect(tableNames).toEqual(['tblX']);
  });

  it('plain string assignment outside a wrapper does NOT emit a vba-sql-table edge', () => {
    // Regression guard: a bare assignment to a String variable must not
    // suddenly start emitting SQL table references just because the wrapper
    // scanner now also consumes `&`-chained literals. The wrapper scanner
    // is anchored to a wrapper-call signature — bare `x = "..."` must not
    // match it.
    const src = [
      'Sub Q()',
      '  Dim msg As String',
      '  msg = "FROM tblShouldNotEmit"',
      '  msg = msg & " and more text"',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Q.bas', src);
    const tableNames = sqlTableNames(r);
    expect(tableNames).not.toContain('tblShouldNotEmit');
    expect(tableNames).toEqual([]);
  });

  it('getdb().Execute with `&`-joined literals emits BOTH tables', () => {
    const src = [
      'Sub Q()',
      '  getdb().Execute "SELECT * FROM tblA" & " JOIN tblB ON tblB.a_id = tblA.id"',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Q.bas', src);
    const tableNames = sqlTableNames(r);
    expect(tableNames).toContain('tblA');
    expect(tableNames).toContain('tblB');
  });

  it('db.OpenRecordset with `&`-joined literals emits BOTH tables', () => {
    const src = [
      'Sub Q()',
      '  db.OpenRecordset "SELECT * FROM tblX" & " JOIN tblY ON tblY.x_id = tblX.id"',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Q.bas', src);
    const tableNames = sqlTableNames(r);
    expect(tableNames).toContain('tblX');
    expect(tableNames).toContain('tblY');
  });
});


// ---------------------------------------------------------------------------
// Issue #42: `DoCmd.RunSQL <identifier>` (variable form) must emit the same
// `vba-sql-table` references the literal form already emits.
//
// Today only `DoCmd.RunSQL "DELETE FROM X"` (literal) and `*db.Execute strSQL`
// (variable form on the `*db` family) are tracked. The dominant Access idiom
// `DoCmd.RunSQL strSQL` is invisible — table impact drops for every procedure
// that builds SQL in a string and runs it through `DoCmd.RunSQL`.
//
// Mirrors the existing `*db.Execute strSQL` coverage at lines 340-354
// (variable form for the `*db` family) but for the `DoCmd.RunSQL` method
// form. The SQL_VAR_ASSIGN_RE → trackSqlVariableAssignment path (Issue #13)
// already populates `sqlVariables` with `&`-accumulate semantics, so the
// concatenated-SQL atom validates that #13's accumulation flows through the
// new path.
// ---------------------------------------------------------------------------

describe('VbaExtractor — DoCmd.RunSQL with variable argument emits SQL table references (Issue #42)', () => {
  function sqlTableNames(r: ReturnType<typeof extract>): string[] {
    return r.edges
      .filter((e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-sql-table')
      .map((e) => r.nodes.find((n) => n.id === e.target)?.name)
      .filter((n): n is string => typeof n === 'string');
  }

  it('DoCmd.RunSQL strSQL with strSQL = "DELETE FROM TbX" emits TbX', () => {
    // Happy path: the variable form `DoCmd.RunSQL strSQL` must resolve
    // `strSQL` against the procedure-local `sqlVariables` map (populated by
    // `trackSqlVariableAssignment`) and emit `TbX` as a vba-sql-table.
    const src = [
      'Sub F()',
      '  Dim strSQL As String',
      '  strSQL = "DELETE FROM TbX"',
      '  DoCmd.RunSQL strSQL',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/F.bas', src);
    const tableNames = sqlTableNames(r);
    expect(tableNames).toContain('TbX');
  });

  it('DoCmd.RunSQL strSQL with `&`-concatenated SQL emits the initial FROM-table (Issue #13 accumulation through the new path)', () => {
    // Validates that the existing `&`-accumulate semantics
    // (`trackSqlVariableAssignment`) flows through the new DoCmd.RunSQL
    // variable path. The first literal declares `TbA`; the chained literal
    // adds a `WHERE` clause (no extra table). Only `TbA` should be emitted.
    const src = [
      'Sub F()',
      '  Dim strSQL As String',
      '  strSQL = "DELETE FROM TbA "',
      '  strSQL = strSQL & " WHERE id = 1"',
      '  DoCmd.RunSQL strSQL',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/F.bas', src);
    const tableNames = sqlTableNames(r);
    expect(tableNames).toContain('TbA');
  });

  it('literal `DoCmd.RunSQL "DELETE FROM TbY"` still emits TbY (regression guard)', () => {
    // The new regex must not break the existing literal-form coverage
    // already exercised at line 1856 ("DoCmd.RunSQL single literal").
    const src = [
      'Sub Q()',
      '  DoCmd.RunSQL "DELETE FROM TbY"',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Q.bas', src);
    const tableNames = sqlTableNames(r);
    expect(tableNames).toEqual(['TbY']);
  });

  it('DoCmd.RunSQL with an undeclared identifier emits zero vba-sql-table edges', () => {
    // Negative: when the captured identifier was never assigned (no row in
    // `sqlVariables`), the new path must stay silent — same behavior as the
    // existing `*db.Execute <undeclared>` path (graceful no-op).
    const src = [
      'Sub F()',
      '  DoCmd.RunSQL undeclared_var',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/F.bas', src);
    const tableNames = sqlTableNames(r);
    expect(tableNames).toEqual([]);
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

// ---------------------------------------------------------------------------
// Antigravity audit Task 3: Variant / untyped local vars must NOT emit a
// qualified-call `calls` edge to a stub named `<receiver>.<member>`. The
// receiver is registered into `localVarTypeMap` as a primitive (`variant`,
// `object`, ...) when the source declares it either as `Dim x` (no `As`
// clause) or `Dim x As Variant|Object|...`; the qualified-call site scan
// then skips ONLY when the receiver is mapped as a primitive, leaving the
// "undeclared receiver → stub → resolver repoints" path intact for
// cross-module qualified calls like `modUtils.Foo(...)`.
// ---------------------------------------------------------------------------

describe('VbaExtractor — Variant / untyped vars do not emit qualified-call stubs', () => {
  it('bare `Dim x` followed by `x.Method(1)` emits zero calls edges to a stub named `x.Method`', () => {
    // Bug fixed: a bare `Dim x` registers `x` as `variant` in
    // `localVarTypeMap`. The qualified paren-form call `x.Method(1)` is
    // then gated and emits no heuristic calls edge to a dead-end
    // `x.Method` stub. (The statement form `x.Method 1` was already silent
    // because `detectQualifiedStatementCall` requires the receiver be in
    // `localVarTypeMap` — the bug lived in the paren-form path.)
    const src = [
      'Sub F()',
      '  Dim x',
      '  x.Method(1)',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const stubEdges = r.edges.filter((e) => {
      if (e.kind !== 'calls') return false;
      const tgt = r.nodes.find((n) => n.id === e.target);
      return tgt?.name === 'x.Method';
    });
    expect(stubEdges).toHaveLength(0);
    const stubNodes = r.nodes.filter((n) => n.name === 'x.Method');
    expect(stubNodes).toHaveLength(0);
  });

  it('explicit `Dim x As Variant` followed by `x.Foo(1)` emits zero calls edges to a stub named `x.Foo`', () => {
    // Bug fixed: an explicit `Dim x As Variant` registers `x` as `variant`
    // in `localVarTypeMap`. The qualified paren-form call `x.Foo(1)` is
    // gated.
    const src = [
      'Sub F2()',
      '  Dim x As Variant',
      '  x.Foo(1)',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const stubEdges = r.edges.filter((e) => {
      if (e.kind !== 'calls') return false;
      const tgt = r.nodes.find((n) => n.id === e.target);
      return tgt?.name === 'x.Foo';
    });
    expect(stubEdges).toHaveLength(0);
    const stubNodes = r.nodes.filter((n) => n.name === 'x.Foo');
    expect(stubNodes).toHaveLength(0);
  });

  it('regression: `Dim m_Srv As Srv` followed by `m_Srv.Registrar` still emits a calls edge to `Srv.Registrar` (resolved)', () => {
    // Regression guard: project-class typed receivers (non-primitive) keep
    // their existing emission behavior. The refined gate skips ONLY when
    // the receiver is mapped as a primitive.
    const src = [
      'Sub F3()',
      '  Dim m_Srv As Srv',
      '  m_Srv.Registrar 1',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const hEdges = r.edges.filter(
      (e) => e.kind === 'calls' && e.provenance === 'heuristic' &&
             e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(hEdges.length).toBeGreaterThanOrEqual(1);
    const target = r.nodes.find((n) => n.id === hEdges[0]?.target);
    expect(target?.name).toBe('Srv.Registrar');
    expect(target?.metadata?.stub).toBe(true);
  });

  it('regression: cross-module qualified call `modUtils.Foo(1)` still emits a calls edge to `modUtils.Foo`', () => {
    // Critical regression guard: the unified `shouldProcessQualifiedCall`
    // gate MUST NOT regress cross-module qualified calls. `modUtils` is
    // NOT in `localVarTypeMap` (it is not declared as a local variable)
    // so the unified gate treats it as a module-name candidate and the
    // heuristic stub is emitted — the post-extraction resolver may
    // repoint it to a real `modUtils.Foo` if one exists.
    //
    // After Issue #40, the statement form `modUtils.Foo arg` follows
    // the SAME unified gate (it now routes through
    // `shouldProcessQualifiedCall` instead of the older
    // `isLocalProjectClassVar`-only path) — so it also emits. This is
    // pinned separately by the "qualified statement-form call on
    // undeclared module receiver emits heuristic edge" test.
    const src = [
      'Sub F4()',
      '  modUtils.Foo(1)',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const hEdges = r.edges.filter(
      (e) => e.kind === 'calls' && e.provenance === 'heuristic' &&
             e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(hEdges.length).toBeGreaterThanOrEqual(1);
    const stubEdge = hEdges.find((e) => {
      const tgt = r.nodes.find((n) => n.id === e.target);
      return tgt?.name === 'modUtils.Foo';
    });
    expect(stubEdge).toBeDefined();
    const target = r.nodes.find((n) => n.id === stubEdge?.target);
    expect(target?.metadata?.stub).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// Issue #46: capture `Set x = New ClassName` late-instantiation.
//
// VBA's dominant idiom is `Set svc = New ACService` followed by
// `svc.Registrar datos`. The PR #61 refined gate already supports
// resolving `svc.Method` when `svc` is in `localVarTypeMap` — but the gate
// is populated only by `Dim` (and `Global`/`Static` since #47). A late
// `Set x = New Foo` without an accompanying typed `Dim x As Foo` left `x`
// invisible to the gate, so `x.Method` was silently dropped.
//
// The fix: scan `Set x = New <Type>[.<Inner>]` lines in the same
// proc-stack order as the `Dim` sweep, emit a `references` edge tagged
// `vba-set-new`, AND register `x` in `localVarTypeMap` so the PR #61
// refined gate lets subsequent qualified calls resolve.
// ---------------------------------------------------------------------------

describe('VbaExtractor — Set x = New ClassName instantiations (Issue #46)', () => {
  it('happy path (statement form): Set svc = New ACService + svc.Registrar "data" emits a vba-set-new references edge AND a calls edge to ACService.Registrar', () => {
    const src = [
      'Sub F()',
      '  Set svc = New ACService',
      '  svc.Registrar "data"',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);

    // 1. references edge to ACService tagged vba-set-new.
    const setRefs = r.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-set-new',
    );
    expect(setRefs.length).toBeGreaterThanOrEqual(1);
    const refTarget = r.nodes.find((n) => n.id === setRefs[0]?.target);
    expect(refTarget?.name).toBe('ACService');

    // 2. calls edge from F to ACService.Registrar (resolved via the PR #61
    //    refined gate, because `svc` is now in localVarTypeMap with outer=ACService).
    const hEdges = r.edges.filter(
      (e) => e.kind === 'calls' && e.provenance === 'heuristic' &&
             e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(hEdges.length).toBeGreaterThanOrEqual(1);
    const callsTarget = r.nodes.find((n) => n.id === hEdges[0]?.target);
    expect(callsTarget?.name).toBe('ACService.Registrar');
  });

  it('happy path (paren form): Set svc = New ACService + svc.Registrar("data") emits the same calls edge to ACService.Registrar', () => {
    const src = [
      'Sub F()',
      '  Set svc = New ACService',
      '  svc.Registrar("data")',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const hEdges = r.edges.filter(
      (e) => e.kind === 'calls' && e.provenance === 'heuristic' &&
             e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(hEdges.length).toBeGreaterThanOrEqual(1);
    const callsTarget = r.nodes.find((n) => n.id === hEdges[0]?.target);
    expect(callsTarget?.name).toBe('ACService.Registrar');
  });

  it('qualified New type: Set rs = New DAO.Recordset emits a vba-set-new references edge to DAO AND keeps rs.MoveNext silent (qualified type suppresses calls)', () => {
    const src = [
      'Sub Q()',
      '  Set rs = New DAO.Recordset',
      '  rs.MoveNext',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);

    // 1. references edge to DAO tagged vba-set-new.
    const setRefs = r.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-set-new',
    );
    expect(setRefs.length).toBeGreaterThanOrEqual(1);
    const refTarget = r.nodes.find((n) => n.id === setRefs[0]?.target);
    expect(refTarget?.name).toBe('DAO');

    // 2. No calls edge for rs.MoveNext or DAO.MoveNext — qualified type
    //    registers `qualified: true` in localVarTypeMap, mirroring
    //    `Dim rs As DAO.Recordset` semantics so the PR #61 gate stays silent.
    const callsEdges = r.edges.filter((e) => {
      if (e.kind !== 'calls' || e.provenance !== 'heuristic') return false;
      const tgt = r.nodes.find((n) => n.id === e.target);
      return tgt?.name === 'rs.MoveNext' || tgt?.name === 'DAO.MoveNext';
    });
    expect(callsEdges).toHaveLength(0);
  });

  it('regression: Dim x As New Foo still emits a vba-name-resolution references edge to Foo (auto-instantiation path unchanged)', () => {
    const src = `Dim x As New Foo`;
    const r = extract('src/modules/Mod.bas', src);
    const refs = r.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(refs).toHaveLength(1);
    const target = r.nodes.find((n) => n.id === refs[0]?.target);
    expect(target?.name).toBe('Foo');
  });

  it('module-scope Set svc = New ACService (no enclosing Sub/Function) emits no vba-set-new references edge', () => {
    // The Set scan lives inside sweepCallsAndSql, which is gated on the
    // proc-stack being non-empty (consistent with how every other call-site
    // detector handles module scope). A bare `Set svc = New ACService` at
    // module scope is silently a no-op.
    const src = `Set svc = New ACService`;
    const r = extract('src/modules/Mod.bas', src);
    const setRefs = r.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-set-new',
    );
    expect(setRefs).toHaveLength(0);
  });
});

// Issue #54: bracketed receivers in qualified calls.
//
// VBA in Access projects supports module/object names with spaces
// (`FUNCIONES UTILES.bas`, `Funciones Generales.bas`). Call sites qualify
// such modules with brackets:
//   `[FUNCIONES UTILES].FormatearFecha(fecha)`     — paren form
//   `[FUNCIONES UTILES].FormatearFecha fecha`     — statement form
//   `Dim x As [Clase Con Espacios]`               — bracketed type in Dim
//
// Before the fix, `CALL_RE`, `detectQualifiedStatementCall`, and the type
// alternative in `DIM_ALL_VARS_RE` only accepted `\p{L}[\p{L}\p{N}_]*` for
// the identifier token — so bracketed-qualified calls were invisible and the
// corpus from a Dysflow-managed project silently dropped cross-module
// references to modules with spaces in their names.
//
// Fix: extend the receiver / type / var-name alternatives to accept the
// bracketed form (`[Name With Spaces]`), and unwrap brackets BEFORE the
// blacklist checks and the `resolveReceiverType` lookup so the stub's
// `qualifiedName` matches the `${name}.${proc}` shape.
// ---------------------------------------------------------------------------

describe('VbaExtractor — bracketed receivers (Issue #54)', () => {
  it('paren form `[FUNCIONES UTILES].FormatearFecha(fecha)` emits a heuristic calls edge to `FUNCIONES UTILES.FormatearFecha`', () => {
    // The bracketed module-name receiver is unwrapped to `FUNCIONES UTILES`
    // before the `localVarTypeMap` lookup and the `qualified` shape is built.
    // `FUNCIONES UTILES` is NOT a file-local declared variable (no Dim for
    // it), so the paren-form path's "undeclared receiver → stub → resolver
    // repoints" gate applies: the heuristic stub is emitted and the post-
    // extraction resolver may repoint it to a real
    // `[FUNCIONES UTILES].FormatearFecha` if one exists.
    const src = [
      'Sub Outer()',
      '  [FUNCIONES UTILES].FormatearFecha(fecha)',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const hEdge = r.edges.find(
      (e) =>
        e.kind === 'calls' &&
        e.provenance === 'heuristic' &&
        e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(hEdge).toBeDefined();
    expect(hEdge?.metadata?.receiverType).toBe('FUNCIONES UTILES');
    expect(hEdge?.metadata?.member).toBe('FormatearFecha');
    const target = r.nodes.find((n) => n.id === hEdge?.target);
    expect(target?.name).toBe('FUNCIONES UTILES.FormatearFecha');
    expect(target?.metadata?.stub).toBe(true);
    // Brackets must NOT leak into the target node name.
    expect(r.nodes.some((n) => n.name === '[FUNCIONES UTILES]')).toBe(false);
    expect(
      r.nodes.some((n) => n.name === '[FUNCIONES UTILES].FormatearFecha'),
    ).toBe(false);
  });

  it('statement form `m_FU.FormatearFecha fecha` (declared via bracketed Dim `As [FUNCIONES UTILES]`) emits edge to `FUNCIONES UTILES.FormatearFecha`', () => {
    // AC #2: when the bracketed type appears in a Dim declaration
    // (`Dim m_FU As [FUNCIONES UTILES]`), DIM_ALL_VARS_RE unwraps the
    // brackets and stores `FUNCIONES UTILES` as the resolved outer type
    // for `m_FU` in `localVarTypeMap`. The subsequent
    // `m_FU.FormatearFecha fecha` call is then gated by
    // `isLocalProjectClassVar('m_FU') === true`, and
    // `resolveReceiverType('m_FU')` returns the unwrapped class name —
    // so the stub's qualifiedName matches `FUNCIONES UTILES.FormatearFecha`.
    //
    // Note: the literal prompt example `[FUNCIONES UTILES].FormatearFecha fecha`
    // (with the bracketed module name as the call receiver, no Dim) goes
    // through the unified `shouldProcessQualifiedCall` gate: the bracket
    // strip inside `isLocalProjectClassVar` does not find
    // `FUNCIONES UTILES` in `localVarTypeMap` (no Dim), so the gate
    // classifies the receiver as an undeclared module-name candidate and
    // EMITS the heuristic `FUNCIONES UTILES.FormatearFecha` stub —
    // matching the paren-form behaviour and letting the post-extraction
    // resolver repoint to the real sibling module if one exists.
    //
    // The `UnknownExternal.Whatever` shape (where the receiver IS a
    // declared external primitive/qualified type) is what stays silent,
    // via the `entry.qualified` branch inside the same unified gate.
    const src = [
      'Sub Outer()',
      '  Dim m_FU As [FUNCIONES UTILES]',
      '  m_FU.FormatearFecha fecha',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const hEdge = r.edges.find(
      (e) =>
        e.kind === 'calls' &&
        e.provenance === 'heuristic' &&
        e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(hEdge).toBeDefined();
    expect(hEdge?.metadata?.receiverType).toBe('FUNCIONES UTILES');
    expect(hEdge?.metadata?.member).toBe('FormatearFecha');
    const target = r.nodes.find((n) => n.id === hEdge?.target);
    expect(target?.name).toBe('FUNCIONES UTILES.FormatearFecha');
    expect(target?.metadata?.stub).toBe(true);
    // Brackets must NOT leak into the target node name.
    expect(r.nodes.some((n) => n.name === '[FUNCIONES UTILES]')).toBe(false);
  });

  it('`Dim x As [Clase Con Espacios]` emits a vba-name-resolution references edge to `Clase Con Espacios`', () => {
    // The TYPE position in DIM_ALL_VARS_RE accepts bracketed identifiers
    // with spaces and the brackets are unwrapped before populating the
    // reference edge's target name.
    const src = `Dim x As [Clase Con Espacios]`;
    const r = extract('src/modules/Mod.bas', src);
    const refs = r.edges.filter(
      (e) =>
        e.kind === 'references' &&
        e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(refs).toHaveLength(1);
    const target = r.nodes.find((n) => n.id === refs[0]?.target);
    expect(target?.name).toBe('Clase Con Espacios');
    // Brackets must NOT leak into the target node name.
    expect(r.nodes.some((n) => n.name === '[Clase Con Espacios]')).toBe(false);
  });

  it('regression: plain (non-bracketed) `Foo.Bar(...)` keeps the existing qualified-call shape', () => {
    // Critical regression guard: the new bracketed alternative must not
    // shift byte-identical behaviour for the dominant un-bracketed shape.
    // `Foo` is undeclared → heuristic stub `Foo.Bar` is emitted unchanged.
    const src = [
      'Sub Outer()',
      '  Foo.Bar(1, 2)',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const hEdge = r.edges.find(
      (e) =>
        e.kind === 'calls' &&
        e.provenance === 'heuristic' &&
        e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(hEdge).toBeDefined();
    expect(hEdge?.metadata?.receiverType).toBe('Foo');
    expect(hEdge?.metadata?.member).toBe('Bar');
    const target = r.nodes.find((n) => n.id === hEdge?.target);
    expect(target?.name).toBe('Foo.Bar');
    expect(target?.metadata?.stub).toBe(true);
  });

  it('regression: cross-module `modUtils.Foo(1)` keeps the existing qualified-call shape', () => {
    // Critical regression guard: the dominant `.bas`-qualified call shape
    // must keep emitting the existing `modUtils.Foo` heuristic stub
    // byte-identical to today. `modUtils` is NOT in `localVarTypeMap`, so
    // the existing undeclared-receiver → heuristic stub path applies.
    const src = [
      'Sub Outer()',
      '  modUtils.Foo(1)',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const hEdge = r.edges.find(
      (e) =>
        e.kind === 'calls' &&
        e.provenance === 'heuristic' &&
        e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(hEdge).toBeDefined();
    expect(hEdge?.metadata?.receiverType).toBe('modUtils');
    expect(hEdge?.metadata?.member).toBe('Foo');
    const target = r.nodes.find((n) => n.id === hEdge?.target);
    expect(target?.name).toBe('modUtils.Foo');
    expect(target?.metadata?.stub).toBe(true);
  });

  it('regression: class-typed `m_Service.Method(1)` keeps the existing resolved-class qualified-call shape', () => {
    // Critical regression guard: the class-typed local-var qualified-call
    // shape must keep resolving the receiver as the declared class type
    // (`Service`, not `m_Service`) so the stub's qualifiedName matches
    // the real `.cls` method's `${className}.${proc}` shape.
    const src = [
      'Sub Outer()',
      '  Dim m_Service As Service',
      '  m_Service.Method(1)',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const hEdge = r.edges.find(
      (e) =>
        e.kind === 'calls' &&
        e.provenance === 'heuristic' &&
        e.metadata?.synthesizedBy === 'vba-name-resolution',
    );
    expect(hEdge).toBeDefined();
    expect(hEdge?.metadata?.receiverType).toBe('Service');
    expect(hEdge?.metadata?.member).toBe('Method');
    const target = r.nodes.find((n) => n.id === hEdge?.target);
    expect(target?.name).toBe('Service.Method');
    expect(target?.metadata?.stub).toBe(true);
  });
});

// Issue #45: single-line `If <cond> Then <call>` (and `Else <call>`) MUST
// emit the same `calls` edge as their multi-line / standalone form.
//
// Bug: `detectStatementCall` extracted the FIRST identifier of the line —
// `If` for `If x Then Foo arg`, which is in `CALL_KEYWORD_BLACKLIST`, so
// the dominant VBA idiom for early-exit guards
// (`If Err.Number <> 0 Then GestionarError`) was silently dropped. The same
// gap affected the qualified path. Issue #45 closes that hole: strip the
// `If … Then` prefix and re-run the statement-form detector on the body,
// splitting on `Else` and on `:` (multi-statement) so all clauses are
// captured. `GoTo` / `Exit` / `Resume` clauses are deliberately silent.
// ---------------------------------------------------------------------------

describe('VbaExtractor — single-line `If … Then` statement-form calls (Issue #45)', () => {
  it('a bare statement-form call after `If … Then` emits a same-file calls edge', () => {
    // `If Err.Number <> 0 Then GestionarError` — the canonical early-exit
    // guard from real Dysflow form fixtures. Pre-#45 this dropped the
    // `GestionarError` call silently because the leading identifier was
    // `If`, which is blacklisted as a control-flow keyword.
    const src = [
      'Public Sub Caller()',
      '    If Err.Number <> 0 Then GestionarError',
      'End Sub',
      'Public Sub GestionarError()',
      '    \' body',
      'End Sub',
    ].join('\n');

    const r = extract('src/modules/modCaller.bas', src);
    const edges = r.edges.filter((e) => {
      if (e.kind !== 'calls') return false;
      const srcNode = r.nodes.find((n) => n.id === e.source);
      const tgt = r.nodes.find((n) => n.id === e.target);
      return srcNode?.name === 'Caller' && tgt?.name === 'GestionarError';
    });
    expect(edges.length).toBeGreaterThanOrEqual(1);
  });

  it('`If x Then Foo Else Bar` emits BOTH Foo and Bar calls edges', () => {
    const src = [
      'Public Sub Caller()',
      '    If x > 0 Then ProcesarDato x, True Else LimpiarDato',
      'End Sub',
      'Public Sub ProcesarDato(ByVal v As Long, ByVal b As Boolean)',
      '    \' body',
      'End Sub',
      'Public Sub LimpiarDato()',
      '    \' body',
      'End Sub',
    ].join('\n');

    const r = extract('src/modules/modCaller.bas', src);
    const caller = r.nodes.find((n) => n.kind === 'function' && n.name === 'Caller');
    const procesar = r.nodes.find((n) => n.kind === 'function' && n.name === 'ProcesarDato');
    const limpiar = r.nodes.find((n) => n.kind === 'function' && n.name === 'LimpiarDato');
    expect(caller).toBeDefined();
    expect(procesar).toBeDefined();
    expect(limpiar).toBeDefined();
    const edges = r.edges.filter(
      (e) => e.kind === 'calls' && e.source === caller?.id,
    );
    const processedEdge = edges.find((e) => e.target === procesar?.id);
    const cleanedEdge = edges.find((e) => e.target === limpiar?.id);
    expect(processedEdge).toBeDefined();
    expect(cleanedEdge).toBeDefined();
  });

  it('`GoTo` clause after `If … Then` is silent (GoTo is a control-flow keyword, not a Sub call)', () => {
    // Defense-in-depth guard at the clause level: even if a project had a
    // Sub named `GoToSomething`, the clause-level splitter explicitly skips
    // any sub-clause that starts with the bare `GoTo` keyword (mirrors the
    // blacklisting in `emitStatementCallEdge`).
    const src = [
      'Public Sub Caller()',
      '    If Err.Number <> 0 Then GoTo fin',
      'fin:',
      '    \' body',
      'End Sub',
    ].join('\n');

    const r = extract('src/modules/modCaller.bas', src);
    // No calls edge with a `GoTo*` target.
    const gotoEdges = r.edges.filter((e) => {
      if (e.kind !== 'calls') return false;
      const tgt = r.nodes.find((n) => n.id === e.target);
      return typeof tgt?.name === 'string' && /^(GoTo|GoToSomething)/.test(tgt.name);
    });
    expect(gotoEdges).toHaveLength(0);
  });

  it('`Exit Sub` clause after `If … Then` is silent (Exit is a control-flow keyword)', () => {
    const src = [
      'Public Sub Caller()',
      '    If Err.Number <> 0 Then Exit Sub',
      '    \' body',
      'End Sub',
    ].join('\n');

    const r = extract('src/modules/modCaller.bas', src);
    const exitEdges = r.edges.filter((e) => {
      if (e.kind !== 'calls') return false;
      const tgt = r.nodes.find((n) => n.id === e.target);
      return typeof tgt?.name === 'string' && /^Exit$/i.test(tgt.name);
    });
    expect(exitEdges).toHaveLength(0);
  });

  it('multi-statement `If x Then DoA: DoB` (colon-separated) emits BOTH calls edges', () => {
    const src = [
      'Public Sub Caller()',
      '    If x > 0 Then DoA: DoB',
      'End Sub',
      'Public Sub DoA()',
      'End Sub',
      'Public Sub DoB()',
      'End Sub',
    ].join('\n');

    const r = extract('src/modules/modCaller.bas', src);
    const caller = r.nodes.find((n) => n.kind === 'function' && n.name === 'Caller');
    const doA = r.nodes.find((n) => n.kind === 'function' && n.name === 'DoA');
    const doB = r.nodes.find((n) => n.kind === 'function' && n.name === 'DoB');
    expect(caller).toBeDefined();
    expect(doA).toBeDefined();
    expect(doB).toBeDefined();
    const edges = r.edges.filter(
      (e) => e.kind === 'calls' && e.source === caller?.id,
    );
    expect(edges.find((e) => e.target === doA?.id)).toBeDefined();
    expect(edges.find((e) => e.target === doB?.id)).toBeDefined();
  });

  it('block-form `If x Then` on its own line with body on next line is unchanged', () => {
    // Regression guard: the `If … Then` matcher ONLY fires when something
    // comes after `Then` on the same line (the single-line If shape). When
    // `Then` is the last token on the line, the body lives on subsequent
    // lines and is picked up by the existing per-line call-site scan that
    // already handles bare `Foo` on a line of its own.
    const src = [
      'Public Sub Caller()',
      '    If x > 0 Then',
      '        Foo',
      '    End If',
      'End Sub',
      'Public Sub Foo()',
      'End Sub',
    ].join('\n');

    const r = extract('src/modules/modCaller.bas', src);
    const caller = r.nodes.find((n) => n.kind === 'function' && n.name === 'Caller');
    const foo = r.nodes.find((n) => n.kind === 'function' && n.name === 'Foo');
    expect(caller).toBeDefined();
    expect(foo).toBeDefined();
    // The existing per-line call-site scan emits the edge from `Caller` to
    // `Foo` regardless of which line `Foo` is on, so this regression
    // guard asserts the same behavior is preserved.
    const edge = r.edges.find(
      (e) => e.kind === 'calls' && e.source === caller?.id && e.target === foo?.id,
    );
    expect(edge).toBeDefined();
  });

  it('`Call` keyword with statement-form arguments after `If … Then` also emits a calls edge', () => {
    // The `Call Foo arg` form (Call keyword, no parens) is the same call
    // shape after the `Call ` keyword is stripped — keep parity with the
    // pre-existing H1 invariant for `Call Sub 1, 2` standalone.
    const src = [
      'Public Sub Caller()',
      '    If x > 0 Then Call GestionarError',
      'End Sub',
      'Public Sub GestionarError()',
      '    \' body',
      'End Sub',
    ].join('\n');

    const r = extract('src/modules/modCaller.bas', src);
    const caller = r.nodes.find((n) => n.kind === 'function' && n.name === 'Caller');
    const gestionar = r.nodes.find((n) => n.kind === 'function' && n.name === 'GestionarError');
    expect(caller).toBeDefined();
    expect(gestionar).toBeDefined();
    const edge = r.edges.find(
      (e) => e.kind === 'calls' && e.source === caller?.id && e.target === gestionar?.id,
    );
    expect(edge).toBeDefined();
  });

  it('a qualified statement-form call after `If … Then` follows the existing qualified-call gate', () => {
    // `m_Srv.Registrar 1` inside an `If … Then` body — pre-#45 this dropped
    // the call silently because the leading identifier was `If`. Post-fix,
    // the qualified detector runs on the clause and emits an edge IF the
    // receiver is in `localVarTypeMap` as a non-primitive (Fix 2 invariant
    // preserved).
    const src = [
      'Sub Outer()',
      '    Dim m_Srv As Srv',
      '    If x > 0 Then m_Srv.Registrar 1',
      'End Sub',
    ].join('\n');

    const r = extract('src/modules/modCaller.bas', src);
    const caller = r.nodes.find((n) => n.kind === 'function' && n.name === 'Outer');
    expect(caller).toBeDefined();
    // Stub for `Srv.Registrar` (resolved class type, Fix 2 / #12a).
    const stubEdge = r.edges.find((e) => {
      if (e.kind !== 'calls') return false;
      if (e.source !== caller?.id) return false;
      const tgt = r.nodes.find((n) => n.id === e.target);
      return tgt?.name === 'Srv.Registrar';
    });
    expect(stubEdge).toBeDefined();
  });
});

// ---------------------------------------------------------------------------
// Issue #43: member calls inside `With <receiver>` blocks.
//
// VBA uses `.Member` inside a With block as shorthand for
// `<receiver>.Member`. Before the fix, the extractor saw a leading dot and
// emitted no call edge, so real form/service helpers hidden behind With blocks
// disappeared from the graph.
// ---------------------------------------------------------------------------

describe('VbaExtractor — With block implicit receiver calls (Issue #43)', () => {
  it('statement-form `.Registrar "data"` inside `With svc` emits a calls edge to the declared class receiver', () => {
    const src = [
      'Sub F()',
      '  Dim svc As ACService',
      '  With svc',
      '    .Registrar "data"',
      '  End With',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const hEdge = r.edges.find((e) => {
      if (e.kind !== 'calls' || e.provenance !== 'heuristic') return false;
      const tgt = r.nodes.find((n) => n.id === e.target);
      return tgt?.name === 'ACService.Registrar';
    });
    expect(hEdge).toBeDefined();
    expect(hEdge?.metadata?.synthesizedBy).toBe('vba-name-resolution');
    expect(hEdge?.metadata?.receiverType).toBe('ACService');
    expect(hEdge?.metadata?.member).toBe('Registrar');
  });

  it('paren-form `.Registrar("data")` inside `With svc` emits the same calls edge', () => {
    const src = [
      'Sub F()',
      '  Dim svc As ACService',
      '  With svc',
      '    .Registrar("data")',
      '  End With',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const target = r.nodes.find((n) => n.name === 'ACService.Registrar');
    expect(target?.metadata?.stub).toBe(true);
    const edge = r.edges.find((e) => e.kind === 'calls' && e.target === target?.id);
    expect(edge).toBeDefined();
  });

  it('property assignment `.Caption = ...` inside a With block stays silent', () => {
    const src = [
      'Sub F()',
      '  Dim label As LabelView',
      '  With label',
      '    .Caption = "Done"',
      '  End With',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const target = r.nodes.find((n) => n.name === 'LabelView.Caption');
    expect(target).toBeUndefined();
  });

  it('runtime typed receiver `With rs As DAO.Recordset` keeps `.MoveNext` silent', () => {
    const src = [
      'Sub F()',
      '  Dim rs As DAO.Recordset',
      '  With rs',
      '    .MoveNext',
      '  End With',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const bogusTargets = r.nodes.filter(
      (n) => n.name === 'rs.MoveNext' || n.name === 'DAO.MoveNext',
    );
    expect(bogusTargets).toHaveLength(0);
  });

  it('nested With blocks restore the outer receiver after `End With`', () => {
    const src = [
      'Sub F()',
      '  Dim outerSvc As OuterService',
      '  Dim innerSvc As InnerService',
      '  With outerSvc',
      '    .Before',
      '    With innerSvc',
      '      .Inside',
      '    End With',
      '    .After',
      '  End With',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const callTargets = r.edges
      .filter((e) => e.kind === 'calls' && e.provenance === 'heuristic')
      .map((e) => r.nodes.find((n) => n.id === e.target)?.name)
      .filter(Boolean)
      .sort();
    expect(callTargets).toContain('OuterService.Before');
    expect(callTargets).toContain('InnerService.Inside');
    expect(callTargets).toContain('OuterService.After');
  });

  it('single-line If clause inside With resolves `.Registrar` through the active receiver', () => {
    const src = [
      'Sub F()',
      '  Dim svc As ACService',
      '  With svc',
      '    If ready Then .Registrar "data"',
      '  End With',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const target = r.nodes.find((n) => n.name === 'ACService.Registrar');
    expect(target).toBeDefined();
  });
});

// ---------------------------------------------------------------------------
// Issue #52 — procedure-local Const must NOT leak to module scope (was:
// one module-level `constant` node per name with module→constant contains
// edge and `visibility: 'public'`, plus a single file-wide `localConstants`
// Map that caused two procs declaring the same Const name to collide on the
// last write for `DoCmd.OpenForm`/`OpenReport`/`OpenQuery` argument
// resolution). Fix: per-proc resolution bucket + a proc-stack-aware
// `currentProcKey` field. Module-level Const behavior is preserved bit-for-
// bit (regression-pinned below) — only the proc-local shape changes.
//
// Reference: the dedicated Enum/Const tests live in
// `__tests__/extraction-vba-enums-consts.test.ts` (REQ-CODE-12 / REQ-CODE-13
// atoms). The atoms in this block pin the Issue #52 fix specifically.
// ---------------------------------------------------------------------------

describe('VbaExtractor — Issue #52: procedure-local Const scoping', () => {
  it('module-level Const still emits a constant node + module contains edge (regression)', () => {
    const src = [
      'Option Explicit',
      'Public Const FORM_EMPLOYEES As String = "frmEmployees"',
      'Public Sub DoSomething()',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/modForms.bas', src);

    const c = r.nodes.find(
      (n) => n.kind === 'constant' && n.name === 'FORM_EMPLOYEES',
    );
    // Pre-fix this was the buggy behavior and post-fix must still hold
    // because the Const lives at module scope (no enclosing proc).
    expect(c).toBeDefined();
    expect(c?.visibility).toBe('public');
    const mod = r.nodes.find((n) => n.kind === 'module');
    expect(mod).toBeDefined();
    const containsEdge = r.edges.find(
      (e) =>
        e.kind === 'contains' && e.source === mod?.id && e.target === c?.id,
    );
    expect(containsEdge).toBeDefined();
  });

  it('procedure-local Const emits NO constant node anywhere (the bug)', () => {
    const src = [
      'Sub Abrir()',
      '    Const FORM_DESTINO As String = "FormDetalle"',
      '    DoCmd.OpenForm FORM_DESTINO',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/modForms.bas', src);

    // The pre-fix extractor emitted a `constant` node named FORM_DESTINO
    // with `visibility: 'public'` and a module→constant contains edge —
    // wrong containment for a local. Post-fix: zero `constant` nodes.
    const consts = r.nodes.filter((n) => n.kind === 'constant');
    expect(consts).toEqual([]);
  });

  it('procedure-local Const still resolves in DoCmd.OpenForm (resolution preserved)', () => {
    const src = [
      'Sub Abrir()',
      '    Const FORM_DESTINO As String = "FormDetalle"',
      '    DoCmd.OpenForm FORM_DESTINO',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/modForms.bas', src);

    const edges = r.edges.filter((e) => e.kind === 'opens-form');
    expect(edges).toHaveLength(1);
    const target = r.nodes.find((n) => n.id === edges[0]?.target);
    expect(target?.name).toBe('FormDetalle');
    expect(edges[0]?.metadata?.targetFormName).toBe('FormDetalle');
  });

  it('two procs with same-named local consts resolve their own OpenForm targets', () => {
    const src = [
      'Sub A()',
      '  Const TARGET As String = "FormA"',
      '  DoCmd.OpenForm TARGET',
      'End Sub',
      'Sub B()',
      '  Const TARGET As String = "FormB"',
      '  DoCmd.OpenForm TARGET',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/modForms.bas', src);

    // Pre-fix the file-wide `localConstants` Map made the second write
    // (`FormB`) overwrite the first, so both OpenForm call sites
    // resolved to `FormB`. Post-fix: each proc keeps its own bucket.
    const edges = r.edges.filter((e) => e.kind === 'opens-form');
    expect(edges).toHaveLength(2);
    const targets = edges
      .map((e) => e.metadata?.targetFormName)
      .sort();
    expect(targets).toEqual(['FormA', 'FormB']);

    // No `constant` nodes for TARGET at all — proc-local consts don't
    // emit module-level symbol nodes anymore.
    const consts = r.nodes.filter(
      (n) => n.kind === 'constant' && n.name === 'TARGET',
    );
    expect(consts).toEqual([]);
  });

  it('mixed-scope consts: module-level emits one node; proc-local shadows it for the inner call', () => {
    const src = [
      'Public Const SHARED_NAME As String = "ModuleShared"',
      'Sub Outer()',
      '    Const SHARED_NAME As String = "ProcShared"',
      '    DoCmd.OpenForm SHARED_NAME',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/modForms.bas', src);

    // Module-level Const still emits one node with the visibility fold.
    const moduleConst = r.nodes.find(
      (n) => n.kind === 'constant' && n.name === 'SHARED_NAME',
    );
    expect(moduleConst).toBeDefined();
    expect(moduleConst?.visibility).toBe('public');

    // The OpenForm call inside Outer() sees the proc-local binding
    // (shadowing the module-level one) — exactly one opens-form edge,
    // resolved to the proc-local value.
    const edges = r.edges.filter((e) => e.kind === 'opens-form');
    expect(edges).toHaveLength(1);
    expect(edges[0]?.metadata?.targetFormName).toBe('ProcShared');

    // Proc-local name does NOT get a second `constant` node.
    expect(
      r.nodes.filter((n) => n.kind === 'constant' && n.name === 'SHARED_NAME'),
    ).toHaveLength(1);
  });

  it('multi-decl Const line at module scope still emits one constant node per name (regression)', () => {
    const r = extract(
      'src/modules/c.bas',
      'Const FORM_EMPLOYEES = "frmEmployees", FORM_ORDERS As String = "frmOrders"',
    );
    const names = r.nodes
      .filter((n) => n.kind === 'constant')
      .map((n) => n.name)
      .sort();
    expect(names).toEqual(['FORM_EMPLOYEES', 'FORM_ORDERS']);
  });
});

// ---------------------------------------------------------------------------
// Issue #50 — TempVars keys as cross-form state nodes.
//
// Access's `TempVars` is a global key-value store that procedures across
// forms read/write to pass state (the canonical Access idiom for
// caller→callee state passing without globals). Each STATIC-LITERAL key
// reference (bang `TempVars!k`, paren `TempVars("k")`, Add `TempVars.Add "k"`)
// gets a `references` edge to a synthetic `class` placeholder node. The
// placeholder id is keyed on `synthetic:tempvar/<key>` (NOT on the
// extraction's `filePath`), so the same key referenced from many files
// collapses to ONE node — the cross-form state premise that lets
// `codegraph_explore` connect producer ⇄ consumer in one hop.
//
// Atoms verified below correspond 1:1 to the issue's acceptance criteria:
//   1. bang write     — `TempVars!MiClave = "valor"`
//   2. paren read     — `Dim x = TempVars("MiClave")` (same MiClave
//                       placeholder as test 1, cross-proc shared id)
//   3. paren write    — `TempVars("MiClave") = 42`
//   4. Add form       — `TempVars.Add "MiClave", "x"` (always write)
//   5. cross-file dedup — write in Form_A.cls + read in Form_B.cls == ONE
//                         placeholder id (deterministic, line-independent)
//   6. dynamic keys  — `TempVars(strNombre)` and `TempVars("a" & s)` emit
//                       ZERO edges and ZERO placeholders (REQ-CODE-4
//                       spirit; unresolvable stays silent)
//   7. cross-cutting  — one proc with two write sites (a, b) and a read of
//                       a → 2 placeholders, 3 edges from the same proc.
// ---------------------------------------------------------------------------

describe('VbaExtractor — Issue #50: TempVars keys modeled as cross-form state nodes', () => {
  it('bang-write form: TempVars!MiClave = "valor" emits one references edge with access=write', () => {
    const src = [
      'Public Sub Escribir()',
      '    TempVars!MiClave = "valor"',
      'End Sub',
    ].join('\n');
    const r = extract('src/forms/FormA.cls', src);

    const escribir = r.nodes.find(
      (n) => n.kind === 'function' && n.name === 'Escribir',
    );
    expect(escribir).toBeDefined();

    const placeholder = r.nodes.find((n) => n.name === 'MiClave' && n.kind === 'class');
    expect(placeholder).toBeDefined();
    // Synthetic file path keeps the id deterministic AND file-independent.
    expect(placeholder?.filePath).toBe('synthetic:tempvar/MiClave');

    const edges = r.edges.filter(
      (e) =>
        e.kind === 'references' &&
        e.metadata?.synthesizedBy === 'vba-tempvar' &&
        e.target === placeholder?.id &&
        e.source === escribir?.id,
    );
    expect(edges).toHaveLength(1);
    expect(edges[0]?.metadata?.access).toBe('write');
  });

  it('paren-read form: Dim x = TempVars("MiClave") emits one references edge with access=read', () => {
    // Same key `MiClave` as test 1 — same placeholder node (cross-proc
    // shared id is the cross-form case).
    const src = [
      'Public Sub Leer()',
      '    Dim x As Variant',
      '    x = TempVars("MiClave")',
      'End Sub',
    ].join('\n');
    const r = extract('src/forms/FormB.cls', src);

    const leer = r.nodes.find(
      (n) => n.kind === 'function' && n.name === 'Leer',
    );
    expect(leer).toBeDefined();

    const placeholder = r.nodes.find((n) => n.name === 'MiClave' && n.kind === 'class');
    expect(placeholder).toBeDefined();
    expect(placeholder?.filePath).toBe('synthetic:tempvar/MiClave');

    const edges = r.edges.filter(
      (e) =>
        e.kind === 'references' &&
        e.metadata?.synthesizedBy === 'vba-tempvar' &&
        e.target === placeholder?.id &&
        e.source === leer?.id,
    );
    expect(edges).toHaveLength(1);
    expect(edges[0]?.metadata?.access).toBe('read');
  });

  it('paren-write form: TempVars("MiClave") = 42 emits one references edge with access=write', () => {
    const src = [
      'Public Sub Escribir2()',
      '    TempVars("MiClave") = 42',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);

    const escribir2 = r.nodes.find(
      (n) => n.kind === 'function' && n.name === 'Escribir2',
    );
    expect(escribir2).toBeDefined();

    const placeholder = r.nodes.find((n) => n.name === 'MiClave' && n.kind === 'class');
    expect(placeholder).toBeDefined();

    const edges = r.edges.filter(
      (e) =>
        e.kind === 'references' &&
        e.metadata?.synthesizedBy === 'vba-tempvar' &&
        e.target === placeholder?.id &&
        e.source === escribir2?.id,
    );
    expect(edges).toHaveLength(1);
    expect(edges[0]?.metadata?.access).toBe('write');
  });

  it('Add form: TempVars.Add "MiClave", "x" emits one edge with access=write and adds a second placeholder for a distinct key', () => {
    const src = [
      'Public Sub Init()',
      '    TempVars.Add "MiClave", "x"',
      '    TempVars.Add "OtraClave", 7',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);

    const init = r.nodes.find(
      (n) => n.kind === 'function' && n.name === 'Init',
    );
    expect(init).toBeDefined();

    const miClave = r.nodes.find((n) => n.name === 'MiClave' && n.kind === 'class');
    const otraClave = r.nodes.find((n) => n.name === 'OtraClave' && n.kind === 'class');
    expect(miClave).toBeDefined();
    expect(otraClave).toBeDefined();
    // Distinct placeholders for distinct keys.
    expect(miClave?.id).not.toBe(otraClave?.id);

    const miClaveEdges = r.edges.filter(
      (e) =>
        e.kind === 'references' &&
        e.metadata?.synthesizedBy === 'vba-tempvar' &&
        e.target === miClave?.id &&
        e.source === init?.id,
    );
    expect(miClaveEdges).toHaveLength(1);
    expect(miClaveEdges[0]?.metadata?.access).toBe('write');

    const otraClaveEdges = r.edges.filter(
      (e) =>
        e.kind === 'references' &&
        e.metadata?.synthesizedBy === 'vba-tempvar' &&
        e.target === otraClave?.id &&
        e.source === init?.id,
    );
    expect(otraClaveEdges).toHaveLength(1);
    expect(otraClaveEdges[0]?.metadata?.access).toBe('write');
  });

  it('cross-file dedup: write in Form_A.cls + read in Form_B.cls collapses to ONE placeholder id and TWO references edges', () => {
    // Two separate extraction calls — the placeholder id must be
    // deterministic across `extract()` invocations so the cross-form
    // case (write in producer, read in consumer) collapses to ONE node
    // when both files reach the same index.
    const fileASrc = [
      'Public Sub SetID()',
      '    TempVars!IDExpediente = 42',
      'End Sub',
    ].join('\n');
    const fileBSrc = [
      'Public Sub Form_Load()',
      '    Me.txtId = TempVars("IDExpediente")',
      'End Sub',
    ].join('\n');

    const aResult = extract('src/forms/FormA.cls', fileASrc);
    const bResult = extract('src/forms/FormB.cls', fileBSrc);

    // The placeholder id should be byte-identical for the same key,
    // regardless of which file emits it.
    const idExpected =
      generateNodeId('synthetic:tempvar/IDExpediente', 'class', 'IDExpediente', 0);

    const placeholderInA = aResult.nodes.find(
      (n) => n.name === 'IDExpediente' && n.kind === 'class',
    );
    const placeholderInB = bResult.nodes.find(
      (n) => n.name === 'IDExpediente' && n.kind === 'class',
    );
    expect(placeholderInA).toBeDefined();
    expect(placeholderInB).toBeDefined();
    expect(placeholderInA?.id).toBe(idExpected);
    expect(placeholderInB?.id).toBe(idExpected);
    expect(placeholderInA?.id).toBe(placeholderInB?.id);

    // Two edges: one write (from `SetID`) + one read (from `Form_Load`).
    const aEdge = aResult.edges.find(
      (e) =>
        e.kind === 'references' &&
        e.metadata?.synthesizedBy === 'vba-tempvar' &&
        e.target === idExpected,
    );
    const bEdge = bResult.edges.find(
      (e) =>
        e.kind === 'references' &&
        e.metadata?.synthesizedBy === 'vba-tempvar' &&
        e.target === idExpected,
    );
    expect(aEdge).toBeDefined();
    expect(bEdge).toBeDefined();
    expect(aEdge?.metadata?.access).toBe('write');
    expect(bEdge?.metadata?.access).toBe('read');
    expect(aEdge?.source).not.toBe(bEdge?.source);
  });

  it('dynamic key (variable arg) stays silent — no edges, no placeholder', () => {
    // REQ-CODE-4 spirit: unresolvable is silent. `TempVars(strNombre)` is
    // unresolvable at static-analysis time (string is a runtime value),
    // so we emit nothing. Different from the regexes deliberately
    // matching ONLY literal `"..."` arg forms.
    const src = [
      'Public Sub DynamicCaller(strNombre As String)',
      '    TempVars(strNombre) = 1',
      '    Dim v As Variant',
      '    v = TempVars(strNombre)',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);

    const tempVarEdges = r.edges.filter(
      (e) => e.metadata?.synthesizedBy === 'vba-tempvar',
    );
    expect(tempVarEdges).toHaveLength(0);
    const placeholders = r.nodes.filter(
      (n) => n.kind === 'class' && n.filePath.startsWith('synthetic:tempvar/'),
    );
    expect(placeholders).toHaveLength(0);
  });

  it('dynamic key (string concatenation) stays silent — no edges, no placeholder', () => {
    // Same as the variable-arg case but for `TempVars("clave" & suffix)`.
    // `TEMP_VAR_PAREN_RE` tolerates only the literal-arg shape
    // (no `&` mid-arg), so the concatenation form stays silent.
    const src = [
      'Public Sub ConcatCaller()',
      '    TempVars("prefijo" & "_x") = 1',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);
    const tempVarEdges = r.edges.filter(
      (e) => e.metadata?.synthesizedBy === 'vba-tempvar',
    );
    expect(tempVarEdges).toHaveLength(0);
  });

  it('cross-cutting smoke: one proc with two writes + one read produces 2 placeholders + 3 edges from the same proc', () => {
    const src = [
      'Public Sub Multi()',
      '    TempVars("a") = 1',
      '    TempVars("b") = 2',
      '    Dim x As Variant',
      '    x = TempVars("a")',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/Mod.bas', src);

    const multi = r.nodes.find(
      (n) => n.kind === 'function' && n.name === 'Multi',
    );
    expect(multi).toBeDefined();

    const a = r.nodes.find((n) => n.name === 'a' && n.kind === 'class');
    const b = r.nodes.find((n) => n.name === 'b' && n.kind === 'class');
    expect(a).toBeDefined();
    expect(b).toBeDefined();

    const edges = r.edges.filter(
      (e) =>
        e.kind === 'references' &&
        e.metadata?.synthesizedBy === 'vba-tempvar' &&
        e.source === multi?.id,
    );
    expect(edges).toHaveLength(3);

    const aWriteEdges = edges.filter((e) => e.target === a?.id && e.metadata?.access === 'write');
    const aReadEdges = edges.filter((e) => e.target === a?.id && e.metadata?.access === 'read');
    const bWriteEdges = edges.filter((e) => e.target === b?.id && e.metadata?.access === 'write');
    const bReadEdges = edges.filter((e) => e.target === b?.id && e.metadata?.access === 'read');

    expect(aWriteEdges).toHaveLength(1);
    expect(aReadEdges).toHaveLength(1);
    expect(bWriteEdges).toHaveLength(1);
    expect(bReadEdges).toHaveLength(0);
  });
});

