/**
 * Issue #254 (task T12 of `docs/vba-node-discovery-plan.md`) — the remaining
 * `DoCmd` verbs that name an Access object.
 *
 * `OpenForm` / `OpenReport` / `OpenQuery` were already modelled. This suite
 * covers the rest of the verbs that carry an object name in a positional
 * argument, all driven by ONE table (verb -> argument position -> reference
 * kind) rather than one bespoke regex per verb.
 *
 * The contract every verb shares:
 *   - a string-literal (or Const-resolved) argument emits exactly ONE
 *     `UnresolvedReference` carrying the verb's `synthesizedBy` tag;
 *   - a dynamic argument (a plain variable, an expression) emits NOTHING —
 *     silence beats a fabricated target;
 *   - NO node is ever synthesized. The referenced artifact either exists in
 *     the indexed source tree (a form, report, query) and the reference
 *     resolves, or it does not (a macro, a table) and the reference stays
 *     visible as an actionable `failed` reference.
 */
import { describe, it, expect } from 'vitest';
import { VbaExtractor } from '../src/extraction/vba-extractor';

function extract(filePath: string, source: string) {
  return new VbaExtractor(filePath, source).extract();
}

/** Wrap one statement in a procedure so the call site has a caller. */
function inSub(statement: string): string {
  return ['Public Sub Run()', `  ${statement}`, 'End Sub'].join('\n');
}

function refsBy(result: ReturnType<typeof extract>, synthesizedBy: string) {
  return result.unresolvedReferences.filter(
    (u) =>
      (u.metadata as Record<string, unknown> | undefined)?.synthesizedBy ===
      synthesizedBy,
  );
}

function synthesizedByOf(meta: unknown): string {
  return String((meta as Record<string, unknown> | undefined)?.synthesizedBy ?? '');
}

/**
 * One row per verb covered by the table-driven dispatch. `literal` is the
 * emitting form, `dynamic` the same call with the object argument replaced
 * by a variable, and `expected` the object name the dispatch must pick out
 * of the argument list.
 */
const VERB_FIXTURES: ReadonlyArray<{
  verb: string;
  synthesizedBy: string;
  literal: string;
  dynamic: string;
  expected: string;
}> = [
  {
    verb: 'RunMacro',
    synthesizedBy: 'vba-runs-macro',
    literal: 'DoCmd.RunMacro "MacroDiaria"',
    dynamic: 'DoCmd.RunMacro strMacro',
    expected: 'MacroDiaria',
  },
  {
    verb: 'OpenTable',
    synthesizedBy: 'vba-opens-table',
    literal: 'DoCmd.OpenTable "tblClientes", acViewNormal, acReadOnly',
    dynamic: 'DoCmd.OpenTable strTabla, acViewNormal, acReadOnly',
    expected: 'tblClientes',
  },
  {
    verb: 'ApplyFilter',
    synthesizedBy: 'vba-applies-filter',
    literal: 'DoCmd.ApplyFilter "fltActivos"',
    dynamic: 'DoCmd.ApplyFilter strFiltro',
    expected: 'fltActivos',
  },
  {
    verb: 'CopyObject',
    synthesizedBy: 'vba-copies-object',
    literal: 'DoCmd.CopyObject , "tblCopia", acTable, "tblOrigen"',
    dynamic: 'DoCmd.CopyObject , "tblCopia", acTable, strOrigen',
    expected: 'tblOrigen',
  },
  {
    verb: 'DeleteObject',
    synthesizedBy: 'vba-deletes-object',
    literal: 'DoCmd.DeleteObject acTable, "tblTemporal"',
    dynamic: 'DoCmd.DeleteObject acTable, strTabla',
    expected: 'tblTemporal',
  },
  {
    verb: 'Rename',
    synthesizedBy: 'vba-renames-object',
    literal: 'DoCmd.Rename "tblNuevo", acTable, "tblViejo"',
    dynamic: 'DoCmd.Rename "tblNuevo", acTable, strViejo',
    expected: 'tblViejo',
  },
  {
    verb: 'SelectObject',
    synthesizedBy: 'vba-selects-object',
    literal: 'DoCmd.SelectObject acForm, "frmClientes", True',
    dynamic: 'DoCmd.SelectObject acForm, strFormulario, True',
    expected: 'frmClientes',
  },
  {
    verb: 'BrowseTo',
    synthesizedBy: 'vba-browses-to',
    literal: 'DoCmd.BrowseTo acBrowseToForm, "frmPrincipal"',
    dynamic: 'DoCmd.BrowseTo acBrowseToForm, strFormulario',
    expected: 'frmPrincipal',
  },
  {
    verb: 'OutputTo',
    synthesizedBy: 'vba-outputs-to',
    literal:
      'DoCmd.OutputTo acOutputReport, "rptVentas", acFormatPDF, "C:\\salida.pdf"',
    dynamic:
      'DoCmd.OutputTo acOutputReport, strInforme, acFormatPDF, "C:\\salida.pdf"',
    expected: 'rptVentas',
  },
  {
    verb: 'SendObject',
    synthesizedBy: 'vba-sends-object',
    literal: 'DoCmd.SendObject acSendReport, "rptMensual", acFormatPDF, "a@b.c"',
    dynamic: 'DoCmd.SendObject acSendReport, strInforme, acFormatPDF, "a@b.c"',
    expected: 'rptMensual',
  },
  {
    verb: 'TransferSpreadsheet',
    synthesizedBy: 'vba-transfers-spreadsheet',
    literal:
      'DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "tblVentas", "C:\\v.xlsx", True',
    dynamic:
      'DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, strTabla, "C:\\v.xlsx", True',
    expected: 'tblVentas',
  },
  {
    verb: 'TransferText',
    synthesizedBy: 'vba-transfers-text',
    literal:
      'DoCmd.TransferText acExportDelim, "EspecVentas", "tblVentas", "C:\\v.csv", True',
    dynamic:
      'DoCmd.TransferText acExportDelim, "EspecVentas", strTabla, "C:\\v.csv", True',
    expected: 'tblVentas',
  },
  {
    verb: 'TransferDatabase',
    synthesizedBy: 'vba-transfers-database',
    literal:
      'DoCmd.TransferDatabase acImport, "Microsoft Access", "C:\\ext.accdb", acTable, "tblRemoto", "tblLocal", False',
    dynamic:
      'DoCmd.TransferDatabase acImport, "Microsoft Access", "C:\\ext.accdb", acTable, strRemoto, "tblLocal", False',
    expected: 'tblRemoto',
  },
];

describe('VbaExtractor — table-driven DoCmd object verbs (issue #254)', () => {
  for (const fixture of VERB_FIXTURES) {
    it(`DoCmd.${fixture.verb} with a literal object argument emits one ${fixture.synthesizedBy} reference`, () => {
      const r = extract('src/modules/modDoCmd.bas', inSub(fixture.literal));
      const refs = refsBy(r, fixture.synthesizedBy);
      expect(refs).toHaveLength(1);
      expect(refs[0]!.referenceName).toBe(fixture.expected);
      expect(refs[0]!.referenceKind).toBe('references');
      const meta = refs[0]!.metadata as Record<string, unknown>;
      expect(meta.docmdVerb).toBe(fixture.verb);
      expect(meta.targetName).toBe(fixture.expected);
    });

    it(`DoCmd.${fixture.verb} with a dynamic object argument stays silent`, () => {
      const r = extract('src/modules/modDoCmd.bas', inSub(fixture.dynamic));
      expect(refsBy(r, fixture.synthesizedBy)).toHaveLength(0);
    });

    it(`DoCmd.${fixture.verb} never synthesizes a node for its target`, () => {
      const r = extract('src/modules/modDoCmd.bas', inSub(fixture.literal));
      expect(r.nodes.filter((n) => n.name === fixture.expected)).toHaveLength(0);
    });
  }
});

describe('VbaExtractor — DoCmd.RunMacro is a reference, never a node (issue #254)', () => {
  it('emits a vba-runs-macro unresolved reference and ZERO extra nodes', () => {
    const src = [
      'Public Sub Lanzar()',
      '  DoCmd.RunMacro "MacroDiaria"',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/modMacros.bas', src);

    const refs = refsBy(r, 'vba-runs-macro');
    expect(refs).toHaveLength(1);
    expect(refs[0]!.referenceName).toBe('MacroDiaria');
    expect(refs[0]!.referenceKind).toBe('references');

    // The only nodes for this file are the file itself, its module and its
    // single Sub — no `macro` node kind exists and no stub may stand in for
    // one.
    const kinds = r.nodes.map((n) => n.kind).sort();
    expect(kinds).toEqual(['file', 'function', 'module']);
    expect(r.nodes.some((n) => n.name === 'MacroDiaria')).toBe(false);
  });

  it('resolves a Const-held macro name to its literal value', () => {
    const src = [
      'Public Sub Lanzar()',
      '  Const MACRO_NOMBRE As String = "MacroNocturna"',
      '  DoCmd.RunMacro MACRO_NOMBRE',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/modMacros.bas', src);
    const refs = refsBy(r, 'vba-runs-macro');
    expect(refs).toHaveLength(1);
    expect(refs[0]!.referenceName).toBe('MacroNocturna');
  });

  it('two calls naming the same macro emit two references and still no node', () => {
    const src = [
      'Public Sub Lanzar()',
      '  DoCmd.RunMacro "MacroDiaria"',
      '  DoCmd.RunMacro "MacroDiaria"',
      'End Sub',
    ].join('\n');
    const r = extract('src/modules/modMacros.bas', src);
    expect(refsBy(r, 'vba-runs-macro')).toHaveLength(2);
    expect(r.nodes.some((n) => n.name === 'MacroDiaria')).toBe(false);
  });
});

describe('VbaExtractor — DoCmd object verbs: shape guards (issue #254)', () => {
  it('a verb name inside a string literal is not a call site', () => {
    const r = extract(
      'src/modules/modDoCmd.bas',
      inSub('Debug.Print "DoCmd.RunMacro ""MacroDiaria"""'),
    );
    expect(refsBy(r, 'vba-runs-macro')).toHaveLength(0);
  });

  it('an unresolvable Const-looking identifier stays silent (no bare-name fallback)', () => {
    const r = extract(
      'src/modules/modDoCmd.bas',
      inSub('DoCmd.OpenTable NOMBRE_TABLA'),
    );
    expect(refsBy(r, 'vba-opens-table')).toHaveLength(0);
  });

  it('an omitted object argument stays silent', () => {
    const r = extract(
      'src/modules/modDoCmd.bas',
      inSub('DoCmd.DeleteObject acTable'),
    );
    expect(refsBy(r, 'vba-deletes-object')).toHaveLength(0);
  });

  it('a doubled-quote escape inside the object name is decoded', () => {
    const r = extract(
      'src/modules/modDoCmd.bas',
      inSub('DoCmd.OpenTable "tbl""raro"""'),
    );
    const refs = refsBy(r, 'vba-opens-table');
    expect(refs).toHaveLength(1);
    expect(refs[0]!.referenceName).toBe('tbl"raro"');
  });

  it('the parenthesised call form is picked up too', () => {
    const r = extract(
      'src/modules/modDoCmd.bas',
      inSub('DoCmd.OpenTable("tblClientes")'),
    );
    const refs = refsBy(r, 'vba-opens-table');
    expect(refs).toHaveLength(1);
    expect(refs[0]!.referenceName).toBe('tblClientes');
  });
});

describe('VbaExtractor — verbs deliberately OUTSIDE the object dispatch (issue #254)', () => {
  it('DoCmd.RunSQL keeps its existing SQL-table modelling and gains no object reference', () => {
    // `RunSQL` names a SQL statement, not an object. The SQL-wrapper sweep
    // already turns it into `vba-sql-table` references; a second emission
    // here would double-count every table.
    const r = extract(
      'src/modules/modDoCmd.bas',
      inSub('DoCmd.RunSQL "DELETE FROM tblTemporal"'),
    );
    const tableEdges = r.edges.filter(
      (e) => synthesizedByOf(e.metadata) === 'vba-sql-table',
    );
    expect(tableEdges).toHaveLength(1);
    const objectRefs = r.unresolvedReferences.filter((u) =>
      synthesizedByOf(u.metadata).startsWith('vba-runs-sql'),
    );
    expect(objectRefs).toHaveLength(0);
  });

  it('DoCmd.RunCommand names an intrinsic command constant, so nothing is emitted', () => {
    const r = extract(
      'src/modules/modDoCmd.bas',
      inSub('DoCmd.RunCommand acCmdSaveRecord'),
    );
    const docmdRefs = r.unresolvedReferences.filter((u) =>
      synthesizedByOf(u.metadata).startsWith('vba-'),
    );
    expect(docmdRefs).toHaveLength(0);
  });
});

describe('VbaExtractor — the pre-existing Open verbs are unchanged (issue #254)', () => {
  it('DoCmd.OpenForm still emits exactly one opens-form edge and one stub', () => {
    const r = extract(
      'src/modules/modDoCmd.bas',
      inSub('DoCmd.OpenForm "frmClientes"'),
    );
    expect(r.edges.filter((e) => e.kind === 'opens-form')).toHaveLength(1);
    expect(r.nodes.filter((n) => n.kind === 'form-layout')).toHaveLength(1);
  });

  it('DoCmd.OpenReport still emits exactly one opens-report edge and one stub', () => {
    const r = extract(
      'src/modules/modDoCmd.bas',
      inSub('DoCmd.OpenReport "rptVentas"'),
    );
    expect(r.edges.filter((e) => e.kind === 'opens-report')).toHaveLength(1);
    expect(r.nodes.filter((n) => n.kind === 'report-layout')).toHaveLength(1);
  });

  it('DoCmd.OpenQuery still emits exactly one dao-query reference', () => {
    const r = extract(
      'src/modules/modDoCmd.bas',
      inSub('DoCmd.OpenQuery "qryVentas"'),
    );
    const refs = refsBy(r, 'vba-opens-query');
    expect(refs).toHaveLength(1);
    expect(refs[0]!.referenceKind).toBe('dao-query');
  });
});
