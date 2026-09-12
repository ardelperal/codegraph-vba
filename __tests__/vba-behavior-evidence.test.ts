/**
 * Issue #299 — bounded, typed behavior evidence for Access controls.
 *
 * Everything here runs against the shared corpus indexed through PRODUCTION
 * extraction and resolution (`helpers/vba-consumer-semantics.ts`): the payload
 * has to come from real indexed facts, not from edges a test wrote to make its
 * own assertion true.
 */
import { afterEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { CodeGraph, CodeGraphBehaviorEvidence } from '../src';
import { getStaticTools, ToolHandler } from '../src/mcp/tools';
import {
  indexConsumerSemanticsCorpus,
  IndexedCorpus,
} from './helpers/vba-consumer-semantics';

let corpus: IndexedCorpus | null = null;
const scratch: Array<() => Promise<void> | void> = [];
const savedToolsEnv = process.env.CODEGRAPH_MCP_TOOLS;

afterEach(async () => {
  if (corpus) {
    await corpus.dispose();
    corpus = null;
  }
  while (scratch.length) await scratch.pop()!();
  if (savedToolsEnv === undefined) delete process.env.CODEGRAPH_MCP_TOOLS;
  else process.env.CODEGRAPH_MCP_TOOLS = savedToolsEnv;
});

/**
 * The consumer's own type, restated here on purpose. It is the contract this
 * repository promises to keep, and importing it from the consumer would make
 * Dysflow a runtime dependency of these tests.
 */
type DysflowBehaviorEvidence = {
  handler: string;
  callPath: string[];
  tables?: string[];
  effects?: string[];
};

function assertDysflowShape(entries: CodeGraphBehaviorEvidence[]): void {
  for (const entry of entries) {
    const consumed: DysflowBehaviorEvidence = entry;
    expect(typeof consumed.handler).toBe('string');
    expect(Array.isArray(consumed.callPath)).toBe(true);
    for (const step of consumed.callPath) expect(typeof step).toBe('string');
    if (consumed.tables !== undefined) {
      expect(Array.isArray(consumed.tables)).toBe(true);
      for (const table of consumed.tables) expect(typeof table).toBe('string');
    }
    if (consumed.effects !== undefined) {
      expect(Array.isArray(consumed.effects)).toBe(true);
      for (const effect of consumed.effects) expect(typeof effect).toBe('string');
    }
    expect(Object.keys(entry).sort()).toEqual(['callPath', 'effects', 'handler', 'tables']);
  }
}

describe('issue #299 — behavior evidence', () => {
  it('behavior evidence matches indexed canonical source', async () => {
    corpus = await indexConsumerSemanticsCorpus();

    const result = corpus.cg.getBehaviorEvidence({
      name: 'btnSave',
      layout: 'Form_Orders',
    });

    expect(result.target).toMatchObject({
      name: 'btnSave',
      kind: 'form-instance-control',
      filePath: 'Form_Orders.form.txt',
      layout: 'Form_Orders',
    });
    expect(result.evidence).toEqual([
      {
        handler: 'btnSave_Click',
        callPath: ['btnSave_Click', 'SaveOrderTotals'],
        tables: ['tblOrderLines', 'tblProducts'],
        effects: [
          'data-access:qryOrderTotals',
          'data-access:tblOrderLines',
          'data-access:tblProducts',
        ],
      },
    ]);
    expect(result.context.handlers).toEqual([
      expect.objectContaining({
        handler: 'btnSave_Click',
        event: 'Click',
        scope: 'control',
        location: 'Form_Orders.cls:14',
      }),
    ]);
    // The tables are reached THROUGH the saved query — the context says so
    // instead of presenting them as something the procedure names directly.
    expect(result.context.data).toContainEqual(
      expect.objectContaining({
        procedure: 'SaveOrderTotals',
        name: 'tblOrderLines',
        throughQuery: 'qryOrderTotals',
        attributedBy: 'edge-source',
      }),
    );
    expect(result.context.notes).toContain('STATIC_SOURCE_EVIDENCE');
  });

  it('payload matches Dysflow compatibility contract', async () => {
    corpus = await indexConsumerSemanticsCorpus();
    for (const request of [
      { name: 'btnSave', layout: 'Form_Orders' },
      { name: 'btnPurge', layout: 'Form_Admin' },
      { name: 'Report_Sales' },
    ]) {
      assertDysflowShape(corpus.cg.getBehaviorEvidence(request).evidence);
    }
  });

  it('ambiguous controls never select arbitrary handlers', async () => {
    corpus = await indexConsumerSemanticsCorpus();

    const ambiguous = corpus.cg.getBehaviorEvidence({ name: 'btnSave' });

    expect(ambiguous.target).toBeNull();
    expect(ambiguous.evidence).toEqual([]);
    expect(ambiguous.context.notes).toContain('AMBIGUOUS_TARGET');
    expect(ambiguous.context.ambiguous[0]!.matches.map((m) => m.layout).sort()).toEqual([
      'Form_Invoices',
      'Form_Orders',
    ]);

    // The same name, scoped, answers for exactly one form.
    const invoices = corpus.cg.getBehaviorEvidence({
      name: 'btnSave',
      layout: 'Form_Invoices',
    });
    expect(invoices.evidence.map((e) => e.callPath)).toEqual([
      ['btnSave_Click', 'SaveInvoiceTotals'],
    ]);
    expect(JSON.stringify(invoices.evidence)).not.toContain('SaveOrderTotals');
  });

  it('reports read and write evidence separately', async () => {
    corpus = await indexConsumerSemanticsCorpus();

    const purge = corpus.cg.getBehaviorEvidence({
      name: 'btnPurge',
      layout: 'Form_Admin',
    });

    expect(purge.evidence).toEqual([
      {
        handler: 'btnPurge_Click',
        callPath: ['btnPurge_Click', 'PurgeOrderLines'],
        tables: ['tblAuditLog'],
        effects: ['opens-form:Form_Invoices', 'write:tblAuditLog'],
      },
    ]);
    expect(purge.context.data).toEqual([
      expect.objectContaining({
        procedure: 'PurgeOrderLines',
        name: 'tblAuditLog',
        access: 'write',
        via: 'vba-sql-table',
        // The reference edge is emitted from the module node by design, so it
        // is attributed to the procedure by its source line.
        attributedBy: 'source-line',
      }),
    ]);
  });

  it('does not report a VBA type reference as a table', async () => {
    corpus = await indexConsumerSemanticsCorpus();

    const evidence = corpus.cg.getBehaviorEvidence({
      name: 'btnSave',
      layout: 'Form_Orders',
    });

    // `Dim rs As DAO.Recordset` is a type reference, not data access.
    expect(evidence.evidence[0]!.tables).not.toContain('DAO');
    expect(evidence.context.data.map((d) => d.name)).not.toContain('DAO');
  });

  it('unresolved effects remain unknown', async () => {
    corpus = await indexConsumerSemanticsCorpus();

    const exported = corpus.cg.getBehaviorEvidence({
      name: 'btnExport',
      layout: 'Form_Orders',
    });

    // The handler exists and is bound; what it goes on to do cannot be
    // answered statically. Empty lists, and the reason reported explicitly.
    expect(exported.evidence).toEqual([
      {
        handler: 'btnExport_Click',
        callPath: ['btnExport_Click'],
        tables: [],
        effects: [],
      },
    ]);
    expect(exported.context.unresolved).toContainEqual(
      expect.objectContaining({
        procedure: 'btnExport_Click',
        referenceName: 'ExportOrdersToExcel',
      }),
    );
  });

  it('represents expression-wired and lifecycle events explicitly', async () => {
    corpus = await indexConsumerSemanticsCorpus();

    const expression = corpus.cg.getBehaviorEvidence({
      name: 'btnAudit',
      layout: 'Form_Admin',
    });
    expect(expression.context.handlers).toEqual([
      expect.objectContaining({
        handler: 'AuditNow',
        event: 'Click',
        scope: 'expression',
        wiredBy: 'vba-expression-handler',
      }),
    ]);
    // No fabricated `btnAudit_Click` procedure name.
    expect(expression.evidence.map((e) => e.handler)).toEqual(['AuditNow']);

    const lifecycle = corpus.cg.getBehaviorEvidence({ name: 'Form_Orders' });
    expect(lifecycle.context.handlers).toEqual([
      expect.objectContaining({ handler: 'Form_Load', event: 'Load', scope: 'form' }),
    ]);

    const report = corpus.cg.getBehaviorEvidence({ name: 'Report_Sales' });
    expect(report.context.handlers).toEqual([
      expect.objectContaining({ handler: 'Report_Open', event: 'Open', scope: 'form' }),
    ]);
  });

  it('reports a control with no bound handler instead of inventing one', async () => {
    corpus = await indexConsumerSemanticsCorpus();

    const unbound = corpus.cg.getBehaviorEvidence({
      name: 'cboCustomer',
      layout: 'Form_Orders',
    });

    expect(unbound.target).not.toBeNull();
    expect(unbound.evidence).toEqual([]);
    expect(unbound.context.notes).toContain('NO_HANDLER_BOUND');
  });

  it('reports a missing target and a missing selector without guessing', async () => {
    corpus = await indexConsumerSemanticsCorpus();

    const missing = corpus.cg.getBehaviorEvidence({ name: 'btnNothingHere' });
    expect(missing.target).toBeNull();
    expect(missing.context.notes).toContain('TARGET_NOT_FOUND');

    const noSelector = corpus.cg.getBehaviorEvidence({});
    expect(noSelector.target).toBeNull();
    expect(noSelector.context.notes).toContain('MISSING_TARGET_SELECTOR');
  });

  it('keeps distinct branches apart instead of flattening them', async () => {
    corpus = await indexConsumerSemanticsCorpus();
    const detalle = corpus.cg.getBehaviorEvidence({
      name: 'Detalle',
      layout: 'Report_Sales',
    });

    // Deterministic and deduplicated: re-running returns the identical payload.
    const again = corpus.cg.getBehaviorEvidence({
      name: 'Detalle',
      layout: 'Report_Sales',
    });
    expect(again).toEqual(detalle);
    expect(detalle.evidence).toEqual([
      {
        handler: 'Detalle_Format',
        callPath: ['Detalle_Format', 'FormatSalesRow'],
        tables: [],
        effects: [],
      },
    ]);
  });

  it('bounded cyclic evidence reports truncation', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'vba-behavior-cycle-'));
    fs.writeFileSync(
      path.join(dir, 'Form_Loop.form.txt'),
      `Version =21
Begin Form
    Begin CommandButton
        Name ="btnGo"
    End
End
`,
    );
    fs.writeFileSync(
      path.join(dir, 'Form_Loop.cls'),
      `Attribute VB_Name = "Form_Loop"
Private Sub btnGo_Click()
    Call Ping
End Sub
`,
    );
    fs.writeFileSync(
      path.join(dir, 'modLoop.bas'),
      `Attribute VB_Name = "modLoop"

Public Sub Ping()
    Call Pong
End Sub

Public Sub Pong()
    Call Ping
    Call Deeper
    Call Wider
End Sub

Public Sub Deeper()
End Sub

Public Sub Wider()
End Sub
`,
    );
    const cg = await CodeGraph.init(dir, { index: false });
    scratch.push(async () => {
      await cg.close();
      fs.rmSync(dir, { recursive: true, force: true });
    });
    await cg.indexAll();

    const cyclic = cg.getBehaviorEvidence({ name: 'btnGo', layout: 'Form_Loop' });
    expect(cyclic.context.truncated.cycle).toBe(true);
    // A cycle ends the path; no name repeats inside one callPath.
    for (const entry of cyclic.evidence) {
      expect(new Set(entry.callPath).size).toBe(entry.callPath.length);
    }

    const shallow = cg.getBehaviorEvidence({
      name: 'btnGo',
      layout: 'Form_Loop',
      maxCallDepth: 2,
    });
    expect(shallow.context.truncated.callDepth).toBe(true);
    for (const entry of shallow.evidence) {
      expect(entry.callPath.length).toBeLessThanOrEqual(2);
    }

    const capped = cg.getBehaviorEvidence({
      name: 'btnGo',
      layout: 'Form_Loop',
      maxResults: 1,
    });
    expect(capped.evidence.length).toBe(1);
    expect(capped.context.truncated.results).toBe(true);
  });

  it('answers without mutating the index', async () => {
    corpus = await indexConsumerSemanticsCorpus();
    const count = () =>
      JSON.stringify([
        corpus!.db.prepare('SELECT count(*) AS n FROM nodes').get(),
        corpus!.db.prepare('SELECT count(*) AS n FROM edges').get(),
        corpus!.db.prepare('SELECT count(*) AS n FROM unresolved_refs').get(),
      ]);

    const before = count();
    corpus.cg.getBehaviorEvidence({ name: 'btnSave', layout: 'Form_Orders' });
    corpus.cg.getBehaviorEvidence({ name: 'Form_Orders' });
    expect(count()).toBe(before);
  });

  describe('MCP adapter', () => {
    it('is registered, opt-in visible, and callable', async () => {
      delete process.env.CODEGRAPH_MCP_TOOLS;
      // The default surface stays explore-only on purpose: an extra listed
      // tool steers agent mis-picks. This one is for a programmatic consumer.
      expect(
        getStaticTools().some((t) => t.name === 'codegraph_behavior_evidence'),
      ).toBe(false);

      process.env.CODEGRAPH_MCP_TOOLS = 'explore,behavior_evidence';
      const tool = getStaticTools().find(
        (t) => t.name === 'codegraph_behavior_evidence',
      );
      expect(tool).toBeDefined();
      expect(tool!.annotations).toMatchObject({ readOnlyHint: true });
      expect(Object.keys(tool!.inputSchema.properties!).sort()).toEqual([
        'layout',
        'maxCallDepth',
        'maxResults',
        'name',
        'nodeId',
        'projectPath',
      ]);
    });

    it('returns the same payload the library API returns', async () => {
      corpus = await indexConsumerSemanticsCorpus();
      const handler = new ToolHandler(corpus.cg);

      const response = await handler.execute('codegraph_behavior_evidence', {
        name: 'btnSave',
        layout: 'Form_Orders',
      });

      expect(response.isError).toBeFalsy();
      const payload = JSON.parse(response.content[0]!.text!);
      expect(payload).toEqual(
        corpus.cg.getBehaviorEvidence({ name: 'btnSave', layout: 'Form_Orders' }),
      );
    });

    it('answers a selector-less call with guidance, not an error', async () => {
      corpus = await indexConsumerSemanticsCorpus();
      const handler = new ToolHandler(corpus.cg);

      const response = await handler.execute('codegraph_behavior_evidence', {});

      expect(response.isError).toBeFalsy();
      expect(JSON.parse(response.content[0]!.text!).context.notes).toContain(
        'MISSING_TARGET_SELECTOR',
      );
    });
  });
});
