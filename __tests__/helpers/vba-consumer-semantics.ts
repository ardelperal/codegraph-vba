/**
 * Shared loader for the VBA consumer-semantics corpus
 * (`__tests__/fixtures/vba-consumer-semantics`).
 *
 * The corpus is the single fixture set behind the Access consumer work:
 * handler tracing and SQL impact regressions (#298), behavior-evidence
 * contract tests (#299), the documented runnable example (#300) and the
 * semantic acceptance corpus (#301). It is copied into a throwaway
 * directory and indexed through PRODUCTION extraction + resolution — never
 * hand-assembled — so a traversal bug cannot hide behind edges written by
 * the test itself.
 *
 * Ground truth (source-reviewed, see the fixture headers):
 *   Form_Orders.btnSave   -> btnSave_Click   -> SaveOrderTotals  -> qryOrderTotals
 *   Form_Orders (form)    -> Form_Load       -> RefreshOrders    -> tblOrderHeaders
 *   Form_Orders.btnExport -> btnExport_Click -> dynamic dispatch + a call to
 *                            ExportOrdersToExcel, which nothing declares
 *                            (deliberately unresolved)
 *   Form_Admin.btnPurge   -> btnPurge_Click  -> PurgeOrderLines  -> writes
 *                            tblAuditLog and opens Form_Invoices
 *   Form_Admin.btnAudit   -> wired by expression (=AuditNow()), reads tblAuditLog
 *   Form_Invoices.btnSave -> btnSave_Click   -> SaveInvoiceTotals-> qryInvoiceTotals
 *   Report_Sales.Detalle  -> Detalle_Format  -> FormatSalesRow
 *   Report_Sales (report) -> Report_Open     -> LoadSalesTotals  -> qrySalesTotals
 *
 * No form or report binds `qryOrderTotals` through RecordSource/RowSource, and
 * the query's tables (`tblOrderLines`, `tblProducts`) are bound by nothing:
 * file scanning alone cannot answer "which form is affected by this query".
 */
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { CodeGraph, DatabaseConnection, getDatabasePath } from '../../src';
import type { SqliteDatabase } from '../../src/db';

export const CORPUS_DIR = path.join(
  __dirname,
  '..',
  'fixtures',
  'vba-consumer-semantics',
);

export interface IndexedCorpus {
  /** Throwaway copy of the corpus that was indexed. */
  readonly dir: string;
  readonly cg: CodeGraph;
  /** Read-only connection to the indexed graph, for helper-level consumers. */
  readonly db: SqliteDatabase;
  dispose(): Promise<void>;
}

/**
 * Copy the corpus into a fresh temp directory and index it with production
 * extraction + resolution. `subset` keeps only the named top-level entries
 * (plus `queries/`) so a test can narrow the corpus without a second fixture.
 */
export async function indexConsumerSemanticsCorpus(
  subset?: readonly string[],
): Promise<IndexedCorpus> {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'vba-consumer-semantics-'));
  fs.cpSync(CORPUS_DIR, dir, { recursive: true });
  if (subset) {
    const keep = new Set<string>([...subset, 'queries']);
    for (const entry of fs.readdirSync(dir)) {
      if (!keep.has(entry)) {
        fs.rmSync(path.join(dir, entry), { recursive: true, force: true });
      }
    }
  }

  const cg = await CodeGraph.init(dir, { index: false });
  await cg.indexAll();

  const readConn = DatabaseConnection.open(getDatabasePath(dir));

  return {
    dir,
    cg,
    db: readConn.getDb(),
    async dispose() {
      readConn.close();
      await cg.close();
      fs.rmSync(dir, { recursive: true, force: true });
    },
  };
}
