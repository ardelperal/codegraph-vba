/**
 * Issue #255 (task T13 of `docs/vba-node-discovery-plan.md`) — the Access
 * domain aggregate functions name a table or a saved query in argument 2.
 *
 * `DLookup`, `DCount`, `DSum`, `DMax`, `DMin`, `DAvg`, `DFirst` and `DLast`
 * all read their DOMAIN from the second argument. None of them was modelled,
 * so a procedure that reads a table only through `DLookup` looked like a
 * procedure that touches no data at all.
 *
 * The contract, which is deliberately T11's contract reused rather than a
 * second one:
 *   - a second argument that is EXACTLY one string literal and parses as a
 *     bare Access object name emits ONE `dao-query` reference tagged
 *     `vba-query-name`, sourced from the CALLING PROCEDURE;
 *   - it resolves to the `query` node when a saved query carries that name,
 *     and declines otherwise — never a fabricated table placeholder;
 *   - a variable or an expression in argument 2 is skipped in SILENCE;
 *   - a project that declares its own `DLookup` keeps resolving the call to
 *     that user code and gains no invented domain reference.
 */
import { describe, it, expect, afterEach } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import { VbaExtractor } from '../src/extraction/vba-extractor';
import { CodeGraph } from '../src';

function extract(filePath: string, source: string) {
  return new VbaExtractor(filePath, source).extract();
}

/** Wrap statements in a procedure so the call site has a caller. */
function inSub(...statements: string[]): string {
  return ['Public Sub Run()', ...statements.map((s) => `  ${s}`), 'End Sub'].join('\n');
}

function domainRefs(result: ReturnType<typeof extract>) {
  return result.unresolvedReferences.filter(
    (u) =>
      (u.metadata as Record<string, unknown> | undefined)?.synthesizedBy ===
      'vba-query-name',
  );
}

function domainNames(result: ReturnType<typeof extract>): string[] {
  return domainRefs(result).map((r) => r.referenceName);
}

describe('issue #255 — the domain of a domain-aggregate call is a reference', () => {
  it('DLookup names its table in argument 2', () => {
    const result = extract(
      'modDatos.bas',
      inSub('x = DLookup("Nombre", "TbUsuarios", "Id=1")'),
    );
    const refs = domainRefs(result);
    expect(refs).toHaveLength(1);
    expect(refs[0]!.referenceName).toBe('TbUsuarios');
    expect(refs[0]!.referenceKind).toBe('dao-query');
  });

  it('DCount names its saved query in argument 2', () => {
    const result = extract('modDatos.bas', inSub('n = DCount("*", "qryPendientes")'));
    expect(domainNames(result)).toEqual(['qryPendientes']);
  });

  it.each([
    ['DLookup', 'x = DLookup("Nombre", "TbUsuarios")'],
    ['DCount', 'x = DCount("*", "TbUsuarios")'],
    ['DSum', 'x = DSum("Importe", "TbUsuarios")'],
    ['DMax', 'x = DMax("Importe", "TbUsuarios")'],
    ['DMin', 'x = DMin("Importe", "TbUsuarios")'],
    ['DAvg', 'x = DAvg("Importe", "TbUsuarios")'],
    ['DFirst', 'x = DFirst("Nombre", "TbUsuarios")'],
    ['DLast', 'x = DLast("Nombre", "TbUsuarios")'],
  ])('%s is covered', (_name, statement) => {
    expect(domainNames(extract('modDatos.bas', inSub(statement)))).toEqual([
      'TbUsuarios',
    ]);
  });

  it('sources the reference from the calling procedure, not the module', () => {
    const result = extract(
      'modDatos.bas',
      inSub('x = DLookup("Nombre", "TbUsuarios", "Id=1")'),
    );
    const proc = result.nodes.find((n) => n.kind === 'function' && n.name === 'Run');
    expect(proc).toBeDefined();
    expect(domainRefs(result)[0]!.fromNodeId).toBe(proc!.id);
  });

  it('emits NO node for the domain name', () => {
    const result = extract(
      'modDatos.bas',
      inSub('x = DLookup("Nombre", "TbUsuarios", "Id=1")'),
    );
    expect(result.nodes.some((n) => n.name === 'TbUsuarios')).toBe(false);
  });

  it('matches case-insensitively, as VBA does', () => {
    const result = extract('modDatos.bas', inSub('x = dlookup("Nombre", "TbUsuarios")'));
    expect(domainNames(result)).toEqual(['TbUsuarios']);
  });

  it('accepts an Access name containing spaces', () => {
    const result = extract(
      'modDatos.bas',
      inSub('n = DCount("*", "Consulta De Pendientes")'),
    );
    expect(domainNames(result)).toEqual(['Consulta De Pendientes']);
  });
});

describe('issue #255 — what the scan deliberately stays silent about', () => {
  it('a variable domain emits nothing', () => {
    const result = extract('modDatos.bas', inSub('x = DLookup("Nombre", tabla)'));
    expect(domainRefs(result)).toHaveLength(0);
  });

  it('a concatenated domain emits nothing', () => {
    const result = extract(
      'modDatos.bas',
      inSub('x = DLookup("Nombre", "Tb" & sufijo, "Id=1")'),
    );
    expect(domainRefs(result)).toHaveLength(0);
  });

  it('a one-argument call emits nothing', () => {
    const result = extract('modDatos.bas', inSub('n = DCount("*")'));
    expect(domainRefs(result)).toHaveLength(0);
  });

  it('a domain spelled as a SQL statement emits nothing', () => {
    const result = extract(
      'modDatos.bas',
      inSub('x = DLookup("Nombre", "SELECT Id FROM TbUsuarios")'),
    );
    expect(domainRefs(result)).toHaveLength(0);
  });

  it('a punctuated literal is neither a query nor a table', () => {
    const result = extract('modDatos.bas', inSub('x = DLookup("Nombre", "no-es!")'));
    expect(domainRefs(result)).toHaveLength(0);
  });

  it('a longer helper name is not read as a domain function', () => {
    const result = extract(
      'modDatos.bas',
      inSub('x = DLookupSeguro("Nombre", "TbUsuarios", "Id=1")'),
    );
    expect(domainRefs(result)).toHaveLength(0);
  });
});

describe('issue #255 — argument 2 is found by walking, not by splitting on commas', () => {
  it('a comma inside the criteria literal does not shift the domain', () => {
    const result = extract(
      'modDatos.bas',
      inSub('x = DLookup("Nombre", "TbUsuarios", "Id In (1,2)")'),
    );
    expect(domainNames(result)).toEqual(['TbUsuarios']);
  });

  it('a comma inside a nested call does not shift the domain', () => {
    const result = extract(
      'modDatos.bas',
      inSub('x = DLookup("Nombre", "TbUsuarios", "Id=" & Nz(pId, 0))'),
    );
    expect(domainNames(result)).toEqual(['TbUsuarios']);
  });

  it('a comma inside the FIRST argument does not shift the domain', () => {
    const result = extract(
      'modDatos.bas',
      inSub('x = DLookup("IIf(a, b, c)", "TbUsuarios")'),
    );
    expect(domainNames(result)).toEqual(['TbUsuarios']);
  });

  it('two calls on one line each name their own domain', () => {
    const result = extract(
      'modDatos.bas',
      inSub('x = DLookup("Nombre", "TbUsuarios") & DCount("*", "qryPendientes")'),
    );
    expect(domainNames(result).sort()).toEqual(['TbUsuarios', 'qryPendientes']);
  });

  it('survives a `_` line continuation, which is joined before scanning', () => {
    const result = extract(
      'modDatos.bas',
      ['Public Sub Run()', '  x = DLookup("Nombre", _', '    "TbUsuarios")', 'End Sub'].join(
        '\n',
      ),
    );
    expect(domainNames(result)).toEqual(['TbUsuarios']);
  });
});

describe('issue #255 — a user-defined DLookup still resolves to user code', () => {
  const source = [
    'Public Sub Run()',
    '  x = DLookup("Nombre", "TbUsuarios", "Id=1")',
    'End Sub',
    '',
    'Public Function DLookup(campo As String, dominio As String, criterio As String) As Variant',
    '  DLookup = Null',
    'End Function',
    '',
  ].join('\n');

  it('emits no domain reference when the file declares its own DLookup', () => {
    expect(domainRefs(extract('modDatos.bas', source))).toHaveLength(0);
  });

  it('keeps the call edge to the local function', () => {
    const result = extract('modDatos.bas', source);
    const caller = result.nodes.find((n) => n.kind === 'function' && n.name === 'Run');
    const callee = result.nodes.find(
      (n) => n.kind === 'function' && n.name === 'DLookup',
    );
    expect(caller).toBeDefined();
    expect(callee).toBeDefined();
    const call = result.edges.find(
      (e) => e.kind === 'calls' && e.source === caller!.id && e.target === callee!.id,
    );
    expect(call).toBeDefined();
  });
});

describe('issue #255 — resolution reuses T11 and never invents a target', () => {
  let tmpDir: string | undefined;

  afterEach(() => {
    if (tmpDir && fs.existsSync(tmpDir)) {
      // Windows holds the SQLite file open a moment after `close()`; a failed
      // temp-dir cleanup must not fail an otherwise-passing assertion.
      try {
        fs.rmSync(tmpDir, { recursive: true, force: true });
      } catch {
        /* best effort — see the Windows notes in CLAUDE.md */
      }
    }
    tmpDir = undefined;
  });

  function writeProject(moduleBody: string, withQuery: boolean): string {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'cg-vba-domain-fn-'));
    fs.mkdirSync(path.join(dir, 'src/modules'), { recursive: true });
    if (withQuery) {
      fs.mkdirSync(path.join(dir, 'src/queries'), { recursive: true });
      fs.writeFileSync(
        path.join(dir, 'src/queries/qryPendientes.sql'),
        'SELECT Id, Nombre FROM TbExpedientes WHERE Estado = 1;\n',
      );
      // A `.sql` only reaches extraction when its directory also holds a
      // `queries.json` — the Dysflow saved-query gate.
      fs.writeFileSync(
        path.join(dir, 'src/queries/queries.json'),
        JSON.stringify({ queries: [{ name: 'qryPendientes' }] }, null, 2),
      );
    }
    fs.writeFileSync(
      path.join(dir, 'src/modules/modDatos.bas'),
      [
        'Attribute VB_Name = "modDatos"',
        'Public Sub Contar()',
        `  ${moduleBody}`,
        'End Sub',
        '',
      ].join('\n'),
    );
    return dir;
  }

  it('DCount resolves to the query node when the saved query exists', async () => {
    tmpDir = writeProject('n = DCount("*", "qryPendientes")', true);
    const cg = CodeGraph.initSync(tmpDir);
    await cg.indexAll();

    const proc = cg.getNodesByKind('function').find((n) => n.name === 'Contar');
    const query = cg.getNodesByKind('query').find((n) => n.name === 'qryPendientes');
    expect(proc).toBeDefined();
    expect(query).toBeDefined();

    const hop = cg.getOutgoingEdges(proc!.id).find((e) => e.target === query!.id);
    expect(hop).toBeDefined();
    expect(hop!.kind).toBe('references');

    cg.close();
  });

  it('DLookup on a table name invents no node and no edge', async () => {
    tmpDir = writeProject('x = DLookup("Nombre", "TbUsuarios", "Id=1")', false);
    const cg = CodeGraph.initSync(tmpDir);
    await cg.indexAll();

    expect(cg.getNodesByKind('query').filter((n) => n.name === 'TbUsuarios')).toEqual([]);
    expect(cg.getNodesByKind('class').filter((n) => n.name === 'TbUsuarios')).toEqual([]);

    cg.close();
  });
});
