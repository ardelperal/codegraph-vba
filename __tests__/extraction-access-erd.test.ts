/**
 * AccessErdExtractor — table structure from the Access ERD export (#257, half B).
 *
 * An Access application can ship a generated structure dump of its backend at
 * `ERD/Estructura_Datos.md`: one `## Tabla:` section per table, a markdown
 * field table per section, and a `(LINKED)` block naming the external backend
 * a linked table really lives in.
 *
 * Only SOME Access projects ship that file, so the whole feature is optional:
 * a project without one must index exactly as it did before — no warning, no
 * error, no empty node. And an `ERD/` folder in a non-Access repo is a very
 * ordinary thing, so the path match alone can never be enough: the extractor
 * also gates on the document's own header line.
 *
 * The header gate has one trap this suite pins deliberately: the real
 * generated file starts with a UTF-8 BOM, so a naive
 * `startsWith('# Estructura de Datos:')` on the first non-empty line fails on
 * the genuine input while passing on every hand-written fixture.
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as os from 'node:os';
import {
  AccessErdExtractor,
  ERD_TABLE_SYNTHESIZED_BY,
  ERD_LINKED_TABLE_SYNTHESIZED_BY,
} from '../src/extraction/access-erd-extractor';
import { isAccessErdFile } from '../src/extraction/grammars';
import { EXTERNAL_BACKEND_SYNTHESIZED_BY } from '../src/extraction/sql-table-scan';
import { clearProjectConfigCache } from '../src/project-config';
import CodeGraph from '../src/index';

const FIXTURE_DIR = path.join(__dirname, 'fixtures', 'access-erd', 'ERD');

function extract(filePath: string, source: string) {
  return new AccessErdExtractor(filePath, source).extract();
}

/** A minimal well-formed export, LF-terminated and BOM-free. */
const SIMPLE_ERD = [
  '# Estructura de Datos: Fixture_Datos.accdb',
  '',
  '## Tabla: TbAnexos',
  '| Campo | Tipo | Longitud |',
  '| :--- | :--- | :--- |',
  '| IDAnexo | 4 | 4 |',
  '| Titulo | 10 | 255 |',
  '',
].join('\n');

describe('isAccessErdFile — path gate', () => {
  it('matches a markdown file directly inside an ERD directory', () => {
    expect(isAccessErdFile('ERD/Estructura_Datos.md')).toBe(true);
    expect(isAccessErdFile('src/backend/ERD/Estructura_Datos.md')).toBe(true);
  });

  it('is case-insensitive on both the directory and the extension', () => {
    expect(isAccessErdFile('erd/estructura_datos.MD')).toBe(true);
  });

  it('accepts Windows-style separators', () => {
    expect(isAccessErdFile('C:\\proj\\ERD\\Estructura_Datos.md')).toBe(true);
  });

  it('rejects markdown outside an ERD directory', () => {
    expect(isAccessErdFile('docs/erd.md')).toBe(false);
    expect(isAccessErdFile('ERDiagrams/x.md')).toBe(false);
  });

  it('rejects a non-markdown file inside an ERD directory', () => {
    expect(isAccessErdFile('ERD/Estructura_Datos.txt')).toBe(false);
  });

  it('rejects a nested markdown file below the ERD directory', () => {
    expect(isAccessErdFile('ERD/notes/deep.md')).toBe(false);
  });

  it('rejects empty input', () => {
    expect(isAccessErdFile('')).toBe(false);
  });
});

describe('AccessErdExtractor — content-shape gate', () => {
  it('accepts the export header on the first non-empty line', () => {
    const r = extract('ERD/Estructura_Datos.md', SIMPLE_ERD);
    expect(r.nodes.some((n) => n.kind === 'class' && n.name === 'TbAnexos')).toBe(true);
    expect(r.errors).toHaveLength(0);
  });

  // The real generated file starts with a UTF-8 BOM. A gate that tests the raw
  // first non-empty line rejects the ONE input this feature exists for.
  it('accepts a BOM-prefixed header (the shape of the real generated file)', () => {
    const r = extract('ERD/Estructura_Datos.md', '\uFEFF' + SIMPLE_ERD);
    expect(r.nodes.some((n) => n.kind === 'class' && n.name === 'TbAnexos')).toBe(true);
    expect(r.errors).toHaveLength(0);
  });

  it('accepts CRLF line endings (the shape of the real generated file)', () => {
    const r = extract('ERD/Estructura_Datos.md', '\uFEFF' + SIMPLE_ERD.replace(/\n/g, '\r\n'));
    const fields = r.nodes.filter((n) => n.kind === 'type_member');
    expect(fields.map((n) => n.name)).toEqual(['IDAnexo', 'Titulo']);
  });

  it('emits nothing and no error for an ordinary markdown file under ERD/', () => {
    const r = extract(
      'ERD/README.md',
      ['# Entity Relationship Diagrams', '', '## Tabla: NotATable', ''].join('\n'),
    );
    expect(r.nodes).toHaveLength(0);
    expect(r.edges).toHaveLength(0);
    expect(r.errors).toHaveLength(0);
  });

  it('emits nothing for an empty file', () => {
    const r = extract('ERD/Estructura_Datos.md', '');
    expect(r.nodes).toHaveLength(0);
    expect(r.errors).toHaveLength(0);
  });

  it('rejects a header that is not on the FIRST non-empty line', () => {
    const r = extract(
      'ERD/Estructura_Datos.md',
      ['Some preamble', '', '# Estructura de Datos: X.accdb', '', '## Tabla: TbX', ''].join('\n'),
    );
    expect(r.nodes).toHaveLength(0);
  });
});

describe('AccessErdExtractor — table and field nodes', () => {
  it('emits a file node for the ERD document', () => {
    const r = extract('ERD/Estructura_Datos.md', SIMPLE_ERD);
    const file = r.nodes.find((n) => n.kind === 'file');
    expect(file).toBeDefined();
    expect(file?.name).toBe('Estructura_Datos.md');
  });

  it('emits one class node per table, tagged as an ERD declaration', () => {
    const r = extract('ERD/Estructura_Datos.md', SIMPLE_ERD);
    const table = r.nodes.find((n) => n.kind === 'class' && n.name === 'TbAnexos');
    expect(table).toBeDefined();
    expect(table?.qualifiedName).toBe('TbAnexos');
    expect(table?.language).toBe('vba');
    expect(table?.metadata?.synthesizedBy).toBe(ERD_TABLE_SYNTHESIZED_BY);
    expect(table?.metadata?.linked).toBe(false);
  });

  it('emits one type_member per field row, contained by its table', () => {
    const r = extract('ERD/Estructura_Datos.md', SIMPLE_ERD);
    const table = r.nodes.find((n) => n.kind === 'class' && n.name === 'TbAnexos')!;
    const fields = r.nodes.filter((n) => n.kind === 'type_member');
    expect(fields.map((n) => n.name)).toEqual(['IDAnexo', 'Titulo']);
    expect(fields.map((n) => n.qualifiedName)).toEqual([
      'TbAnexos.IDAnexo',
      'TbAnexos.Titulo',
    ]);
    for (const field of fields) {
      expect(
        r.edges.some(
          (e) => e.kind === 'contains' && e.source === table.id && e.target === field.id,
        ),
      ).toBe(true);
    }
  });

  it('carries the Access type, length and ordinal position on each field', () => {
    const r = extract('ERD/Estructura_Datos.md', SIMPLE_ERD);
    const titulo = r.nodes.find((n) => n.name === 'Titulo')!;
    expect(titulo.metadata).toMatchObject({ accessType: 10, length: 255, position: 2 });
    const id = r.nodes.find((n) => n.name === 'IDAnexo')!;
    expect(id.metadata).toMatchObject({ accessType: 4, length: 4, position: 1 });
  });

  it('keeps a non-numeric type or length as its raw text', () => {
    const r = extract(
      'ERD/Estructura_Datos.md',
      [
        '# Estructura de Datos: X.accdb',
        '## Tabla: TbX',
        '| Campo | Tipo | Longitud |',
        '| :--- | :--- | :--- |',
        '| Nombre | Texto | corto |',
        '',
      ].join('\n'),
    );
    const field = r.nodes.find((n) => n.kind === 'type_member')!;
    expect(field.metadata).toMatchObject({ accessType: 'Texto', length: 'corto' });
  });

  it('never treats the header or separator row as a field', () => {
    const r = extract('ERD/Estructura_Datos.md', SIMPLE_ERD);
    const names = r.nodes.filter((n) => n.kind === 'type_member').map((n) => n.name);
    expect(names).not.toContain('Campo');
    expect(names).not.toContain(':---');
  });

  it('emits a table with no fields when its markdown table is empty', () => {
    const r = extract(
      'ERD/Estructura_Datos.md',
      [
        '# Estructura de Datos: X.accdb',
        '## Tabla: TbEmpty',
        '| Campo | Tipo | Longitud |',
        '| :--- | :--- | :--- |',
        '',
      ].join('\n'),
    );
    expect(r.nodes.some((n) => n.kind === 'class' && n.name === 'TbEmpty')).toBe(true);
    expect(r.nodes.filter((n) => n.kind === 'type_member')).toHaveLength(0);
  });

  it('collapses a table name declared twice in one document onto one node', () => {
    const r = extract(
      'ERD/Estructura_Datos.md',
      [
        '# Estructura de Datos: X.accdb',
        '## Tabla: TbDup',
        '| Campo | Tipo | Longitud |',
        '| :--- | :--- | :--- |',
        '| A | 4 | 4 |',
        '## Tabla: TbDup',
        '| Campo | Tipo | Longitud |',
        '| :--- | :--- | :--- |',
        '| B | 4 | 4 |',
        '',
      ].join('\n'),
    );
    const ids = new Set(
      r.nodes.filter((n) => n.kind === 'class' && n.name === 'TbDup').map((n) => n.id),
    );
    expect(ids.size).toBe(1);
  });
});

describe('AccessErdExtractor — linked tables and external backends', () => {
  const LINKED = [
    '# Estructura de Datos: X.accdb',
    '',
    '## Tabla: TbAplicaciones',
    '> (LINKED) **Tabla Vinculada**',
    '> *Origen:* TbAplicacionesOrigen',
    '> *Conexión:* MS Access;PWD=***;DATABASE=\\\\datoste\\Apps\\Lanzadera_Datos.accdb',
    '',
    '| Campo | Tipo | Longitud |',
    '| :--- | :--- | :--- |',
    '',
  ].join('\n');

  it('emits a references edge from the linked table to its external backend', () => {
    const r = extract('ERD/Estructura_Datos.md', LINKED);
    const table = r.nodes.find((n) => n.kind === 'class' && n.name === 'TbAplicaciones')!;
    const backend = r.nodes.find((n) => n.kind === 'file' && n.metadata?.external === true)!;
    expect(backend).toBeDefined();
    expect(backend.metadata?.backendPath).toBe(
      '//datoste/apps/lanzadera_datos.accdb',
    );
    expect(backend.name).toBe('lanzadera_datos.accdb');

    const edge = r.edges.find(
      (e) => e.kind === 'references' && e.source === table.id && e.target === backend.id,
    );
    expect(edge).toBeDefined();
    expect(edge?.metadata?.synthesizedBy).toBe(ERD_LINKED_TABLE_SYNTHESIZED_BY);
    expect(edge?.metadata?.originTable).toBe('TbAplicacionesOrigen');
    expect(edge?.metadata?.backendPath).toBe('//datoste/apps/lanzadera_datos.accdb');
  });

  // #256 already models an external Access backend as a `file` node keyed on
  // the normalized path with `metadata.external`. This half must reuse that
  // node, not invent a second representation.
  it('reuses the #256 external-backend node shape, so both halves converge', () => {
    const r = extract('ERD/Estructura_Datos.md', LINKED);
    const backend = r.nodes.find((n) => n.kind === 'file' && n.metadata?.external === true)!;
    expect(backend.language).toBe('vba');
    expect(backend.qualifiedName).toBe('//datoste/apps/lanzadera_datos.accdb');
    expect(backend.filePath).toBe(
      'synthetic:external-db///datoste/apps/lanzadera_datos.accdb',
    );
    // The linked-table edge is its own mechanism, so it keeps its own tag.
    expect(ERD_LINKED_TABLE_SYNTHESIZED_BY).not.toBe(EXTERNAL_BACKEND_SYNTHESIZED_BY);
  });

  it('marks the linked table on the table node itself', () => {
    const r = extract('ERD/Estructura_Datos.md', LINKED);
    const table = r.nodes.find((n) => n.kind === 'class' && n.name === 'TbAplicaciones')!;
    expect(table.metadata?.linked).toBe(true);
    expect(table.metadata?.originTable).toBe('TbAplicacionesOrigen');
    expect(table.metadata?.backendPath).toBe('//datoste/apps/lanzadera_datos.accdb');
  });

  it('collapses two linked tables sharing one backend onto a single node', () => {
    const two = [
      '# Estructura de Datos: X.accdb',
      '## Tabla: TbA',
      '> (LINKED) **Tabla Vinculada**',
      '> *Origen:* TbA',
      '> *Conexión:* MS Access;DATABASE=\\\\srv\\Apps\\Other.accdb',
      '| Campo | Tipo | Longitud |',
      '| :--- | :--- | :--- |',
      '## Tabla: TbB',
      '> (LINKED) **Tabla Vinculada**',
      '> *Origen:* TbB',
      '> *Conexión:* MS Access;DATABASE=\\\\SRV\\apps\\OTHER.accdb',
      '| Campo | Tipo | Longitud |',
      '| :--- | :--- | :--- |',
      '',
    ].join('\n');
    const r = extract('ERD/Estructura_Datos.md', two);
    const backends = r.nodes.filter((n) => n.kind === 'file' && n.metadata?.external === true);
    expect(new Set(backends.map((n) => n.id)).size).toBe(1);
    expect(
      r.edges.filter((e) => e.metadata?.synthesizedBy === ERD_LINKED_TABLE_SYNTHESIZED_BY),
    ).toHaveLength(2);
  });

  it('emits no backend edge when the connection string names no database', () => {
    const r = extract(
      'ERD/Estructura_Datos.md',
      [
        '# Estructura de Datos: X.accdb',
        '## Tabla: TbOdbc',
        '> (LINKED) **Tabla Vinculada**',
        '> *Origen:* dbo_TbOdbc',
        '> *Conexión:* ODBC;DSN=Something;UID=user',
        '| Campo | Tipo | Longitud |',
        '| :--- | :--- | :--- |',
        '',
      ].join('\n'),
    );
    expect(r.nodes.some((n) => n.metadata?.external === true)).toBe(false);
    expect(
      r.edges.some((e) => e.metadata?.synthesizedBy === ERD_LINKED_TABLE_SYNTHESIZED_BY),
    ).toBe(false);
    // The table itself still exists and is still flagged as linked.
    const table = r.nodes.find((n) => n.kind === 'class' && n.name === 'TbOdbc')!;
    expect(table.metadata?.linked).toBe(true);
  });
});

describe('AccessErdExtractor — the real generated file shape (fixture)', () => {
  it('parses the BOM + CRLF fixture that mirrors the real export', () => {
    const file = path.join(FIXTURE_DIR, 'Estructura_Datos.md');
    const source = fs.readFileSync(file, 'utf8');
    // Guard the guard: the fixture must actually carry the BOM.
    expect(source.charCodeAt(0)).toBe(0xfeff);

    const r = extract('ERD/Estructura_Datos.md', source);
    const tables = r.nodes.filter((n) => n.kind === 'class');
    expect(tables.map((n) => n.name).sort()).toEqual(['TbAnexos', 'TbAplicaciones']);
    expect(r.nodes.filter((n) => n.kind === 'type_member')).toHaveLength(2);
    expect(
      r.edges.filter((e) => e.metadata?.synthesizedBy === ERD_LINKED_TABLE_SYNTHESIZED_BY),
    ).toHaveLength(1);
    expect(r.errors).toHaveLength(0);
  });

  it('ignores the non-Access README that sits beside it', () => {
    const source = fs.readFileSync(path.join(FIXTURE_DIR, 'README.md'), 'utf8');
    const r = extract('ERD/README.md', source);
    expect(r.nodes).toHaveLength(0);
    expect(r.errors).toHaveLength(0);
  });
});

describe('Access ERD — end-to-end indexing', () => {
  let dir: string;

  beforeEach(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'cg-access-erd-'));
    clearProjectConfigCache();
  });

  afterEach(async () => {
    clearProjectConfigCache();
    await new Promise((resolve) => setTimeout(resolve, 100));
    try {
      fs.rmSync(dir, { recursive: true, force: true });
    } catch {
      // ignore cleanup EPERM on Windows
    }
  });

  function writeErd(relDir: string, contents: string): void {
    const full = path.join(dir, relDir);
    fs.mkdirSync(full, { recursive: true });
    fs.writeFileSync(path.join(full, 'Estructura_Datos.md'), contents);
  }

  it('indexes tables and fields from an ERD export', async () => {
    writeErd('ERD', '\uFEFF' + SIMPLE_ERD.replace(/\n/g, '\r\n'));

    const cg = CodeGraph.initSync(dir);
    await cg.indexAll();

    const table = cg.getNodesByName('TbAnexos').find((n) => n.kind === 'class');
    expect(table).toBeDefined();
    const fields = cg.getNodesByName('Titulo').filter((n) => n.kind === 'type_member');
    expect(fields).toHaveLength(1);
    expect(fields[0]!.qualifiedName).toBe('TbAnexos.Titulo');

    await cg.destroy();
  }, 60_000);

  it('converges a SQL table reference onto the ERD table node', async () => {
    writeErd('ERD', '\uFEFF' + SIMPLE_ERD.replace(/\n/g, '\r\n'));
    fs.writeFileSync(
      path.join(dir, 'Datos.cls'),
      [
        'Public Sub CargarAnexos()',
        '    CurrentDb.Execute "SELECT * FROM TbAnexos"',
        'End Sub',
      ].join('\n'),
    );

    const cg = CodeGraph.initSync(dir);
    await cg.indexAll();

    // Exactly ONE TbAnexos node survives, and it is the ERD one.
    const tables = cg.getNodesByName('TbAnexos').filter((n) => n.kind === 'class');
    expect(tables).toHaveLength(1);
    expect(tables[0]!.metadata?.synthesizedBy).toBe(ERD_TABLE_SYNTHESIZED_BY);

    // It owns the ERD fields...
    const contained = cg
      .getOutgoingEdges(tables[0]!.id)
      .filter((e) => e.kind === 'contains')
      .map((e) => cg.getNode(e.target)?.name);
    expect(contained).toContain('Titulo');

    // ...and carries the SQL reference from the .cls.
    const callers = cg.getCallers(tables[0]!.id);
    const sqlRef = callers.find((c) => c.edge.metadata?.synthesizedBy === 'vba-sql-table');
    expect(sqlRef).toBeDefined();
    expect(sqlRef?.edge.metadata?.repointDecision).toBe('reponted-to-real');

    await cg.destroy();
  }, 60_000);

  it('repoints nothing when two ERD files declare the same table', async () => {
    writeErd('ERD', '\uFEFF' + SIMPLE_ERD.replace(/\n/g, '\r\n'));
    writeErd('backend/ERD', '\uFEFF' + SIMPLE_ERD.replace(/\n/g, '\r\n'));
    fs.writeFileSync(
      path.join(dir, 'Datos.cls'),
      [
        'Public Sub CargarAnexos()',
        '    CurrentDb.Execute "SELECT * FROM TbAnexos"',
        'End Sub',
      ].join('\n'),
    );

    const cg = CodeGraph.initSync(dir);
    await cg.indexAll();

    const tables = cg.getNodesByName('TbAnexos').filter((n) => n.kind === 'class');
    const erdTables = tables.filter(
      (n) => n.metadata?.synthesizedBy === ERD_TABLE_SYNTHESIZED_BY,
    );
    // Both ERD declarations survive, and so does the untouched SQL placeholder.
    expect(erdTables).toHaveLength(2);
    expect(tables.length).toBe(3);

    const placeholder = tables.find(
      (n) => n.metadata?.synthesizedBy !== ERD_TABLE_SYNTHESIZED_BY,
    )!;
    const callers = cg.getCallers(placeholder.id);
    const sqlRef = callers.find((c) => c.edge.metadata?.synthesizedBy === 'vba-sql-table');
    expect(sqlRef?.edge.metadata?.repointDecision).toBe('declined-ambiguous');

    await cg.destroy();
  }, 60_000);

  it('ignores a non-Access markdown file under an ERD directory', async () => {
    const erd = path.join(dir, 'ERD');
    fs.mkdirSync(erd, { recursive: true });
    fs.writeFileSync(
      path.join(erd, 'README.md'),
      ['# Entity Relationship Diagrams', '', '## Tabla: NotATable', ''].join('\n'),
    );

    const cg = CodeGraph.initSync(dir);
    const result = await cg.indexAll();

    expect(cg.getNodesByName('NotATable')).toHaveLength(0);
    expect(result.errors).toHaveLength(0);

    await cg.destroy();
  }, 60_000);

  it('leaves a project with no ERD directory exactly as it was', async () => {
    fs.writeFileSync(
      path.join(dir, 'Datos.cls'),
      [
        'Public Sub CargarAnexos()',
        '    CurrentDb.Execute "SELECT * FROM TbAnexos"',
        'End Sub',
      ].join('\n'),
    );

    const cg = CodeGraph.initSync(dir);
    const result = await cg.indexAll();

    expect(result.errors).toHaveLength(0);
    // The SQL placeholder is untouched — no ERD to promote it to.
    const tables = cg.getNodesByName('TbAnexos').filter((n) => n.kind === 'class');
    expect(tables).toHaveLength(1);
    expect(tables[0]!.metadata?.synthesizedBy).not.toBe(ERD_TABLE_SYNTHESIZED_BY);
    expect(
      cg.getNodesByName('TbAnexos').some((n) => n.kind === 'type_member'),
    ).toBe(false);

    await cg.destroy();
  }, 60_000);
});
