/**
 * Issue #322: a CP1252 Access ERD export must decode like the rest of the
 * export, so its table names still join the graph.
 *
 * #53 gave `.bas` / `.cls` / `.form.txt` / `.report.txt` / `.sql` an
 * encoding-robust read because Access writes Windows-1252 and accented
 * identifiers are everywhere in Spanish-language Access projects. The Access
 * structure export (`ERD/*.md`) was left on the plain UTF-8 path.
 *
 * The damage was not cosmetic. The `.sql` naming a table DID get the CP1252
 * fallback, so it emitted a placeholder named `TbSituación`; the ERD declaring
 * that same table did NOT, so it declared `TbSituaci<U+FFFD>n`. The two never
 * matched: the declaration never joined the placeholder, and the table's
 * fields hung off an orphan node. This suite pins both halves — the decode,
 * and the join it exists to make possible.
 */
import { describe, it, expect, afterEach } from 'vitest';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as os from 'node:os';
import { readVbaSource } from '../src/extraction/vba-source';
import { usesAccessEncoding } from '../src/extraction';
import CodeGraph from '../src/index';

/** The generated export's real byte shape: BOM, CRLF, `# Estructura de Datos:`. */
const ERD_DOC = [
  '# Estructura de Datos: Datos.accdb',
  '',
  '## Tabla: TbSituación',
  '| Campo | Tipo | Longitud |',
  '| :--- | :--- | :--- |',
  '| Código | 4 | 4 |',
  '| Descripción | 10 | 255 |',
  '',
].join('\r\n');

/** Same document, CP1252 on disk — one byte per accented character. */
const cp1252 = (text: string) => Buffer.from(text, 'latin1');
const utf8 = (text: string) => Buffer.from(text, 'utf8');

const openProjects: Array<{ cg: CodeGraph; dir: string }> = [];

afterEach(async () => {
  while (openProjects.length > 0) {
    const { cg, dir } = openProjects.pop()!;
    try {
      await cg.close();
    } catch {
      // ignore close errors
    }
    await new Promise((r) => setTimeout(r, 50));
    try {
      fs.rmSync(dir, { recursive: true, force: true });
    } catch {
      // ignore cleanup EPERM on Windows
    }
  }
});

describe('Access ERD export — encoding (#322)', () => {
  it('routes the ERD export by path shape, and nothing else markdown', () => {
    // `.md` must never become a family EXTENSION — that would pull every
    // markdown file in every repo through the CP1252 fallback. Only a document
    // directly inside an `ERD/` directory qualifies.
    for (const p of [
      'ERD/Estructura_Datos.md',
      'src/backend/erd/Estructura_Datos.md',
      'C:\\proj\\ERD\\Estructura_Datos.md',
    ]) {
      expect(usesAccessEncoding(p), p).toBe(true);
    }
    for (const p of [
      'docs/architecture.md',
      'README.md',
      'ERD/notes/deep.md',
      'index.ts',
    ]) {
      expect(usesAccessEncoding(p), p).toBe(false);
    }
    // The extension half still routes on its own.
    expect(usesAccessEncoding('src/modules/constantes.bas')).toBe(true);
  });

  it('decodes a CP1252 export to the same names as its UTF-8 twin', () => {
    const asUtf8 = readVbaSource('ERD/Estructura_Datos.md', {
      readFile: () => utf8(ERD_DOC),
    });
    let fallbacks = 0;
    const asCp1252 = readVbaSource('ERD/Estructura_Datos.md', {
      readFile: () => cp1252(ERD_DOC),
      onFallback: () => { fallbacks++; },
    });

    expect(asCp1252.text).toBe(asUtf8.text);
    expect(asCp1252.text).toContain('TbSituación');
    expect(asCp1252.text).toContain('Código');
    // The fallback is what did the work; a silent pass would mean the bytes
    // happened to be valid UTF-8 and the test proves nothing.
    expect(fallbacks).toBe(1);
  });

  it('leaves a UTF-8 export with a BOM byte-identical', () => {
    const withBom = Buffer.concat([
      Buffer.from([0xef, 0xbb, 0xbf]),
      utf8(ERD_DOC),
    ]);
    let fallbacks = 0;
    const r = readVbaSource('ERD/Estructura_Datos.md', {
      readFile: () => withBom,
      onFallback: () => { fallbacks++; },
    });

    expect(r.text).toBe(ERD_DOC);
    expect(r.bomStripped).toBe(true);
    expect(fallbacks).toBe(0);
  });

  it('joins a CP1252 ERD table to the query that references it', async () => {
    // The failure this issue is really about: both files are CP1252 on disk,
    // but only one of them used to be decoded as such, so the two spellings of
    // the same table name never met.
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'cg-erd-enc-'));
    fs.mkdirSync(path.join(dir, 'ERD'), { recursive: true });
    fs.mkdirSync(path.join(dir, 'queries'), { recursive: true });

    fs.writeFileSync(path.join(dir, 'ERD', 'Estructura_Datos.md'), cp1252(ERD_DOC));
    fs.writeFileSync(
      path.join(dir, 'queries', 'Q_Situaciones.sql'),
      cp1252('SELECT * FROM TbSituación;\r\n'),
    );
    fs.writeFileSync(
      path.join(dir, 'queries', 'queries.json'),
      '[{ "name": "Q_Situaciones", "file": "Q_Situaciones.sql" }]',
    );

    const cg = await CodeGraph.init(dir, { index: false });
    openProjects.push({ cg, dir });
    await cg.indexAll();

    // The ERD declared the table and its fields under their real names. A
    // replacement character anywhere here is the exact signature of the bug.
    const erdNodes = cg.getNodesInFile('ERD/Estructura_Datos.md');
    const names = erdNodes.map((n) => n.name);
    expect(names).toContain('TbSituación');
    expect(names).toContain('Código');
    expect(names).toContain('Descripción');
    expect(names.filter((n) => (n ?? '').includes('�'))).toEqual([]);

    // And the two sides meet: the saved query's table reference resolves onto
    // the ERD's own declaration rather than a second, differently-spelled node.
    const table = erdNodes.find(
      (n) => n.kind === 'class' && n.name === 'TbSituación',
    );
    expect(table).toBeDefined();

    const query = cg
      .getNodesInFile('queries/Q_Situaciones.sql')
      .find((n) => n.kind === 'query');
    expect(query).toBeDefined();

    const toTable = cg
      .getOutgoingEdges(query!.id)
      .filter((e) => e.target === table!.id);
    expect(toTable.length).toBeGreaterThan(0);
  });
});
