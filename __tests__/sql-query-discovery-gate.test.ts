/**
 * Discovery gate for Dysflow saved queries.
 *
 * A `.sql` file is indexed as an Access query ONLY when its directory also
 * contains a `queries.json` manifest (Dysflow's marker). This protects every
 * non-Access repo: an ordinary `.sql` migration with no sibling manifest stays
 * ignored exactly as it was before SQL support existed.
 *
 * This exercises the real directory-discovery walk via CodeGraph.indexAll()
 * against a temp project with two query-bearing dirs — one with the manifest,
 * one without.
 */
import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { CodeGraph } from '../src';

describe('SQL query discovery gate (sibling queries.json)', () => {
  let dir = '';
  let cg: CodeGraph | null = null;

  beforeAll(async () => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'cg-sql-gate-'));

    // Dir WITH the manifest → indexed as a query.
    const withManifest = path.join(dir, 'with');
    fs.mkdirSync(withManifest, { recursive: true });
    fs.writeFileSync(
      path.join(withManifest, 'Managed.sql'),
      'SELECT * FROM TbManaged;\n',
    );
    fs.writeFileSync(
      path.join(withManifest, 'queries.json'),
      '[{ "name": "Managed", "file": "Managed.sql" }]\n',
    );

    // Dir WITHOUT the manifest → must be ignored (non-Access repo case).
    const without = path.join(dir, 'plain');
    fs.mkdirSync(without, { recursive: true });
    fs.writeFileSync(
      path.join(without, 'Migration.sql'),
      'SELECT * FROM TbMigration;\n',
    );

    cg = await CodeGraph.init(dir, { index: false });
    await cg.indexAll();
  }, 60_000);

  afterAll(async () => {
    if (cg) {
      try {
        await cg.close();
      } catch {
        // ignore
      }
    }
    if (dir && fs.existsSync(dir)) {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  });

  it('indexes a .sql that has a sibling queries.json', () => {
    if (!cg) return;
    const managed = cg.searchNodes('Managed', {
      languages: ['sql'],
      kinds: ['query'],
    });
    expect(managed.length).toBeGreaterThan(0);
  });

  it('ignores a .sql with no sibling queries.json', () => {
    if (!cg) return;
    const migration = cg.searchNodes('Migration', { languages: ['sql'] });
    expect(migration).toHaveLength(0);
    // And its table is not pulled in either.
    const tbl = cg.searchNodes('TbMigration', {});
    expect(tbl).toHaveLength(0);
  });
});
