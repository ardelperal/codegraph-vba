/**
 * The MCP `initialize` instructions must describe the VBA/Access surfaces the
 * fork actually indexes (issue #320).
 *
 * `src/mcp/server-instructions.ts` is the single source of truth for
 * agent-facing guidance (#529) — it is the first and only thing an agent is
 * told about what lives in this index. Prose there is cheap to write and easy
 * to let rot, so these tests do NOT assert on wording. They run the real
 * extractors, collect the node kinds and `synthesizedBy` tags those extractors
 * emit, and require the instructions to name each one. Rename a tag in an
 * extractor and this suite fails until the guidance follows.
 */
import { describe, it, expect } from 'vitest';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { SERVER_INSTRUCTIONS } from '../src/mcp/server-instructions';
import { SqlQueryExtractor } from '../src/extraction/sql-query-extractor';
import { VbaTestSequenceExtractor } from '../src/extraction/vba-test-sequence-extractor';
import { VbaTestManifestExtractor } from '../src/extraction/vba-test-manifest-extractor';
import {
  AccessErdExtractor,
  ERD_TABLE_SYNTHESIZED_BY,
  ERD_LINKED_TABLE_SYNTHESIZED_BY,
} from '../src/extraction/access-erd-extractor';
import { EXTERNAL_BACKEND_SYNTHESIZED_BY } from '../src/extraction/sql-table-scan';

/** Every `synthesizedBy` tag on the nodes, edges and refs of a result. */
function tagsOf(result: {
  nodes: Array<{ metadata?: Record<string, unknown> }>;
  edges: Array<{ metadata?: Record<string, unknown> }>;
  unresolvedReferences: Array<{ metadata?: Record<string, unknown> }>;
}): string[] {
  const out = new Set<string>();
  for (const item of [...result.nodes, ...result.edges, ...result.unresolvedReferences]) {
    const tag = item.metadata?.synthesizedBy;
    if (typeof tag === 'string') out.add(tag);
  }
  return [...out];
}

describe('MCP server instructions — VBA/Access coverage (#320)', () => {
  it('names every tag and node kind a saved Access query emits', () => {
    const r = new SqlQueryExtractor(
      'queries/Q_Ventas.sql',
      'SELECT * FROM TbVentas INNER JOIN TbClientes ON 1 = 1;\n',
    ).extract();

    // Sanity: the fixture really does produce a query node with table refs, so
    // an extractor that silently stopped emitting them cannot pass this test.
    expect(r.nodes.some((n) => n.kind === 'query')).toBe(true);
    const tags = tagsOf(r);
    expect(tags).toContain('sql-query-table');

    for (const tag of tags) {
      expect(SERVER_INSTRUCTIONS, `undocumented tag: ${tag}`).toContain(tag);
    }
    expect(SERVER_INSTRUCTIONS).toContain('query');
    expect(SERVER_INSTRUCTIONS).toContain('queries.json');
    expect(SERVER_INSTRUCTIONS).toContain(EXTERNAL_BACKEND_SYNTHESIZED_BY);
  });

  it('names the test-sequence surface, not just the manifests', () => {
    const r = new VbaTestSequenceExtractor(
      'tests/sequences/cache.json',
      JSON.stringify({
        runnerPolicy: { tool: 'dysflow', sequential: true },
        procedures: ['Test_Cache_RunSlice'],
      }),
    ).extract();

    const tags = tagsOf(r);
    expect(tags).toContain('vba-test-sequence');
    for (const tag of tags) {
      expect(SERVER_INSTRUCTIONS, `undocumented tag: ${tag}`).toContain(tag);
    }

    // The metadata an agent has to know exists to act on a sequence.
    const ref = r.unresolvedReferences[0];
    for (const key of Object.keys(ref?.metadata ?? {})) {
      if (key === 'synthesizedBy') continue;
      expect(SERVER_INSTRUCTIONS, `undocumented metadata key: ${key}`).toContain(key);
    }
  });

  it('still names the manifest surface it already documented', () => {
    // Guard against a future edit trimming the paragraph the sequences one
    // now sits beside.
    const r = new VbaTestManifestExtractor(
      'tests/tests.vba.smoke.json',
      JSON.stringify({ tests: [{ procedure: 'Test_X_RunAll' }] }),
    ).extract();
    for (const tag of tagsOf(r)) {
      expect(SERVER_INSTRUCTIONS, `undocumented tag: ${tag}`).toContain(tag);
    }
  });

  it('names every tag the Access ERD export emits', () => {
    const fixture = path.join(
      __dirname,
      'fixtures',
      'access-erd',
      'ERD',
      'Estructura_Datos.md',
    );
    const r = new AccessErdExtractor(
      'ERD/Estructura_Datos.md',
      fs.readFileSync(fixture, 'utf-8'),
    ).extract();

    // The fixture carries both a plain and a LINKED table, so both tags are
    // exercised rather than assumed.
    const tags = tagsOf(r);
    expect(tags).toContain(ERD_TABLE_SYNTHESIZED_BY);
    expect(tags).toContain(ERD_LINKED_TABLE_SYNTHESIZED_BY);

    for (const tag of tags) {
      expect(SERVER_INSTRUCTIONS, `undocumented tag: ${tag}`).toContain(tag);
    }
    // The node kinds an agent would have to ask for by name.
    expect(SERVER_INSTRUCTIONS).toContain('type_member');
    expect(SERVER_INSTRUCTIONS).toContain('ERD/');
  });
});
