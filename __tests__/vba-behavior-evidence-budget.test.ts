import { describe, expect, it, vi } from 'vitest';
import { buildBehaviorEvidence } from '../src/graph/behavior-evidence';
import { QueryBuilder } from '../src/db/queries';
import { Node } from '../src/types';

// Synthetic DAGs isolate traversal complexity independently of VBA extraction.
function graph(layers: number, duplicateNames = false) {
  const nodes = new Map<string, Node>();
  for (let layer = 0; layer < layers; layer++) {
    for (let branch = 0; branch < (layer === 0 ? 1 : 2); branch++) {
      const id = `${layer.toString().padStart(2, '0')}-${branch}`;
      nodes.set(id, { id, name: duplicateNames ? `Step${layer}` : id,
        qualifiedName: id, kind: 'function', filePath: 'Synthetic.bas',
        language: 'vba', startLine: layer + 1, endLine: layer + 1,
        startColumn: 0, endColumn: 0 });
    }
  }
  const expand = vi.fn((id: string, kinds: string[]) => {
    if (!kinds.includes('calls')) return [];
    const next = Number(id.split('-')[0]) + 1;
    return [...nodes.values()].filter(n => Number(n.id.split('-')[0]) === next)
      .map(n => ({ source: id, target: n.id, kind: 'calls' }));
  });
  const queries = { getNodeById: (id: string) => nodes.get(id),
    getOutgoingEdges: expand, getIncomingEdges: () => [],
    getUnresolvedReferencesForFile: () => [] } as unknown as QueryBuilder;
  return { queries, expansions: () => expand.mock.calls.filter(c => c[1].includes('calls')).length };
}

describe('issue #309 — traversal work budget', () => {
  it.each([false, true])('bounds layered DAG expansion even with duplicate payloads (%s)', duplicates => {
    const fixture = graph(14, duplicates);
    const result = buildBehaviorEvidence(fixture.queries, { nodeId: '00-0', maxResults: 1, maxCallDepth: 20 });
    expect(result.evidence).toHaveLength(1);
    expect(result.context.truncated.results).toBe(true);
    expect(fixture.expansions()).toBeLessThanOrEqual(40);
    expect(buildBehaviorEvidence(fixture.queries, { nodeId: '00-0', maxResults: 1, maxCallDepth: 20 })).toEqual(result);
  });

  it('does not report truncation when all paths fit exactly', () => {
    const { queries } = graph(2);
    const result = buildBehaviorEvidence(queries, { nodeId: '00-0', maxResults: 2 });
    expect(result.evidence).toHaveLength(2);
    expect(result.context.truncated.results).toBe(false);
  });

  it('does not report truncation for a fully consumed duplicate path', () => {
    const { queries } = graph(2, true);
    const result = buildBehaviorEvidence(queries, { nodeId: '00-0', maxResults: 1 });
    expect(result.evidence).toHaveLength(1);
    expect(result.context.truncated.results).toBe(false);
  });
});
