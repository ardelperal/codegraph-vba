/**
 * DB primitives for VBA call-stub resolution (vba-graph-connectivity-fixes,
 * issue #12). Direct `DatabaseConnection` + `QueryBuilder` unit tests
 * (pattern from `db-perf.test.ts`) — supplementary coverage; the
 * load-bearing e2e coverage lives in `extraction-vba-realfixtures.test.ts`
 * per the Windows CI VBA regression subset requirement.
 *
 * `getVbaCallStubs()` finds candidate stub nodes via a JOIN against
 * `edges.metadata` rather than `nodes.metadata`: the stub flag is a
 * relationship fact about an unresolved call edge, not an intrinsic property
 * of the target symbol.
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import { DatabaseConnection } from '../src/db';
import { QueryBuilder } from '../src/db/queries';
import { Node } from '../src/types';

function makeVbaFunctionNode(id: string, name = id): Node {
  return {
    id,
    kind: 'function',
    name,
    qualifiedName: name,
    filePath: 'src/modules/Mod.bas',
    language: 'vba',
    startLine: 1,
    endLine: 1,
    startColumn: 0,
    endColumn: 0,
    updatedAt: Date.now(),
  };
}

describe('VBA call-stub DB primitives', () => {
  let dir: string;
  let db: DatabaseConnection;
  let q: QueryBuilder;

  beforeEach(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'db-vba-stub-'));
    db = DatabaseConnection.initialize(path.join(dir, 'test.db'));
    q = new QueryBuilder(db.getDb());
  });

  afterEach(() => {
    db.close();
    if (fs.existsSync(dir)) fs.rmSync(dir, { recursive: true, force: true });
  });

  describe('getVbaCallStubs()', () => {
    it('returns a function/vba node targeted by a calls edge with metadata.stub===true', () => {
      q.insertNodes([makeVbaFunctionNode('caller'), makeVbaFunctionNode('stub', 'Foo.Bar')]);
      q.insertEdges([
        {
          source: 'caller',
          target: 'stub',
          kind: 'calls',
          provenance: 'heuristic',
          metadata: { synthesizedBy: 'vba-name-resolution', stub: true, receiverType: 'Foo', member: 'Bar' },
        },
      ]);

      const stubs = q.getVbaCallStubs();
      expect(stubs.map((n) => n.id)).toContain('stub');
    });

    it('does NOT return a node whose incoming calls edge has metadata.stub===false', () => {
      q.insertNodes([makeVbaFunctionNode('caller'), makeVbaFunctionNode('real', 'Foo.Bar')]);
      q.insertEdges([
        {
          source: 'caller',
          target: 'real',
          kind: 'calls',
          provenance: 'heuristic',
          metadata: { synthesizedBy: 'vba-name-resolution', stub: false, receiverType: 'Foo', member: 'Bar' },
        },
      ]);

      const stubs = q.getVbaCallStubs();
      expect(stubs.map((n) => n.id)).not.toContain('real');
    });

    it('does NOT return a node with no metadata on its incoming calls edge', () => {
      q.insertNodes([makeVbaFunctionNode('caller'), makeVbaFunctionNode('plain', 'Plain')]);
      q.insertEdges([{ source: 'caller', target: 'plain', kind: 'calls' }]);

      const stubs = q.getVbaCallStubs();
      expect(stubs.map((n) => n.id)).not.toContain('plain');
    });

    it('does NOT return a non-vba or non-function node even if targeted by a stub-flagged edge', () => {
      const classNode: Node = {
        ...makeVbaFunctionNode('caller'),
      };
      const nonFunctionStub: Node = {
        id: 'notafunc',
        kind: 'class',
        name: 'Foo',
        qualifiedName: 'Foo',
        filePath: 'src/classes/Foo.cls',
        language: 'vba',
        startLine: 1,
        endLine: 1,
        startColumn: 0,
        endColumn: 0,
        updatedAt: Date.now(),
      };
      q.insertNodes([classNode, nonFunctionStub]);
      q.insertEdges([
        {
          source: 'caller',
          target: 'notafunc',
          kind: 'calls',
          metadata: { synthesizedBy: 'vba-name-resolution', stub: true },
        },
      ]);

      const stubs = q.getVbaCallStubs();
      expect(stubs.map((n) => n.id)).not.toContain('notafunc');
    });
  });

  describe('getVbaReferenceStubs()', () => {
    it('returns a class node targeted by a references edge from a VBA module', () => {
      const caller: Node = { ...makeVbaFunctionNode('caller'), language: 'vba' };
      const stub: Node = { ...makeVbaFunctionNode('stub', 'MyEnum'), kind: 'class', language: 'vba' };
      q.insertNodes([caller, stub]);
      q.insertEdges([
        {
          source: 'caller',
          target: 'stub',
          kind: 'references',
          provenance: 'heuristic',
        },
      ]);

      const stubs = q.getVbaReferenceStubs();
      expect(stubs.map((n) => n.id)).toContain('stub');
    });

    it('does NOT return a node targeted by a non-references edge', () => {
      const caller: Node = { ...makeVbaFunctionNode('caller'), language: 'vba' };
      const stub: Node = { ...makeVbaFunctionNode('stub', 'MyEnum'), kind: 'class', language: 'vba' };
      q.insertNodes([caller, stub]);
      q.insertEdges([
        {
          source: 'caller',
          target: 'stub',
          kind: 'calls',
          provenance: 'heuristic',
        },
      ]);

      const stubs = q.getVbaReferenceStubs();
      expect(stubs.map((n) => n.id)).not.toContain('stub');
    });
  });

  describe('node metadata persistence', () => {
    it('round-trips node metadata through the nodes table', () => {
      q.insertNode({
        ...makeVbaFunctionNode('decl', 'GetTickCount'),
        kind: 'declare',
        metadata: {
          dll: 'kernel32',
          declareKind: 'function',
          ptrSafe: true,
        },
      });

      const node = q.getNodeById('decl');

      expect(node?.metadata).toEqual({
        dll: 'kernel32',
        declareKind: 'function',
        ptrSafe: true,
      });
    });
  });

  describe('repointEdgeTarget()', () => {
    it('updates target + metadata in place, leaves other columns untouched', () => {
      q.insertNodes([
        makeVbaFunctionNode('caller'),
        makeVbaFunctionNode('stub', 'Foo.Bar'),
        makeVbaFunctionNode('real', 'Bar'),
      ]);
      q.insertEdges([
        {
          source: 'caller',
          target: 'stub',
          kind: 'calls',
          provenance: 'heuristic',
          line: 42,
          column: 3,
          metadata: { synthesizedBy: 'vba-name-resolution', stub: true, receiverType: 'Foo', member: 'Bar' },
        },
      ]);
      const [edge] = q.getIncomingEdges('stub', ['calls']);
      expect(edge?.id).toBeDefined();

      q.repointEdgeTarget(
        edge!.id!,
        'real',
        JSON.stringify({ synthesizedBy: 'vba-name-resolution', stub: false, receiverType: 'Foo', member: 'Bar' }),
      );

      const outgoing = q.getOutgoingEdges('caller');
      expect(outgoing).toHaveLength(1);
      expect(outgoing[0]?.target).toBe('real');
      expect(outgoing[0]?.metadata?.stub).toBe(false);
      // Untouched columns.
      expect(outgoing[0]?.line).toBe(42);
      expect(outgoing[0]?.column).toBe(3);
      expect(outgoing[0]?.kind).toBe('calls');
      expect(outgoing[0]?.provenance).toBe('heuristic');
    });
  });

  describe('deleteEdgeById()', () => {
    it('removes exactly that row and leaves sibling edges intact', () => {
      q.insertNodes([makeVbaFunctionNode('caller'), makeVbaFunctionNode('t1'), makeVbaFunctionNode('t2')]);
      q.insertEdges([
        { source: 'caller', target: 't1', kind: 'calls' },
        { source: 'caller', target: 't2', kind: 'calls' },
      ]);
      const edges = q.getOutgoingEdges('caller');
      expect(edges).toHaveLength(2);
      const toDelete = edges.find((e) => e.target === 't1')!;

      q.deleteEdgeById(toDelete.id!);

      const remaining = q.getOutgoingEdges('caller');
      expect(remaining).toHaveLength(1);
      expect(remaining[0]?.target).toBe('t2');
    });
  });

  describe('edgeExists()', () => {
    it('returns true when a matching (source,target,kind) row exists', () => {
      q.insertNodes([makeVbaFunctionNode('a'), makeVbaFunctionNode('b')]);
      q.insertEdges([{ source: 'a', target: 'b', kind: 'calls' }]);
      expect(q.edgeExists('a', 'b', 'calls')).toBe(true);
    });

    it('returns false when no matching row exists', () => {
      q.insertNodes([makeVbaFunctionNode('a'), makeVbaFunctionNode('b')]);
      expect(q.edgeExists('a', 'b', 'calls')).toBe(false);
    });

    it('returns false when the kind does not match', () => {
      q.insertNodes([makeVbaFunctionNode('a'), makeVbaFunctionNode('b')]);
      q.insertEdges([{ source: 'a', target: 'b', kind: 'references' }]);
      expect(q.edgeExists('a', 'b', 'calls')).toBe(false);
    });
  });
});
