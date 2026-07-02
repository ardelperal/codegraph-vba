import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import { DatabaseConnection } from '../src/db';
import { QueryBuilder } from '../src/db/queries';
import { Node } from '../src/types';
import {
  parseSignatureParams,
  reconstructSQL,
  traverseGraph
} from '../src/utils/backtrace-helpers';

function makeVbaNode(id: string, name: string, kind: string): Node {
  return {
    id,
    kind,
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

describe('VBA Handler Backtrace Helpers', () => {
  describe('parseSignatureParams', () => {
    it('should extract custom type parameters and filter out primitive types', () => {
      const signature = 'Public Sub ProcessOrder(ByRef ctx As OrderContext, ByVal flags As Long)';
      const result = parseSignatureParams(signature);
      expect(result).toEqual([{ name: 'ctx', type: 'OrderContext' }]);
    });
  });

  describe('reconstructSQL', () => {
    it('should reconstruct a multiline SQL query string from VBA lines', () => {
      const lines = [
        'db.Execute "INSERT INTO Log (Msg) " & _',
        '           "VALUES (\'Order Processed\')"'
      ];
      const result = reconstructSQL(lines);
      expect(result).toBe("INSERT INTO Log (Msg) VALUES ('Order Processed')");
    });
  });

  describe('traverseGraph', () => {
    let dir: string;
    let db: DatabaseConnection;
    let q: QueryBuilder;

    beforeEach(() => {
      dir = fs.mkdtempSync(path.join(os.tmpdir(), 'db-vba-backtrace-'));
      db = DatabaseConnection.initialize(path.join(dir, 'test.db'));
      q = new QueryBuilder(db.getDb());
    });

    afterEach(() => {
      db.close();
      if (fs.existsSync(dir)) fs.rmSync(dir, { recursive: true, force: true });
    });

    it('should resolve complete call paths from control to event to method', () => {
      q.insertNodes([
        makeVbaNode('btn1', 'btnSave', 'control'),
        makeVbaNode('btn1_click', 'btnSave_Click', 'event'),
        makeVbaNode('save_method', 'SaveRecord', 'function'),
      ]);
      q.insertEdges([
        { source: 'btn1', target: 'btn1_click', kind: 'defines-event' },
        { source: 'btn1_click', target: 'save_method', kind: 'calls' },
      ]);

      const result = traverseGraph(db.getDb(), 'btn1');
      expect(result.cycle_detected).toBe(false);
      expect(result.warnings).toEqual([]);
      expect(result.tree).toEqual({
        id: 'btn1',
        name: 'btnSave',
        kind: 'control',
        children: [
          {
            id: 'btn1_click',
            name: 'btnSave_Click',
            kind: 'event',
            children: [
              {
                id: 'save_method',
                name: 'SaveRecord',
                kind: 'function',
                children: [],
              },
            ],
          },
        ],
      });
    });

    it('should detect cycles and terminate traversal to avoid infinite loop', () => {
      q.insertNodes([
        makeVbaNode('nodeA', 'FuncA', 'function'),
        makeVbaNode('nodeB', 'FuncB', 'function'),
      ]);
      q.insertEdges([
        { source: 'nodeA', target: 'nodeB', kind: 'calls' },
        { source: 'nodeB', target: 'nodeA', kind: 'calls' },
      ]);

      const result = traverseGraph(db.getDb(), 'nodeA');
      expect(result.cycle_detected).toBe(true);
      expect(result.tree?.id).toBe('nodeA');
      // Children of nodeA contains nodeB, but nodeB's children shouldn't loop back indefinitely
      const childB = result.tree?.children[0];
      expect(childB?.id).toBe('nodeB');
      expect(childB?.children).toEqual([
        {
          id: 'nodeA',
          name: 'FuncA',
          kind: 'function',
          children: [],
        }
      ]);
    });

    it('should cap depth and add warning when maxDepth is exceeded', () => {
      q.insertNodes([
        makeVbaNode('nodeA', 'FuncA', 'function'),
        makeVbaNode('nodeB', 'FuncB', 'function'),
        makeVbaNode('nodeC', 'FuncC', 'function'),
      ]);
      q.insertEdges([
        { source: 'nodeA', target: 'nodeB', kind: 'calls' },
        { source: 'nodeB', target: 'nodeC', kind: 'calls' },
      ]);

      // maxDepth = 1 limits path to nodeA -> nodeB (nodeB's children not processed/traversed)
      const result = traverseGraph(db.getDb(), 'nodeA', 1);
      expect(result.warnings).toContain('MAX_DEPTH_EXCEEDED');
      expect(result.tree).toEqual({
        id: 'nodeA',
        name: 'FuncA',
        kind: 'function',
        children: [
          {
            id: 'nodeB',
            name: 'FuncB',
            kind: 'function',
            children: [],
          }
        ],
      });
    });
  });
});
