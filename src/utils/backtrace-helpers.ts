import { SqliteDatabase } from '../db';

/**
 * VBA Handler Backtrace Helper Functions
 */

export interface VariableInfo {
  name: string;
  type: string;
}

/**
 * Parses parameters from a VBA subroutine/function signature,
 * extracting parameters of custom types (filtering out primitive types).
 */
export function parseSignatureParams(signature: string): VariableInfo[] {
  const primitiveTypes = new Set([
    'long', 'integer', 'string', 'boolean', 'double', 'single',
    'byte', 'currency', 'date', 'variant', 'object',
    'longlong', 'longptr', 'decimal'
  ]);
  const regex = /(?:ByVal|ByRef)?\s*(\w+)\s+As\s+(\w+)/gi;
  const result: VariableInfo[] = [];
  
  const matches = signature.matchAll(regex);
  for (const m of matches) {
    const name = m[1];
    const type = m[2];
    if (name !== undefined && type !== undefined) {
      if (!primitiveTypes.has(type.toLowerCase())) {
        result.push({ name, type });
      }
    }
  }
  
  return result;
}

/**
 * Reconstructs a multiline SQL query string concatenated using VBA line continuation
 * and string concatenation operators. Accumulates up to a limit of 200 characters.
 */
export function reconstructSQL(lines: string[]): string {
  let accumulated = "";
  const stringRegex = /"((?:[^"\\]|\\.)*)"/g;
  
  for (const line of lines) {
    let match;
    stringRegex.lastIndex = 0;
    while ((match = stringRegex.exec(line)) !== null) {
      let content = match[1];
      if (content !== undefined) {
        content = content.replace(/\\'/g, "'").replace(/\\"/g, '"').replace(/\\\\/g, '\\');
        accumulated += content;
        if (accumulated.length >= 200) {
          break;
        }
      }
    }
    if (accumulated.length >= 200) {
      break;
    }
  }
  
  if (accumulated.length > 200) {
    accumulated = accumulated.slice(0, 200);
  }
  
  return accumulated;
}

export interface TraversalNode {
  id: string;
  name: string;
  kind: string;
  children: TraversalNode[];
}

export interface TraversalResult {
  tree: TraversalNode | null;
  cycle_detected: boolean;
  warnings: string[];
}

/**
 * Traverses VBA call/relationship graphs starting from a specific node ID,
 * tracking cycles and capping search depth.
 */
export function traverseGraph(
  db: SqliteDatabase,
  startNodeId: string,
  maxDepth = 10
): TraversalResult {
  if (!db) {
    return {
      tree: null,
      cycle_detected: false,
      warnings: ['DATABASE_NOT_PROVIDED']
    };
  }

  const warnings: string[] = [];
  let cycleDetected = false;

  try {
    // Optimization: Prepare statements once to avoid parsing cost inside the recursion
    const stmtNode = db.prepare('SELECT name, kind FROM nodes WHERE id = ?');
    const stmtEdges = db.prepare('SELECT target FROM edges WHERE source = ?');

    function helper(nodeId: string, depth: number, visited: Set<string>): TraversalNode | null {
      const nodeRow = stmtNode.get(nodeId) as { name: string; kind: string } | undefined;
      if (!nodeRow) {
        return null;
      }

      const node: TraversalNode = {
        id: nodeId,
        name: nodeRow.name,
        kind: nodeRow.kind,
        children: []
      };

      // Cycle detection
      if (visited.has(nodeId)) {
        cycleDetected = true;
        return node;
      }

      // Depth capping
      if (depth >= maxDepth) {
        if (!warnings.includes('MAX_DEPTH_EXCEEDED')) {
          warnings.push('MAX_DEPTH_EXCEEDED');
        }
        return node;
      }

      // Fetch outgoing target nodes
      const edges = stmtEdges.all(nodeId) as { target: string }[];

      // Traverse children with updated path history
      const newVisited = new Set(visited);
      newVisited.add(nodeId);

      for (const edge of edges) {
        const child = helper(edge.target, depth + 1, newVisited);
        if (child) {
          node.children.push(child);
        }
      }

      return node;
    }

    const tree = helper(startNodeId, 0, new Set<string>());

    return {
      tree,
      cycle_detected: cycleDetected,
      warnings
    };
  } catch (error) {
    const errorMsg = error instanceof Error ? error.message : String(error);
    return {
      tree: null,
      cycle_detected: false,
      warnings: [`TRAVERSAL_ERROR: ${errorMsg}`]
    };
  }
}
