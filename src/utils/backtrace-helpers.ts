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
 * Node kinds that represent a UI object an event can be wired to.
 *
 * The canonical Access kinds (`form-layout`, `report-layout`,
 * `form-instance-control`) are what the extractor emits today; `control`,
 * `form` and `report` are kept so graphs written before those kinds existed
 * still traverse.
 */
const UI_NODE_KINDS = new Set([
  'form-instance-control',
  'form-layout',
  'report-layout',
  'control',
  'form',
  'report',
]);

/**
 * Edge kinds that always represent an execution step away from the current
 * node. `calls` is the real one; `defines-event` is a legacy shape some older
 * graphs stored outgoing from the control.
 */
const EXECUTION_EDGE_KINDS = ['calls', 'defines-event'];

/**
 * Node kinds a `references` edge may land on and still be a call.
 *
 * VBA's statement-form Sub call is genuinely ambiguous at parse time: `Foo`
 * on its own line could be a `Const` read, so the extractor deliberately
 * classifies it as `unqualified-ident` (issue #265) and the resolver stores
 * the resolved edge as `references`, not `calls`. That bare form is the
 * dominant call style in Access code-behind, so a consumer that follows only
 * `calls` reaches the handler and stops — the empty trace this helper exists
 * to prevent. Constraining the target kind is what keeps the rule honest:
 * containment, typing and data references (tables, queries, classes) land on
 * other kinds and stay out of the trace.
 */
const CALLABLE_TARGET_KINDS = ['function', 'method'];

/**
 * Traverses the VBA execution graph from a node id, following the semantics
 * the production extractor actually stores.
 *
 * Two directions are involved, which is the whole point of this helper:
 *
 * - An `event-handler` edge is stored HANDLER -> UI object (both for the
 *   `<Control>_<Event>` code-behind convention and for `=Expression()`
 *   wiring, which the resolver repoints the same way). So reaching a
 *   handler from its control or layout means following that edge BACKWARDS.
 * - A `calls` edge is stored CALLER -> CALLEE, so the rest of the path is
 *   followed forwards.
 *
 * Starting from a handler still works: it simply has no incoming
 * `event-handler` edge to expand, and its callees are found the usual way.
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
    const stmtCallees = db.prepare(
      `SELECT e.target AS id
         FROM edges e
         JOIN nodes n ON n.id = e.target
        WHERE e.source = ?
          AND (e.kind IN (${EXECUTION_EDGE_KINDS.map(() => '?').join(', ')})
               OR (e.kind = 'references'
                   AND n.kind IN (${CALLABLE_TARGET_KINDS.map(() => '?').join(', ')})))
        ORDER BY e.target`
    );
    // Inverse of the stored binding: the handler is the SOURCE of the
    // `event-handler` edge and the control/layout is its target.
    const stmtHandlers = db.prepare(
      "SELECT source AS id FROM edges WHERE target = ? AND kind = 'event-handler' ORDER BY source"
    );

    function nextNodeIds(nodeId: string, kind: string): string[] {
      const ids: string[] = [];
      const seen = new Set<string>();
      const push = (row: { id: string }) => {
        if (row.id !== undefined && !seen.has(row.id)) {
          seen.add(row.id);
          ids.push(row.id);
        }
      };

      if (UI_NODE_KINDS.has(kind)) {
        (stmtHandlers.all(nodeId) as { id: string }[]).forEach(push);
      }
      (
        stmtCallees.all(
          nodeId,
          ...EXECUTION_EDGE_KINDS,
          ...CALLABLE_TARGET_KINDS,
        ) as { id: string }[]
      ).forEach(push);

      return ids;
    }

    function helper(nodeId: string, depth: number, visited: Set<string>): TraversalNode | null {
      const nodeRow = stmtNode.get(nodeId) as { name: string; kind: string } | undefined;
      if (!nodeRow) {
        if (depth === 0) {
          warnings.push('START_NODE_NOT_FOUND');
        }
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

      // Traverse children with updated path history
      const newVisited = new Set(visited);
      newVisited.add(nodeId);

      for (const childId of nextNodeIds(nodeId, nodeRow.kind)) {
        const child = helper(childId, depth + 1, newVisited);
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
