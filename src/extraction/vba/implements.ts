/**
 * `Implements IFoo` sweep (REQ-CODE-5). Emits one `interface` node per
 * `Implements` declaration plus a pending `implements` edge whose source is
 * rewired to the module/class node once the orchestrator creates it.
 */
import { Edge } from '../../types';
import { generateNodeId } from '../tree-sitter-helpers';
import { VbaClassifier } from './context';
import { defineRule, runRules, VbaExtractionRule } from './rules';

/** Implements regex. */
const IMPLEMENTS_RE = /^\s*Implements\s+(\p{L}[\p{L}\p{N}_]*)/iu;

/**
 * Issue #153: the declarative rule table for the implements concern.
 * One rule — `implements` — matching a leading `Implements <Name>`
 * declaration. The body has no inter-line state to track, so the
 * orchestrator walks the table and bumps `this.count` on every
 * non-null emit. Keeps the inline `classifyLine` below honest: any
 * branch not represented in this table is dead code.
 */
export const RULES: readonly VbaExtractionRule<unknown>[] = [
  defineRule({
    id: 'implements',
    description:
      'Match a leading `Implements <Name>` declaration; emit an `interface` node + a pending `implements` edge.',
    pattern: IMPLEMENTS_RE,
    emit: (m, ctx, line, lineNum) => {
      const name = m[1] ?? '';
      if (!name) return null;
      const targetId = generateNodeId(ctx.filePath, 'interface', name, lineNum);
      ctx.nodes.push({
        id: targetId,
        kind: 'interface',
        name,
        qualifiedName: name,
        filePath: ctx.filePath,
        language: 'vba',
        startLine: lineNum,
        endLine: lineNum,
        startColumn: 0,
        endColumn: line.length,
        updatedAt: Date.now(),
      });
      const edge: Edge = {
        source: '', // placeholder; rewritten after module node exists
        target: targetId,
        kind: 'implements',
        // S4 fix: `Implements IFoo` is a static, source-declared fact —
        // not a guess. Use the `parser` provenance (generalizes
        // `tree-sitter` for non-tree-sitter extractors like our regex
        // sweepers) instead of `heuristic`, which is reserved for
        // guessed/inferred edges.
        provenance: 'parser',
        line: lineNum,
        column: 0,
      };
      ctx.edges.push(edge);
      ctx.pendingModuleOrClassSource.push(edge);
      return { name };
    },
  }),
];

/**
 * Issue #83: factory for the `Implements` classifier. Stateless per-line.
 *
 * The body walks the declarative `RULES` table — the inline regex
 * match that lived here before #153 is now a single entry in
 * `RULES`. This is the canonical pattern every classifier will
 * converge on as part of the refactor; the legacy "one giant
 * if/else cascade" is gone.
 */
export function createImplementsClassifier(): VbaClassifier {
  return {
    name: 'implements',
    count: 0,
    classifyLine(line, i, ctx) {
      const lineNum = i + 1;
      this.count += runRules(RULES, ctx, line, line, lineNum, {});
    },
  };
}
