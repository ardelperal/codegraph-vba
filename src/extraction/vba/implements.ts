/**
 * `Implements IFoo` sweep (REQ-CODE-5). Emits one `interface` node per
 * `Implements` declaration plus a pending `implements` edge whose source is
 * rewired to the module/class node once the orchestrator creates it.
 */
import { Edge } from '../../types';
import { generateNodeId } from '../tree-sitter-helpers';
import { VbaExtractorContext, VbaClassifier } from './context';

/** Implements regex. */
const IMPLEMENTS_RE = /^\s*Implements\s+(\p{L}[\p{L}\p{N}_]*)/iu;

/**
 * Issue #83: factory for the `Implements` classifier. Stateless per-line.
 */
export function createImplementsClassifier(): VbaClassifier {
  return {
    name: 'implements',
    count: 0,
    classifyLine(line, i, ctx) {
      const m = IMPLEMENTS_RE.exec(line);
      if (!m) return;
      const name = m[1] ?? '';
      if (!name) return;
      const lineNum = i + 1;
      const targetId = generateNodeId(
        ctx.filePath,
        'interface',
        name,
        lineNum,
      );
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
      this.count++;
    },
  };
}

/**
 * Backward-compat wrapper (see procedures.ts). Returns the classifier's
 * `count` so the orchestrator can decide `hasAnySymbols`.
 */
export function sweepImplements(ctx: VbaExtractorContext, src: string): number {
  const cls = createImplementsClassifier();
  const lines = src.split('\n');
  for (let i = 0; i < lines.length; i++) {
    cls.classifyLine(lines[i] ?? '', i, ctx);
  }
  return cls.count;
}
