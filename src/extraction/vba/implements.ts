/**
 * `Implements IFoo` sweep (REQ-CODE-5). Emits one `interface` node per
 * `Implements` declaration plus a pending `implements` edge whose source is
 * rewired to the module/class node once the orchestrator creates it.
 */
import { Edge } from '../../types';
import { generateNodeId } from '../tree-sitter-helpers';
import { VbaExtractorContext } from './context';

/** Implements regex. */
const IMPLEMENTS_RE = /^\s*Implements\s+(\p{L}[\p{L}\p{N}_]*)/iu;

export function sweepImplements(ctx: VbaExtractorContext, src: string): number {
  const lines = src.split('\n');
  let count = 0;
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i] ?? '';
    const m = IMPLEMENTS_RE.exec(line);
    if (!m) continue;
    const name = m[1] ?? '';
    if (!name) continue;
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
    count++;
  }
  return count;
}
