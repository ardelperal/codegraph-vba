import type { ProcInfo, VbaExtractorContext } from './context';

export const VBA_FILESYSTEM_STATEMENT_SYNTHESIZED_BY =
  'vba-filesystem-statement' as const;

export type VbaFilesystemOperation = 'kill' | 'open' | 'close';

const IDENT = String.raw`\p{L}[\p{L}\p{N}_]*`;
const FILE_NUMBER = String.raw`#?\s*(?:\d+|${IDENT})`;

const KILL_STATEMENT_RE = /^\s*Kill\s+(?!\s*\()([^,]+?)\s*$/iu;
const OPEN_STATEMENT_RE = new RegExp(
  String.raw`^\s*Open\s+(?!\s*\()(.+?)\s+For\s+(Append|Binary|Input|Output|Random)` +
    String.raw`(?:\s+Access\s+(?:Read|Write|Read\s+Write))?` +
    String.raw`(?:\s+Lock\s+(?:Read|Write|Read\s+Write|Shared))?` +
    String.raw`\s+As\s+${FILE_NUMBER}(?:\s+Len\s*=\s*.+)?\s*$`,
  'iu',
);
const CLOSE_STATEMENT_RE = new RegExp(
  String.raw`^\s*Close(?:\s+${FILE_NUMBER}(?:\s*,\s*${FILE_NUMBER})*)?\s*$`,
  'iu',
);

/**
 * Parse only the intrinsic grammar-shaped statement forms. The input must be a
 * single string-masked statement clause, so quoted keywords remain opaque while
 * valid string operands retain their token shape.
 */
export function detectVbaFilesystemStatement(
  statement: string,
): VbaFilesystemOperation | null {
  if (KILL_STATEMENT_RE.test(statement)) return 'kill';
  if (OPEN_STATEMENT_RE.test(statement)) return 'open';
  if (CLOSE_STATEMENT_RE.test(statement)) return 'close';
  return null;
}

/**
 * The shared provenance gate used by both resolution and runtime
 * classification. Name, kind, language, family, and operation must agree; a
 * copied or partially forged metadata stamp is rejected.
 */
export function isExactVbaFilesystemStatementReference(ref: {
  language?: string | null;
  referenceKind?: string | null;
  referenceName?: string | null;
  metadata?: Record<string, unknown> | null;
}): boolean {
  if ((ref.language ?? '').toLowerCase() !== 'vba') return false;
  if (ref.referenceKind !== 'calls') return false;

  const operation = (ref.referenceName ?? '').toLowerCase();
  if (operation !== 'kill' && operation !== 'open' && operation !== 'close') {
    return false;
  }

  return (
    ref.metadata?.synthesizedBy === VBA_FILESYSTEM_STATEMENT_SYNTHESIZED_BY &&
    ref.metadata.runtimeFamily === 'filesystem' &&
    ref.metadata.operation === operation
  );
}

/** Emit one typed unresolved runtime reference from the containing procedure. */
export function emitVbaFilesystemStatementReference(
  ctx: VbaExtractorContext,
  caller: ProcInfo,
  statement: string,
  lineNum: number,
): boolean {
  const operation = detectVbaFilesystemStatement(statement);
  if (!operation) return false;

  const referenceName = operation[0]!.toUpperCase() + operation.slice(1);
  ctx.unresolvedReferences.push({
    fromNodeId: ctx.findOrCreateFunctionNodeId(caller),
    referenceName,
    referenceKind: 'calls',
    line: lineNum,
    column: Math.max(0, statement.search(/\S/)),
    filePath: ctx.filePath,
    language: 'vba',
    metadata: {
      synthesizedBy: VBA_FILESYSTEM_STATEMENT_SYNTHESIZED_BY,
      runtimeFamily: 'filesystem',
      operation,
    },
  });
  return true;
}
