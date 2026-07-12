/**
 * Roadmap #26 declaration sweep:
 * - Event declarations become `event` nodes and `RaiseEvent` can point to them.
 * - Type...End Type blocks become `type` + `type_member` nodes.
 * - Win32 API Declare statements become `declare` nodes, while still being
 *   cached by name so normal call-site scanning can emit `calls` edges.
 */
import { Node } from '../../types';
import { generateNodeId } from '../tree-sitter-helpers';
import { foldVisibility } from './text-utils';
import { VbaExtractorContext, VbaClassifier } from './context';

/** `[visibility] Event <Name>(...)` custom event declaration. */
const EVENT_DECL_RE =
  /^\s*((?:Public|Private|Friend)\s+)?Event\s+(\p{L}[\p{L}\p{N}_]*)\b/iu;

/** `[visibility] Type <Name>` user-defined type block start. */
const TYPE_START_RE =
  /^\s*((?:Public|Private|Friend)\s+)?Type\s+(\p{L}[\p{L}\p{N}_]*)\b/iu;

/** `End Type` user-defined type block end. */
const TYPE_END_RE = /^\s*End\s+Type\b/iu;

/** `<MemberName> As <Type>` inside a user-defined type block. */
const TYPE_MEMBER_RE =
  /^\s*(\p{L}[\p{L}\p{N}_]*)\s*(?:\([^)]*\))?\s+As\s+(.+?)\s*$/iu;

/** `[visibility] Declare [PtrSafe] Sub|Function <Name> Lib "dll" [Alias "x"] ...` */
const DLL_DECLARE_RE =
  /^\s*((?:Public|Private)\s+)?Declare\s+(PtrSafe\s+)?(Sub|Function)\s+(\p{L}[\p{L}\p{N}_]*)\s+Lib\s+"([^"]+)"(?:\s+Alias\s+"([^"]+)")?/iu;

/**
 * Issue #83: factory for the events/types/declares classifier. Closure
 * state: `currentType` (the open `Type ... End Type` block).
 */
export function createEventsTypesDeclaresClassifier(): VbaClassifier {
  let currentType: { id: string; name: string } | null = null;
  const cls: VbaClassifier = {
    name: 'eventsTypesDeclares',
    count: 0,
    classifyLine(line, i, ctx) {
      const lineNum = i + 1;

      if (currentType) {
        if (TYPE_END_RE.test(line)) {
          currentType = null;
          return;
        }
        const member = TYPE_MEMBER_RE.exec(line);
        if (member) {
          const memberName = member[1] ?? '';
          const memberType = (member[2] ?? '').trim();
          if (!memberName) return;
          const memberId = generateNodeId(ctx.filePath, 'type_member', memberName, lineNum);
          ctx.nodes.push({
            id: memberId,
            kind: 'type_member',
            name: memberName,
            qualifiedName: `${currentType.name}.${memberName}`,
            filePath: ctx.filePath,
            language: 'vba',
            startLine: lineNum,
            endLine: lineNum,
            startColumn: 0,
            endColumn: line.length,
            metadata: { memberType },
            updatedAt: Date.now(),
          });
          ctx.edges.push({
            source: currentType.id,
            target: memberId,
            kind: 'type-member',
            provenance: 'parser',
          });
        }
        return;
      }

      const eventDecl = EVENT_DECL_RE.exec(line);
      if (eventDecl) {
        const visibility = foldVisibility(eventDecl[1] ?? '');
        const name = eventDecl[2] ?? '';
        if (!name) return;
        const eventId = generateNodeId(ctx.filePath, 'event', name, lineNum);
        const eventNode: Node = {
          id: eventId,
          kind: 'event',
          name,
          signature: line.trim(),
          qualifiedName: ctx.classNamePrefix ? `${ctx.classNamePrefix}.${name}` : name,
          filePath: ctx.filePath,
          language: 'vba',
          startLine: lineNum,
          endLine: lineNum,
          startColumn: 0,
          endColumn: line.length,
          visibility,
          updatedAt: Date.now(),
        };
        ctx.nodes.push(eventNode);
        ctx.localEvents.set(name.toLowerCase(), eventNode);
        ctx.pushContainsFromModule(eventId);
        this.count++;
        return;
      }

      const typeStart = TYPE_START_RE.exec(line);
      if (typeStart) {
        const visibility = foldVisibility(typeStart[1] ?? '');
        const name = typeStart[2] ?? '';
        if (!name) return;
        const typeId = generateNodeId(ctx.filePath, 'type', name, lineNum);
        ctx.nodes.push({
          id: typeId,
          kind: 'type',
          name,
          qualifiedName: ctx.classNamePrefix ? `${ctx.classNamePrefix}.${name}` : name,
          filePath: ctx.filePath,
          language: 'vba',
          startLine: lineNum,
          endLine: lineNum,
          startColumn: 0,
          endColumn: line.length,
          visibility,
          updatedAt: Date.now(),
        });
        ctx.pushContainsFromModule(typeId);
        currentType = { id: typeId, name };
        this.count++;
        return;
      }

      const declaration = DLL_DECLARE_RE.exec(line);
      if (declaration) {
        const visibility = foldVisibility(declaration[1] ?? '');
        const ptrSafe = !!declaration[2];
        const declareKind = (declaration[3] ?? '').toLowerCase();
        const name = declaration[4] ?? '';
        const dll = declaration[5] ?? '';
        const aliasName = declaration[6] ?? undefined;
        if (!name) return;
        const declareId = generateNodeId(ctx.filePath, 'declare', name, lineNum);
        const declareNode: Node = {
          id: declareId,
          kind: 'declare',
          name,
          qualifiedName: ctx.classNamePrefix ? `${ctx.classNamePrefix}.${name}` : name,
          filePath: ctx.filePath,
          language: 'vba',
          startLine: lineNum,
          endLine: lineNum,
          startColumn: 0,
          endColumn: line.length,
          visibility,
          metadata: {
            isDeclare: true,
            dll,
            declareKind,
            ptrSafe,
            ...(aliasName ? { aliasName } : {}),
          },
          updatedAt: Date.now(),
        };
        ctx.nodes.push(declareNode);
        if (!ctx.functionNodeByName.has(name)) {
          ctx.functionNodeByName.set(name, declareNode);
        }
        ctx.pushContainsFromModule(declareId);
        this.count++;
      }
    },
  };
  return cls;
}

/**
 * Backward-compat wrapper (see procedures.ts). Returns the classifier's
 * `count` so the orchestrator can decide `hasAnySymbols`.
 */
export function sweepEventsTypesAndDeclares(ctx: VbaExtractorContext, src: string): number {
  const cls = createEventsTypesDeclaresClassifier();
  const lines = src.split('\n');
  for (let i = 0; i < lines.length; i++) {
    cls.classifyLine(lines[i] ?? '', i, ctx);
  }
  return cls.count;
}
