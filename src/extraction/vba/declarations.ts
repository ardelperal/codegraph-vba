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
import { VbaClassifier } from './context';
import { defineRule, matchRule, VbaExtractionRule } from './rules';

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
 * Issue #153: the declarative rule table for the events/types/declares
 * concern. Five rules, partitioned by where they fire:
 *
 *  - `event-decl`   (outside-type-block) — match an `Event <Name>` decl
 *  - `type-start`   (outside-type-block) — match a `Type <Name>` header;
 *                  sets `ctx.vbaDeclTypeBlock` on a successful match
 *  - `type-end`     (inside-type-block)  — match `End Type`; clears
 *                  `ctx.vbaDeclTypeBlock` on a successful match
 *  - `type-member`  (inside-type-block)  — match `<Member> As <Type>`
 *                  inside an open type block
 *  - `dll-declare`  (outside-type-block) — match a `[Private|Public]
 *                  Declare [PtrSafe] Sub|Function <Name> Lib "<dll>"
 *                  [Alias "<x>"]` Win32 API declaration
 *
 * The `requires` field encodes the type-block precondition. The
 * dispatcher (the factory's `classifyLine`) filters rules whose
 * precondition does not hold, so an `End Type` line never
 * accidentally re-enters the `event-decl` / `dll-declare` paths.
 *
 * Inter-line state lives on `ctx.vbaDeclTypeBlock` (replaces the
 * pre-#153 factory closure variable) so the emit functions can
 * read/write it without taking a mutable closure reference.
 */
export const RULES: readonly VbaExtractionRule<unknown>[] = [
  defineRule({
    id: 'event-decl',
    description:
      'Match an `[Public|Private|Friend] Event <Name>(...)` custom event declaration; emit an `event` node + `contains` edge from the module/class.',
    pattern: EVENT_DECL_RE,
    requires: 'outside-type-block',
    emit: (m, ctx, line, lineNum) => {
      const visibility = foldVisibility(m[1] ?? '');
      const name = m[2] ?? '';
      if (!name) return null;
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
      return { name };
    },
  }),
  defineRule({
    id: 'type-start',
    description:
      'Match a `[Public|Private|Friend] Type <Name>` UDT block header; emit a `type` node + `contains` edge and mark `ctx.vbaDeclTypeBlock` so subsequent member lines route to `type-member` / `End Type` lines route to `type-end`.',
    pattern: TYPE_START_RE,
    requires: 'outside-type-block',
    emit: (m, ctx, line, lineNum) => {
      const visibility = foldVisibility(m[1] ?? '');
      const name = m[2] ?? '';
      if (!name) return null;
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
      // Mark the open type block on ctx so the next non-End line is
      // dispatched to `type-member` / `type-end` instead of the
      // outside-type-block rules.
      ctx.vbaDeclTypeBlock = { id: typeId, name };
      return { name, typeId };
    },
  }),
  defineRule({
    id: 'type-end',
    description:
      'Match `End Type` while inside a type block; emit nothing and clear `ctx.vbaDeclTypeBlock` so subsequent lines route back to the outside-type-block rules.',
    pattern: TYPE_END_RE,
    requires: 'inside-type-block',
    emit: (_m, ctx) => {
      ctx.vbaDeclTypeBlock = null;
      return { kind: 'end-type-block' as const };
    },
  }),
  defineRule({
    id: 'type-member',
    description:
      'Match `<MemberName> [As <Type>]` inside an open type block; emit a `type_member` node + a `type-member` edge from the open type to the member.',
    pattern: TYPE_MEMBER_RE,
    requires: 'inside-type-block',
    emit: (m, ctx, line, lineNum) => {
      const memberName = m[1] ?? '';
      const memberType = (m[2] ?? '').trim();
      if (!memberName || !ctx.vbaDeclTypeBlock) return null;
      const memberId = generateNodeId(ctx.filePath, 'type_member', memberName, lineNum);
      ctx.nodes.push({
        id: memberId,
        kind: 'type_member',
        name: memberName,
        qualifiedName: `${ctx.vbaDeclTypeBlock.name}.${memberName}`,
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
        source: ctx.vbaDeclTypeBlock.id,
        target: memberId,
        kind: 'type-member',
        provenance: 'parser',
      });
      return { name: memberName };
    },
  }),
  defineRule({
    id: 'dll-declare',
    description:
      'Match a `[Public|Private] Declare [PtrSafe] Sub|Function <Name> Lib "<dll>" [Alias "<x>"]` Win32 API declaration; emit a `declare` node + `contains` edge and register the name in `functionNodeByName` so call-site scanning can resolve it.',
    pattern: DLL_DECLARE_RE,
    requires: 'outside-type-block',
    emit: (m, ctx, line, lineNum) => {
      const visibility = foldVisibility(m[1] ?? '');
      const ptrSafe = !!m[2];
      const declareKind = (m[3] ?? '').toLowerCase();
      const name = m[4] ?? '';
      const dll = m[5] ?? '';
      const aliasName = m[6] ?? undefined;
      if (!name) return null;
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
      return { name };
    },
  }),
];

/**
 * Issue #83: factory for the events/types/declares classifier.
 *
 * Inter-line state (`ctx.vbaDeclTypeBlock`) lives on `ctx` — the
 * declarative RULES table's `emit` functions read/write it directly.
 * The dispatcher below honours each rule's `requires` precondition
 * so the per-line dispatch mirrors the legacy cascade exactly:
 *
 *   if (currentType) {
 *     try type-end, then type-member; return early
 *   } else {
 *     try event-decl, type-start, dll-declare
 *   }
 */
export function createEventsTypesDeclaresClassifier(): VbaClassifier {
  const cls: VbaClassifier = {
    name: 'eventsTypesDeclares',
    count: 0,
    classifyLine(line, i, ctx) {
      const lineNum = i + 1;
      for (const rule of RULES) {
        // Gate by the rule's `requires` precondition. The
        // `outside-type-block` / `inside-type-block` values are the
        // concrete preconditions the dispatcher knows about today;
        // unknown strings are treated as "always" so adding new
        // preconditions does not require touching this loop.
        if (rule.requires === 'inside-type-block' && !ctx.vbaDeclTypeBlock) continue;
        if (rule.requires === 'outside-type-block' && ctx.vbaDeclTypeBlock) continue;
        const m = matchRule(rule.pattern, line);
        if (!m) continue;
        const result = rule.emit(m, ctx, line, lineNum);
        if (result !== null && result !== undefined) {
          this.count += rule.count ? rule.count(result as never) : 1;
        }
        // `type-end` cleared the type block — nothing else on this
        // line can apply (a VBA `End Type` line cannot also be an
        // `Event <Name>` declaration). The legacy cascade did the
        // same short-circuit via `return`. Break out of the loop
        // explicitly so a later refactor cannot accidentally let an
        // outside-type-block rule fire on the same line.
        if (rule.id === 'type-end') break;
      }
    },
  };
  return cls;
}
