# Exploration: VBA API Declarations and Extraction Improvements

This document explores the architectural design, implementation details, and technical considerations for enhancing the VBA extraction parser (`vba-extractor.ts`, `vba-preprocess.ts`) to support external DLL declarations, conditional compilation preprocessing, customized DB variable SQL execution, and constant resolution in `DoCmd.OpenForm` calls.

---

## 1. External DLL Declarations (`Declare` Statements)

### Current Limitations
Currently, `vba-extractor.ts` ignores `Declare` (and `Declare PtrSafe`) statements, treating them as compiler hints/bindings rather than user procedures. This leaves external Win32 / DLL functions completely absent from the CodeGraph, so calls to them (e.g., `Sleep 1000`) cannot be resolved to any local declaration node.

### Proposed Solution
DLL declarations act as local procedure declarations within a module/class. We will extract them as `function` nodes and insert them into the local procedure list so that call site matching works seamlessly.

#### Syntax & Regex Design
VBA DLL declarations follow this syntax:
```vba
[Public | Private] Declare [PtrSafe] Sub|Function name Lib "libname" [Alias "aliasname"] [([arglist])] [As type]
```
Since line continuations are already joined by the preprocessor, we can match these on a single physical line using the following regex:
```typescript
private static readonly DLL_DECLARE_RE =
  /^\s*((?:Public|Private)\s+)?Declare\s+(?:PtrSafe\s+)?(Sub|Function)\s+(\p{L}[\p{L}\p{N}_]*)\s+Lib\s+/iu;
```
* **Group 1**: Visibility prefix (optional, default to `Public` at module level, `Private` at class level).
* **Group 2**: Procedure kind (`Sub` or `Function`).
* **Group 3**: Function/Procedure name.

#### Mapping to AST Nodes
When a match is found during `sweepProcedures`, we create a `ProcInfo` object and a `Node` of kind `'function'`:
* `visibility`: Normalized to `'public'` or `'private'`.
* `kind`: `'sub'` or `'function'`.
* `startLine`: `lineNum`.
* `endLine`: `lineNum` (single-line declaration, so no `End Sub` block exists).
* `metadata`: `{ isDeclare: true }` (to distinguish DLL imports from user-defined procedures).

We add the `ProcInfo` to `this.localProcs` so that call site scans successfully resolve references to this DLL function.

---

## 2. Conditional Compilation (`#If ... #Else`)

### Current Limitations
When a module contains both standard (32-bit) and `PtrSafe` (64-bit) declarations of the same external DLL function under `#If ... #Else` blocks, scanning both blocks emits duplicate function nodes and breaks lookup logic (which relies on unique names).

### Proposed Solution
Introduce a conditional compilation preprocessor in `vba-preprocess.ts` that evaluates compilation directives and filters out inactive branches *before* any extraction sweeps are performed.

#### Preprocessor Algorithm
To maintain absolute line-number alignment (critical for indexing and node ranges), we replace discarded branch lines and the directive lines themselves with empty lines (`""`).

We evaluate conditions using standard modern Office defaults:
* `VBA7 = true`
* `Win64 = true`
* `Mac = false`

```typescript
export function preprocessConditionalCompilation(src: string): string {
  if (!src) return src;
  const lines = src.split('\n');
  const out: string[] = [];
  
  interface ConditionFrame {
    active: boolean;
    parentActive: boolean;
    trueBranchTaken: boolean;
  }
  const stack: ConditionFrame[] = [];
  
  const IF_DIR = /^\s*#If\s+(.+?)\s+Then\s*$/i;
  const ELSEIF_DIR = /^\s*#ElseIf\s+(.+?)\s+Then\s*$/i;
  const ELSE_DIR = /^\s*#Else\s*$/i;
  const ENDIF_DIR = /^\s*#(?:End\s*If|EndIf)\s*$/i;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i] ?? '';
    
    // Directive matching
    const ifMatch = IF_DIR.exec(line);
    if (ifMatch) {
      const parentActive = stack.length > 0 ? stack[stack.length - 1]!.active : true;
      const condVal = evaluateCondition(ifMatch[1] ?? '');
      const active = parentActive && condVal;
      stack.push({ active, parentActive, trueBranchTaken: active });
      out.push('');
      continue;
    }
    
    const elseIfMatch = ELSEIF_DIR.exec(line);
    if (elseIfMatch) {
      if (stack.length > 0) {
        const frame = stack[stack.length - 1]!;
        const parentActive = frame.parentActive;
        const condVal = evaluateCondition(elseIfMatch[1] ?? '');
        const active = parentActive && !frame.trueBranchTaken && condVal;
        frame.active = active;
        if (active) frame.trueBranchTaken = true;
      }
      out.push('');
      continue;
    }
    
    const elseMatch = ELSE_DIR.exec(line);
    if (elseMatch) {
      if (stack.length > 0) {
        const frame = stack[stack.length - 1]!;
        const parentActive = frame.parentActive;
        const active = parentActive && !frame.trueBranchTaken;
        frame.active = active;
        if (active) frame.trueBranchTaken = true;
      }
      out.push('');
      continue;
    }
    
    const endIfMatch = ENDIF_DIR.exec(line);
    if (endIfMatch) {
      stack.pop();
      out.push('');
      continue;
    }
    
    const isCurrentActive = stack.length > 0 ? stack[stack.length - 1]!.active : true;
    out.push(isCurrentActive ? line : '');
  }
  
  return out.join('\n');
}

function evaluateCondition(expr: string): boolean {
  let normalized = expr.toLowerCase().trim();
  normalized = normalized.replace(/\bvba7\b/g, 'true');
  normalized = normalized.replace(/\bwin64\b/g, 'true');
  normalized = normalized.replace(/\bmac\b/g, 'false');
  normalized = normalized.replace(/\band\b/g, '&&');
  normalized = normalized.replace(/\bor\b/g, '||');
  normalized = normalized.replace(/\bnot\b/g, '!');
  
  const safeChars = /^[truefals&\(\)\|\!\s]+$/;
  if (!safeChars.test(normalized)) return false;
  
  try {
    return Function(`return (${normalized});`)();
  } catch {
    return false;
  }
}
```

This preprocessor is integrated into `VbaExtractor.extract()`:
```typescript
const joined = joinLineContinuations(this.source);
const preprocessed = preprocessConditionalCompilation(joined);
const uncommented = stripVbaComments(preprocessed);
```
By resolving directives first, all downstream sweeps only see the correct target branch, solving symbol collisions at the root level.

---

## 3. Customized DB Variable SQL Extraction

### Current Limitations
Currently, SQL extraction logic is hardcoded to recognize execution calls specifically against `db`, `getdb()`, or `CurrentDb`. Real-world codebases utilize customized names (such as parameters prefixing `p_db` or module variables `m_Db`).

### Proposed Solution
Modify SQL regex patterns to support arbitrary variable identifiers ending with `db` (case-insensitive), covering standard conventions (like `p_db`, `m_Db`, `g_db`, `db`).

#### Pattern Modifications
1. **`SQL_VAR_EXEC_RE`**:
   Change `(?:getdb\(\)|CurrentDb|db)` to match any identifier ending in `db`:
   ```typescript
   private static readonly SQL_VAR_EXEC_RE =
     /\b(?:getdb\(\)|CurrentDb|(?:\p{L}[\p{L}\p{N}_]*)?db)\.(?:OpenRecordset|Execute)\s*\(?\s*(\p{L}[\p{L}\p{N}_]*)\s*\)?/giu;
   ```

2. **`SQL_WRAPPERS`**:
   Update `SQL_WRAPPERS` to include generic prefix-optional `db` matches at the end of the matching chain (so specific wrappers like `getdb()` are evaluated first):
   ```typescript
   private static readonly SQL_WRAPPERS: ReadonlyArray<{ name: string; re: RegExp }> = [
     { name: 'DoCmd.RunSQL', re: /\bDoCmd\.RunSQL\s+"((?:[^"]|"")*)"/gi },
     { name: 'CurrentDb.OpenRecordset', re: /\bCurrentDb\.OpenRecordset\s+"((?:[^"]|"")*)"/gi },
     { name: 'CurrentDb.Execute', re: /\bCurrentDb\.Execute\s+"((?:[^"]|"")*)"/gi },
     { name: 'getdb().Execute', re: /\bgetdb\(\)\.Execute\s+"((?:[^"]|"")*)"/gi },
     { name: 'getdb().OpenRecordset', re: /\bgetdb\(\)\.OpenRecordset\s+"((?:[^"]|"")*)"/gi },
     // Custom db variables (e.g. p_db, m_Db, db)
     { name: 'db.OpenRecordset', re: /\b(?:\p{L}[\p{L}\p{N}_]*)?db\.OpenRecordset\s+"((?:[^"]|"")*)"/giu },
     { name: 'db.Execute', re: /\b(?:\p{L}[\p{L}\p{N}_]*)?db\.Execute\s+"((?:[^"]|"")*)"/giu },
   ];
   ```

---

## 4. Constant Resolution inside `DoCmd.OpenForm`

### Current Limitations
`DoCmd.OpenForm` modeling captures only literal string forms (e.g., `DoCmd.OpenForm "FormName"`). Variable-form calls (e.g., `DoCmd.OpenForm FORM_NAME_CONST`) are ignored.

### Proposed Solution
Modify `DoCmd.OpenForm` syntax parsing to match both string literals and identifier arguments. Track constant declarations in the file, and resolve the identifier argument to its literal value when possible.

#### Syntax & Regex Design
We extend `OPEN_FORM_RE` to match either a quoted literal (group 1) or an identifier (group 2):
```typescript
private static readonly OPEN_FORM_RE =
  /\bDoCmd\.OpenForm\s+(?:"([^"]+)"|(\p{L}[\p{L}\p{N}_]*))/gu;
```

#### Constant Tracking
Add `private localConstants = new Map<string, string>()` to the `VbaExtractor` class. In `sweepEnumsAndConsts`, extract values of declared constants:
```typescript
const constDecl = VbaExtractor.CONST_DECL_RE.exec(line);
if (constDecl) {
  const body = constDecl[2] ?? '';
  const localConstRe = /(\p{L}[\p{L}\p{N}_]*)\s*(?:As\s+[\p{L}\p{N}_.]+\s*)?=\s*(?:"((?:[^"]|"")*)"|([^,]+))/giu;
  let cm: RegExpExecArray | null;
  while ((cm = localConstRe.exec(body)) !== null) {
    const constName = cm[1] ?? '';
    if (!constName) continue;
    
    let constValue = '';
    if (cm[2] !== undefined) {
      constValue = cm[2].replace(/""/g, '"');
    } else if (cm[3] !== undefined) {
      constValue = cm[3].trim();
    }
    
    this.localConstants.set(constName.toLowerCase(), constValue);
    // Continue standard node generation...
  }
}
```

Since `sweepEnumsAndConsts` runs before `sweepCallsAndSql`, `localConstants` will be fully populated when call scanning begins.

#### Call Site Resolution
In `scanOpenFormCalls`:
```typescript
private scanOpenFormCalls(line: string, caller: ProcInfo, lineNum: number): void {
  const localRe = new RegExp(
    VbaExtractor.OPEN_FORM_RE.source,
    VbaExtractor.OPEN_FORM_RE.flags,
  );
  let m: RegExpExecArray | null;
  while ((m = localRe.exec(line)) !== null) {
    let targetFormName = '';
    if (m[1] !== undefined) {
      targetFormName = m[1].trim();
    } else if (m[2] !== undefined) {
      const constName = m[2].trim();
      const resolved = this.localConstants.get(constName.toLowerCase());
      if (resolved !== undefined) {
        targetFormName = resolved.trim();
      } else {
        // Fall back to using the constant name itself as a target stub
        targetFormName = constName;
      }
    }
    if (!targetFormName) continue;
    this.emitOpensFormEdge(caller, targetFormName, lineNum, m.index);
  }
}
```

This guarantees that local constant definitions resolve to their form name targets, and unresolved global constants still emit an `opens-form` edge to a stub named after the constant itself.
