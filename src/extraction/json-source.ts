/**
 * Decoding helpers for the JSON artifacts Dysflow writes beside a VBA export
 * (`tests(.<slice>)*.json` test manifests, `sequences/*.json` test sequences).
 *
 * These are the only indexed files that are parsed as JSON rather than scanned
 * as text, and they are deliberately NOT part of `isVbaFamilyFile` — JSON is
 * UTF-8 by specification, so the CP1252 fallback in `readVbaSource` would turn
 * a mis-encoded manifest into silently wrong data instead of a loud parse
 * error. What they do share with the rest of the VBA family is the byte-order
 * mark.
 */

/**
 * Strip a leading UTF-8 BOM from already-decoded JSON text.
 *
 * A `.json` is read with a plain `fs.readFile(path, 'utf-8')`, and Node keeps
 * the BOM as a leading `\uFEFF` rather than consuming it. `JSON.parse` rejects
 * that character outright, so a BOM-carrying manifest used to contribute
 * nothing to the graph but a `warning`-severity parse error — every `Test_*`
 * link it declared silently vanished (issue #316). On Windows this is the
 * common case, not an exotic one: PowerShell 5.1 writes the BOM by default for
 * `-Encoding UTF8` in both `Set-Content` and `Out-File`.
 *
 * Only a BOM at position 0 is removed; a `\uFEFF` anywhere else is data and is
 * left for `JSON.parse` to judge.
 */
export function stripJsonBom(text: string): string {
  return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
}
