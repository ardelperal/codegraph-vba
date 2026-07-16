/**
 * Canonical Access form/report identity used by SourceObject and DoCmd targets.
 * Access object names are case-insensitive and code modules conventionally add
 * a `Form_` / `Report_` prefix that UI properties omit.
 */
export function normalizeAccessObjectName(name: string): string {
  return name
    .trim()
    .replace(/^(?:form|report)[._]/i, '')
    .toLocaleLowerCase('en-US');
}
