/**
 * Access event properties shared by form-layout expression extraction and
 * code-behind naming-convention validation. Keep this as the single source of
 * truth when adding events. Reference: Microsoft Access object event docs:
 * https://learn.microsoft.com/en-us/office/vba/api/overview/access/object-model
 */
export const ACCESS_EVENT_PROPERTIES = new Map<string, string>([
  ['onclick', 'Click'],
  ['ondblclick', 'DblClick'],
  ['oncurrent', 'Current'],
  ['onload', 'Load'],
  ['onopen', 'Open'],
  ['onclose', 'Close'],
  ['onactivate', 'Activate'],
  ['ondeactivate', 'Deactivate'],
  ['onunload', 'Unload'],
  ['afterupdate', 'AfterUpdate'],
  ['beforeupdate', 'BeforeUpdate'],
  ['onchange', 'Change'],
  ['onenter', 'Enter'],
  ['onexit', 'Exit'],
  ['ongotfocus', 'GotFocus'],
  ['onlostfocus', 'LostFocus'],
  ['onmousedown', 'MouseDown'],
  ['onmousemove', 'MouseMove'],
  ['onmouseup', 'MouseUp'],
  ['onkeydown', 'KeyDown'],
  ['onkeypress', 'KeyPress'],
  ['onkeyup', 'KeyUp'],
  ['ontimer', 'Timer'],
  ['onnotinlist', 'NotInList'],
  ['ondirty', 'Dirty'],
  ['onundo', 'Undo'],
  ['ondelete', 'Delete'],
  ['beforeinsert', 'BeforeInsert'],
  ['afterinsert', 'AfterInsert'],
  ['onerror', 'Error'],
  ['onfilter', 'Filter'],
  ['onapplyfilter', 'ApplyFilter'],
  ['onresize', 'Resize'],
  ['onnodata', 'NoData'],
  ['onformat', 'Format'],
  ['onprint', 'Print'],
  ['onpage', 'Page'],
  ['onretreat', 'Retreat'],
]);

/**
 * Closed set of Access event procedure suffixes. Values from form/report
 * event properties above are combined with control-specific events whose
 * property is not necessarily present in an exported layout. VBA matching is
 * case-insensitive; consumers should compare lower-cased names.
 */
export const ACCESS_EVENT_NAMES: ReadonlySet<string> = new Set([
  ...ACCESS_EVENT_PROPERTIES.values(),
  'Updated',
  'DropButtonClick',
  'BeforeDelConfirm',
  'AfterDelConfirm',
  'MouseWheel',
  'BeforeDragOver',
  'BeforeDropOrPaste',
  'BeforePrint',
  'BeforeRender',
  'AfterRender',
  'AfterLayout',
  'AfterFinalRender',
  'BeforeQuery',
  'BeforeScreenTip',
  'CommandBeforeExecute',
  'CommandChecked',
  'CommandEnabled',
  'CommandExecute',
  'DataChange',
  'DataSetChange',
  'OnConnect',
  'OnDisconnect',
  'PivotTableChange',
  'Query',
  'SelectionChange',
  'ViewChange',
]);

const ACCESS_EVENT_NAMES_CASE_INSENSITIVE = new Set(
  [...ACCESS_EVENT_NAMES].map((name) => name.toLowerCase()),
);

export function isAccessEventName(name: string): boolean {
  return ACCESS_EVENT_NAMES_CASE_INSENSITIVE.has(name.toLowerCase());
}
