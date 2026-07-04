/**
 * Issue #53: encoding-robust source reading for VBA-family files.
 *
 * Access `SaveAsText` / VBIDE `Export` write Windows-1252 (ANSI). dysflow
 * converts everything to UTF-8-no-BOM on export, but codegraph-vba can be
 * pointed at exports from older dysflow runs, other tools, or hand-copied
 * modules that are still CP1252 on disk. Today every file is read as UTF-8
 * unconditionally; a CP1252 file with accented identifiers (ubiquitous in
 * Spanish-language Access code: `Sub ActualizarSituación`, enum member `Sí`,
 * module headers `' MÓDULO:`) decodes to replacement characters, `\p{L}`
 * regexes stop matching, and the symbol silently disappears from the graph.
 * A BOM-carrying UTF-8 file additionally corrupts the first line (breaks
 * `Attribute VB_Name` detection → wrong module name).
 *
 * Scope (do NOT change decoding for tree-sitter languages): only
 * `.bas`/`.cls`/`.frm`/`.dsr`/`.form.txt`/`.report.txt`/`.sql` files get
 * the BOM strip + CP1252 fallback; everything else stays byte-identical to
 * the existing `fs.readFile(path, 'utf-8')` path.
 */
import * as fs from 'fs';

/**
 * Family of extensions routed through Dysflow's VBA pipeline (or the
 * SqlQueryExtractor's saved-query parser). The matching is intentionally
 * case-insensitive because file systems vary (NTFS preserves case, ext4
 * on Linux through WSL may not, and Dysflow itself is rename-tolerant).
 */
const VBA_FAMILY_RE = /\.(?:bas|cls|frm|dsr|form\.txt|report\.txt|sql)$/i;

export function isVbaFamilyFile(filePath: string): boolean {
  return VBA_FAMILY_RE.test(filePath);
}

/**
 * UTF-8 BOM bytes (EF BB BF) — the WHATWG UTF-8 decoder consumes these
 * silently (no U+FEFF written), but `readFileSync(path, 'utf-8')` on
 * Node does NOT strip the BOM and emits `\uFEFF` as the first character,
 * which breaks `Attribute VB_Name` matching.
 */
const UTF8_BOM = [0xef, 0xbb, 0xbf] as const;

export interface ReadVbaSourceOptions {
  /**
   * Dependency injection for the byte read — defaults to `fs.readFileSync(path)`.
   * Tests inject in-memory buffers via this parameter so no temp files are
   * needed.
   */
  readFile?: (filePath: string) => Buffer;
  /**
   * Optional callback invoked when the CP1252 fallback fires. Receives the
   * file path and the underlying decoder error so the CLI can surface a
   * low-severity warning. Extraction continues silently otherwise.
   */
  onFallback?: (filePath: string, reason: string) => void;
}

export interface ReadVbaSourceResult {
  /**
   * Decoded source text with leading UTF-8 BOM stripped. Use this exactly
   * as you would use the result of `fs.readFileSync(path, 'utf-8')`.
   */
  text: string;
  /**
   * True iff a leading UTF-8 BOM was stripped from the input bytes.
   * Callers that care (e.g. post-processing that compares against a
   * hex-signatured reference) can read this.
   */
  bomStripped: boolean;
}

/**
 * Read a VBA-family source file with encoding-robust decode:
 *  1. Strip leading UTF-8 BOM (bytes EF BB BF) if present.
 *  2. Try UTF-8 with `fatal: true` — a single byte sequence like 0xF3 in
 *     the middle of `ActualizarSituación` (CP1252 'ó') throws a `TypeError`
 *     and we fall back to `windows-1252`.
 *  3. Return the decoded text.
 *
 * Performance: the BOM peek is 3 byte comparisons; the UTF-8 sniff is `O(fileSize)`
 * but only runs the decoder once on the happy path. CP1252 fallback only
 * fires when UTF-8 actually fails — valid UTF-8 (dysflow, modern editors,
 * anything written today) is byte-identical to `readFileSync(path, 'utf-8')`
 * with the BOM difference only. The windows-1252 decoder is built into
 * Node's WHATWG encodings spec — no extra dep.
 *
 * @param filePath absolute path to read.
 * @param opts optional dependencies; see {@link ReadVbaSourceOptions}.
 */
export function readVbaSource(
  filePath: string,
  opts: ReadVbaSourceOptions = {},
): ReadVbaSourceResult {
  const readFile = opts.readFile ?? ((p: string) => fs.readFileSync(p));
  const onFallback = opts.onFallback;
  const buf = readFile(filePath);

  // BOM strip: only at start, byte-exact. Encoding BOMs in the middle of a
  // file are not stripped (they'd be data, not marker).
  const bomStripped =
    buf.length >= 3 &&
    buf[0] === UTF8_BOM[0] &&
    buf[1] === UTF8_BOM[1] &&
    buf[2] === UTF8_BOM[2];
  const body = bomStripped ? buf.subarray(3) : buf;

  // Cheap UTF-8 sniff with the fatal flag — throws on invalid sequences.
  // Node's WHATWG TextDecoder supports this and gives us a deterministic
  // signal instead of silent replacement-char substitution.
  try {
    const text = new TextDecoder('utf-8', { fatal: true }).decode(body);
    return { text, bomStripped };
  } catch (err) {
    const reason = err instanceof Error ? err.message : String(err);
    // CP1252 covers the full Latin-1 range used by Access exports:
    // Spanish ('áéíóúñ¿¡'), French ('éèàç'), German ('äöüß'), Italian
    // ('àèéìòù'), Portuguese ('ãõç'), etc. — and isLosslessRecovery from
    // `new TextDecoder('windows-1252').decode(bytes)` for these byte
    // values, so this fallback turns every byte-for-byte CP1252 source into
    // a valid JavaScript string the extractor can regex.
    const text = new TextDecoder('windows-1252').decode(body);
    onFallback?.(filePath, reason);
    return { text, bomStripped };
  }
}

/**
 * String-level BOM strip — survives when the upstream `fsp.readFile(path, 'utf-8')`
 * already decoded the bytes (it's the only way the `\uFEFF` char survives in
 * a string). Used as a defensive last resort when a read site has NOT gone
 * through `readVbaSource` and the source text starts with the BOM marker.
 *
 * Most call sites should prefer `readVbaSource` for fresh reads; this is
 * a string post-process for callers that already have a `string` in hand.
 */
export function stripUtf8Bom(text: string): string {
  return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
}
