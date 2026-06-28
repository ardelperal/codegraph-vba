/**
 * detectVbaFormFile() — extension-detection tests.
 *
 * The helper recognizes two-segment extensions `.form.txt` and
 * `.report.txt` (Dysflow SaveAsText output for Access form/report UI).
 * A naive `path.extname()` returns `.txt`, which would falsely classify
 * every text file as VBA; this helper is the gate that fixes it.
 *
 * IMPORTANT: path.extname() on Windows preserves leading separators as
 * part of the basename — only `lastIndexOf('.')` over the basename is
 * safe for two-segment detection.
 */
import { describe, it, expect } from 'vitest';
import { detectVbaFormFile } from '../src/extraction/grammars';

describe('detectVbaFormFile', () => {
  it('returns true for a forward-slash path ending in .form.txt', () => {
    expect(detectVbaFormFile('src/forms/Form_Main.form.txt')).toBe(true);
  });

  it('returns true for a forward-slash path ending in .report.txt', () => {
    expect(detectVbaFormFile('src/reports/Report_Orders.report.txt')).toBe(true);
  });

  it('returns false for a single-segment .form extension', () => {
    // .form (no .txt) is not a Dysflow SaveAsText file
    expect(detectVbaFormFile('src/forms/Form_Main.form')).toBe(false);
  });

  it('returns false for a plain .txt document', () => {
    // Critical regression trap: a naive `path.extname()` collapses to .txt
    expect(detectVbaFormFile('notes/Document.txt')).toBe(false);
  });

  it('returns true for a backslash path ending in .form.txt', () => {
    // Windows path separators must be honored
    expect(detectVbaFormFile('src\\forms\\Form_Main.form.txt')).toBe(true);
  });

  it('returns true for a backslash path ending in .report.txt', () => {
    expect(detectVbaFormFile('src\\reports\\Report_Orders.report.txt')).toBe(true);
  });

  it('returns false for an unrelated extension', () => {
    expect(detectVbaFormFile('src/modules/modHelpers.bas')).toBe(false);
    expect(detectVbaFormFile('Form_Main.cls')).toBe(false);
  });

  it('returns false for an empty string', () => {
    expect(detectVbaFormFile('')).toBe(false);
  });
});