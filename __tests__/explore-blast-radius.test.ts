/**
 * codegraph_explore blast-radius section.
 *
 * explore now appends a compact, always-on "Blast radius" for the entry
 * symbols: who depends on each (locations only — no source) and which test
 * files cover it, so the agent knows what to update/verify before editing
 * without a separate impact call. Symbols with no dependents are skipped, and
 * the section is omitted entirely when nothing qualifies.
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import CodeGraph from '../src/index';
import { ToolHandler } from '../src/mcp/tools';

describe('codegraph_explore — blast radius', () => {
  let testDir: string;
  let cg: CodeGraph;
  let handler: ToolHandler;

  beforeEach(async () => {
    testDir = fs.mkdtempSync(path.join(os.tmpdir(), 'codegraph-blast-'));
    const src = path.join(testDir, 'src');
    fs.mkdirSync(src, { recursive: true });

    // `target` is depended on by a sibling (caller) and a test file.
    fs.writeFileSync(
      path.join(src, 'feature.ts'),
      `export function target() { return 1; }\n` +
      `export function caller() { return target(); }\n`,
    );
    fs.writeFileSync(
      path.join(src, 'feature.test.ts'),
      `import { target } from './feature';\n` +
      `export function checkTarget() { return target(); }\n`,
    );
    // A leaf with no dependents — must NOT show up in the blast radius.
    fs.writeFileSync(
      path.join(src, 'leaf.ts'),
      `export function lonelyLeaf() { return 42; }\n`,
    );

    cg = CodeGraph.initSync(testDir, { config: { include: ['**/*.ts'], exclude: [] } });
    await cg.indexAll();
    handler = new ToolHandler(cg);
  });

  afterEach(() => {
    if (cg) cg.destroy();
    if (fs.existsSync(testDir)) fs.rmSync(testDir, { recursive: true, force: true });
  });

  it('lists dependents (locations only) and covering tests for an entry symbol', async () => {
    const res = await handler.execute('codegraph_explore', { query: 'target' });
    const text = res.content[0].text;

    expect(text).toContain('**Blast radius');
    expect(text).toContain('`target`');
    expect(text).toMatch(/caller/); // a caller count is reported
    // It names WHERE (the caller file) — not the caller's source body.
    expect(text).toContain('feature.ts');
    // Test coverage is surfaced (either the covering test file, or the warning).
    expect(text).toMatch(/tests:.*feature\.test\.ts|no covering tests/);
  });

  it('omits symbols that have no dependents from the blast radius', async () => {
    const res = await handler.execute('codegraph_explore', { query: 'lonelyLeaf' });
    const text = res.content[0].text;
    // lonelyLeaf has zero callers — it must never appear under a blast-radius bullet.
    expect(text).not.toMatch(/Blast radius[\s\S]*`lonelyLeaf`/);
  });

  it('resolves VBA reference stubs so test-file callers appear in blast radius', async () => {
    // VBA fixture: Module1 defines an enum, Test_Module1 references it via Dim.
    const vbaDir = fs.mkdtempSync(path.join(os.tmpdir(), 'codegraph-blast-vba-'));
    const vbaSrc = path.join(vbaDir, 'src');
    fs.mkdirSync(vbaSrc, { recursive: true });

    fs.writeFileSync(
      path.join(vbaSrc, 'Module1.bas'),
      `Option Explicit\n\nPublic Enum MyEnum\n    Foo = 1\n    Bar = 2\nEnd Enum\n\nPublic Sub DoStuff()\n    Dim x As Long\n    x = 5\nEnd Sub\n`,
    );
    fs.writeFileSync(
      path.join(vbaSrc, 'Test_Module1.bas'),
      `Option Explicit\n\nPublic Sub Test_DoStuff()\n    Dim result As MyEnum\n    result = Foo\n    DoStuff\nEnd Sub\n`,
    );

    const vbaCg = CodeGraph.initSync(vbaDir, { config: { include: ['**/*.bas', '**/*.cls'], exclude: [] } });
    await vbaCg.indexAll();
    const vbaHandler = new ToolHandler(vbaCg);

    try {
      const res = await vbaHandler.execute('codegraph_explore', { query: 'MyEnum' });
      const text = res.content[0].text;

      // The blast radius must show the test file as a covering test.
      expect(text).toContain('**Blast radius');
      expect(text).toContain('`MyEnum`');
      expect(text).toMatch(/tests:.*Test_Module1\.bas/);
      // Must NOT show the false "no covering tests" warning.
      expect(text).not.toMatch(/no covering tests/);
    } finally {
      vbaCg.destroy();
      fs.rmSync(vbaDir, { recursive: true, force: true });
    }
  });
});
