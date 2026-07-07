import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as os from 'node:os';
import { loadVbaConfig, clearProjectConfigCache } from '../src/project-config';
import CodeGraph from '../src/index';

describe('VBA targets config loader', () => {
  let dir: string;

  beforeEach(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'cg-vba-config-'));
    clearProjectConfigCache();
  });

  afterEach(() => {
    clearProjectConfigCache();
    fs.rmSync(dir, { recursive: true, force: true });
  });

  const writeConfig = (obj: unknown) =>
    fs.writeFileSync(
      path.join(dir, 'codegraph.json'),
      typeof obj === 'string' ? obj : JSON.stringify(obj)
    );

  it('returns empty targets when there is no codegraph.json (the default)', () => {
    expect(loadVbaConfig(dir)).toEqual({});
  });

  it('loads a well-formed vba.targets block', () => {
    writeConfig({
      vba: {
        targets: {
          VBA7: true,
          Win64: false,
          Win32: true,
        },
      },
    });
    expect(loadVbaConfig(dir)).toEqual({
      targets: {
        VBA7: true,
        Win64: false,
        Win32: true,
      },
    });
  });

  it('ignores invalid types under targets (warns and skips individual keys)', () => {
    writeConfig({
      vba: {
        targets: {
          VBA7: true,
          Win64: 'not-a-boolean',
          Win32: false,
          Mac: 42,
        },
      },
    });
    expect(loadVbaConfig(dir)).toEqual({
      targets: {
        VBA7: true,
        Win32: false,
      },
    });
  });

  it('ignores a non-object vba or targets value without throwing', () => {
    writeConfig({
      vba: 'should-be-an-object',
    });
    expect(loadVbaConfig(dir)).toEqual({});

    clearProjectConfigCache();
    writeConfig({
      vba: {
        targets: 'should-be-an-object',
      },
    });
    expect(loadVbaConfig(dir)).toEqual({});
  });

  it('merges local and root project config where root overrides local', () => {
    // Write local config
    const localDir = path.join(dir, '.codegraph-vba');
    fs.mkdirSync(localDir, { recursive: true });
    fs.writeFileSync(
      path.join(localDir, 'config.json'),
      JSON.stringify({
        vba: {
          targets: {
            VBA7: false,
            Win64: true,
            Mac: true,
          },
        },
      })
    );

    // Only local config exists initially
    expect(loadVbaConfig(dir)).toEqual({
      targets: {
        VBA7: false,
        Win64: true,
        Mac: true,
      },
    });

    // Write root config to override
    writeConfig({
      vba: {
        targets: {
          VBA7: true,
          Win64: false,
        },
      },
    });

    // Merged: root overrides local where conflict; other keys preserved
    expect(loadVbaConfig(dir)).toEqual({
      targets: {
        VBA7: true,
        Win64: false,
        Mac: true,
      },
    });
  });
});

describe('VBA targets config end-to-end integration', () => {
  let dir: string;

  beforeEach(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'cg-vba-integration-'));
    clearProjectConfigCache();
  });

  afterEach(async () => {
    clearProjectConfigCache();
    // A small delay on Windows to allow file handles to release
    await new Promise((resolve) => setTimeout(resolve, 100));
    try {
      fs.rmSync(dir, { recursive: true, force: true });
    } catch {
      // ignore cleanup EPERM on Windows
    }
  });

  it('correctly threads custom targets through indexAll and conditional compilation', async () => {
    // Write Class module with #If Win64 guard
    fs.writeFileSync(
      path.join(dir, 'MyClass.cls'),
      [
        'VERSION 1.0 CLASS',
        '#If Win64 Then',
        'Public Sub ActiveWin64Sub()',
        'End Sub',
        '#Else',
        'Public Sub InactiveWin64Sub()',
        'End Sub',
        '#End If',
      ].join('\n')
    );

    // 1. Run with Win64: false in codegraph.json
    fs.writeFileSync(
      path.join(dir, 'codegraph.json'),
      JSON.stringify({
        vba: {
          targets: {
            Win64: false,
          },
        },
      })
    );

    const cg1 = CodeGraph.initSync(dir);
    await cg1.indexAll();

    const nodes1 = cg1.getNodesByKind('function');
    const qualifiedNames1 = nodes1.map((n) => n.qualifiedName);
    
    // InactiveWin64Sub should be active (Win64 is false)
    expect(qualifiedNames1).toContain('MyClass.InactiveWin64Sub');
    expect(qualifiedNames1).not.toContain('MyClass.ActiveWin64Sub');

    await cg1.destroy();
    fs.rmSync(path.join(dir, '.codegraph-vba'), { recursive: true, force: true });
    clearProjectConfigCache();

    // 2. Run with Win64: true in codegraph.json
    fs.writeFileSync(
      path.join(dir, 'codegraph.json'),
      JSON.stringify({
        vba: {
          targets: {
            Win64: true,
          },
        },
      })
    );

    const cg2 = CodeGraph.initSync(dir);
    await cg2.indexAll();

    const nodes2 = cg2.getNodesByKind('function');
    const qualifiedNames2 = nodes2.map((n) => n.qualifiedName);

    // ActiveWin64Sub should be active (Win64 is true)
    expect(qualifiedNames2).toContain('MyClass.ActiveWin64Sub');
    expect(qualifiedNames2).not.toContain('MyClass.InactiveWin64Sub');

    await cg2.destroy();
  });
});
