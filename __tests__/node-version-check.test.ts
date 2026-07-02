/**
 * Pin the Node-25 block banner content. The banner replaced a soft
 * `console.warn` because the warning was scrolling off-screen before
 * the OOM crash 30 seconds later, generating duplicate bug reports
 * (#54, #81, #140). The recipe and override env var below are
 * load-bearing — if any of them get edited away, this test catches it.
 */

import { describe, it, expect } from 'vitest';
import {
  buildNode25BlockBanner,
  buildNodeTooOldBanner,
  MIN_NODE_MAJOR,
  MIN_NODE_MINOR,
  isBelowMinimumNodeVersion,
} from '../src/bin/node-version-check';

describe('buildNode25BlockBanner', () => {
  it('embeds the reported Node version in the header', () => {
    expect(buildNode25BlockBanner('25.9.0')).toContain(
      'Unsupported Node.js version: 25.9.0'
    );
  });

  it('names the V8 turboshaft WASM root cause and the OOM symptom', () => {
    const banner = buildNode25BlockBanner('25.7.0');
    expect(banner).toContain('V8 WASM JIT');
    expect(banner).toContain('turboshaft');
    expect(banner).toContain('Fatal process out of memory: Zone');
  });

  it('points users to Node 22 LTS via nvm and Homebrew', () => {
    const banner = buildNode25BlockBanner('25.7.0');
    expect(banner).toContain('Node.js 22 LTS');
    expect(banner).toContain('nvm install 22');
    expect(banner).toContain('brew install node@22');
  });

  it('documents the CODEGRAPH_ALLOW_UNSAFE_NODE override', () => {
    const banner = buildNode25BlockBanner('25.7.0');
    expect(banner).toContain('CODEGRAPH_ALLOW_UNSAFE_NODE=1');
  });

  it('links to issue #81 for the root-cause writeup', () => {
    expect(buildNode25BlockBanner('25.7.0')).toContain(
      'github.com/colbymchenry/codegraph/issues/81'
    );
  });
});

describe('buildNodeTooOldBanner', () => {
  it('embeds the reported Node version in the header', () => {
    expect(buildNodeTooOldBanner('18.20.0')).toContain(
      'Unsupported Node.js version: 18.20.0'
    );
  });

  it('states the supported floor matching MIN_NODE_MAJOR', () => {
    expect(MIN_NODE_MAJOR).toBe(22);
    expect(MIN_NODE_MINOR).toBe(5);
    expect(buildNodeTooOldBanner('20.0.0')).toContain('requires Node.js 22.5 or newer');
    expect(buildNodeTooOldBanner('22.4.0')).toContain('built-in node:sqlite');
  });

  it('rejects Node 20 and Node 22 before node:sqlite became available', () => {
    expect(isBelowMinimumNodeVersion('20.19.0')).toBe(true);
    expect(isBelowMinimumNodeVersion('22.4.0')).toBe(true);
    expect(isBelowMinimumNodeVersion('22.5.0')).toBe(false);
    expect(isBelowMinimumNodeVersion('24.0.0')).toBe(false);
  });

  it('points users to Node 22 LTS via nvm and Homebrew', () => {
    const banner = buildNodeTooOldBanner('16.0.0');
    expect(banner).toContain('Node.js 22 LTS');
    expect(banner).toContain('nvm install 22');
    expect(banner).toContain('brew install node@22');
  });

  it('documents the CODEGRAPH_ALLOW_UNSAFE_NODE override', () => {
    expect(buildNodeTooOldBanner('18.0.0')).toContain('CODEGRAPH_ALLOW_UNSAFE_NODE=1');
  });
});
