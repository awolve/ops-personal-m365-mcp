/**
 * Unit tests for signature handling (myoffice bug #2: a --no-html send must
 * never carry the HTML signature as literal text).
 * Run with: npm test
 */

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, mkdirSync, writeFileSync } from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';

// The signature module resolves ~/.config/myoffice-mcp at import time, so the
// fake home must be in place before the import.
const home = mkdtempSync(join(tmpdir(), 'myoffice-sig-'));
process.env.HOME = home;
const configDir = join(home, '.config', 'myoffice-mcp');
mkdirSync(configDir, { recursive: true });

const HTML_SIGNATURE =
  '<div class="cortex-signature"><table><tr><td><img src="data:image/svg+xml;base64,AAAA"></td></tr></table></div>';
writeFileSync(join(configDir, 'signature.html'), HTML_SIGNATURE + '\n');

const { appendSignature, getSignature } = await import('../src/utils/signature.js');

test('html send gets the html signature', () => {
  const out = appendSignature('Hello', true, 'standard');
  assert.ok(out.includes(HTML_SIGNATURE));
});

test('plain-text send with only an html signature installed contains no markup', () => {
  const out = appendSignature('Alert: spec service unreachable', false, 'standard');
  assert.equal(out, 'Alert: spec service unreachable');
  assert.ok(!out.includes('<'), `plain text body carried markup: ${out}`);
});

test('plain-text send with an explicit minimal style still appends no html', () => {
  const out = appendSignature('Hi', false, 'minimal');
  assert.ok(!out.includes('<'));
});

test('style none appends nothing in either format', () => {
  assert.equal(appendSignature('Hi', true, 'none'), 'Hi');
  assert.equal(appendSignature('Hi', false, 'none'), 'Hi');
});

test('plain-text send uses the text rendering once one is installed', () => {
  writeFileSync(join(configDir, 'signature.txt'), 'Björn Allvin\nCTO, Awolve\n');
  assert.equal(getSignature('standard', 'text'), 'Björn Allvin\nCTO, Awolve');
  const out = appendSignature('Hi', false, 'standard');
  assert.equal(out, 'Hi\n\n--\nBjörn Allvin\nCTO, Awolve');
  assert.ok(!out.includes('<'));
});
