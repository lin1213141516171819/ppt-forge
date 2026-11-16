import assert from 'node:assert/strict';
import { mkdtempSync, mkdirSync, writeFileSync, rmSync, utimesSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { tmpdir } from 'node:os';
import test from 'node:test';
import { apply } from '../index.js';

let handler;
apply({
  effect: (callback) => callback(),
  tools: { register: () => () => {} },
  inject: (_deps, callback) => callback({
    effect: (effect) => effect(),
    webServer: { register: (spec) => { handler = spec.handler; return () => {}; } },
  }),
});

test('history includes project folders and legacy decks, sorted newest first', async () => {
  const root = mkdtempSync(join(tmpdir(), 'ppt-forge-test-'));
  try {
    const output = join(root, 'ppt-forge-output');
    const project = join(output, 'report-20261116-120000');
    mkdirSync(project, { recursive: true });
    writeFileSync(join(output, 'legacy.pptx'), 'older');
    utimesSync(join(output, 'legacy.pptx'), 1, 1);
    writeFileSync(join(project, 'deck.PPTX'), 'newer');
    writeFileSync(join(project, 'outline.json'), '{}');
    mkdirSync(join(output, 'directory.pptx'));
    let body;
    await handler({ method: 'GET', url: '/ppt-forge/history?cwd=' + encodeURIComponent(root) }, {
      writeHead: (status) => assert.equal(status, 200),
      end: (text) => { body = JSON.parse(text); },
    });
    assert.deepEqual(body.files.map((file) => file.name), [join('report-20261116-120000', 'deck.PPTX'), 'legacy.pptx']);
    assert.equal(body.files[0].size, 5);
  } finally {
    assert.equal(dirname(root), tmpdir());
    rmSync(root, { recursive: true, force: true });
  }
});

test('history returns an empty list before any decks exist', async () => {
  const root = mkdtempSync(join(tmpdir(), 'ppt-forge-test-'));
  try {
    let body;
    await handler({ method: 'GET', url: '/ppt-forge/history?cwd=' + encodeURIComponent(root) }, {
      writeHead: (status) => assert.equal(status, 200),
      end: (text) => { body = JSON.parse(text); },
    });
    assert.deepEqual(body.files, []);
  } finally {
    assert.equal(dirname(root), tmpdir());
    rmSync(root, { recursive: true, force: true });
  }
});
