import { test } from 'node:test';
import assert from 'node:assert/strict';
import { normalizeWinding, parsePathData, parsePaths } from '../src/plugin/vector';

test('parses Figma-style absolute path data', () => {
  const cmds = parsePathData('M 0 0 L 100 0 C 120 0 140 20 140 40 Q 140 60 120 60 Z');
  assert.deepEqual(cmds, [
    { op: 'M', x: 0, y: 0 },
    { op: 'L', x: 100, y: 0 },
    { op: 'C', x1: 120, y1: 0, x2: 140, y2: 20, x: 140, y: 40 },
    { op: 'Q', x1: 140, y1: 60, x: 120, y: 60 },
    { op: 'Z' },
  ]);
});

test('implicit repeated coordinates after M become line-tos; commas tolerated', () => {
  const cmds = parsePathData('M0,0 10,0 10,10Z');
  assert.deepEqual(cmds, [
    { op: 'M', x: 0, y: 0 },
    { op: 'L', x: 10, y: 0 },
    { op: 'L', x: 10, y: 10 },
    { op: 'Z' },
  ]);
});

test('rejects relative or arc commands so the caller rasterizes', () => {
  assert.equal(parsePathData('M 0 0 a 5 5 0 1 0 10 0'), null);
  assert.equal(parsePathData('m 0 0 l 1 1'), null);
});

test('concatenates multiple sub-paths', () => {
  const cmds = parsePaths([
    { windingRule: 'NONZERO', data: 'M 0 0 L 1 0 Z' },
    { windingRule: 'NONZERO', data: 'M 5 5 L 6 5 Z' },
  ]);
  assert.equal(cmds?.length, 6);
  assert.deepEqual(cmds?.[3], { op: 'M', x: 5, y: 5 });
});

test('even-odd rings get alternating orientation so non-zero renders the hole', () => {
  // Two same-direction squares: outer 0..10, inner 3..7 (both clockwise in y-down space).
  const cmds = parsePathData('M 0 0 L 10 0 L 10 10 L 0 10 Z M 3 3 L 7 3 L 7 7 L 3 7 Z')!;
  const out = normalizeWinding(cmds);
  const inner = out.slice(5);
  assert.deepEqual(inner.map((c) => (c.op === 'Z' ? 'Z' : `${c.op}${(c as { x: number }).x},${(c as { y: number }).y}`)), ['M3,7', 'L7,7', 'L7,3', 'L3,3', 'Z']);
  // Outer subpath untouched, and a single subpath is returned as is.
  assert.deepEqual(out.slice(0, 5), cmds.slice(0, 5));
  assert.deepEqual(normalizeWinding(cmds.slice(0, 5)), cmds.slice(0, 5));
});

test('reversing keeps Bézier geometry (control points swapped)', () => {
  const cmds = parsePathData('M 0 0 L 10 0 L 10 10 L 0 10 Z M 2 2 C 6 2 8 4 8 8 L 2 8 Z')!;
  const out = normalizeWinding(cmds);
  const inner = out.slice(5);
  assert.deepEqual(inner[0], { op: 'M', x: 2, y: 8 });
  assert.deepEqual(inner[1], { op: 'L', x: 8, y: 8 });
  assert.deepEqual(inner[2], { op: 'C', x1: 8, y1: 4, x2: 6, y2: 2, x: 2, y: 2 });
  assert.deepEqual(inner[3], { op: 'Z' });
});
