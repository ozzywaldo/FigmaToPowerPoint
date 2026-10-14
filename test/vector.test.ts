import { test } from 'node:test';
import assert from 'node:assert/strict';
import { parsePathData, parsePaths } from '../src/plugin/vector';

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
