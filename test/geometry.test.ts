import { test } from 'node:test';
import assert from 'node:assert/strict';
import { apply, invert, multiply, placementFromTransform, rectInSlide, type Mat } from '../src/plugin/geometry';

const rot = (deg: number, tx = 0, ty = 0): Mat => {
  // Figma convention: rotation is CCW on screen → [[cos, sin, tx], [-sin, cos, ty]]
  const r = (deg * Math.PI) / 180;
  return [
    [Math.cos(r), Math.sin(r), tx],
    [-Math.sin(r), Math.cos(r), ty],
  ];
};

test('invert ∘ multiply is identity', () => {
  const m = rot(37, 120, -40);
  const id = multiply(invert(m), m);
  assert.ok(Math.abs(id[0][0] - 1) < 1e-9 && Math.abs(id[1][1] - 1) < 1e-9);
  assert.ok(Math.abs(id[0][1]) < 1e-9 && Math.abs(id[1][0]) < 1e-9);
  assert.ok(Math.abs(id[0][2]) < 1e-9 && Math.abs(id[1][2]) < 1e-9);
});

test('unrotated node maps straight through', () => {
  const p = placementFromTransform([[1, 0, 100], [0, 1, 50]], 200, 80);
  assert.deepEqual(p, { x: 100, y: 50, w: 200, h: 80, rotation: 0, flipH: false, flipV: false });
});

test('Figma CCW rotation becomes PowerPoint clockwise rotation about the centre', () => {
  // A 200×80 box rotated 30° CCW (Figma) around its top-left at (100, 50).
  const m = rot(30, 100, 50);
  const p = placementFromTransform(m, 200, 80);
  assert.equal(p.rotation, 330); // -30° clockwise → 330
  // Centre must coincide with the transformed centre.
  const c = apply(m, 100, 40);
  // placement rounds to 3 decimals
  assert.ok(Math.abs(p.x + 100 - c.x) < 2e-3);
  assert.ok(Math.abs(p.y + 40 - c.y) < 2e-3);
  assert.equal(p.flipH, false);
});

test('horizontal flip is detected and rotation stays consistent', () => {
  // R(φ)·F with φ = 20° clockwise: d = cos φ, b = -sin φ, a = -cos φ, c = -sin φ
  const r = (20 * Math.PI) / 180;
  const m: Mat = [
    [-Math.cos(r), -Math.sin(r), 10],
    [-Math.sin(r), Math.cos(r), 20],
  ];
  const p = placementFromTransform(m, 100, 50);
  assert.equal(p.flipH, true);
  assert.equal(p.rotation, 20);
});

test('rectInSlide expresses absolute bounds in slide space', () => {
  const slideInv = invert([[1, 0, 1000], [0, 1, 500]]);
  const r = rectInSlide(slideInv, { x: 1100, y: 600, width: 50, height: 25 });
  assert.deepEqual(r, { x: 100, y: 100, w: 50, h: 25 });
});
