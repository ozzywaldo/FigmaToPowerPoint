/**
 * Affine helpers for turning Figma transforms into PowerPoint placements.
 *
 * Figma: `[[a, b, tx], [c, d, ty]]` maps local (u, v) → (a·u + b·v + tx, c·u + d·v + ty).
 * Rotation is counter-clockwise positive (`rotation = atan2(-c, a)`), flips show up as a
 * negative determinant. PowerPoint stores an unrotated box plus a clockwise rotation about
 * the box centre, and optional flips applied before the rotation.
 */
import type { ElementBase } from '../shared/types';

export type Mat = Transform;

export const IDENTITY: Mat = [
  [1, 0, 0],
  [0, 1, 0],
];

export function multiply(m: Mat, n: Mat): Mat {
  // (m ∘ n)(p) = m(n(p))
  const [[a1, b1, tx1], [c1, d1, ty1]] = m;
  const [[a2, b2, tx2], [c2, d2, ty2]] = n;
  return [
    [a1 * a2 + b1 * c2, a1 * b2 + b1 * d2, a1 * tx2 + b1 * ty2 + tx1],
    [c1 * a2 + d1 * c2, c1 * b2 + d1 * d2, c1 * tx2 + d1 * ty2 + ty1],
  ];
}

export function invert(m: Mat): Mat {
  const [[a, b, tx], [c, d, ty]] = m;
  const det = a * d - b * c;
  if (Math.abs(det) < 1e-12) return IDENTITY;
  const ia = d / det;
  const ib = -b / det;
  const ic = -c / det;
  const id = a / det;
  return [
    [ia, ib, -(ia * tx + ib * ty)],
    [ic, id, -(ic * tx + id * ty)],
  ];
}

export function apply(m: Mat, x: number, y: number): { x: number; y: number } {
  return { x: m[0][0] * x + m[0][1] * y + m[0][2], y: m[1][0] * x + m[1][1] * y + m[1][2] };
}

export type Placement = Pick<ElementBase, 'x' | 'y' | 'w' | 'h' | 'rotation' | 'flipH' | 'flipV'>;

/**
 * Given a node's transform *relative to the slide* and its unrotated size, return the
 * PowerPoint-style placement: top-left of the unrotated box, clockwise rotation, flipH.
 */
export function placementFromTransform(rel: Mat, w: number, h: number): Placement {
  const [[a, b], [c, d]] = rel;
  const det = a * d - b * c;
  const flipH = det < 0;
  // For M = R(φ)·F (F = optional horizontal flip), d = cos φ and b = -sin φ in both cases,
  // where φ is the clockwise screen rotation PowerPoint wants.
  let deg = (Math.atan2(-b, d) * 180) / Math.PI;
  if (Math.abs(deg) < 0.01) deg = 0;
  if (deg < 0) deg += 360;
  const centre = apply(rel, w / 2, h / 2);
  return {
    x: round(centre.x - w / 2),
    y: round(centre.y - h / 2),
    w: round(w),
    h: round(h),
    rotation: round(deg),
    flipH,
    flipV: false,
  };
}

/** Axis-aligned box of an absolute rect expressed in slide space. */
export function rectInSlide(slideInv: Mat, r: Rect): { x: number; y: number; w: number; h: number } {
  const pts = [
    apply(slideInv, r.x, r.y),
    apply(slideInv, r.x + r.width, r.y),
    apply(slideInv, r.x, r.y + r.height),
    apply(slideInv, r.x + r.width, r.y + r.height),
  ];
  const xs = pts.map((p) => p.x);
  const ys = pts.map((p) => p.y);
  const minX = Math.min(...xs);
  const minY = Math.min(...ys);
  return { x: round(minX), y: round(minY), w: round(Math.max(...xs) - minX), h: round(Math.max(...ys) - minY) };
}

export function round(n: number, places = 3): number {
  const f = 10 ** places;
  return Math.round(n * f) / f;
}
