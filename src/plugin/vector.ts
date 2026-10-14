/** Figma vector path data → intermediate path commands. */
import type { PathCommand } from '../shared/types';

/**
 * Figma emits absolute SVG path data using only M, L, C, Q and Z, space separated.
 * We parse a superset (commas, repeated implicit commands) to be safe, and refuse anything else
 * (arcs, relative commands) so the caller can fall back to rasterizing.
 */
export function parsePathData(data: string): PathCommand[] | null {
  const tokens = data.replace(/,/g, ' ').match(/[MLCQZmlcqz]|-?\d*\.?\d+(?:e[-+]?\d+)?/gi);
  if (!tokens) return [];
  const out: PathCommand[] = [];
  let i = 0;
  let cmd = '';
  const num = (): number => {
    const v = Number(tokens[i++]);
    if (Number.isNaN(v)) throw new Error('bad number');
    return v;
  };
  try {
    while (i < tokens.length) {
      const t = tokens[i];
      if (/^[A-Za-z]$/.test(t)) {
        cmd = t;
        i++;
        if (cmd === 'Z' || cmd === 'z') {
          out.push({ op: 'Z' });
          continue;
        }
      }
      switch (cmd) {
        case 'M':
          out.push({ op: 'M', x: num(), y: num() });
          cmd = 'L'; // subsequent implicit pairs are line-tos
          break;
        case 'L':
          out.push({ op: 'L', x: num(), y: num() });
          break;
        case 'C':
          out.push({ op: 'C', x1: num(), y1: num(), x2: num(), y2: num(), x: num(), y: num() });
          break;
        case 'Q':
          out.push({ op: 'Q', x1: num(), y1: num(), x: num(), y: num() });
          break;
        default:
          return null; // relative or arc commands: not expected from Figma
      }
    }
  } catch {
    return null;
  }
  return out;
}

export function parsePaths(paths: readonly VectorPath[]): PathCommand[] | null {
  const all: PathCommand[] = [];
  for (const p of paths) {
    const cmds = parsePathData(p.data);
    if (cmds === null) return null;
    all.push(...cmds);
  }
  return all;
}

// ---- even-odd normalisation -----------------------------------------------------------------
//
// DrawingML custom geometry has no fill-rule attribute, so a path that relies on even-odd (every
// glyph with a counter, most logos) must be rewritten so that it renders identically under
// non-zero winding: subpaths are oriented by nesting depth (outer one way, holes the other way).
// That is also correct under even-odd, so the result is renderer-independent.

type Point = { x: number; y: number };

export function normalizeWinding(commands: PathCommand[]): PathCommand[] {
  const subpaths = splitSubpaths(commands);
  if (subpaths.length < 2) return commands;
  const polys = subpaths.map(flatten);
  const areas = polys.map(signedArea);
  const out: PathCommand[] = [];
  subpaths.forEach((sp, i) => {
    if (polys[i].length < 3 || !isClosed(sp)) {
      out.push(...sp);
      return;
    }
    let depth = 0;
    for (let j = 0; j < polys.length; j++) {
      if (j !== i && polys[j].length >= 3 && isClosed(subpaths[j]) && Math.abs(areas[j]) > Math.abs(areas[i]) && contains(polys[j], polys[i][0])) depth++;
    }
    const wantPositive = depth % 2 === 0;
    const isPositive = areas[i] > 0;
    out.push(...(wantPositive === isPositive ? sp : reverseSubpath(sp)));
  });
  return out;
}

function splitSubpaths(commands: PathCommand[]): PathCommand[][] {
  const result: PathCommand[][] = [];
  let current: PathCommand[] = [];
  for (const c of commands) {
    if (c.op === 'M' && current.length) {
      result.push(current);
      current = [];
    }
    current.push(c);
  }
  if (current.length) result.push(current);
  return result;
}

function isClosed(sp: PathCommand[]): boolean {
  return sp[sp.length - 1]?.op === 'Z';
}

function flatten(sp: PathCommand[]): Point[] {
  const pts: Point[] = [];
  let cur: Point = { x: 0, y: 0 };
  for (const c of sp) {
    if (c.op === 'M' || c.op === 'L') {
      cur = { x: c.x, y: c.y };
      pts.push(cur);
    } else if (c.op === 'C') {
      for (let t = 0.2; t <= 1.0001; t += 0.2) {
        const u = 1 - t;
        pts.push({
          x: u * u * u * cur.x + 3 * u * u * t * c.x1 + 3 * u * t * t * c.x2 + t * t * t * c.x,
          y: u * u * u * cur.y + 3 * u * u * t * c.y1 + 3 * u * t * t * c.y2 + t * t * t * c.y,
        });
      }
      cur = { x: c.x, y: c.y };
    } else if (c.op === 'Q') {
      for (let t = 0.25; t <= 1.0001; t += 0.25) {
        const u = 1 - t;
        pts.push({ x: u * u * cur.x + 2 * u * t * c.x1 + t * t * c.x, y: u * u * cur.y + 2 * u * t * c.y1 + t * t * c.y });
      }
      cur = { x: c.x, y: c.y };
    }
  }
  return pts;
}

function signedArea(poly: Point[]): number {
  let a = 0;
  for (let i = 0; i < poly.length; i++) {
    const p = poly[i];
    const q = poly[(i + 1) % poly.length];
    a += p.x * q.y - q.x * p.y;
  }
  return a / 2;
}

function contains(poly: Point[], p: Point): boolean {
  let inside = false;
  for (let i = 0, j = poly.length - 1; i < poly.length; j = i++) {
    const a = poly[i];
    const b = poly[j];
    if (a.y > p.y !== b.y > p.y && p.x < ((b.x - a.x) * (p.y - a.y)) / (b.y - a.y) + a.x) inside = !inside;
  }
  return inside;
}

/** Reverse a closed subpath, swapping Bézier control points so the curve shape is unchanged. */
function reverseSubpath(sp: PathCommand[]): PathCommand[] {
  const segs = sp.filter((c) => c.op !== 'M' && c.op !== 'Z');
  const start = sp[0] as Extract<PathCommand, { op: 'M' }>;
  const points: Point[] = [start, ...segs.map((s) => ({ x: (s as { x: number }).x, y: (s as { y: number }).y }))];
  const out: PathCommand[] = [{ op: 'M', x: points[points.length - 1].x, y: points[points.length - 1].y }];
  for (let i = segs.length - 1; i >= 0; i--) {
    const s = segs[i];
    const to = points[i];
    if (s.op === 'C') out.push({ op: 'C', x1: s.x2, y1: s.y2, x2: s.x1, y2: s.y1, x: to.x, y: to.y });
    else if (s.op === 'Q') out.push({ op: 'Q', x1: s.x1, y1: s.y1, x: to.x, y: to.y });
    else out.push({ op: 'L', x: to.x, y: to.y });
  }
  out.push({ op: 'Z' });
  return out;
}
