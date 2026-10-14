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
