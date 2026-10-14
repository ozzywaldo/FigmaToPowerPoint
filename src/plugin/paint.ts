/** Paint, stroke and effect conversion. */
import type { Hex, Shadow, SolidFill, Stroke } from '../shared/types';

export function toHex(c: RGB): Hex {
  const h = (v: number) =>
    Math.max(0, Math.min(255, Math.round(v * 255)))
      .toString(16)
      .padStart(2, '0');
  return (h(c.r) + h(c.g) + h(c.b)).toUpperCase();
}

export function visiblePaints(paints: readonly Paint[] | PluginAPI['mixed']): Paint[] {
  if (!Array.isArray(paints)) return [];
  return paints.filter((p) => p.visible !== false && (p.opacity ?? 1) > 0);
}

/**
 * Collapse a paint stack into one solid fill, if the stack is representable as one.
 * Returns `undefined` for "no fill", and `null` when the stack needs rasterizing.
 * Gradient/image paints are reported through `onGradient` so the caller can decide
 * (text uses the first gradient stop, shapes rasterize).
 */
export function solidFromPaints(
  paints: readonly Paint[] | PluginAPI['mixed'],
  opacity: number,
  allowGradientFallback = false,
): SolidFill | undefined | null {
  const visible = visiblePaints(paints);
  if (visible.length === 0) return undefined;
  const top = visible[visible.length - 1];
  if (top.type === 'SOLID') {
    // If anything under a fully opaque solid is hidden anyway, we can still flatten.
    const alpha = (top.opacity ?? 1) * opacity;
    if (visible.length > 1 && alpha < 1) return null;
    return { color: toHex(top.color), alpha: clamp01(alpha) };
  }
  if (allowGradientFallback && top.type.startsWith('GRADIENT')) {
    const g = top as GradientPaint;
    const stop = g.gradientStops[0];
    if (stop) return { color: toHex(stop.color), alpha: clamp01(stop.color.a * (top.opacity ?? 1) * opacity) };
  }
  return null;
}

export function hasImagePaint(paints: readonly Paint[] | PluginAPI['mixed']): boolean {
  return visiblePaints(paints).some((p) => p.type === 'IMAGE' || p.type === 'VIDEO');
}

export function isGradient(paints: readonly Paint[] | PluginAPI['mixed']): boolean {
  return visiblePaints(paints).some((p) => p.type.startsWith('GRADIENT'));
}

export function strokeFrom(node: GeometryMixin & MinimalStrokesMixin, opacity: number): Stroke | undefined | null {
  const visible = visiblePaints(node.strokes);
  if (visible.length === 0) return undefined;
  const weight = typeof node.strokeWeight === 'number' ? node.strokeWeight : maxStrokeWeight(node);
  if (!weight) return undefined;
  const top = visible[visible.length - 1];
  if (top.type !== 'SOLID') return null;
  return {
    color: toHex(top.color),
    alpha: clamp01((top.opacity ?? 1) * opacity),
    widthPx: weight,
    dashed: Array.isArray(node.dashPattern) && node.dashPattern.length > 0,
  };
}

function maxStrokeWeight(node: MinimalStrokesMixin): number {
  const n = node as Partial<IndividualStrokesMixin>;
  return Math.max(n.strokeTopWeight ?? 0, n.strokeRightWeight ?? 0, n.strokeBottomWeight ?? 0, n.strokeLeftWeight ?? 0);
}

/**
 * Effects. Returns `{shadow}` when the effects are representable (none, or a single drop
 * shadow), or `null` when the node needs rasterizing (inner shadows, blurs, multiple shadows).
 */
export function effectsFrom(node: BlendMixin, opacity: number): { shadow?: Shadow } | null {
  const effects = (node.effects ?? []).filter((e) => e.visible !== false);
  if (effects.length === 0) return {};
  if (effects.length > 1) return null;
  const e = effects[0];
  if (e.type !== 'DROP_SHADOW') return null;
  return {
    shadow: {
      color: toHex(e.color),
      alpha: clamp01(e.color.a * opacity),
      offsetXPx: e.offset.x,
      offsetYPx: e.offset.y,
      blurPx: e.radius,
    },
  };
}

export function clamp01(n: number): number {
  return Math.max(0, Math.min(1, n));
}
