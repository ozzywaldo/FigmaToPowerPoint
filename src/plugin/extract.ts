/**
 * Figma node tree → Slide model.
 *
 * The walker keeps layers editable whenever PowerPoint can represent them faithfully
 * (solid fills, single drop shadows, rectangles/ellipses/lines/vector paths, text) and
 * rasterizes everything else (gradients, image fills, blurs, masks, blend modes, clipped
 * overflow). Every fallback is reported as a warning so the designer can decide whether to
 * restructure the layer or accept the image.
 *
 * Layer-name conventions (case-insensitive, at the start of the name):
 *   [PNG]  / [JPG] / [IMG]   force the layer to export as an image (PNG/JPG/auto)
 *   notes…                   a text layer that becomes the slide's speaker notes
 *   [SKIP]                   leave the layer out of the export entirely
 */
import type {
  Element,
  ExportOptions,
  ImageElement,
  Link,
  PathElement,
  ShapeElement,
  Slide,
  Stroke,
  TextElement,
  Warning,
} from '../shared/types';
import { invert, multiply, placementFromTransform, rectInSlide, round, type Mat } from './geometry';
import { effectsFrom, hasImagePaint, solidFromPaints, strokeFrom, visiblePaints } from './paint';
import { extractTextContent } from './text';
import { parsePaths } from './vector';

export type SlideRoot = FrameNode | ComponentNode | InstanceNode | SlideNode;

export interface ExtractContext {
  options: ExportOptions;
  /** ids of every slide in the export, for slide-to-slide links. */
  slideIds: Set<string>;
  warnings: Warning[];
}

interface Walk {
  ctx: ExtractContext;
  root: SlideRoot;
  slide: Slide;
  slideInv: Mat;
  warn: (node: SceneNode, message: string) => void;
}

type Container = FrameNode | ComponentNode | InstanceNode | GroupNode | SectionNode | SlideNode;

const FORCE_IMAGE = /^\[(png|jpg|jpeg|img|image)\]/i;
const SKIP = /^\[skip\]/i;

export function isSlideRoot(node: BaseNode): node is SlideRoot {
  return node.type === 'FRAME' || node.type === 'COMPONENT' || node.type === 'INSTANCE' || node.type === 'SLIDE';
}

export async function extractSlide(root: SlideRoot, ctx: ExtractContext): Promise<Slide> {
  const slide: Slide = { id: root.id, name: root.name, wPx: round(root.width), hPx: round(root.height), elements: [] };
  const w: Walk = {
    ctx,
    root,
    slide,
    slideInv: invert(root.absoluteTransform),
    warn: (node, message) => ctx.warnings.push({ slideName: root.name, nodeName: node.name, message }),
  };

  if (Math.abs(root.rotation) > 0.01) {
    w.warn(root, 'Slide frames should not be rotated; child positions may be off.');
  }

  collectNotes(root, w);

  if (!ctx.options.editable) {
    const img = await rasterNode(root, w, 1);
    if (img) slide.elements.push(img);
    return slide;
  }

  const bg = solidFromPaints(root.fills, 1);
  if (bg) {
    slide.background = bg;
  } else if (bg === null) {
    const shell = await rasterShell(root, w, 1);
    if (shell) slide.elements.push(shell);
  }
  if (visiblePaints(root.strokes).length) w.warn(root, 'Strokes on the slide frame itself are not exported.');

  for (const child of root.children) await walk(child, w, 1);
  return slide;
}

// ---- notes --------------------------------------------------------------------------------

function isNotesLayer(node: SceneNode, options: ExportOptions): boolean {
  const prefix = options.notesLayerName.trim().toLowerCase();
  return node.type === 'TEXT' && prefix.length > 0 && node.name.trim().toLowerCase().startsWith(prefix);
}

function collectNotes(root: SlideRoot, w: Walk): void {
  const notes = root.findAll((n) => isNotesLayer(n, w.ctx.options)) as TextNode[];
  const text = notes
    .map((n) => n.characters.trim())
    .filter(Boolean)
    .join('\n\n');
  if (text) w.slide.notes = text;
}

// ---- walker --------------------------------------------------------------------------------

async function walk(node: SceneNode, w: Walk, opacity: number): Promise<void> {
  if (!node.visible || SKIP.test(node.name)) return;
  if (isNotesLayer(node, w.ctx.options)) return;
  if (node.type === 'SLICE') return;

  if (FORCE_IMAGE.test(node.name)) return push(w, await rasterNode(node, w, opacity));

  if ('blendMode' in node && node.blendMode !== 'NORMAL' && node.blendMode !== 'PASS_THROUGH') {
    w.warn(node, `Blend mode ${node.blendMode} was exported as an image.`);
    return push(w, await rasterNode(node, w, opacity));
  }

  switch (node.type) {
    case 'TEXT':
      return push(w, await textElement(node, w, opacity));
    case 'RECTANGLE':
      return push(w, await rectElement(node, w, opacity));
    case 'ELLIPSE':
      return push(w, await ellipseElement(node, w, opacity));
    case 'LINE':
      return push(w, (await lineElement(node, w, opacity)) ?? (await rasterNode(node, w, opacity)));
    case 'VECTOR':
    case 'STAR':
    case 'POLYGON':
    case 'BOOLEAN_OPERATION':
      return push(w, await vectorElement(node, w, opacity));
    case 'FRAME':
    case 'COMPONENT':
    case 'INSTANCE':
    case 'GROUP':
    case 'SECTION':
    case 'SLIDE':
      return container(node, w, opacity);
    case 'COMPONENT_SET':
      return;
    default:
      if ('exportAsync' in node) {
        w.warn(node, `${node.type} layers are exported as images.`);
        return push(w, await rasterNode(node, w, opacity));
      }
  }
}

function push(w: Walk, el: Element | null | undefined): void {
  if (el) w.slide.elements.push(el);
}

async function container(node: Container, w: Walk, parentOpacity: number): Promise<void> {
  const opacity = parentOpacity * ('opacity' in node ? node.opacity : 1);
  const isRootLike = node.type === 'SLIDE';

  if (node.children.some((c) => 'isMask' in c && c.isMask)) {
    w.warn(node, 'Contains a mask; exported as one image.');
    return push(w, await rasterNode(node, w, parentOpacity));
  }

  if (!isRootLike && node.type !== 'GROUP' && node.type !== 'SECTION') {
    if (node.clipsContent && childrenOverflow(node)) {
      w.warn(node, 'Clips overflowing children; exported as one image (PowerPoint cannot clip).');
      return push(w, await rasterNode(node, w, parentOpacity));
    }
    const fill = solidFromPaints(node.fills, opacity);
    const stroke = strokeFrom(node, opacity);
    const fx = effectsFrom(node, opacity);
    if (fill === null || stroke === null || fx === null) {
      const shell = await rasterShell(node, w, parentOpacity);
      push(w, shell);
    } else if (fill || stroke || fx.shadow) {
      const el = await shapeElement(node, w, 'rect', { fill, stroke, shadow: fx.shadow });
      el.cornerRadiusPx = cornerRadius(node, w);
      push(w, el);
    }
  }

  for (const child of node.children) await walk(child, w, opacity);
}

function childrenOverflow(node: FrameNode | ComponentNode | InstanceNode): boolean {
  const box = node.absoluteBoundingBox;
  if (!box) return false;
  const tol = 0.5;
  return node.children.some((c) => {
    if (!c.visible) return false;
    const b = renderBounds(c);
    if (!b) return false;
    return (
      b.x < box.x - tol ||
      b.y < box.y - tol ||
      b.x + b.width > box.x + box.width + tol ||
      b.y + b.height > box.y + box.height + tol
    );
  });
}

// ---- element builders ---------------------------------------------------------------------

async function base(node: SceneNode, w: Walk): Promise<Omit<ShapeElement, 'type' | 'shape'>> {
  const rel = multiply(w.slideInv, node.absoluteTransform);
  return {
    id: node.id,
    name: node.name,
    ...placementFromTransform(rel, node.width, node.height),
    link: await linkFromReactions(node, w),
  };
}

async function shapeElement(
  node: SceneNode & GeometryMixin & MinimalStrokesMixin,
  w: Walk,
  shape: ShapeElement['shape'],
  parts: Pick<ShapeElement, 'fill' | 'stroke' | 'shadow'>,
): Promise<ShapeElement> {
  const el: ShapeElement = { type: 'shape', shape, ...(await base(node, w)), ...parts };
  // Figma strokes default to INSIDE; PowerPoint's are centred. Nudge the box so edges match.
  if (parts.stroke && shape !== 'line') {
    const half = parts.stroke.widthPx / 2;
    const delta = node.strokeAlign === 'INSIDE' ? -half : node.strokeAlign === 'OUTSIDE' ? half : 0;
    if (delta !== 0) {
      el.x = round(el.x - delta);
      el.y = round(el.y - delta);
      el.w = round(Math.max(0, el.w + 2 * delta));
      el.h = round(Math.max(0, el.h + 2 * delta));
    }
  }
  return el;
}

async function rectElement(node: RectangleNode, w: Walk, opacity: number): Promise<Element | null> {
  const own = opacity * node.opacity;
  const fill = solidFromPaints(node.fills, own);
  const stroke = strokeFrom(node, own);
  const fx = effectsFrom(node, own);
  if (fill === null || stroke === null || fx === null) {
    if (!hasImagePaint(node.fills)) w.warn(node, 'Gradient, multi-layer fill or unsupported effect; exported as an image.');
    return rasterNode(node, w, opacity);
  }
  if (!fill && !stroke) return null;
  const el = await shapeElement(node, w, 'rect', { fill, stroke, shadow: fx.shadow });
  el.cornerRadiusPx = cornerRadius(node, w);
  return el;
}

function cornerRadius(node: CornerMixin & Partial<RectangleCornerMixin> & SceneNode, w: Walk): number | undefined {
  if (typeof node.cornerRadius === 'number') return node.cornerRadius || undefined;
  const radii = [node.topLeftRadius, node.topRightRadius, node.bottomRightRadius, node.bottomLeftRadius].filter(
    (r): r is number => typeof r === 'number',
  );
  const max = Math.max(0, ...radii);
  if (max > 0) w.warn(node, 'Mixed corner radii; PowerPoint rounded rectangles use one radius (largest used).');
  return max || undefined;
}

async function ellipseElement(node: EllipseNode, w: Walk, opacity: number): Promise<Element | null> {
  const arc = node.arcData;
  const full = arc.innerRadius === 0 && Math.abs(Math.abs(arc.endingAngle - arc.startingAngle) - Math.PI * 2) < 1e-4;
  if (!full) return vectorElement(node, w, opacity);
  const own = opacity * node.opacity;
  const fill = solidFromPaints(node.fills, own);
  const stroke = strokeFrom(node, own);
  const fx = effectsFrom(node, own);
  if (fill === null || stroke === null || fx === null) {
    if (!hasImagePaint(node.fills)) w.warn(node, 'Gradient, multi-layer fill or unsupported effect; exported as an image.');
    return rasterNode(node, w, opacity);
  }
  if (!fill && !stroke) return null;
  return shapeElement(node, w, 'ellipse', { fill, stroke, shadow: fx.shadow });
}

async function lineElement(node: LineNode, w: Walk, opacity: number): Promise<ShapeElement | null> {
  const own = opacity * node.opacity;
  const stroke = strokeFrom(node, own);
  const fx = effectsFrom(node, own);
  if (stroke === null || fx === null) return null;
  if (!stroke) return null;
  applyArrows(stroke, node.strokeCap);
  const el = await shapeElement(node, w, 'line', { stroke, shadow: fx.shadow });
  el.h = 0;
  return el;
}

function applyArrows(stroke: Stroke, cap: StrokeCap | PluginAPI['mixed']): void {
  if (typeof cap === 'string' && cap.startsWith('ARROW')) {
    stroke.arrowStart = true;
    stroke.arrowEnd = true;
  }
}

async function vectorElement(
  node: VectorNode | StarNode | PolygonNode | BooleanOperationNode | EllipseNode,
  w: Walk,
  opacity: number,
): Promise<Element | null> {
  if (w.ctx.options.vectors === 'raster') return rasterNode(node, w, opacity);
  const own = opacity * node.opacity;
  const fill = solidFromPaints(node.fills, own);
  const stroke = strokeFrom(node, own);
  const fx = effectsFrom(node, own);
  const paths = node.type === 'VECTOR' && node.fillGeometry.length === 0 ? node.vectorPaths : node.fillGeometry;
  const commands = paths.some((p) => p.windingRule === 'EVENODD') ? null : parsePaths(paths);
  if (fill === null || stroke === null || fx === null || commands === null || commands.length === 0) {
    w.warn(node, 'Vector could not be converted to an editable shape; exported as an image.');
    return rasterNode(node, w, opacity);
  }
  if (!fill && !stroke) return null;
  if (stroke) applyArrows(stroke, node.strokeCap);
  const el: PathElement = { type: 'path', ...(await base(node, w)), commands, fill, stroke, shadow: fx.shadow };
  return el;
}

async function textElement(node: TextNode, w: Walk, opacity: number): Promise<Element | null> {
  const own = opacity * node.opacity;
  const content = extractTextContent(node, {
    opacity: own,
    // Text segments are read synchronously; node links are resolved to slides afterwards.
    resolveLink: (t) => (t ? (t.type === 'URL' ? { kind: 'url', url: t.value } : { kind: 'slide', slideId: t.value }) : undefined),
    warn: w.warn,
  });
  if (content === undefined) return null;
  const fx = effectsFrom(node, own);
  if (content === null || fx === null) {
    if (fx === null) w.warn(node, 'Text with blur/inner shadow effects was exported as an image.');
    return rasterNode(node, w, opacity);
  }
  const outline = strokeFrom(node, own);
  if (outline === null) w.warn(node, 'Non-solid text stroke was dropped.');
  const el: TextElement = {
    type: 'text',
    ...(await base(node, w)),
    ...content,
    shadow: fx.shadow,
    outline: outline ?? undefined,
  };
  for (const p of el.paragraphs) {
    for (const r of p.runs) {
      if (r.link?.kind === 'slide') r.link = await slideLink(r.link.slideId, w);
    }
  }
  return el;
}

// ---- raster fallbacks ----------------------------------------------------------------------

function pickFormat(node: SceneNode, w: Walk, opacity: number): 'PNG' | 'JPG' {
  const forced = node.name.match(FORCE_IMAGE)?.[1]?.toLowerCase();
  if (forced === 'png') return 'PNG';
  if (forced === 'jpg' || forced === 'jpeg') return 'JPG';
  const pref = w.ctx.options.imageFormat;
  if (pref === 'png') return 'PNG';
  if (pref === 'jpg') return 'JPG';
  // auto: JPG only when we are confident there is no transparency.
  if (node.type !== 'RECTANGLE' || opacity < 1 || node.opacity < 1) return 'PNG';
  if (typeof node.cornerRadius === 'number' ? node.cornerRadius > 0 : true) return 'PNG';
  if ((node.effects ?? []).some((e) => e.visible !== false)) return 'PNG';
  const fills = visiblePaints(node.fills);
  const opaqueImage = fills.length > 0 && fills.every((p) => p.type === 'IMAGE' && (p.opacity ?? 1) === 1 && p.scaleMode !== 'FIT');
  return opaqueImage ? 'JPG' : 'PNG';
}

async function rasterNode(node: SceneNode, w: Walk, parentOpacity: number): Promise<ImageElement | null> {
  const bounds = renderBounds(node);
  if (!bounds || bounds.width < 0.5 || bounds.height < 0.5) return null;
  const format = pickFormat(node, w, parentOpacity);
  let bytes: Uint8Array;
  try {
    bytes = await node.exportAsync({
      format,
      constraint: { type: 'SCALE', value: w.ctx.options.imageScale },
      useAbsoluteBounds: false,
    });
  } catch (e) {
    w.warn(node, `Image export failed: ${(e as Error).message}`);
    return null;
  }
  return {
    type: 'image',
    id: node.id,
    name: node.name,
    ...rectInSlide(w.slideInv, bounds),
    rotation: 0,
    flipH: false,
    flipV: false,
    format: format === 'JPG' ? 'jpg' : 'png',
    bytes,
    opacity: parentOpacity,
    link: await linkFromReactions(node, w),
  };
}

/** Export a container's own fills/strokes/effects without its children. */
async function rasterShell(node: Container & ChildrenMixin, w: Walk, parentOpacity: number): Promise<ImageElement | null> {
  const hidden: SceneNode[] = [];
  try {
    for (const c of node.children) {
      if (c.visible) {
        c.visible = false;
        hidden.push(c);
      }
    }
    return await rasterNode(node, w, parentOpacity);
  } catch (e) {
    w.warn(node, `Could not export frame background: ${(e as Error).message}`);
    return null;
  } finally {
    for (const c of hidden) c.visible = true;
  }
}

// ---- links --------------------------------------------------------------------------------

async function linkFromReactions(node: SceneNode, w: Walk): Promise<Link | undefined> {
  if (!('reactions' in node)) return undefined;
  for (const reaction of node.reactions) {
    const actions = reaction.actions ?? (reaction.action ? [reaction.action] : []);
    for (const action of actions) {
      if (action.type === 'URL' && action.url) return { kind: 'url', url: action.url };
      if (action.type === 'NODE' && action.destinationId) {
        const link = await slideLink(action.destinationId, w);
        if (link) return link;
      }
    }
  }
  return undefined;
}

/** A link may target a slide or something inside one; resolve it to the containing slide. */
async function slideLink(nodeId: string, w: Walk): Promise<Link | undefined> {
  if (w.ctx.slideIds.has(nodeId)) return { kind: 'slide', slideId: nodeId };
  let n: BaseNode | null = await figma.getNodeByIdAsync(nodeId);
  while (n) {
    if (w.ctx.slideIds.has(n.id)) return { kind: 'slide', slideId: n.id };
    n = n.parent;
  }
  return undefined;
}

/** Render bounds include shadows/outside strokes; some node types only expose the plain box. */
function renderBounds(node: SceneNode): Rect | null {
  if ('absoluteRenderBounds' in node) return node.absoluteRenderBounds;
  return node.absoluteBoundingBox;
}
