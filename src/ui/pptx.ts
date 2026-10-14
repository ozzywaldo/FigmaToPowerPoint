/**
 * Slide model → pptxgenjs presentation.
 *
 * Pure: takes the pptxgenjs constructor as an argument so it runs identically in the plugin
 * iframe and under Node for tests. Everything here is about unit conversion and mapping the
 * model onto pptxgenjs' option bags; no Figma types are involved.
 */
import type PptxGenJS from 'pptxgenjs';
import type {
  Element,
  ExportOptions,
  ImageElement,
  Link,
  PathElement,
  Shadow,
  ShapeElement,
  Slide,
  SolidFill,
  Stroke,
  TextElement,
} from '../shared/types';
import { isBold, resolveFontFace } from './fonts';

export interface DeckInput {
  /** Dominant frame size; slides of other sizes are letterboxed into the layout. */
  wPx: number;
  hPx: number;
  slides: Slide[];
  options: ExportOptions;
}

type PptxCtor = new () => PptxGenJS;
type PptxSlide = ReturnType<PptxGenJS['addSlide']>;
type TextProps = PptxGenJS.TextProps;
type TextPropsOptions = PptxGenJS.TextPropsOptions;
type ShapeProps = PptxGenJS.ShapeProps;

/** Unit converter for one slide: Figma px → inches / points, including letterbox offsets. */
interface Units {
  x: (px: number) => number;
  y: (px: number) => number;
  len: (px: number) => number;
  pt: (px: number) => number;
}

export function layoutSize(deck: Pick<DeckInput, 'wPx' | 'hPx' | 'options'>): { width: number; height: number } {
  const width = deck.options.slideWidthIn;
  const ratio =
    deck.options.layout === '16:9' ? 9 / 16 : deck.options.layout === '4:3' ? 3 / 4 : deck.hPx / Math.max(1, deck.wPx);
  return { width, height: round(width * ratio, 4) };
}

export function buildPresentation(Pptx: PptxCtor, deck: DeckInput): PptxGenJS {
  const pptx = new Pptx();
  const layout = layoutSize(deck);
  pptx.defineLayout({ name: 'FIGMA', width: layout.width, height: layout.height });
  pptx.layout = 'FIGMA';
  pptx.title = deck.options.fileName || 'Deck';

  const slideNumbers = new Map(deck.slides.map((s, i) => [s.id, i + 1]));

  for (const slide of deck.slides) {
    const scale = Math.min(layout.width / slide.wPx, layout.height / slide.hPx);
    const ox = (layout.width - slide.wPx * scale) / 2;
    const oy = (layout.height - slide.hPx * scale) / 2;
    const u: Units = {
      x: (px) => round(ox + px * scale, 5),
      y: (px) => round(oy + px * scale, 5),
      len: (px) => round(px * scale, 5),
      pt: (px) => round(px * scale * 72, 2),
    };
    const ps = pptx.addSlide();
    if (slide.background) ps.background = { color: slide.background.color, transparency: transparency(slide.background.alpha) };
    if (slide.notes) ps.addNotes(slide.notes);
    for (const el of slide.elements) addElement(ps, el, u, deck.options, slideNumbers);
  }
  return pptx;
}

function addElement(ps: PptxSlide, el: Element, u: Units, options: ExportOptions, slideNumbers: Map<string, number>): void {
  switch (el.type) {
    case 'shape':
      return addShape(ps, el, u, slideNumbers);
    case 'path':
      return addPath(ps, el, u, slideNumbers);
    case 'image':
      return addImage(ps, el, u, slideNumbers);
    case 'text':
      return addText(ps, el, u, options, slideNumbers);
  }
}

// ---- shapes ------------------------------------------------------------------------------

function placement(el: Element, u: Units): Pick<ShapeProps, 'x' | 'y' | 'w' | 'h' | 'rotate' | 'flipH' | 'flipV'> {
  return {
    x: u.x(el.x),
    y: u.y(el.y),
    w: u.len(el.w),
    h: u.len(el.h),
    rotate: el.rotation || undefined,
    flipH: el.flipH || undefined,
    flipV: el.flipV || undefined,
  };
}

function fillProps(fill: SolidFill | undefined): PptxGenJS.ShapeFillProps | undefined {
  if (!fill) return undefined;
  return { color: fill.color, transparency: transparency(fill.alpha) };
}

function lineProps(stroke: Stroke | undefined, u: Units): PptxGenJS.ShapeLineProps | undefined {
  if (!stroke) return undefined;
  return {
    color: stroke.color,
    transparency: transparency(stroke.alpha),
    width: Math.max(0.25, u.pt(stroke.widthPx)),
    dashType: stroke.dashed ? 'dash' : 'solid',
    beginArrowType: stroke.arrowStart ? 'triangle' : undefined,
    endArrowType: stroke.arrowEnd ? 'triangle' : undefined,
  };
}

function shadowProps(shadow: Shadow | undefined, u: Units): PptxGenJS.ShadowProps | undefined {
  if (!shadow) return undefined;
  const offset = Math.hypot(shadow.offsetXPx, shadow.offsetYPx);
  let angle = (Math.atan2(shadow.offsetYPx, shadow.offsetXPx) * 180) / Math.PI;
  if (angle < 0) angle += 360;
  return {
    type: 'outer',
    color: shadow.color,
    opacity: round(shadow.alpha, 3),
    blur: u.pt(shadow.blurPx),
    offset: u.pt(offset),
    angle: round(angle, 1),
  };
}

function hyperlink(link: Link | undefined, slideNumbers: Map<string, number>): PptxGenJS.HyperlinkProps | undefined {
  if (!link) return undefined;
  if (link.kind === 'url') return { url: link.url };
  const n = slideNumbers.get(link.slideId);
  return n ? { slide: n } : undefined;
}

function addShape(ps: PptxSlide, el: ShapeElement, u: Units, slideNumbers: Map<string, number>): void {
  const props: ShapeProps = {
    ...placement(el, u),
    fill: fillProps(el.fill),
    line: lineProps(el.stroke, u),
    shadow: shadowProps(el.shadow, u),
    hyperlink: hyperlink(el.link, slideNumbers),
    objectName: el.name,
  };
  if (el.shape === 'line') {
    props.h = 0;
    ps.addShape('line', props);
    return;
  }
  if (el.shape === 'ellipse') {
    ps.addShape('ellipse', props);
    return;
  }
  const radius = el.cornerRadiusPx ? Math.min(el.cornerRadiusPx, Math.min(el.w, el.h) / 2) : 0;
  if (radius > 0) {
    props.rectRadius = u.len(radius);
    ps.addShape('roundRect', props);
  } else {
    ps.addShape('rect', props);
  }
}

function addPath(ps: PptxSlide, el: PathElement, u: Units, slideNumbers: Map<string, number>): void {
  const points: NonNullable<ShapeProps['points']> = el.commands.map((c) => {
    switch (c.op) {
      case 'M':
        return { x: u.len(c.x), y: u.len(c.y), moveTo: true };
      case 'L':
        return { x: u.len(c.x), y: u.len(c.y) };
      case 'C':
        return {
          x: u.len(c.x),
          y: u.len(c.y),
          curve: { type: 'cubic', x1: u.len(c.x1), y1: u.len(c.y1), x2: u.len(c.x2), y2: u.len(c.y2) },
        };
      case 'Q':
        return { x: u.len(c.x), y: u.len(c.y), curve: { type: 'quadratic', x1: u.len(c.x1), y1: u.len(c.y1) } };
      case 'Z':
        return { close: true };
    }
  });
  ps.addShape('custGeom' as PptxGenJS.SHAPE_NAME, {
    ...placement(el, u),
    points,
    fill: fillProps(el.fill),
    line: lineProps(el.stroke, u),
    shadow: shadowProps(el.shadow, u),
    hyperlink: hyperlink(el.link, slideNumbers),
    objectName: el.name,
  });
}

function addImage(ps: PptxSlide, el: ImageElement, u: Units, slideNumbers: Map<string, number>): void {
  const mime = el.format === 'jpg' ? 'image/jpeg' : 'image/png';
  ps.addImage({
    data: `${mime};base64,${toBase64(el.bytes)}`,
    ...placement(el, u),
    transparency: el.opacity < 1 ? transparency(el.opacity) : undefined,
    hyperlink: hyperlink(el.link, slideNumbers),
    objectName: el.name,
  });
}

// ---- text --------------------------------------------------------------------------------

function addText(ps: PptxSlide, el: TextElement, u: Units, options: ExportOptions, slideNumbers: Map<string, number>): void {
  const runs: TextProps[] = [];
  el.paragraphs.forEach((p, pi) => {
    const last = pi === el.paragraphs.length - 1;
    const paragraphOpts: TextPropsOptions = {
      align: el.align,
      bullet: p.list === 'bullet' ? true : p.list === 'number' ? { type: 'number' } : undefined,
      indentLevel: p.list !== 'none' ? Math.max(0, p.indent - 1) : p.indent > 0 ? p.indent : undefined,
      paraSpaceAfter: p.spacingAfterPx ? u.pt(p.spacingAfterPx) : undefined,
    };
    if (p.runs.length === 0) {
      runs.push({ text: '', options: { ...paragraphOpts, breakLine: !last } });
      return;
    }
    p.runs.forEach((r, ri) => {
      const fontSize = u.pt(r.fontSizePx);
      const o: TextPropsOptions = {
        ...paragraphOpts,
        fontFace: resolveFontFace(r.fontFamily, r.fontStyle, options.fontFace),
        fontSize,
        bold: isBold(r.fontWeight, r.fontStyle) || undefined,
        italic: r.italic || undefined,
        color: r.color,
        transparency: r.alpha < 1 ? transparency(r.alpha) : undefined,
        underline: r.underline ? { style: 'sng' } : undefined,
        strike: r.strike ? 'sngStrike' : undefined,
        charSpacing: r.letterSpacing
          ? round(r.letterSpacing.unit === 'px' ? u.pt(r.letterSpacing.value) : (fontSize * r.letterSpacing.value) / 100, 2)
          : undefined,
        hyperlink: hyperlink(r.link, slideNumbers),
        softBreakBefore: r.softBreakBefore || undefined,
        breakLine: ri === p.runs.length - 1 && !last ? true : undefined,
      };
      runs.push({ text: r.text, options: stripUndefined(o) });
    });
  });

  const box: TextPropsOptions = {
    ...placement(el, u),
    margin: 0,
    inset: 0,
    isTextBox: true,
    fit: 'none',
    wrap: el.autoResize !== 'width_height',
    valign: el.valign,
    align: el.align,
    shadow: shadowProps(el.shadow, u),
    outline: el.outline ? { color: el.outline.color, size: u.pt(el.outline.widthPx) } : undefined,
    objectName: el.name,
  };
  if (el.lineHeight) {
    if (el.lineHeight.unit === 'px') box.lineSpacing = u.pt(el.lineHeight.value);
    else box.lineSpacingMultiple = round(el.lineHeight.value / 100, 3);
  }
  ps.addText(runs, stripUndefined(box));
}

// ---- helpers -----------------------------------------------------------------------------

function transparency(alpha: number): number {
  return round(Math.max(0, Math.min(100, (1 - alpha) * 100)), 1);
}

function round(n: number, places: number): number {
  const f = 10 ** places;
  return Math.round(n * f) / f;
}

function stripUndefined<T extends object>(o: T): T {
  for (const k of Object.keys(o) as (keyof T)[]) if (o[k] === undefined) delete o[k];
  return o;
}

export function toBase64(bytes: Uint8Array): string {
  let binary = '';
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode.apply(null, Array.from(bytes.subarray(i, i + chunk)));
  }
  return btoa(binary);
}
