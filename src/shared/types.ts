/**
 * The intermediate deck model.
 *
 * The plugin sandbox (src/plugin) turns Figma nodes into this model; the UI (src/ui) turns the
 * model into a .pptx. Keeping the two sides decoupled is deliberate: it lets you add new emitters
 * (Google Slides, Keynote, HTML…) or new extraction rules without touching the other half.
 *
 * Units: every geometric value is in Figma pixels, relative to the slide's top-left corner.
 * The emitter owns the conversion to inches / points.
 */

/** 6-digit uppercase hex without '#', e.g. "1A1A1A". */
export type Hex = string;

export interface SolidFill {
  color: Hex;
  /** 0..1, already multiplied by the node's (and ancestors') layer opacity. */
  alpha: number;
}

export interface Stroke {
  color: Hex;
  alpha: number;
  widthPx: number;
  dashed: boolean;
  arrowStart?: boolean;
  arrowEnd?: boolean;
}

export interface Shadow {
  color: Hex;
  alpha: number;
  offsetXPx: number;
  offsetYPx: number;
  blurPx: number;
}

export type Link = { kind: 'url'; url: string } | { kind: 'slide'; slideId: string };

export interface ElementBase {
  id: string;
  name: string;
  /** Unrotated box, slide-relative px. */
  x: number;
  y: number;
  w: number;
  h: number;
  /** Clockwise degrees in [0, 360) — PowerPoint's convention. 0 for most nodes. */
  rotation: number;
  flipH: boolean;
  flipV: boolean;
  link?: Link;
}

export interface ShapeElement extends ElementBase {
  type: 'shape';
  shape: 'rect' | 'ellipse' | 'line';
  fill?: SolidFill;
  stroke?: Stroke;
  cornerRadiusPx?: number;
  shadow?: Shadow;
}

export type PathCommand =
  | { op: 'M'; x: number; y: number }
  | { op: 'L'; x: number; y: number }
  | { op: 'C'; x1: number; y1: number; x2: number; y2: number; x: number; y: number }
  | { op: 'Q'; x1: number; y1: number; x: number; y: number }
  | { op: 'Z' };

/** A vector converted to an editable PowerPoint freeform. Path coords are px relative to the element box. */
export interface PathElement extends ElementBase {
  type: 'path';
  commands: PathCommand[];
  fill?: SolidFill;
  stroke?: Stroke;
  shadow?: Shadow;
}

export interface ImageElement extends ElementBase {
  type: 'image';
  format: 'png' | 'jpg' | 'svg';
  bytes: Uint8Array;
  /** 0..1 opacity inherited from ancestors (the export already bakes in the node's own opacity). */
  opacity: number;
}

export interface LengthValue {
  unit: 'px' | 'percent';
  value: number;
}

export interface TextRun {
  text: string;
  fontFamily: string;
  /** Figma style name, e.g. "Semi Bold Italic". */
  fontStyle: string;
  fontWeight: number;
  italic: boolean;
  fontSizePx: number;
  color: Hex;
  alpha: number;
  underline: boolean;
  strike: boolean;
  letterSpacing?: LengthValue;
  link?: Link;
  /** A Figma soft line break (U+2028) precedes this run: new line, same paragraph. */
  softBreakBefore?: boolean;
}

export interface Paragraph {
  runs: TextRun[];
  list: 'none' | 'bullet' | 'number';
  /** Figma indentation level (0 = none; list items start at 1). */
  indent: number;
  spacingAfterPx: number;
}

export interface TextElement extends ElementBase {
  type: 'text';
  paragraphs: Paragraph[];
  align: 'left' | 'center' | 'right' | 'justify';
  valign: 'top' | 'middle' | 'bottom';
  autoResize: 'none' | 'height' | 'width_height' | 'truncate';
  /** Single value per text box — PowerPoint has no per-run line height. */
  lineHeight?: LengthValue;
  shadow?: Shadow;
  /** Text stroke, mapped to a PowerPoint text outline. */
  outline?: Stroke;
}

export type Element = ShapeElement | PathElement | ImageElement | TextElement;

export interface Slide {
  id: string;
  name: string;
  wPx: number;
  hPx: number;
  background?: SolidFill;
  notes?: string;
  elements: Element[];
}

export interface SlideSummary {
  id: string;
  name: string;
  wPx: number;
  hPx: number;
  /** Figma Slides "skip" state. Skipped slides are unchecked by default. */
  skipped: boolean;
}

export interface Warning {
  slideName: string;
  nodeName: string;
  message: string;
}

export interface ExportOptions {
  /** 'frame' derives the slide height from the frame aspect; the others letterbox into a fixed ratio. */
  layout: 'frame' | '16:9' | '4:3';
  /**
   * Physical slide width. 20 in maps 1920 px at 96 DPI (1 px = 9525 EMU = 0.75 pt), the convention
   * Pitchdeck uses; 13.333 in is PowerPoint's own widescreen default (1 px = 0.5 pt).
   */
  slideWidthIn: number;
  /** false = every slide becomes a single flattened image. */
  editable: boolean;
  /** 'shapes' converts vectors to editable freeforms; 'svg' embeds them as SVG (crisp, PowerPoint 2016+); 'raster' as PNG. */
  vectors: 'shapes' | 'svg' | 'raster';
  /** Export scale for rasterized layers. */
  imageScale: 1 | 2 | 3 | 4;
  /** 'auto' picks JPG for opaque rectangular photos and PNG otherwise. */
  imageFormat: 'auto' | 'png' | 'jpg';
  order: 'position' | 'layers' | 'name' | 'selection';
  /** A text layer whose name starts with this (case-insensitive) becomes speaker notes instead of slide content. */
  notesLayerName: string;
  /**
   * 'family' → "Inter" (bold/italic as flags); 'family-style' → "Inter SemiBold" for weights PowerPoint
   * cannot express as flags; 'full' → always "Inter Regular" / "Inter Semi Bold Italic" (Pitchdeck's convention).
   */
  fontFace: 'family' | 'family-style' | 'full';
  /**
   * Percentage added to the width of fixed-width text boxes. PowerPoint's glyph advances differ a
   * little from Figma's, so lines can wrap early (positive values widen) or late (negative narrows).
   */
  textWidthAdjustPct: number;
  /**
   * Multiplier on every font size, letter spacing and line height. Figma's own PowerPoint export
   * uses 0.95 to stop PowerPoint, whose glyph advances run wider than Figma's, from wrapping early.
   */
  fontScale: number;
  fileName: string;
}

export const DEFAULT_OPTIONS: ExportOptions = {
  layout: 'frame',
  slideWidthIn: 20,
  editable: true,
  vectors: 'shapes',
  imageScale: 2,
  imageFormat: 'auto',
  order: 'position',
  notesLayerName: 'notes',
  fontFace: 'family-style',
  textWidthAdjustPct: 0,
  fontScale: 1,
  fileName: '',
};

// ---- Message protocol between the UI iframe and the plugin sandbox ------------------------------

export type UiToPluginMessage =
  | { type: 'ui-ready' }
  | { type: 'export'; options: ExportOptions; slideIds: string[] }
  | { type: 'cancel' }
  | { type: 'save-options'; options: ExportOptions }
  | { type: 'focus-slide'; slideId: string }
  | { type: 'notify'; message: string; error?: boolean };

export type PluginToUiMessage =
  | {
      type: 'init';
      editorType: string;
      fileName: string;
      options: ExportOptions | null;
      slides: SlideSummary[];
    }
  | { type: 'slides'; slides: SlideSummary[] }
  | { type: 'deck-start'; total: number; wPx: number; hPx: number }
  | { type: 'slide'; slide: Slide; index: number; total: number }
  | { type: 'deck-end'; warnings: Warning[]; cancelled: boolean }
  | { type: 'error'; message: string };
