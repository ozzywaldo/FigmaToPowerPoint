/** Text node → TextElement. */
import type { Link, Paragraph, TextElement, TextRun } from '../shared/types';
import { clamp01, hasImagePaint, isGradient, solidFromPaints, toHex, visiblePaints } from './paint';

export interface TextContext {
  opacity: number;
  resolveLink: (target: HyperlinkTarget | null) => Link | undefined;
  warn: (node: SceneNode, message: string) => void;
}

const SEGMENT_FIELDS = [
  'fontName',
  'fontSize',
  'fontWeight',
  'fontStyle',
  'fills',
  'textDecoration',
  'letterSpacing',
  'lineHeight',
  'hyperlink',
  'listOptions',
  'indentation',
  'textCase',
  'paragraphSpacing',
] as const;

type Segment = Pick<StyledTextSegment, (typeof SEGMENT_FIELDS)[number] | 'characters' | 'start' | 'end'>;

/** Returns `null` when the text must be rasterized (image fills), `undefined` when it is empty. */
export function extractTextContent(
  node: TextNode,
  ctx: TextContext,
): Pick<TextElement, 'paragraphs' | 'align' | 'valign' | 'autoResize' | 'lineHeight'> | null | undefined {
  if (!node.characters) return undefined;
  if (hasImagePaint(node.fills)) {
    ctx.warn(node, 'Text with an image fill was exported as an image.');
    return null;
  }
  const segments = node.getStyledTextSegments([...SEGMENT_FIELDS]) as Segment[];
  if (segments.some((s) => isGradient(s.fills))) {
    ctx.warn(node, 'Gradient text fill was flattened to the first gradient stop.');
  }

  // Paragraph-level properties (list, indent, spacing) belong to the first segment that lands
  // in a paragraph — a segment ending in "\n" says nothing about the paragraph that follows it.
  const paragraphs: Paragraph[] = [];
  let current: Paragraph | null = null;
  const paragraphFor = (seg: Segment): Paragraph => {
    if (!current) {
      current = newParagraph(seg);
      paragraphs.push(current);
    }
    return current;
  };
  for (const seg of segments) {
    const parts = seg.characters.split('\n');
    parts.forEach((part, idx) => {
      if (idx > 0) {
        paragraphFor(seg); // the "\n" closes a paragraph, even an empty one
        current = null;
      }
      const lines = part.split('\u2028');
      lines.forEach((line, li) => {
        if (line.length === 0 && li === 0) return;
        const run = runFrom(seg, line, ctx);
        if (li > 0) run.softBreakBefore = true;
        paragraphFor(seg).runs.push(run);
      });
    });
  }
  if (!current) paragraphs.push(newParagraph(segments[segments.length - 1]));

  return {
    paragraphs,
    align: mapAlign(node.textAlignHorizontal),
    valign: node.textAlignVertical === 'CENTER' ? 'middle' : node.textAlignVertical === 'BOTTOM' ? 'bottom' : 'top',
    autoResize: mapAutoResize(node.textAutoResize),
    lineHeight: mapLineHeight(node, segments, ctx),
  };
}

function newParagraph(seg: Segment | undefined): Paragraph {
  const list = seg?.listOptions?.type;
  return {
    runs: [],
    list: list === 'ORDERED' ? 'number' : list === 'UNORDERED' ? 'bullet' : 'none',
    indent: seg?.indentation ?? 0,
    spacingAfterPx: seg?.paragraphSpacing ?? 0,
  };
}

function runFrom(seg: Segment, text: string, ctx: TextContext): TextRun {
  const fill = solidFromPaints(seg.fills, ctx.opacity, true) ?? firstSolid(seg.fills, ctx.opacity);
  const style = seg.fontName.style;
  return {
    text: applyCase(text, seg.textCase),
    fontFamily: seg.fontName.family,
    fontStyle: style,
    fontWeight: seg.fontWeight,
    italic: seg.fontStyle === 'ITALIC' || /italic|oblique/i.test(style),
    fontSizePx: seg.fontSize,
    color: fill.color,
    alpha: fill.alpha,
    underline: seg.textDecoration === 'UNDERLINE',
    strike: seg.textDecoration === 'STRIKETHROUGH',
    letterSpacing:
      seg.letterSpacing.value !== 0
        ? { unit: seg.letterSpacing.unit === 'PERCENT' ? 'percent' : 'px', value: seg.letterSpacing.value }
        : undefined,
    link: ctx.resolveLink(seg.hyperlink),
  };
}

function firstSolid(paints: readonly Paint[], opacity: number): { color: string; alpha: number } {
  const solid = visiblePaints(paints).find((p): p is SolidPaint => p.type === 'SOLID');
  if (solid) return { color: toHex(solid.color), alpha: clamp01((solid.opacity ?? 1) * opacity) };
  return { color: '000000', alpha: visiblePaints(paints).length ? opacity : 0 };
}

function applyCase(text: string, textCase: TextCase): string {
  switch (textCase) {
    case 'UPPER':
    case 'SMALL_CAPS':
    case 'SMALL_CAPS_FORCED':
      return text.toUpperCase();
    case 'LOWER':
      return text.toLowerCase();
    case 'TITLE':
      return text.replace(/(^|\s)(\p{L})/gu, (_m, sp: string, ch: string) => sp + ch.toUpperCase());
    default:
      return text;
  }
}

function mapAlign(a: TextNode['textAlignHorizontal']): TextElement['align'] {
  return a === 'CENTER' ? 'center' : a === 'RIGHT' ? 'right' : a === 'JUSTIFIED' ? 'justify' : 'left';
}

function mapAutoResize(a: TextNode['textAutoResize']): TextElement['autoResize'] {
  return a === 'HEIGHT' ? 'height' : a === 'WIDTH_AND_HEIGHT' ? 'width_height' : a === 'TRUNCATE' ? 'truncate' : 'none';
}

function mapLineHeight(node: TextNode, segments: Segment[], ctx: TextContext): TextElement['lineHeight'] {
  const lh = node.lineHeight !== figma.mixed ? node.lineHeight : segments[0]?.lineHeight;
  if (node.lineHeight === figma.mixed) {
    ctx.warn(node, 'Mixed line heights in one text layer; PowerPoint supports one per text box. The first was used.');
  }
  if (!lh || lh.unit === 'AUTO') return undefined;
  return { unit: lh.unit === 'PERCENT' ? 'percent' : 'px', value: lh.value };
}
