/**
 * Headless entry: runs the extractor without the panel, for scripted use (e.g. the Figma MCP
 * `use_figma` tool, or automated fidelity runs). Returns slides with image bytes base64-encoded
 * so the result survives JSON serialisation; `tools/` turns it back into a .pptx.
 */
import { DEFAULT_OPTIONS, type ExportOptions, type Slide, type Warning } from '../shared/types';
import { extractSlide, isSlideRoot, type SlideRoot } from './extract';

export interface HeadlessSlide extends Omit<Slide, 'elements'> {
  elements: Array<Record<string, unknown>>;
}

export interface HeadlessResult {
  wPx: number;
  hPx: number;
  slides: HeadlessSlide[];
  warnings: Warning[];
}

export interface HeadlessOptions {
  /** false → image elements carry no bytes (fetch them separately by node id); keeps payloads small. */
  imageBytes?: boolean;
  /** Return only elements [from, to) of each slide, for paging through large slides. */
  elementRange?: [number, number];
  /** Decimal places kept for freeform path coordinates (px). 0 is visually lossless at slide scale. */
  pathDecimals?: number;
}

export async function extractHeadless(
  slideIds: string[],
  overrides: Partial<ExportOptions> = {},
  opts: HeadlessOptions = {},
): Promise<HeadlessResult> {
  const options: ExportOptions = { ...DEFAULT_OPTIONS, ...overrides };
  const roots: SlideRoot[] = [];
  for (const id of slideIds) {
    const n = await figma.getNodeByIdAsync(id);
    if (n && isSlideRoot(n)) roots.push(n);
  }
  const warnings: Warning[] = [];
  const slideIdSet = new Set(roots.map((r) => r.id));
  const slides: HeadlessSlide[] = [];
  for (const root of roots) {
    const slide = await extractSlide(root, { options, slideIds: slideIdSet, warnings });
    if (root.type === 'SLIDE' && 'speakerNotes' in root) {
      const notes = String((root as unknown as { speakerNotes: string }).speakerNotes ?? '').trim();
      if (notes) slide.notes = slide.notes ? `${slide.notes}\n\n${notes}` : notes;
    }
    if (opts.pathDecimals !== undefined) {
      const f = 10 ** opts.pathDecimals;
      for (const el of slide.elements) {
        if (el.type === 'path') {
          for (const c of el.commands) {
            for (const k of Object.keys(c) as (keyof typeof c)[]) {
              if (typeof c[k] === 'number') (c as unknown as Record<string, number>)[k] = Math.round((c[k] as number) * f) / f;
            }
          }
        }
      }
    }
    const [from, to] = opts.elementRange ?? [0, slide.elements.length];
    slides.push({
      ...slide,
      elements: slide.elements.slice(from, to).map((el) =>
        el.type === 'image'
          ? { ...el, bytes: undefined, bytesB64: opts.imageBytes === false ? undefined : toBase64(el.bytes) }
          : (el as unknown as Record<string, unknown>),
      ),
    });
  }
  const first = roots[0];
  return { wPx: first ? Math.round(first.width) : 1920, hPx: first ? Math.round(first.height) : 1080, slides, warnings };
}

function toBase64(bytes: Uint8Array): string {
  let binary = '';
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) binary += String.fromCharCode.apply(null, Array.from(bytes.subarray(i, i + chunk)));
  return btoa(binary);
}

/** Compact JSON: drops default-valued keys so payloads fit small transport limits. Inverse: expandCompact(). */
export function compactJson(result: HeadlessResult): string {
  return JSON.stringify(result, (key, value) => {
    if (value === false || value === undefined || value === null) return undefined;
    if ((key === 'rotation' || key === 'indent' || key === 'spacingAfterPx') && value === 0) return undefined;
    if ((key === 'alpha' || key === 'opacity') && value === 1) return undefined;
    if (key === 'list' && value === 'none') return undefined;
    if (key === 'fontWeight' && value === 400) return undefined;
    if (key === 'fontStyle' && value === 'Regular') return undefined;
    if (typeof value === 'number') return Math.round(value * 100) / 100;
    return value;
  });
}
