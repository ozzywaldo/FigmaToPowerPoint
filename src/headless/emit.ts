/**
 * Node entry: turns headless extraction results (JSON with base64 images) into a .pptx.
 *   node dist/emit.mjs out.pptx [--options '{"fontScale":0.95}'] part1.json part2.json …
 * Each part is a HeadlessResult; slides are concatenated in argument order.
 */
import fs from 'node:fs/promises';
import PptxGenJS from 'pptxgenjs';
import { DEFAULT_OPTIONS, type ExportOptions, type Slide } from '../shared/types';
import { buildPresentation } from '../ui/pptx';
import type { HeadlessResult } from '../plugin/headless';

const args = process.argv.slice(2);
const out = args.shift();
if (!out) throw new Error('usage: emit.mjs out.pptx [--options json] parts…');
let overrides: Partial<ExportOptions> = {};
const oi = args.indexOf('--options');
if (oi >= 0) {
  overrides = JSON.parse(args[oi + 1]);
  args.splice(oi, 2);
}
// --images dir: PNG/JPG files named after the node id (':' → '-') for image elements without bytes.
let imagesDir: string | undefined;
const ii = args.indexOf('--images');
if (ii >= 0) {
  imagesDir = args[ii + 1];
  args.splice(ii, 2);
}
const missing: string[] = [];
async function imageBytes(el: Record<string, unknown>): Promise<{ bytes: Uint8Array; format: 'png' | 'jpg' } | null> {
  if (typeof el.bytesB64 === 'string') return { bytes: new Uint8Array(Buffer.from(el.bytesB64, 'base64')), format: el.format as 'png' | 'jpg' };
  if (!imagesDir) return null;
  const base = `${imagesDir}/${String(el.id).replace(/[:;]/g, '-')}`;
  for (const ext of ['png', 'jpg', 'jpeg']) {
    try {
      const bytes = new Uint8Array(await fs.readFile(`${base}.${ext}`));
      return { bytes, format: ext === 'png' ? 'png' : 'jpg' };
    } catch {
      /* try next */
    }
  }
  missing.push(String(el.id));
  return null;
}

const slides: Slide[] = [];
let wPx = 1920;
let hPx = 1080;
for (const file of args) {
  const part = JSON.parse(await fs.readFile(file, 'utf8')) as HeadlessResult;
  wPx = part.wPx;
  hPx = part.hPx;
  for (const s of part.slides) {
    const elements: Record<string, unknown>[] = [];
    for (const el of s.elements) {
      const e = expandCompact(el);
      if (e.type === 'image') {
        const img = await imageBytes(e);
        if (!img) continue;
        const { bytesB64, ...rest } = e as { bytesB64?: string } & Record<string, unknown>;
        elements.push({ ...rest, bytes: img.bytes, format: img.format });
      } else {
        elements.push(e);
      }
    }
    slides.push({ ...s, elements: elements as unknown as Slide['elements'] });
  }
}
if (missing.length) console.warn(`no image file for ${missing.length} image element(s): ${missing.join(', ')}`);

/** Inverse of compactJson(): restore default-valued keys the extractor omitted. */
function expandCompact(el: Record<string, unknown>): Record<string, unknown> {
  const e: Record<string, unknown> = { rotation: 0, flipH: false, flipV: false, ...el };
  if (e.type === 'image') e.opacity ??= 1;
  if (e.type === 'text') {
    for (const p of e.paragraphs as Record<string, unknown>[]) {
      p.list ??= 'none';
      p.indent ??= 0;
      p.spacingAfterPx ??= 0;
      for (const r of p.runs as Record<string, unknown>[]) {
        r.fontWeight ??= 400;
        r.fontStyle ??= 'Regular';
        r.alpha ??= 1;
        r.italic ??= false;
        r.underline ??= false;
        r.strike ??= false;
      }
    }
  }
  for (const k of ['fill', 'stroke', 'shadow']) {
    const v = e[k] as Record<string, unknown> | undefined;
    if (v && v.alpha === undefined) v.alpha = 1;
  }
  return e;
}
const options: ExportOptions = { ...DEFAULT_OPTIONS, fileName: 'deck', ...overrides };
const pptx = buildPresentation(PptxGenJS, { wPx, hPx, slides, options });
const buf = (await pptx.write({ outputType: 'nodebuffer' })) as Buffer;
await fs.writeFile(out, buf);
console.log(`${out}: ${slides.length} slides, ${(buf.length / 1048576).toFixed(1)} MB`);
