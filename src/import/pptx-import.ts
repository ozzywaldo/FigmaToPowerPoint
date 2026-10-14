/**
 * .pptx → Slide model.
 *
 * Reads a PowerPoint file back into the same intermediate model the Figma extractor produces.
 * Two uses: (1) round-trip testing — push a deck exported by another tool (e.g. Pitchdeck) through
 * our emitter and compare renders; (2) the seed of a PPTX *import* feature, since the model can be
 * turned into Figma nodes just as it is turned into pptx.
 *
 * Coverage is deliberately the subset our emitter writes: text boxes, preset rect/roundRect/
 * ellipse/line shapes with solid fills and lines, pictures (PNG/JPEG; the PNG fallback of SVGs),
 * groups, rotation/flips, outer shadows, hyperlinks, notes, and solid slide backgrounds.
 */
import JSZip from 'jszip';
import { XMLParser } from 'fast-xml-parser';
import type { Element, ImageElement, Link, Paragraph, Shadow, ShapeElement, Slide, SolidFill, Stroke, TextElement, TextRun } from '../shared/types';

export interface ImportedDeck {
  wPx: number;
  hPx: number;
  slides: Slide[];
}

export interface ImportOptions {
  /** Pixels per inch used to map EMU → px. 96 makes a 20 in deck 1920 px wide (Pitchdeck's convention). */
  dpi?: number;
}

const EMU_PER_IN = 914400;

// preserveOrder nodes: { tag: XNode[], ':@'?: { '@_name': value } } or { '#text': string }
type XNode = Record<string, unknown>;

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  preserveOrder: true,
  removeNSPrefix: false,
  parseTagValue: false,
  trimValues: false,
});

export async function importPptx(bytes: Uint8Array | ArrayBuffer, options: ImportOptions = {}): Promise<ImportedDeck> {
  const dpi = options.dpi ?? 96;
  const zip = await JSZip.loadAsync(bytes);
  const read = (p: string): Promise<string | null> => zip.file(p)?.async('string') ?? Promise.resolve(null);

  const pres = parse(await must(read('ppt/presentation.xml'), 'ppt/presentation.xml'));
  const sldSz = findDeep(pres, 'p:sldSz');
  const cx = num(attr(sldSz, 'cx'));
  const cy = num(attr(sldSz, 'cy'));
  const emuPerPx = EMU_PER_IN / dpi;
  const u = {
    px: (emu: number) => round(emu / emuPerPx),
    ptToPx: (pt: number) => round((pt / 72) * dpi),
  };

  // Slide order comes from presentation.xml's sldIdLst + its rels.
  const presRels = parseRels(await must(read('ppt/_rels/presentation.xml.rels'), 'presentation rels'));
  const sldIds = findAllDeep(pres, 'p:sldId').map((n) => str(attr(n, 'r:id')));
  const slidePaths = sldIds.map((rid) => resolvePath('ppt/', presRels.get(rid) ?? ''));
  const slideIdByPath = new Map(slidePaths.map((p, i) => [p, `s${i + 1}`]));

  const slides: Slide[] = [];
  for (let i = 0; i < slidePaths.length; i++) {
    const slidePath = slidePaths[i];
    const xml = await must(read(slidePath), slidePath);
    const relsPath = slidePath.replace(/slides\/(slide\d+\.xml)$/, 'slides/_rels/$1.rels');
    const rels = parseRels((await read(relsPath)) ?? '');
    const relTargets = new Map([...rels].map(([k, v]) => [k, resolvePath(path.dirname(slidePath) + '/', v)]));
    const doc = parse(xml);
    const ctx: Ctx = {
      u,
      zip,
      rels: relTargets,
      relsRaw: rels,
      slideIdByPath,
      elements: [],
      xf: { ox: 0, oy: 0, sx: 1, sy: 1 },
    };
    const spTree = findDeep(doc, 'p:spTree');
    if (spTree) await walkTree(spTree, ctx);
    const slide: Slide = {
      id: `s${i + 1}`,
      name: `Slide ${i + 1}`,
      wPx: u.px(cx),
      hPx: u.px(cy),
      elements: ctx.elements,
    };
    const bg = findDeep(findDeep(doc, 'p:bg'), 'a:solidFill');
    if (bg) slide.background = solidFill(bg) ?? undefined;
    const notesRel = [...rels.entries()].find(([, t]) => /notesSlide\d+\.xml$/.test(t));
    if (notesRel) {
      const notesXml = await read(resolvePath(path.dirname(slidePath) + '/', notesRel[1]));
      if (notesXml) {
        const notes = notesText(parse(notesXml));
        if (notes) slide.notes = notes;
      }
    }
    slides.push(slide);
  }
  return { wPx: u.px(cx), hPx: u.px(cy), slides };
}

// ---- shapes -------------------------------------------------------------------------------

interface Units {
  px: (emu: number) => number;
  ptToPx: (pt: number) => number;
}
interface GroupXf {
  ox: number;
  oy: number;
  sx: number;
  sy: number;
}
interface Ctx {
  u: Units;
  zip: JSZip;
  rels: Map<string, string>;
  relsRaw: Map<string, string>;
  slideIdByPath: Map<string, string>;
  elements: Element[];
  xf: GroupXf;
}

async function walkTree(tree: XNode, ctx: Ctx): Promise<void> {
  for (const node of kids(tree)) {
    const tag = tagOf(node);
    if (tag === 'p:sp' || tag === 'p:cxnSp') addShapeOrText(node, ctx);
    else if (tag === 'p:pic') await addPicture(node, ctx);
    else if (tag === 'p:grpSp') await addGroup(node, ctx);
  }
}

async function addGroup(node: XNode, ctx: Ctx): Promise<void> {
  const xfrm = findDeep(child(node, 'p:grpSpPr'), 'a:xfrm');
  const off = child(xfrm, 'a:off');
  const ext = child(xfrm, 'a:ext');
  const chOff = child(xfrm, 'a:chOff');
  const chExt = child(xfrm, 'a:chExt');
  const parent = ctx.xf;
  const sx = chExt && num(attr(chExt, 'cx')) ? num(attr(ext, 'cx')) / num(attr(chExt, 'cx')) : 1;
  const sy = chExt && num(attr(chExt, 'cy')) ? num(attr(ext, 'cy')) / num(attr(chExt, 'cy')) : 1;
  const inner: GroupXf = {
    sx: parent.sx * sx,
    sy: parent.sy * sy,
    ox: parent.ox + parent.sx * (num(attr(off, 'x')) - sx * num(attr(chOff, 'x'))),
    oy: parent.oy + parent.sy * (num(attr(off, 'y')) - sy * num(attr(chOff, 'y'))),
  };
  ctx.xf = inner;
  try {
    await walkTree(node, ctx);
  } finally {
    ctx.xf = parent;
  }
}

function placement(node: XNode, ctx: Ctx, prTag: string) {
  const xfrm = findDeep(child(node, prTag), 'a:xfrm');
  const off = child(xfrm, 'a:off');
  const ext = child(xfrm, 'a:ext');
  const x = ctx.xf.ox + ctx.xf.sx * num(attr(off, 'x'));
  const y = ctx.xf.oy + ctx.xf.sy * num(attr(off, 'y'));
  const w = ctx.xf.sx * num(attr(ext, 'cx'));
  const h = ctx.xf.sy * num(attr(ext, 'cy'));
  let rotation = num(attr(xfrm, 'rot')) / 60000;
  rotation = ((rotation % 360) + 360) % 360;
  return {
    x: ctx.u.px(x),
    y: ctx.u.px(y),
    w: ctx.u.px(w),
    h: ctx.u.px(h),
    rotation: round(rotation),
    flipH: attr(xfrm, 'flipH') === '1',
    flipV: attr(xfrm, 'flipV') === '1',
  };
}

function nameOf(node: XNode): { id: string; name: string } {
  const cNvPr = findDeep(node, 'p:cNvPr');
  return { id: `n${str(attr(cNvPr, 'id'))}`, name: str(attr(cNvPr, 'name')) };
}

function addShapeOrText(node: XNode, ctx: Ctx): void {
  const spPr = child(node, 'p:spPr');
  const base = { ...nameOf(node), ...placement(node, ctx, 'p:spPr'), link: linkOf(findDeep(child(node, 'p:nvSpPr') ?? child(node, 'p:nvCxnSpPr'), 'a:hlinkClick'), ctx) };
  const fill = solidFill(child(spPr, 'a:solidFill'));
  const stroke = strokeOf(child(spPr, 'a:ln'), ctx);
  const shadow = shadowOf(findDeep(child(spPr, 'a:effectLst'), 'a:outerShdw'), ctx);
  const txBody = child(node, 'p:txBody');
  const paragraphs = txBody ? paragraphsOf(txBody, ctx) : [];
  const hasText = paragraphs.some((p) => p.runs.some((r) => r.text.length > 0));

  const prst = str(attr(child(spPr, 'a:prstGeom'), 'prst'));
  if (fill || stroke) {
    const shape: ShapeElement['shape'] = prst === 'ellipse' ? 'ellipse' : prst === 'line' || prst === 'straightConnector1' ? 'line' : 'rect';
    const el: ShapeElement = { type: 'shape', shape, ...base, fill: fill ?? undefined, stroke: stroke ?? undefined, shadow: shadow ?? undefined };
    if (prst === 'roundRect') {
      const gd = findDeep(child(spPr, 'a:prstGeom'), 'a:gd');
      const adj = gd ? num(str(attr(gd, 'fmla')).replace('val ', '')) : 16667;
      el.cornerRadiusPx = round((adj / 100000) * Math.min(el.w, el.h));
    }
    if (!hasText) {
      ctx.elements.push(el);
      return;
    }
    ctx.elements.push({ ...el, link: undefined });
  }
  if (!hasText) return;

  const bodyPr = child(txBody!, 'a:bodyPr');
  const anchor = str(attr(bodyPr, 'anchor'));
  const firstPPr = findDeep(child(txBody!, 'a:p'), 'a:pPr');
  const el: TextElement = {
    type: 'text',
    ...base,
    paragraphs,
    align: alignOf(str(attr(firstPPr, 'algn'))),
    valign: anchor === 'ctr' ? 'middle' : anchor === 'b' ? 'bottom' : 'top',
    autoResize: attr(bodyPr, 'wrap') === 'none' ? 'width_height' : 'height',
    lineHeight: lineHeightOf(child(firstPPr, 'a:lnSpc'), ctx),
    shadow: shadow ?? undefined,
  };
  ctx.elements.push(el);
}

async function addPicture(node: XNode, ctx: Ctx): Promise<void> {
  const blip = findDeep(child(node, 'p:blipFill'), 'a:blip');
  const rid = str(attr(blip, 'r:embed'));
  const target = ctx.rels.get(rid);
  if (!target) return;
  const file = ctx.zip.file(target);
  if (!file) return;
  const ext = target.toLowerCase().split('.').pop();
  if (ext !== 'png' && ext !== 'jpg' && ext !== 'jpeg') return; // SVG without fallback etc.
  const bytes = await file.async('uint8array');
  const alpha = findDeep(blip, 'a:alphaModFix');
  const el: ImageElement = {
    type: 'image',
    ...nameOf(node),
    ...placement(node, ctx, 'p:spPr'),
    format: ext === 'png' ? 'png' : 'jpg',
    bytes,
    opacity: alpha ? num(attr(alpha, 'amt')) / 100000 : 1,
    link: linkOf(findDeep(child(node, 'p:nvPicPr'), 'a:hlinkClick'), ctx),
  };
  ctx.elements.push(el);
}

// ---- text ---------------------------------------------------------------------------------

function paragraphsOf(txBody: XNode, ctx: Ctx): Paragraph[] {
  return childrenOf(txBody, 'a:p').map((p) => {
    const pPr = child(p, 'a:pPr');
    const list: Paragraph['list'] = child(pPr, 'a:buAutoNum') ? 'number' : child(pPr, 'a:buChar') ? 'bullet' : 'none';
    const lvl = num(attr(pPr, 'lvl'));
    const spcAft = findDeep(child(pPr, 'a:spcAft'), 'a:spcPts');
    const runs: TextRun[] = [];
    let softBreak = false;
    for (const n of kids(p)) {
      const tag = tagOf(n);
      if (tag === 'a:br') {
        softBreak = true;
      } else if (tag === 'a:r') {
        const run = runOf(n, ctx);
        if (softBreak) run.softBreakBefore = true;
        softBreak = false;
        runs.push(run);
      }
    }
    return {
      runs,
      list,
      indent: list === 'none' ? lvl : lvl + 1,
      spacingAfterPx: spcAft ? ctx.u.ptToPx(num(attr(spcAft, 'val')) / 100) : 0,
    };
  });
}

function runOf(r: XNode, ctx: Ctx): TextRun {
  const rPr = child(r, 'a:rPr');
  const sizePt = num(attr(rPr, 'sz')) / 100 || 18;
  const face = str(attr(child(rPr, 'a:latin'), 'typeface')) || 'Calibri';
  const { family, style } = splitFontFace(face);
  const bold = attr(rPr, 'b') === '1';
  const italic = attr(rPr, 'i') === '1' || /italic|oblique/i.test(style);
  const fill = solidFill(child(rPr, 'a:solidFill')) ?? { color: '000000', alpha: 1 };
  const spc = num(attr(rPr, 'spc'));
  return {
    text: childrenOf(r, 'a:t').map((t) => textOf(t)).join(''),
    fontFamily: family,
    fontStyle: style,
    fontWeight: bold ? 700 : weightOf(style),
    italic,
    fontSizePx: ctx.u.ptToPx(sizePt),
    color: fill.color,
    alpha: fill.alpha,
    underline: !!attr(rPr, 'u') && attr(rPr, 'u') !== 'none',
    strike: !!attr(rPr, 'strike') && attr(rPr, 'strike') !== 'noStrike',
    letterSpacing: spc ? { unit: 'px', value: ctx.u.ptToPx(spc / 100) } : undefined,
    link: linkOf(child(rPr, 'a:hlinkClick'), ctx),
  };
}

const STYLE_WORDS = /\s+(Regular|Bold|Italic|Light|Medium|Semi ?Bold|Demi ?Bold|Extra ?Bold|Ultra ?Bold|Extra ?Light|Ultra ?Light|Thin|Black|Heavy|Book|Oblique|Condensed|Display|Text)+$/i;

/** "Lato Regular" → { family: "Lato", style: "Regular" }; "Instrument Serif Italic" → { "Instrument Serif", "Italic" }. */
export function splitFontFace(face: string): { family: string; style: string } {
  let family = face.trim();
  const styles: string[] = [];
  let m: RegExpMatchArray | null;
  while ((m = family.match(STYLE_WORDS))) {
    styles.unshift(m[1]);
    family = family.slice(0, m.index).trim();
  }
  return { family: family || face, style: styles.join(' ') || 'Regular' };
}

function weightOf(style: string): number {
  const s = style.toLowerCase().replace(/\s/g, '');
  if (s.includes('thin')) return 100;
  if (s.includes('extralight') || s.includes('ultralight')) return 200;
  if (s.includes('light')) return 300;
  if (s.includes('medium')) return 500;
  if (s.includes('semibold') || s.includes('demibold')) return 600;
  if (s.includes('extrabold') || s.includes('ultrabold')) return 800;
  if (s.includes('black') || s.includes('heavy')) return 900;
  if (s.includes('bold')) return 700;
  return 400;
}

function alignOf(algn: string): TextElement['align'] {
  return algn === 'ctr' ? 'center' : algn === 'r' ? 'right' : algn === 'just' ? 'justify' : 'left';
}

function lineHeightOf(lnSpc: XNode | undefined, ctx: Ctx): TextElement['lineHeight'] {
  if (!lnSpc) return undefined;
  const pts = child(lnSpc, 'a:spcPts');
  if (pts) return { unit: 'px', value: ctx.u.ptToPx(num(attr(pts, 'val')) / 100) };
  const pct = child(lnSpc, 'a:spcPct');
  // Inverse of the emitter: Figma % of font size ≈ PowerPoint % of natural line height × 1.2
  if (pct) return { unit: 'percent', value: round((num(attr(pct, 'val')) / 100000) * 120) };
  return undefined;
}

function notesText(doc: XNode[]): string {
  const out: string[] = [];
  for (const sp of findAllDeep(doc, 'p:sp')) {
    const ph = findDeep(sp, 'p:ph');
    if (!ph || attr(ph, 'type') !== 'body') continue;
    for (const p of findAllDeep(sp, 'a:p')) out.push(findAllDeep(p, 'a:t').map(textOf).join(''));
  }
  return out.join('\n').trim();
}

// ---- paints -------------------------------------------------------------------------------

function solidFill(node: XNode | undefined): SolidFill | null {
  if (!node) return null;
  const clr = child(node, 'a:srgbClr');
  if (!clr) return null;
  const alphaNode = child(clr, 'a:alpha');
  return { color: str(attr(clr, 'val')).toUpperCase(), alpha: alphaNode ? num(attr(alphaNode, 'val')) / 100000 : 1 };
}

function strokeOf(ln: XNode | undefined, ctx: Ctx): Stroke | null {
  if (!ln || child(ln, 'a:noFill')) return null;
  const fill = solidFill(child(ln, 'a:solidFill'));
  if (!fill) return null;
  const wEmu = num(attr(ln, 'w')) || 12700;
  const dash = str(attr(child(ln, 'a:prstDash'), 'val'));
  return {
    color: fill.color,
    alpha: fill.alpha,
    widthPx: ctx.u.px(wEmu),
    dashed: !!dash && dash !== 'solid',
    arrowStart: !!child(ln, 'a:headEnd') && attr(child(ln, 'a:headEnd'), 'type') !== 'none',
    arrowEnd: !!child(ln, 'a:tailEnd') && attr(child(ln, 'a:tailEnd'), 'type') !== 'none',
  };
}

function shadowOf(shdw: XNode | undefined, ctx: Ctx): Shadow | null {
  if (!shdw) return null;
  const fill = solidFill(shdw) ?? (child(shdw, 'a:prstClr') ? { color: '000000', alpha: 1 } : null);
  if (!fill) return null;
  const dist = ctx.u.px(num(attr(shdw, 'dist')));
  const dir = (num(attr(shdw, 'dir')) / 60000) * (Math.PI / 180);
  const alphaNode = findDeep(shdw, 'a:alpha');
  return {
    color: fill.color,
    alpha: alphaNode ? num(attr(alphaNode, 'val')) / 100000 : fill.alpha,
    offsetXPx: round(Math.cos(dir) * dist),
    offsetYPx: round(Math.sin(dir) * dist),
    blurPx: ctx.u.px(num(attr(shdw, 'blurRad'))),
  };
}

function linkOf(hlink: XNode | undefined, ctx: Ctx): Link | undefined {
  if (!hlink) return undefined;
  const rid = str(attr(hlink, 'r:id'));
  const action = str(attr(hlink, 'action'));
  if (action.includes('hlinksldjump')) {
    const target = ctx.rels.get(rid);
    const id = target ? ctx.slideIdByPath.get(target) : undefined;
    return id ? { kind: 'slide', slideId: id } : undefined;
  }
  const url = ctx.relsRaw.get(rid);
  return url ? { kind: 'url', url } : undefined;
}

// ---- xml helpers --------------------------------------------------------------------------

function parse(xml: string): XNode[] {
  return parser.parse(xml) as XNode[];
}

function parseRels(xml: string): Map<string, string> {
  const map = new Map<string, string>();
  if (!xml) return map;
  for (const rel of findAllDeep(parse(xml), 'Relationship')) map.set(str(attr(rel, 'Id')), str(attr(rel, 'Target')));
  return map;
}

function tagOf(n: XNode): string {
  return Object.keys(n).find((k) => k !== ':@') ?? '';
}
function kids(n: XNode | undefined): XNode[] {
  if (!n) return [];
  const v = n[tagOf(n)];
  return Array.isArray(v) ? (v as XNode[]) : [];
}
function attr(n: XNode | undefined, name: string): string | undefined {
  const a = n?.[':@'] as Record<string, unknown> | undefined;
  const v = a?.[`@_${name}`];
  return v === undefined ? undefined : String(v);
}
function child(n: XNode | undefined, tag: string): XNode | undefined {
  return kids(n).find((c) => tagOf(c) === tag);
}
function childrenOf(n: XNode | undefined, tag: string): XNode[] {
  return kids(n).filter((c) => tagOf(c) === tag);
}
function findDeep(root: XNode | XNode[] | undefined, tag: string): XNode | undefined {
  const list = Array.isArray(root) ? root : root ? kids(root) : [];
  for (const c of list) {
    if (tagOf(c) === tag) return c;
    const hit = findDeep(c, tag);
    if (hit) return hit;
  }
  return undefined;
}
function findAllDeep(root: XNode | XNode[] | undefined, tag: string, out: XNode[] = []): XNode[] {
  const list = Array.isArray(root) ? root : root ? kids(root) : [];
  for (const c of list) {
    if (tagOf(c) === tag) out.push(c);
    findAllDeep(c, tag, out);
  }
  return out;
}
function textOf(n: XNode): string {
  return kids(n)
    .map((c) => (c['#text'] !== undefined ? String(c['#text']) : ''))
    .join('');
}
function num(v: unknown): number {
  const n = Number(v);
  return Number.isFinite(n) ? n : 0;
}
function str(v: unknown): string {
  return v === undefined || v === null ? '' : String(v);
}
function round(n: number): number {
  return Math.round(n * 1000) / 1000;
}
async function must(v: Promise<string | null> | string | null, what: string): Promise<string> {
  const r = await v;
  if (r === null) throw new Error(`pptx is missing ${what}`);
  return r;
}

const path = {
  dirname: (p: string) => p.replace(/\/[^/]*$/, ''),
};

/** Resolve "../media/x.png" against "ppt/slides/". */
function resolvePath(baseDir: string, target: string): string {
  if (/^[a-z]+:/i.test(target)) return target;
  const parts = (baseDir + target).split('/');
  const out: string[] = [];
  for (const part of parts) {
    if (part === '..') out.pop();
    else if (part !== '.' && part !== '') out.push(part);
  }
  return out.join('/');
}
