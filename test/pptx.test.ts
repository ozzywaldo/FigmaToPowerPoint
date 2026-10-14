import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import JSZip from 'jszip';
import PptxGenJS from 'pptxgenjs';
import { buildPresentation, layoutSize, toBase64 } from '../src/ui/pptx';
import { resolveFontFace, isBold } from '../src/ui/fonts';
import { DEFAULT_OPTIONS, type Slide } from '../src/shared/types';

// 1×1 transparent PNG
const PNG = Uint8Array.from(
  atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='),
  (c) => c.charCodeAt(0),
);

const base = { rotation: 0, flipH: false, flipV: false };

const slides: Slide[] = [
  {
    id: 's1',
    name: 'Title',
    wPx: 1920,
    hPx: 1080,
    background: { color: '101418', alpha: 1 },
    notes: 'Open with the mission statement.',
    elements: [
      {
        type: 'text',
        id: 't1',
        name: 'Headline',
        x: 120,
        y: 200,
        w: 1200,
        h: 160,
        ...base,
        align: 'left',
        valign: 'top',
        autoResize: 'height',
        lineHeight: { unit: 'percent', value: 110 },
        paragraphs: [
          {
            list: 'none',
            indent: 0,
            spacingAfterPx: 16,
            runs: [
              { text: 'Autonomy ', fontFamily: 'Inter', fontStyle: 'Semi Bold', fontWeight: 600, italic: false, fontSizePx: 96, color: 'FFFFFF', alpha: 1, underline: false, strike: false },
              { text: 'at scale', fontFamily: 'Inter', fontStyle: 'Light Italic', fontWeight: 300, italic: true, fontSizePx: 96, color: 'FFB000', alpha: 0.8, underline: false, strike: false, letterSpacing: { unit: 'percent', value: -2 } },
            ],
          },
          { list: 'bullet', indent: 1, spacingAfterPx: 0, runs: [{ text: 'Point one', fontFamily: 'Inter', fontStyle: 'Regular', fontWeight: 400, italic: false, fontSizePx: 40, color: 'DDDDDD', alpha: 1, underline: false, strike: false, link: { kind: 'url', url: 'https://example.com' } }] },
          { list: 'number', indent: 1, spacingAfterPx: 0, runs: [{ text: 'Go to slide 2', fontFamily: 'Inter', fontStyle: 'Regular', fontWeight: 400, italic: false, fontSizePx: 40, color: 'DDDDDD', alpha: 1, underline: true, strike: false, link: { kind: 'slide', slideId: 's2' } }] },
        ],
      },
      { type: 'shape', shape: 'rect', id: 'r1', name: 'Card', x: 1400, y: 200, w: 400, h: 300, ...base, fill: { color: '1F2A37', alpha: 1 }, stroke: { color: '3B82F6', alpha: 1, widthPx: 2, dashed: true }, cornerRadiusPx: 24, shadow: { color: '000000', alpha: 0.4, offsetXPx: 0, offsetYPx: 8, blurPx: 24 } },
      { type: 'shape', shape: 'ellipse', id: 'e1', name: 'Dot', x: 100, y: 900, w: 60, h: 60, rotation: 45, flipH: false, flipV: false, fill: { color: 'FF0000', alpha: 0.5 } },
      { type: 'shape', shape: 'line', id: 'l1', name: 'Rule', x: 120, y: 400, w: 1200, h: 0, ...base, stroke: { color: 'FFFFFF', alpha: 1, widthPx: 2, dashed: false, arrowEnd: true } },
      { type: 'path', id: 'v1', name: 'Chevron', x: 1500, y: 700, w: 200, h: 200, ...base, fill: { color: '00FF88', alpha: 1 }, commands: [{ op: 'M', x: 0, y: 0 }, { op: 'L', x: 200, y: 100 }, { op: 'C', x1: 200, y1: 150, x2: 100, y2: 200, x: 0, y: 200 }, { op: 'Z' }] },
      { type: 'image', id: 'i1', name: 'Photo', x: 200, y: 600, w: 400, h: 225, ...base, format: 'png', bytes: PNG, opacity: 1 },
    ],
  },
  { id: 's2', name: 'Odd size', wPx: 1280, hPx: 1280, elements: [{ type: 'shape', shape: 'rect', id: 'r2', name: 'Full', x: 0, y: 0, w: 1280, h: 1280, ...base, fill: { color: '00FF00', alpha: 1 } }] },
];

async function writeDeck(): Promise<JSZip> {
  // Pin the layout so the expected EMU/pt numbers below stay meaningful regardless of defaults.
  const options = { ...DEFAULT_OPTIONS, slideWidthIn: 13.333, fontFace: 'family' as const, fileName: 'Test deck' };
  const pptx = buildPresentation(PptxGenJS, { wPx: 1920, hPx: 1080, slides, options });
  const data = (await pptx.write({ outputType: 'nodebuffer' })) as Buffer;
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'deck-'));
  await fs.writeFile(path.join(dir, 'test.pptx'), data);
  return JSZip.loadAsync(data);
}

test('layout follows the dominant frame aspect, or a forced ratio', () => {
  assert.deepEqual(layoutSize({ wPx: 1920, hPx: 1080, options: { ...DEFAULT_OPTIONS, slideWidthIn: 13.333 } }), { width: 13.333, height: 7.4998 });
  assert.deepEqual(layoutSize({ wPx: 1920, hPx: 1080, options: DEFAULT_OPTIONS }), { width: 20, height: 11.25 });
  assert.deepEqual(layoutSize({ wPx: 1000, hPx: 1000, options: { ...DEFAULT_OPTIONS, layout: '4:3', slideWidthIn: 10 } }), { width: 10, height: 7.5 });
});

test('writes a valid pptx with every element kind', async () => {
  const zip = await writeDeck();
  const names = Object.keys(zip.files);
  assert.ok(names.includes('ppt/slides/slide1.xml'));
  assert.ok(names.includes('ppt/slides/slide2.xml'));
  assert.ok(names.includes('ppt/notesSlides/notesSlide1.xml'));
  assert.ok(names.some((n) => /ppt\/media\/image.*\.png$/.test(n)), 'image media stored');

  const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
  // text runs & styling
  assert.match(xml, /Autonomy /);
  assert.match(xml, /typeface="Inter"/);
  assert.match(xml, /sz="4800"/); // 96 px × 0.5 pt/px
  assert.match(xml, /b="1"/);
  assert.match(xml, /i="1"/);
  assert.match(xml, /u="sng"/);
  assert.match(xml, /<a:buChar/);
  assert.match(xml, /<a:buAutoNum/);
  assert.match(xml, /lnSpc><a:spcPct val="91667"/); // 110% of font size ÷ 1.2 natural line height
  assert.match(xml, /spc="-96" kern="0"/); // -2% of 48pt
  // shapes
  assert.match(xml, /prst="roundRect"/);
  assert.match(xml, /prst="ellipse"/);
  assert.match(xml, /prst="line"/);
  assert.match(xml, /<a:custGeom>/);
  assert.match(xml, /<a:cubicBezTo>/);
  assert.match(xml, /<a:tailEnd type="triangle"/);
  assert.match(xml, /prstDash val="dash"/);
  assert.match(xml, /<a:outerShdw/);
  assert.match(xml, /rot="2700000"/); // 45° ellipse
  assert.match(xml, /<a:alpha val="50000"\/>/); // 50% fill
  // links
  assert.match(xml, /ppaction:\/\/hlinksldjump/);
  const rels = await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string');
  assert.match(rels, /https:\/\/example\.com/);
  assert.match(rels, /slide2\.xml/);
  // background + notes
  assert.match(xml, /<p:bg>[\s\S]*101418/);
  const notes = await zip.file('ppt/notesSlides/notesSlide1.xml')!.async('string');
  assert.match(notes, /mission statement/);
});

test('letterboxes slides of a different size into the deck layout', async () => {
  const zip = await writeDeck();
  const xml = await zip.file('ppt/slides/slide2.xml')!.async('string');
  // 1280×1280 into 13.333×7.5in → scale by height: 7.5in square, centred: x = (13.333-7.5)/2 = 2.9165in = 2666490 EMU
  // The first <a:off> belongs to the slide's group transform; the shape's is the next one.
  const offs = [...xml.matchAll(/<a:off x="(\d+)" y="(\d+)"\/>/g)];
  const off = offs[1];
  assert.ok(off, 'shape offset present');
  assert.ok(Math.abs(Number(off![1]) - 2666490) < 2000, `x offset ${off![1]}`);
  assert.ok(Number(off![2]) < 2000);
});

test('font face resolution', () => {
  assert.equal(resolveFontFace('Inter', 'Semi Bold Italic', 'family'), 'Inter');
  assert.equal(resolveFontFace('Inter', 'Semi Bold Italic', 'family-style'), 'Inter Semi Bold');
  assert.equal(resolveFontFace('Inter', 'Bold', 'family-style'), 'Inter');
  assert.equal(resolveFontFace('Lato', 'Regular', 'full'), 'Lato Regular');
  assert.equal(resolveFontFace('Instrument Serif', 'Italic', 'full'), 'Instrument Serif Italic');
  assert.equal(isBold(600, 'Semi Bold'), true);
  assert.equal(isBold(0, 'Black'), true);
  assert.equal(isBold(400, 'Regular'), false);
});

test('base64 helper handles large buffers', () => {
  const big = new Uint8Array(200_000).map((_, i) => i % 251);
  assert.equal(toBase64(big), Buffer.from(big).toString('base64'));
});
