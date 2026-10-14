/**
 * Round-trip a real deck through importer → emitter → importer.
 * Set DECK_PPTX=/path/to/deck.pptx to run it (decks are not committed); DECK_OUT to keep the
 * re-emitted file for visual comparison with tools/compare.mjs.
 */
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import PptxGenJS from 'pptxgenjs';
import { importPptx, splitFontFace } from '../src/import/pptx-import';
import { buildPresentation } from '../src/ui/pptx';
import { DEFAULT_OPTIONS } from '../src/shared/types';

test('font face splitting', () => {
  assert.deepEqual(splitFontFace('Lato Regular'), { family: 'Lato', style: 'Regular' });
  assert.deepEqual(splitFontFace('Instrument Serif Italic'), { family: 'Instrument Serif', style: 'Italic' });
  assert.deepEqual(splitFontFace('Inter Semi Bold Italic'), { family: 'Inter', style: 'Semi Bold Italic' });
  assert.deepEqual(splitFontFace('Calibri'), { family: 'Calibri', style: 'Regular' });
});

const deckPath = process.env.DECK_PPTX;
test('round-trips a real deck through import → emit → import', { skip: !deckPath && 'set DECK_PPTX to run' }, async () => {
  const src = await importPptx(await fs.readFile(deckPath!));
  assert.ok(src.slides.length > 0);
  // DECK_OPTIONS='{"textWidthAdjustPct":-3}' overrides emitter options for experiments.
  const overrides = process.env.DECK_OPTIONS ? JSON.parse(process.env.DECK_OPTIONS) : {};
  const pptx = buildPresentation(PptxGenJS, { ...src, options: { ...DEFAULT_OPTIONS, slideWidthIn: src.wPx / 96, fontFace: 'full', fileName: 'roundtrip', ...overrides } });
  const out = (await pptx.write({ outputType: 'nodebuffer' })) as Buffer;
  const dir = process.env.DECK_OUT ?? (await fs.mkdtemp(path.join(os.tmpdir(), 'roundtrip-')));
  await fs.mkdir(dir, { recursive: true });
  const file = path.join(dir, 'roundtrip.pptx');
  await fs.writeFile(file, out);
  console.log(`re-emitted deck: ${file}`);

  const back = await importPptx(out);
  assert.equal(back.slides.length, src.slides.length);
  assert.equal(back.wPx, src.wPx);
  for (let i = 0; i < src.slides.length; i++) {
    const a = src.slides[i];
    const b = back.slides[i];
    assert.equal(b.elements.length, a.elements.length, `slide ${i + 1} element count`);
    assert.equal(b.notes ?? '', a.notes ?? '', `slide ${i + 1} notes`);
    for (let j = 0; j < a.elements.length; j++) {
      const ea = a.elements[j];
      const eb = b.elements[j];
      assert.equal(eb.type, ea.type, `slide ${i + 1} element ${j} type`);
      for (const k of ['x', 'y', 'w', 'h', 'rotation'] as const) {
        assert.ok(Math.abs(eb[k] - ea[k]) <= 0.51, `slide ${i + 1} ${ea.name} ${k}: ${ea[k]} → ${eb[k]}`);
      }
      if (ea.type === 'text' && eb.type === 'text') {
        const ta = ea.paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join('\n');
        const tb = eb.paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join('\n');
        assert.equal(tb, ta, `slide ${i + 1} ${ea.name} text`);
        const fa = ea.paragraphs[0]?.runs[0];
        const fb = eb.paragraphs[0]?.runs[0];
        if (fa && fb) {
          assert.ok(Math.abs(fa.fontSizePx - fb.fontSizePx) < 0.05, `font size ${fa.fontSizePx} → ${fb.fontSizePx}`);
          assert.equal(fb.fontFamily, fa.fontFamily);
          assert.equal(fb.color, fa.color);
        }
      }
    }
  }
});
