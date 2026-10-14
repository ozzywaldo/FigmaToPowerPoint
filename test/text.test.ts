import { test } from 'node:test';
import assert from 'node:assert/strict';

// The text extractor only touches `figma.mixed`; stub the global before importing it.
(globalThis as Record<string, unknown>).figma = { mixed: Symbol('mixed') };
const { extractTextContent } = await import('../src/plugin/text');

const inter = (style: string) => ({ family: 'Inter', style });
const black = [{ type: 'SOLID', color: { r: 0, g: 0, b: 0 }, opacity: 1, visible: true }];

function segment(characters: string, over: Record<string, unknown> = {}) {
  return {
    characters,
    start: 0,
    end: characters.length,
    fontName: inter('Regular'),
    fontSize: 32,
    fontWeight: 400,
    fontStyle: 'REGULAR',
    fills: black,
    textDecoration: 'NONE',
    letterSpacing: { unit: 'PERCENT', value: 0 },
    lineHeight: { unit: 'AUTO' },
    hyperlink: null,
    listOptions: { type: 'NONE' },
    indentation: 0,
    textCase: 'ORIGINAL',
    paragraphSpacing: 0,
    ...over,
  };
}

function fakeText(segments: unknown[], over: Record<string, unknown> = {}) {
  return {
    name: 'Title',
    characters: segments.map((s) => (s as { characters: string }).characters).join(''),
    fills: black,
    getStyledTextSegments: () => segments,
    textAlignHorizontal: 'LEFT',
    textAlignVertical: 'TOP',
    textAutoResize: 'WIDTH_AND_HEIGHT',
    lineHeight: { unit: 'PERCENT', value: 120 },
    ...over,
  } as unknown as TextNode;
}

const ctx = { opacity: 1, resolveLink: () => undefined, warn: () => {} };

test('splits paragraphs on newlines and keeps run styling', () => {
  const node = fakeText([
    segment('Hello ', { fontName: inter('Bold'), fontWeight: 700 }),
    segment('world\nSecond line', { textDecoration: 'UNDERLINE' }),
  ]);
  const t = extractTextContent(node, ctx);
  assert.ok(t);
  assert.equal(t!.paragraphs.length, 2);
  assert.equal(t!.paragraphs[0].runs.length, 2);
  assert.equal(t!.paragraphs[0].runs[0].fontWeight, 700);
  assert.equal(t!.paragraphs[0].runs[1].underline, true);
  assert.equal(t!.paragraphs[1].runs[0].text, 'Second line');
  assert.deepEqual(t!.lineHeight, { unit: 'percent', value: 120 });
  assert.equal(t!.autoResize, 'width_height');
});

test('soft line breaks stay inside the paragraph', () => {
  const t = extractTextContent(fakeText([segment('a b')]), ctx)!;
  assert.equal(t.paragraphs.length, 1);
  assert.equal(t.paragraphs[0].runs[1].softBreakBefore, true);
});

test('lists, indentation, text case and gradient fallback', () => {
  const gradient = [{ type: 'GRADIENT_LINEAR', visible: true, opacity: 1, gradientStops: [{ color: { r: 1, g: 0, b: 0, a: 1 }, position: 0 }] }];
  const t = extractTextContent(
    fakeText([
      segment('one\n', { listOptions: { type: 'UNORDERED' }, indentation: 1 }),
      segment('two', { listOptions: { type: 'ORDERED' }, indentation: 2, textCase: 'UPPER', fills: gradient }),
    ]),
    ctx,
  )!;
  assert.equal(t.paragraphs[0].list, 'bullet');
  assert.equal(t.paragraphs[1].list, 'number');
  assert.equal(t.paragraphs[1].indent, 2);
  assert.equal(t.paragraphs[1].runs[0].text, 'TWO');
  assert.equal(t.paragraphs[1].runs[0].color, 'FF0000');
});

test('empty text returns undefined, image-filled text asks for raster', () => {
  assert.equal(extractTextContent(fakeText([segment('')], { characters: '' }), ctx), undefined);
  const img = [{ type: 'IMAGE', visible: true, opacity: 1 }];
  assert.equal(extractTextContent(fakeText([segment('x')], { fills: img }), ctx), null);
});
