/** Plugin panel. Vanilla DOM on purpose: nothing to learn, easy to rip apart. */
import './ui.css';
import PptxGenJS from 'pptxgenjs';
import { DEFAULT_OPTIONS, type ExportOptions, type PluginToUiMessage, type Slide, type SlideSummary, type UiToPluginMessage, type Warning } from '../shared/types';
import { buildPresentation } from './pptx';

interface State {
  editorType: string;
  fileName: string;
  options: ExportOptions;
  slides: SlideSummary[];
  excluded: Set<string>;
  phase: 'idle' | 'extracting' | 'writing' | 'done' | 'error';
  progress: { index: number; total: number; name: string } | null;
  status: string;
  warnings: Warning[];
  deck: { wPx: number; hPx: number; slides: Slide[] } | null;
}

const state: State = {
  editorType: 'figma',
  fileName: '',
  options: { ...DEFAULT_OPTIONS },
  slides: [],
  excluded: new Set(),
  phase: 'idle',
  progress: null,
  status: '',
  warnings: [],
  deck: null,
};

const app = document.getElementById('app')!;

function post(msg: UiToPluginMessage): void {
  parent.postMessage({ pluginMessage: msg }, '*');
}

window.onmessage = async (ev: MessageEvent<{ pluginMessage?: PluginToUiMessage }>) => {
  const msg = ev.data.pluginMessage;
  if (!msg) return;
  switch (msg.type) {
    case 'init':
      state.editorType = msg.editorType;
      state.fileName = msg.fileName;
      state.options = { ...DEFAULT_OPTIONS, ...(msg.options ?? {}) };
      if (!state.options.fileName) state.options.fileName = msg.fileName;
      setSlides(msg.slides);
      break;
    case 'slides':
      if (state.phase === 'extracting') return;
      setSlides(msg.slides);
      break;
    case 'deck-start':
      state.deck = { wPx: msg.wPx, hPx: msg.hPx, slides: [] };
      state.progress = { index: 0, total: msg.total, name: '' };
      break;
    case 'slide':
      state.deck?.slides.push(msg.slide);
      state.progress = { index: msg.index + 1, total: msg.total, name: msg.slide.name };
      state.status = `Reading ${msg.index + 1}/${msg.total} · ${msg.slide.name}`;
      break;
    case 'deck-end':
      state.warnings = msg.warnings;
      if (msg.cancelled) {
        state.phase = 'idle';
        state.status = 'Export cancelled.';
      } else {
        await writeFile();
      }
      break;
    case 'error':
      state.phase = 'error';
      state.status = msg.message;
      break;
  }
  render();
};

function setSlides(slides: SlideSummary[]): void {
  state.slides = slides;
  // Keep manual exclusions for slides that still exist; newly listed skipped slides start excluded.
  const ids = new Set(slides.map((s) => s.id));
  for (const id of [...state.excluded]) if (!ids.has(id)) state.excluded.delete(id);
  for (const s of slides) if (s.skipped && !seen.has(s.id)) state.excluded.add(s.id);
  for (const s of slides) seen.add(s.id);
}
const seen = new Set<string>();

async function writeFile(): Promise<void> {
  if (!state.deck) return;
  state.phase = 'writing';
  state.status = 'Writing PowerPoint file…';
  render();
  try {
    const pptx = buildPresentation(PptxGenJS, { ...state.deck, options: state.options });
    const blob = (await pptx.write({ outputType: 'blob', compression: true })) as Blob;
    const name = (state.options.fileName || state.fileName || 'deck').replace(/\.pptx$/i, '') + '.pptx';
    download(blob, name);
    state.phase = 'done';
    state.status = `Exported ${state.deck.slides.length} slide${state.deck.slides.length === 1 ? '' : 's'} → ${name} (${(blob.size / 1048576).toFixed(1)} MB)`;
    post({ type: 'notify', message: `Exported ${name}` });
  } catch (e) {
    state.phase = 'error';
    state.status = `Could not write file: ${(e as Error).message}`;
  } finally {
    state.deck = null;
  }
}

function download(blob: Blob, name: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    a.remove();
    URL.revokeObjectURL(url);
  }, 2000);
}

function startExport(): void {
  const ids = state.slides.filter((s) => !state.excluded.has(s.id)).map((s) => s.id);
  if (!ids.length) {
    state.status = 'Nothing to export: select at least one slide.';
    render();
    return;
  }
  state.phase = 'extracting';
  state.warnings = [];
  state.status = 'Reading slides…';
  state.progress = { index: 0, total: ids.length, name: '' };
  post({ type: 'export', options: state.options, slideIds: ids });
  render();
}

function updateOption<K extends keyof ExportOptions>(key: K, value: ExportOptions[K]): void {
  state.options = { ...state.options, [key]: value };
  post({ type: 'save-options', options: state.options });
  render();
}

// ---- rendering ---------------------------------------------------------------------------

function render(): void {
  const busy = state.phase === 'extracting' || state.phase === 'writing';
  const included = state.slides.filter((s) => !state.excluded.has(s.id));
  const dominant = dominantSize(included);
  app.replaceChildren(
    h('div', { class: 'header' }, [
      h('h1', {}, ['Deck Export']),
      h('span', { class: 'meta' }, [
        state.slides.length
          ? `${included.length}/${state.slides.length} slides · ${dominant ? `${dominant.w}×${dominant.h}` : ''}`
          : state.editorType === 'slides'
            ? 'No slides found'
            : 'No frames on this page',
      ]),
    ]),
    renderSlideList(busy, dominant),
    renderOptions(busy),
    renderFooter(busy),
  );
}

function renderSlideList(busy: boolean, dominant: { w: number; h: number } | null): HTMLElement {
  if (!state.slides.length) {
    return h('div', { class: 'slides' }, [
      h('div', { class: 'empty' }, ['Select frames on the canvas, or open a page with top-level frames. Each frame becomes a slide.']),
    ]);
  }
  const rows = state.slides.map((s, i) => {
    const checked = !state.excluded.has(s.id);
    const outlier = dominant !== null && (s.wPx !== dominant.w || s.hPx !== dominant.h);
    return h('label', { class: `slide-row${s.skipped ? ' skipped' : ''}` }, [
      h('input', {
        type: 'checkbox',
        checked: checked ? 'checked' : null,
        disabled: busy ? 'disabled' : null,
        onchange: (e: Event) => {
          if ((e.target as HTMLInputElement).checked) state.excluded.delete(s.id);
          else state.excluded.add(s.id);
          render();
        },
      }),
      h('span', { class: 'idx' }, [String(i + 1)]),
      h('span', { class: 'name', title: 'Show on canvas', onclick: (e: Event) => { e.preventDefault(); post({ type: 'focus-slide', slideId: s.id }); } }, [s.name]),
      h('span', { class: `size${outlier ? ' outlier' : ''}`, title: outlier ? 'Different size: will be fitted into the deck size' : '' }, [`${s.wPx}×${s.hPx}`]),
    ]);
  });
  return h('div', {}, [
    h('div', { class: 'list-tools' }, [
      h('a', { onclick: () => { state.excluded.clear(); render(); } }, ['Include all']),
      h('a', { onclick: () => { state.slides.forEach((s) => state.excluded.add(s.id)); render(); } }, ['Exclude all']),
      h('span', { style: 'margin-left:auto' }, [state.editorType === 'slides' ? 'Figma Slides' : 'Order: ' + orderLabel(state.options.order)]),
    ]),
    h('div', { class: 'slides' }, rows),
  ]);
}

function orderLabel(o: ExportOptions['order']): string {
  return { position: 'canvas position', layers: 'layer order', name: 'name', selection: 'selection' }[o];
}

function renderOptions(busy: boolean): HTMLElement {
  const o = state.options;
  const sel = <K extends keyof ExportOptions>(key: K, choices: [ExportOptions[K], string][]) =>
    h(
      'select',
      {
        disabled: busy ? 'disabled' : null,
        onchange: (e: Event) => {
          const picked = choices[(e.target as HTMLSelectElement).selectedIndex];
          if (picked) updateOption(key, picked[0]);
        },
      },
      choices.map(([v, label]) => h('option', { value: String(v), selected: o[key] === v ? 'selected' : null }, [label])),
    );
  const field = (label: string, control: HTMLElement, hint?: string) =>
    h('div', { class: 'field' }, [h('label', {}, [label]), control, ...(hint ? [h('div', { class: 'hint' }, [hint])] : [])]);

  return h('div', { class: 'options' }, [
    field(
      'Slide size',
      h('div', { class: 'inline' }, [
        sel('layout', [
          ['frame', 'Match frames'],
          ['16:9', '16:9'],
          ['4:3', '4:3'],
        ]),
        h('input', {
          type: 'number', step: '0.001', min: '1', max: '56', value: String(o.slideWidthIn), title: 'Slide width in inches',
          disabled: busy ? 'disabled' : null,
          onchange: (e: Event) => updateOption('slideWidthIn', Math.max(1, Number((e.target as HTMLInputElement).value) || 13.333)),
        }),
        h('span', { style: 'color:var(--text-3)' }, ['in']),
      ]),
      '20 in = 96 DPI (1920 px → 1 px = 0.75 pt, as Pitchdeck). 13.333 in is PowerPoint’s own default.',
    ),
    field('Layers', sel('editable', [[true, 'Editable text & shapes'], [false, 'Flatten each slide to an image']])),
    field('Vectors', sel('vectors', [['shapes', 'Editable freeform shapes'], ['svg', 'SVG (crisp, convertible in PowerPoint)'], ['raster', 'PNG images']])),
    field('Images', h('div', { class: 'inline' }, [
      sel('imageScale', [[1, '1×'], [2, '2×'], [3, '3×'], [4, '4×']]),
      sel('imageFormat', [['auto', 'Auto PNG/JPG'], ['png', 'Always PNG'], ['jpg', 'Always JPG']]),
    ])),
    field('Slide order', sel('order', [['position', 'Canvas position'], ['layers', 'Layer order'], ['name', 'Name (natural sort)'], ['selection', 'Selection order']])),
    field('Notes layer', h('input', {
      type: 'text', value: o.notesLayerName, placeholder: 'notes', disabled: busy ? 'disabled' : null,
      onchange: (e: Event) => updateOption('notesLayerName', (e.target as HTMLInputElement).value),
    }), 'Text layers whose name starts with this become speaker notes.'),
    field('Font names', sel('fontFace', [['family-style', 'Family + weight (Inter SemiBold)'], ['family', 'Family only (Inter)'], ['full', 'Full name (Inter Regular) — Pitchdeck style']])),
    field('Text box width', h('div', { class: 'inline' }, [
      h('input', {
        type: 'number', step: '0.5', min: '-20', max: '20', value: String(o.textWidthAdjustPct), disabled: busy ? 'disabled' : null,
        onchange: (e: Event) => updateOption('textWidthAdjustPct', Number((e.target as HTMLInputElement).value) || 0),
      }),
      h('span', { style: 'color:var(--text-3)' }, ['%']),
    ]), 'Widen (+) or narrow (−) fixed-width text boxes if PowerPoint wraps lines differently than Figma.'),
    field('File name', h('input', {
      type: 'text', value: o.fileName, placeholder: state.fileName, disabled: busy ? 'disabled' : null,
      onchange: (e: Event) => updateOption('fileName', (e.target as HTMLInputElement).value),
    })),
  ]);
}

function renderFooter(busy: boolean): HTMLElement {
  const children: (HTMLElement | string)[] = [];
  if (busy && state.progress) {
    const pct = state.phase === 'writing' ? 100 : Math.round((state.progress.index / Math.max(1, state.progress.total)) * 100);
    children.push(h('div', { class: 'progress' }, [h('div', { style: `width:${pct}%` })]));
  }
  children.push(h('div', { class: `status${state.phase === 'error' ? ' error' : state.phase === 'done' ? ' ok' : ''}` }, [state.status]));
  if (state.warnings.length && !busy) {
    const bySlide = new Map<string, Warning[]>();
    for (const w of state.warnings) bySlide.set(w.slideName, [...(bySlide.get(w.slideName) ?? []), w]);
    children.push(
      h('details', { class: 'warnings' }, [
        h('summary', {}, [`${state.warnings.length} layer${state.warnings.length === 1 ? '' : 's'} exported as images or adjusted`]),
        ...[...bySlide.entries()].map(([slide, ws]) =>
          h('div', {}, [h('b', {}, [slide]), h('ul', { style: 'margin:0;padding:0' }, ws.map((w) => h('li', {}, [`${w.nodeName}: ${w.message}`])))]),
        ),
      ]),
    );
  }
  children.push(
    h('div', { class: 'row' }, [
      busy
        ? h('button', { onclick: () => { post({ type: 'cancel' }); state.status = 'Cancelling…'; render(); }, disabled: state.phase === 'writing' ? 'disabled' : null }, ['Cancel'])
        : h('button', { class: 'primary', onclick: startExport, disabled: state.slides.length ? null : 'disabled' }, ['Export to PowerPoint']),
    ]),
  );
  return h('div', { class: 'footer' }, children);
}

function dominantSize(slides: SlideSummary[]): { w: number; h: number } | null {
  const counts = new Map<string, { w: number; h: number; n: number }>();
  for (const s of slides) {
    const k = `${s.wPx}x${s.hPx}`;
    const e = counts.get(k) ?? { w: s.wPx, h: s.hPx, n: 0 };
    e.n++;
    counts.set(k, e);
  }
  return [...counts.values()].sort((a, b) => b.n - a.n)[0] ?? null;
}

type Attrs = Record<string, string | null | ((e: Event) => void)>;
function h(tag: string, attrs: Attrs = {}, children: (HTMLElement | string)[] = []): HTMLElement {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (v === null || v === undefined) continue;
    if (typeof v === 'function') (el as unknown as Record<string, unknown>)[k] = v;
    else el.setAttribute(k, v);
  }
  for (const c of children) el.append(c);
  return el;
}

render();
post({ type: 'ui-ready' });
