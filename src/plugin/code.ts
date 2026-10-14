/** Plugin entry point (runs in Figma's sandbox). Talks to the UI iframe via postMessage. */
import { DEFAULT_OPTIONS, type ExportOptions, type PluginToUiMessage, type SlideSummary, type UiToPluginMessage, type Warning } from '../shared/types';
import { extractSlide, isSlideRoot, type SlideRoot } from './extract';

const STORAGE_KEY = 'deck-export.options.v1';

let options: ExportOptions = { ...DEFAULT_OPTIONS };
let cancelled = false;
let exporting = false;

figma.showUI(__html__, { themeColors: true, width: 380, height: 640 });

function send(msg: PluginToUiMessage): void {
  figma.ui.postMessage(msg);
}

figma.ui.onmessage = async (msg: UiToPluginMessage) => {
  switch (msg.type) {
    case 'ui-ready': {
      const stored = (await figma.clientStorage.getAsync(STORAGE_KEY)) as Partial<ExportOptions> | undefined;
      options = { ...DEFAULT_OPTIONS, ...(stored ?? {}), fileName: stored?.fileName || figma.root.name };
      send({
        type: 'init',
        editorType: figma.editorType,
        fileName: figma.root.name,
        options,
        slides: summarize(orderedCandidates()),
      });
      break;
    }
    case 'save-options':
      options = { ...DEFAULT_OPTIONS, ...msg.options };
      await figma.clientStorage.setAsync(STORAGE_KEY, options);
      send({ type: 'slides', slides: summarize(orderedCandidates()) });
      break;
    case 'export':
      options = { ...DEFAULT_OPTIONS, ...msg.options };
      await runExport(msg.slideIds);
      break;
    case 'cancel':
      cancelled = true;
      break;
    case 'focus-slide': {
      const node = await figma.getNodeByIdAsync(msg.slideId);
      if (node && isSlideRoot(node)) {
        figma.currentPage.selection = [node];
        figma.viewport.scrollAndZoomIntoView([node]);
      }
      break;
    }
    case 'notify':
      figma.notify(msg.message, { error: msg.error });
      break;
  }
};

figma.on('selectionchange', () => {
  if (!exporting) send({ type: 'slides', slides: summarize(orderedCandidates()) });
});
figma.on('currentpagechange', () => {
  if (!exporting) send({ type: 'slides', slides: summarize(orderedCandidates()) });
});

// ---- slide discovery ---------------------------------------------------------------------------

/**
 * Which frames are slides:
 *  - the current selection (frames, components, instances, Figma Slides, or sections/rows containing them), else
 *  - in Figma Slides: every slide in the grid, else
 *  - every top-level frame/component/instance on the page, including those inside sections.
 */
function candidates(): SlideRoot[] {
  const selected = expand(figma.currentPage.selection);
  if (selected.length) return selected;
  if (figma.editorType === 'slides') {
    return figma.currentPage.findAllWithCriteria({ types: ['SLIDE'] }) as SlideRoot[];
  }
  return expand(figma.currentPage.children);
}

function expand(nodes: readonly SceneNode[]): SlideRoot[] {
  const out: SlideRoot[] = [];
  const seen = new Set<string>();
  const visit = (n: SceneNode): void => {
    if (isSlideRoot(n)) {
      if (!seen.has(n.id)) {
        seen.add(n.id);
        out.push(n);
      }
    } else if (n.type === 'SECTION' || n.type === 'SLIDE_GRID' || n.type === 'SLIDE_ROW') {
      n.children.forEach(visit);
    }
  };
  nodes.forEach(visit);
  return out;
}

function orderedCandidates(): SlideRoot[] {
  const nodes = candidates();
  switch (options.order) {
    case 'name':
      return [...nodes].sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' }));
    case 'position':
      return byPosition(nodes);
    case 'selection':
      return figma.currentPage.selection.length ? nodes : byPosition(nodes);
    case 'layers':
    default:
      return nodes;
  }
}

/** Reading order: cluster into rows (overlapping vertical extents), then left to right. */
function byPosition(nodes: SlideRoot[]): SlideRoot[] {
  const boxed = nodes
    .map((n) => ({ n, box: n.absoluteBoundingBox }))
    .filter((e): e is { n: SlideRoot; box: Rect } => e.box !== null)
    .sort((a, b) => a.box.y - b.box.y || a.box.x - b.box.x);
  const rows: { bottom: number; items: typeof boxed }[] = [];
  for (const e of boxed) {
    const row = rows[rows.length - 1];
    if (row && e.box.y < row.bottom - e.box.height / 2) {
      row.items.push(e);
      row.bottom = Math.max(row.bottom, e.box.y + e.box.height);
    } else {
      rows.push({ bottom: e.box.y + e.box.height, items: [e] });
    }
  }
  return rows.flatMap((r) => r.items.sort((a, b) => a.box.x - b.box.x).map((e) => e.n));
}

function summarize(nodes: SlideRoot[]): SlideSummary[] {
  return nodes.map((n) => ({
    id: n.id,
    name: n.name,
    wPx: Math.round(n.width),
    hPx: Math.round(n.height),
    skipped: n.type === 'SLIDE' ? n.isSkippedSlide : false,
  }));
}

// ---- export --------------------------------------------------------------------------------------

async function runExport(slideIds: string[]): Promise<void> {
  if (exporting) return;
  exporting = true;
  cancelled = false;
  const warnings: Warning[] = [];
  try {
    const byId = new Map(orderedCandidates().map((n) => [n.id, n]));
    const roots = slideIds.map((id) => byId.get(id)).filter((n): n is SlideRoot => !!n);
    if (roots.length === 0) {
      send({ type: 'error', message: 'No slides to export. Select frames on the canvas or open a page with top-level frames.' });
      return;
    }
    const size = dominantSize(roots);
    send({ type: 'deck-start', total: roots.length, wPx: size.w, hPx: size.h });
    const slideIdSet = new Set(roots.map((r) => r.id));
    for (let i = 0; i < roots.length; i++) {
      if (cancelled) break;
      const slide = await extractSlide(roots[i], { options, slideIds: slideIdSet, warnings });
      send({ type: 'slide', slide, index: i, total: roots.length });
    }
    send({ type: 'deck-end', warnings, cancelled });
  } catch (e) {
    send({ type: 'error', message: (e as Error).message || String(e) });
  } finally {
    exporting = false;
  }
}

/** Most common frame size; outliers get letterboxed into it by the emitter. */
function dominantSize(nodes: SlideRoot[]): { w: number; h: number } {
  const counts = new Map<string, { w: number; h: number; n: number }>();
  for (const node of nodes) {
    const key = `${Math.round(node.width)}x${Math.round(node.height)}`;
    const e = counts.get(key) ?? { w: Math.round(node.width), h: Math.round(node.height), n: 0 };
    e.n++;
    counts.set(key, e);
  }
  return [...counts.values()].sort((a, b) => b.n - a.n)[0];
}
