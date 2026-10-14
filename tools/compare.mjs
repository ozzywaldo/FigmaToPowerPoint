#!/usr/bin/env node
/**
 * Visual regression harness: renders a .pptx (via LibreOffice) and compares it, slide by slide,
 * against a reference — either a PDF exported from Figma or another .pptx (e.g. Pitchdeck's).
 *
 *   node tools/compare.mjs --candidate out.pptx --reference figma.pdf [--out report] [--width 1280]
 *
 * Needs `soffice` (LibreOffice) and `pdftoppm` (poppler) on PATH. For the real answer, export the
 * candidate from PowerPoint itself as PDF and pass that as --candidate: LibreOffice is only a proxy
 * for PowerPoint's text layout.
 *
 * Output: <out>/index.html with a sortable table, <out>/slide-NN.png (reference | candidate | diff)
 * and <out>/report.json. Metrics per slide: RMSE over RGB (0..1, lower is better) and the share of
 * pixels whose colour moved by more than --threshold (default 32/255).
 */
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { PNG } from 'pngjs';

const args = parseArgs(process.argv.slice(2));
const candidate = must(args.candidate, '--candidate <file.pptx|file.pdf>');
const reference = must(args.reference, '--reference <figma.pdf|other.pptx>');
const out = path.resolve(args.out ?? 'compare-report');
const width = Number(args.width ?? 1280);
const threshold = Number(args.threshold ?? 32);
fs.mkdirSync(out, { recursive: true });

const candPdf = toPdf(candidate, path.join(out, 'render-candidate'));
const refPdf = toPdf(reference, path.join(out, 'render-reference'));
const candPages = rasterize(candPdf, path.join(out, 'png-candidate'), width);
const refPages = rasterize(refPdf, path.join(out, 'png-reference'), width);
const n = Math.min(candPages.length, refPages.length);
if (candPages.length !== refPages.length) {
  console.warn(`page count differs: candidate ${candPages.length}, reference ${refPages.length}; comparing first ${n}`);
}

const rows = [];
for (let i = 0; i < n; i++) {
  const a = PNG.sync.read(fs.readFileSync(refPages[i]));
  const b = PNG.sync.read(fs.readFileSync(candPages[i]));
  const w = Math.min(a.width, b.width);
  const h = Math.min(a.height, b.height);
  const diff = new PNG({ width: w, height: h });
  let sq = 0;
  let moved = 0;
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const ia = (y * a.width + x) * 4;
      const ib = (y * b.width + x) * 4;
      const id = (y * w + x) * 4;
      let maxd = 0;
      for (let c = 0; c < 3; c++) {
        const d = a.data[ia + c] - b.data[ib + c];
        sq += d * d;
        maxd = Math.max(maxd, Math.abs(d));
      }
      if (maxd > threshold) moved++;
      // diff view: reference faint, changes in red
      const grey = 255 - Math.round((255 - (a.data[ia] + a.data[ia + 1] + a.data[ia + 2]) / 3) * 0.15);
      diff.data[id] = maxd > threshold ? 220 : grey;
      diff.data[id + 1] = maxd > threshold ? 40 : grey;
      diff.data[id + 2] = maxd > threshold ? 40 : grey;
      diff.data[id + 3] = 255;
    }
  }
  const rmse = Math.sqrt(sq / (w * h * 3)) / 255;
  const movedPct = (moved / (w * h)) * 100;
  const sheet = new PNG({ width: w * 3 + 8, height: h });
  sheet.data.fill(128);
  PNG.bitblt(a, sheet, 0, 0, w, h, 0, 0);
  PNG.bitblt(b, sheet, 0, 0, w, h, w + 4, 0);
  PNG.bitblt(diff, sheet, 0, 0, w, h, 2 * w + 8, 0);
  const file = `slide-${String(i + 1).padStart(2, '0')}.png`;
  fs.writeFileSync(path.join(out, file), PNG.sync.write(sheet));
  rows.push({ slide: i + 1, rmse: round(rmse, 4), movedPct: round(movedPct, 2), file });
}

const summary = {
  candidate: path.resolve(candidate),
  reference: path.resolve(reference),
  width,
  threshold,
  slides: rows,
  meanRmse: round(rows.reduce((s, r) => s + r.rmse, 0) / Math.max(1, rows.length), 4),
  meanMovedPct: round(rows.reduce((s, r) => s + r.movedPct, 0) / Math.max(1, rows.length), 2),
};
fs.writeFileSync(path.join(out, 'report.json'), JSON.stringify(summary, null, 2));
fs.writeFileSync(path.join(out, 'index.html'), html(summary));

console.log(`slide  rmse    moved%`);
for (const r of [...rows].sort((x, y) => y.rmse - x.rmse)) {
  console.log(`${String(r.slide).padStart(5)}  ${r.rmse.toFixed(4)}  ${r.movedPct.toFixed(2).padStart(6)}`);
}
console.log(`mean   ${summary.meanRmse.toFixed(4)}  ${summary.meanMovedPct.toFixed(2).padStart(6)}`);
console.log(`report: ${path.join(out, 'index.html')}`);

// ---- helpers ------------------------------------------------------------------------------

function toPdf(file, dir) {
  if (file.toLowerCase().endsWith('.pdf')) return file;
  fs.mkdirSync(dir, { recursive: true });
  const r = spawnSync('soffice', ['--headless', '--convert-to', 'pdf', '--outdir', dir, file], { encoding: 'utf8', timeout: 600000 });
  if (r.error || r.status !== 0) fail(`LibreOffice conversion failed for ${file}: ${r.error?.message ?? r.stderr}`);
  const pdf = path.join(dir, path.basename(file).replace(/\.[^.]+$/, '.pdf'));
  if (!fs.existsSync(pdf)) fail(`LibreOffice produced no PDF for ${file}`);
  return pdf;
}

function rasterize(pdf, dir, w) {
  fs.mkdirSync(dir, { recursive: true });
  const r = spawnSync('pdftoppm', ['-png', '-scale-to-x', String(w), '-scale-to-y', '-1', pdf, path.join(dir, 'p')], { encoding: 'utf8', timeout: 600000 });
  if (r.error || r.status !== 0) fail(`pdftoppm failed for ${pdf}: ${r.error?.message ?? r.stderr}`);
  return fs
    .readdirSync(dir)
    .filter((f) => f.endsWith('.png'))
    .sort((x, y) => Number(x.match(/(\d+)\.png$/)[1]) - Number(y.match(/(\d+)\.png$/)[1]))
    .map((f) => path.join(dir, f));
}

function html(s) {
  const rows = s.slides
    .map(
      (r) =>
        `<tr><td>${r.slide}</td><td>${r.rmse.toFixed(4)}</td><td>${r.movedPct.toFixed(2)}%</td><td><a href="${r.file}"><img src="${r.file}" loading="lazy"></a></td></tr>`,
    )
    .join('\n');
  return `<!doctype html><meta charset="utf-8"><title>Deck comparison</title>
<style>body{font:13px/1.4 system-ui;margin:24px;color:#222}table{border-collapse:collapse}td,th{padding:6px 10px;border-bottom:1px solid #ddd;vertical-align:top;text-align:left}img{width:900px;display:block}code{font-size:12px}</style>
<h1>Deck comparison</h1>
<p>reference: <code>${esc(s.reference)}</code><br>candidate: <code>${esc(s.candidate)}</code></p>
<p>mean RMSE <b>${s.meanRmse}</b> · mean moved pixels <b>${s.meanMovedPct}%</b> · threshold ${s.threshold}/255 · ${s.width}px wide. Each strip: reference | candidate | diff (red = moved).</p>
<table><tr><th>slide</th><th>RMSE</th><th>moved</th><th></th></tr>${rows}</table>`;
}

function esc(t) {
  return String(t).replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' })[c]);
}
function round(n, p) {
  return Math.round(n * 10 ** p) / 10 ** p;
}
function must(v, what) {
  if (!v) fail(`missing ${what}`);
  return v;
}
function fail(msg) {
  console.error(msg);
  process.exit(1);
}
function parseArgs(argv) {
  const o = {};
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a.startsWith('--')) o[a.slice(2)] = argv[i + 1] && !argv[i + 1].startsWith('--') ? argv[++i] : true;
  }
  return o;
}
