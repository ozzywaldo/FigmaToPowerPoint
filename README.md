# Deck Export — Figma → PowerPoint

A Figma plugin that turns frames (or Figma Slides) into an **editable** `.pptx`: real text boxes,
real shapes, real vector freeforms, speaker notes and hyperlinks — not a stack of screenshots.

It is a clean-room take on the export half of Hypermatic's *Pitchdeck Presentation Studio*,
built to be hacked on: the extraction rules, the unit math and the PowerPoint emitter are three
small, independent modules with tests, so you can bend any of them without understanding the rest.

```
Figma nodes ──(src/plugin)──▶ Slide model (src/shared/types.ts) ──(src/ui/pptx.ts)──▶ .pptx
```

## Getting started

```bash
npm install
npm run build        # dist/code.js + dist/ui.html
npm run watch        # rebuild on change
npm run check        # typecheck + build + tests
```

Then in the Figma desktop app: **Plugins → Development → Import plugin from manifest…** and pick
`manifest.json`. Run it from the Plugins menu in a Figma design file or a Figma Slides file.

The `id` in `manifest.json` is a placeholder. If Figma rejects it, create an empty plugin via
**Plugins → Development → New plugin…**, copy the `id` Figma generates into this manifest, and
import again. (The id is only used to identify the plugin locally and when publishing.)

## How it works

### Which frames are slides

1. Whatever is selected (frames, components, instances, slides, or sections containing them), else
2. in Figma Slides, every slide in the grid, else
3. every top-level frame / component / instance on the current page, including those inside sections.

The panel lists them in export order with checkboxes. Figma Slides marked *skip* start unchecked.
Clicking a slide name scrolls the canvas to it.

### Slide size

The deck uses the most common frame size. The default width of `20 in` is a 96 DPI mapping: a
1920 px frame becomes exactly 9525 EMU per pixel and 0.75 pt per pixel (a 72 px headline is 54 pt).
That is the convention Pitchdeck's exports use, so decks from both tools line up. `13.333 in` gives
PowerPoint's own widescreen default instead (1 px = 0.5 pt). Frames of other sizes are fitted
proportionally and centred (letterboxed), never cropped. You can also force 16:9 or 4:3.

### What stays editable

| Figma | PowerPoint |
| --- | --- |
| Text (per-run font, size, weight, italic, colour, opacity, underline, strike, letter spacing, case, bullets, numbered lists, indentation, paragraph spacing, line height, alignment, hyperlinks) | Text box with runs and paragraphs |
| Rectangle (solid fill, solid stroke, corner radius, drop shadow) | `rect` / `roundRect` |
| Ellipse | `ellipse` (arcs/rings become freeforms) |
| Line (incl. arrow caps) | `line` |
| Vector, star, polygon, boolean | Freeform (`custGeom`) with cubic/quadratic curves |
| Frame / component / instance background | `roundRect` behind its children |
| Rotation, horizontal flip | `rot` / `flipH` about the box centre |
| Prototype link to a URL or to another slide | Hyperlink on the shape, text run or image |
| Text layer named `notes…` | Speaker notes |

### What becomes an image (and tells you so)

Gradients, image fills, blurs, inner shadows, multiple effects, masks, blend modes, text with image
fills, and frames that *clip* overflowing children (PowerPoint has no clipping). Each fallback is
listed in the panel after export, by slide and layer, so you can restructure the layer or accept it.
Images export at the chosen scale (default 2×) as PNG, or JPG when the layer is an opaque,
square-cornered photo (configurable).

### Layer-name conventions

| Prefix | Effect |
| --- | --- |
| `[PNG]`, `[JPG]`, `[IMG]` | Force the layer to export as an image (PNG / JPG / auto) |
| `notes` (configurable) | Text layer becomes the slide's speaker notes, not slide content |
| `[SKIP]` | Leave the layer out |

### Fonts

PowerPoint only has *family + bold + italic*. Intermediate weights exist only as separately named
faces on the viewer's machine ("Inter SemiBold"). The **Font names** option chooses between
emitting the plain family or `Family Weight`; brand-specific renames live in `FONT_FAMILY_MAP`
in `src/ui/fonts.ts`.

Design for PowerPoint's limits while you're still in Figma: one line height per text layer,
auto-height text boxes (fixed boxes smaller than their text drift vertically), whole-pixel positions.

## Project layout

```
manifest.json            Figma plugin manifest (design + slides editors, no network)
build.mjs                esbuild: sandbox bundle + single-file UI with inlined JS/CSS
src/shared/types.ts      Slide model + UI↔plugin message protocol — the contract
src/plugin/code.ts       Sandbox entry: slide discovery, ordering, export loop, settings
src/plugin/extract.ts    Node walker: Figma node → Slide elements, with raster fallbacks
src/plugin/text.ts       Styled text segments → paragraphs/runs
src/plugin/paint.ts      Fills, strokes, effects → solid colours / shadows
src/plugin/vector.ts     Figma path data → path commands
src/plugin/geometry.ts   Affine transforms → PowerPoint placement (rotation/flip)
src/ui/ui.ts, ui.css     Panel (vanilla DOM, Figma theme tokens)
src/ui/pptx.ts           Slide model → pptxgenjs (pure; runs in Node for tests)
src/ui/fonts.ts          Font face resolution + brand overrides
src/import/pptx-import.ts .pptx → slide model (round-trip tests, seed of a PPTX import feature)
tools/compare.mjs        Visual regression harness (LibreOffice render + per-slide diff report)
test/*.test.ts           node:test suites; pptx.test.ts writes a real deck and inspects its XML,
                         roundtrip.test.ts re-emits a real deck when DECK_PPTX is set
```

### Extending it

- **New export target** (Google Slides, Keynote via PPTX, HTML…): write another emitter next to
  `src/ui/pptx.ts` that consumes `Slide[]`. The plugin side doesn't change.
- **New extraction rule** (e.g. tables, a component that should become a native PowerPoint
  chart, slide-number placeholders): add a case to `walk()` in `src/plugin/extract.ts` and, if it
  needs a new element kind, add it to `types.ts` and the emitter.
- **Templates / masters**: pptxgenjs supports `defineSlideMaster`; a brand master with
  placeholders is a natural next step in `buildPresentation`.
- **Notes stored in the file**: swap the layer-name convention for `node.setPluginData` plus a
  notes field in the panel.

## Fidelity: measuring it, and what we learned from Pitchdeck's output

`npm run compare -- --candidate deck.pptx --reference figma.pdf --out report` renders the deck with
LibreOffice, rasterises both sides at the same width and writes `report/index.html`: per slide, a
strip of *reference | candidate | diff* plus two numbers, RMSE over RGB (0 = identical) and the
share of pixels that moved. The reference can also be another `.pptx`. It needs `soffice` and
`pdftoppm` on PATH. **For the real verdict, open the candidate in PowerPoint, save it as PDF and
pass that PDF as `--candidate`** — LibreOffice is only a proxy for PowerPoint's text layout.

`src/import/pptx-import.ts` reads a `.pptx` back into the slide model. `DECK_PPTX=deck.pptx
DECK_OUT=out npm test` pushes a real deck through importer → emitter → importer and asserts that
geometry, text, fonts and notes survive; the re-emitted file lands in `out/` for the harness.
`DECK_OPTIONS='{"textWidthAdjustPct":-2}'` overrides emitter options for experiments.

Calibrating against a 31-slide deck exported by Pitchdeck and the matching Figma PDF:

| Pitchdeck convention | Here |
| --- | --- |
| 20 in × 11.25 in slide (96 DPI, 1 px = 9525 EMU) | same, default |
| Text boxes: zero insets, `wrap="square"`, anchor from Figma vertical align | same |
| Font faces written as full names ("Lato Regular", "Instrument Serif Italic") | `fontFace: 'full'`; default is family + weight, which fontconfig/LibreOffice resolves and PowerPoint lists |
| Letter spacing: `spc` = tracking in 1/100 pt, `kern="0"` | same |
| Line height: pixels → exact `spcPts`; percent → `spcPct` = Figma % ÷ 1.2 | same |
| Every slide's frame fill rasterised to a full-bleed image | solid fills become a slide background; only non-solid fills are rasterised |
| Vectors and logos as SVG with PNG fallback | `vectors: 'svg'` does the same; the default emits editable freeforms |
| Groups with masks/clipping, gradients, textures → PNG | same, each reported as a warning |

Round-tripping that deck through this emitter and rendering both files gives pixel-identical slides
on 25 of 31 (mean RMSE 0.002; the rest differ only where SVGs were swapped for their PNG fallback).
Against the Figma PDF both score a mean RMSE of about 0.09 with ~4 % of pixels moved. Inspecting
those slides shows the residual is glyph anti-aliasing between renderers plus occasional line-wrap
differences where the renderer's advances are a few percent narrower than Figma's (uppercase,
tracked body text is the worst case). Narrowing boxes (`textWidthAdjustPct`) did not help in the
LibreOffice proxy; the knob exists for calibrating against PowerPoint itself.

## Known gaps vs. Pitchdeck

Pitchdeck also does PDF, Keynote/Google Slides, PSD-for-Canva, hosted web presentations with
animations, PPTX *import*, charts from CSV and embedded media. None of that is here. Within the
PowerPoint path, this plugin does not yet: convert gradients to native PowerPoint gradients (it
rasterizes), handle per-corner radii (largest wins), map Figma text truncation, or emit slide
masters.

## Verification

`npm test` builds a two-slide deck with every element type, checks the slide XML (fonts, sizes,
bullets, freeform curves, shadows, links, notes, letterboxing) and the ZIP structure. The resulting
file also opens in python-pptx and renders in LibreOffice. Rendering in PowerPoint itself should be
checked on your machine; fidelity notes are in *Fonts* above.
