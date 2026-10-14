// Build script: bundles the plugin sandbox code and inlines the UI into a single HTML file,
// which is what Figma requires. `node build.mjs --watch` rebuilds on change,
// `node build.mjs --tests` bundles the test suite for `node --test`.
import * as esbuild from 'esbuild';
import fs from 'node:fs/promises';
import path from 'node:path';

const args = new Set(process.argv.slice(2));
const watch = args.has('--watch');
const tests = args.has('--tests');

/** @type {import('esbuild').BuildOptions} */
const codeOptions = {
  entryPoints: ['src/plugin/code.ts'],
  bundle: true,
  outfile: 'dist/code.js',
  // The Figma sandbox is a plain JS VM (no DOM, no Node). Keep syntax conservative.
  target: 'es2019',
  format: 'iife',
  platform: 'neutral',
  mainFields: ['module', 'main'],
  legalComments: 'none',
  logLevel: 'info',
};

/** Inlines the bundled JS + CSS into src/ui/ui.html → dist/ui.html. */
const inlineUiPlugin = {
  name: 'inline-ui-html',
  setup(build) {
    build.onEnd(async (result) => {
      if (result.errors.length || !result.outputFiles) return;
      const js = result.outputFiles.find((f) => f.path.endsWith('.js'))?.text ?? '';
      const css = result.outputFiles.find((f) => f.path.endsWith('.css'))?.text ?? '';
      const template = await fs.readFile('src/ui/ui.html', 'utf8');
      // A literal "</script>" inside the bundle would terminate the inline script early.
      const safeJs = js.replace(/<\/script/gi, '<\\/script');
      const html = template
        .replace('/*__CSS__*/', () => css)
        .replace('/*__JS__*/', () => safeJs);
      await fs.mkdir('dist', { recursive: true });
      await fs.writeFile('dist/ui.html', html);
      console.log(`  dist/ui.html  ${(html.length / 1024).toFixed(1)}kb`);
    });
  },
};

/** @type {import('esbuild').BuildOptions} */
const uiOptions = {
  entryPoints: ['src/ui/ui.ts'],
  bundle: true,
  write: false,
  outdir: 'dist/.ui',
  target: 'es2020',
  format: 'iife',
  platform: 'browser',
  define: { 'process.env.NODE_ENV': '"production"' },
  legalComments: 'none',
  logLevel: 'info',
  plugins: [inlineUiPlugin],
};

async function buildTests() {
  const files = (await fs.readdir('test')).filter((f) => f.endsWith('.test.ts'));
  await fs.rm('dist-test', { recursive: true, force: true });
  await esbuild.build({
    entryPoints: files.map((f) => path.join('test', f)),
    bundle: true,
    outdir: 'dist-test',
    outExtension: { '.js': '.mjs' },
    platform: 'node',
    format: 'esm',
    target: 'node20',
    packages: 'external',
    logLevel: 'info',
  });
}

if (tests) {
  await buildTests();
} else if (watch) {
  const [codeCtx, uiCtx] = await Promise.all([esbuild.context(codeOptions), esbuild.context(uiOptions)]);
  await Promise.all([codeCtx.watch(), uiCtx.watch()]);
  console.log('watching…');
} else {
  await Promise.all([esbuild.build(codeOptions), esbuild.build(uiOptions)]);
}
