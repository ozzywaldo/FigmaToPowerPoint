// Bundles the headless extractor as a single expression string for scripted execution.
import * as esbuild from 'esbuild';
import fs from 'node:fs/promises';
const r = await esbuild.build({
  entryPoints: ['src/plugin/headless.ts'],
  bundle: true,
  write: false,
  format: 'iife',
  globalName: '__headless',
  target: 'es2019',
  platform: 'neutral',
  minify: true,
  legalComments: 'none',
});
await fs.mkdir('dist', { recursive: true });
await fs.writeFile('dist/headless.js', r.outputFiles[0].text);
console.log(`dist/headless.js ${r.outputFiles[0].text.length} chars`);
await esbuild.build({
  entryPoints: ['src/headless/emit.ts'],
  bundle: true,
  outfile: 'dist/emit.mjs',
  platform: 'node',
  format: 'esm',
  target: 'node20',
  packages: 'external',
  logLevel: 'info',
});
