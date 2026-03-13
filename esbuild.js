// =============================================================================
// esbuild.js — Bundle the r-kernel extension for VSCode
//
// Usage:
//   node esbuild.js           # production build
//   node esbuild.js --watch   # watch mode for development
// =============================================================================

const esbuild = require('esbuild');
const path    = require('path');

const watch = process.argv.includes('--watch');
const prod  = process.argv.includes('--prod') || process.env.NODE_ENV === 'production';

/** @type {import('esbuild').BuildOptions} */
const options = {
  entryPoints: ['src/extension.ts'],
  bundle:      true,
  outfile:     'out/extension.js',
  external: [
    // VSCode API — provided at runtime by the extension host
    'vscode',
    // Node.js built-ins bundled by the host
    'path', 'fs', 'os', 'child_process', 'readline', 'events',
    'stream', 'util', 'net', 'http', 'https', 'url', 'crypto',
    'buffer', 'assert', 'tty',
  ],
  platform:   'node',
  target:     'node18',
  format:     'cjs',
  sourcemap:  !prod,
  minify:     prod,
  logLevel:   'info',
  metafile:   false,
  define: {
    'process.env.NODE_ENV': prod ? '"production"' : '"development"',
  },
};

async function main() {
  if (watch) {
    const ctx = await esbuild.context(options);
    await ctx.watch();
    console.log('[r-kernel] Watching for changes…');
  } else {
    const result = await esbuild.build(options);
    if (result.errors.length) process.exit(1);
    console.log('[r-kernel] Build complete →', options.outfile);
  }
}

main().catch(e => { console.error(e); process.exit(1); });
