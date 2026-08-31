import * as esbuild from 'esbuild';

const isWatch = process.argv.includes('--watch');

const sharedConfig = {
  bundle: true,
  platform: 'browser',
  target: 'es2022',
  format: 'esm',
  sourcemap: true,
  minify: !isWatch,
  metafile: true,
  external: ['esbuild-wasm', 'typescript'],
  define: {
    'process.env.NODE_ENV': isWatch ? '"development"' : '"production"',
    '__VERSION__': JSON.stringify('1.0.0'),
    '__BUILD_TIME__': JSON.stringify(new Date().toISOString())
  }
};

// Main entry
const mainConfig = {
  ...sharedConfig,
  entryPoints: ['src/index.ts'],
  outfile: 'dist/codbsharepoint.mjs',
  format: 'esm'
};

// CommonJS fallback
const cjsConfig = {
  ...sharedConfig,
  entryPoints: ['src/index.ts'],
  outfile: 'dist/codbsharepoint.js',
  format: 'cjs',
  platform: 'browser'
};

// Tools module
const toolsConfig = {
  ...sharedConfig,
  entryPoints: ['src/tools/index.ts'],
  outfile: 'dist/codbsharepoint-tools.mjs',
  format: 'esm'
};

async function build() {
  try {
    if (isWatch) {
      const ctx = await esbuild.context(mainConfig);
      await ctx.watch();
      console.log('Watching for changes...');
    } else {
      await Promise.all([
        esbuild.build(mainConfig),
        esbuild.build(cjsConfig),
        esbuild.build(toolsConfig)
      ]);
      console.log('Build complete!');
    }
  } catch (err) {
    console.error(err);
    process.exit(1);
  }
}

build();
