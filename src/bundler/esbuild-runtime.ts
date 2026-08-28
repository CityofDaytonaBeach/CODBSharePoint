// ============================================================================
// Esbuild Runtime Wrapper
// Lazily loads esbuild-wasm (offline-friendly) to perform real JSX / TypeScript
// transforms and module bundling. Falls back to a no-op when esbuild-wasm is
// not available (e.g. standalone CDN builds) so the rest of the SDK still works.
// ============================================================================

export interface TransformOptions {
  loader: 'ts' | 'tsx';
  minify?: boolean;
  sourceMap?: boolean;
  target?: string;
}

export interface TransformResult {
  ok: boolean;
  code?: string;
  map?: string;
  error?: string;
}

export interface BundleInput {
  path: string;
  content: string;
}

export interface FileMapEntry {
  path: string;
  content: string;
}

export interface BuildOptions {
  bundle?: boolean;
  format?: 'iife' | 'esm' | 'cjs';
  minify?: boolean;
  sourceMap?: boolean;
  target?: string;
  platform?: 'browser' | 'neutral' | 'node';
  external?: string[];
  define?: Record<string, string>;
}

export interface BuildResult {
  ok: boolean;
  code?: string;
  map?: string;
  error?: string;
}

type EsbuildModule = typeof import('esbuild-wasm');

let cachedModule: EsbuildModule | undefined;
let initialized = false;
let initPromise: Promise<void> | undefined;
let availabilityError: string | undefined;

// Injected by exports.initializeEsbuild() for bundlers/CDN that use WasmURL.
let wasmURLOverride: string | undefined;

async function resolveModule(): Promise<EsbuildModule | undefined> {
  if (cachedModule) return cachedModule;
  if (availabilityError !== undefined) return undefined;

  try {
    const mod = (await import(/* webpackIgnore: true */ 'esbuild-wasm')) as EsbuildModule;
    // In browser builds via bundlers the package resolves; in pure CDN ESM the
    // bare specifier cannot resolve — we catch that above and fall back.
    cachedModule = mod;
    return mod;
  } catch (err) {
    availabilityError = err instanceof Error ? err.message : String(err);
    cachedModule = undefined;
    return undefined;
  }
}

async function ensureInit(mod: EsbuildModule): Promise<void> {
  if (initialized) return;
  if (!initPromise) {
    initPromise = (async () => {
      const opts: Record<string, unknown> = {};
      if (wasmURLOverride) opts.wasmURL = wasmURLOverride;
      await mod.initialize(opts);
      initialized = true;
    })();
  }
  await initPromise;
}

export function setWasmURL(url: string): void {
  wasmURLOverride = url;
}

export async function isAvailable(): Promise<boolean> {
  const mod = await resolveModule();
  return !!mod;
}

export async function transformContent(
  content: string,
  options: TransformOptions
): Promise<TransformResult> {
  const mod = await resolveModule();
  if (!mod) {
    return { ok: false, error: 'esbuild-wasm is not available' };
  }

  try {
    await ensureInit(mod);
    const result = await mod.transform(content, {
      loader: options.loader,
      minify: options.minify ?? false,
      sourcemap: options.sourceMap ?? false,
      target: options.target || 'es2022'
    });
    return {
      ok: true,
      code: result.code,
      map: result.map
    };
  } catch (err) {
    return { ok: false, error: err instanceof Error ? err.message : String(err) };
  }
}

export async function bundleContent(
  entry: BundleInput,
  options: BuildOptions = {}
): Promise<BuildResult> {
  const mod = await resolveModule();
  if (!mod) {
    return { ok: false, error: 'esbuild-wasm is not available' };
  }

  try {
    await ensureInit(mod);
    const result = await mod.build({
      stdin: {
        contents: entry.content,
        loader: /\.tsx$/.test(entry.path) ? 'tsx' : /\.ts$/.test(entry.path) ? 'ts' : 'js',
        resolveDir: '.'
      },
      bundle: options.bundle ?? true,
      format: options.format || 'iife',
      minify: options.minify ?? false,
      sourcemap: options.sourceMap ?? false,
      target: options.target || 'es2022',
      platform: options.platform || 'browser',
      external: options.external || [],
      define: options.define,
      write: false
    });

    const out = result.outputFiles?.[0];
    if (!out) {
      return { ok: false, error: 'esbuild produced no output' };
    }
    const text = typeof out.text === 'string' ? out.text : new TextDecoder().decode(out.contents);
    return { ok: true, code: text };
  } catch (err) {
    return { ok: false, error: err instanceof Error ? err.message : String(err) };
  }
}

// ----------------------------------------------------------------------------
// VFS-backed bundling
// Bundles a module graph from an in-memory file map (no filesystem needed),
// marking SPFx/React externals and CSS imports as external.
// ----------------------------------------------------------------------------

const DEFAULT_EXTERNALS = [
  'react',
  'react-dom',
  /^@microsoft\//
];

export interface BundleVFSResult {
  ok: boolean;
  code?: string;
  error?: string;
}

export async function bundleFromVFS(
  entryPath: string,
  files: FileMapEntry[],
  options: BuildOptions = {}
): Promise<BundleVFSResult> {
  const mod = await resolveModule();
  if (!mod) {
    return { ok: false, error: 'esbuild-wasm is not available' };
  }

  try {
    await ensureInit(mod);

    const fileMap = new Map<string, string>();
    for (const f of files) {
      fileMap.set(f.path, f.content);
      fileMap.set(f.path.replace(/\.js$/, ''), f.content);
    }

    const isExternal = (specifier: string): boolean => {
      if (options.external?.includes(specifier)) return true;
      if (specifier === 'react' || specifier === 'react-dom') return true;
      if (specifier.startsWith('@microsoft/')) return true;
      if (specifier.endsWith('.css') || specifier.endsWith('.scss')) return true;
      return false;
    };

    const resolveCandidate = (base: string): string | undefined => {
      if (fileMap.has(base)) return base;
      if (!base.endsWith('.js') && fileMap.has(`${base}.js`)) return `${base}.js`;
      if (!base.endsWith('.json') && fileMap.has(`${base}.json`)) return `${base}.json`;
      return undefined;
    };

    const normalize = (dir: string, specifier: string): string => {
      const parts = (dir === '' ? [] : dir.split('/')).concat(specifier.split('/'));
      const out: string[] = [];
      for (const p of parts) {
        if (p === '' || p === '.') continue;
        if (p === '..') { out.pop(); continue; }
        out.push(p);
      }
      return out.join('/');
    };

    const plugin: any = {
      name: 'codb-vfs',
      setup(build: any) {
        build.onResolve({ filter: /.*/ }, (args: { path: string; importer: string; resolveDir: string; kind: string }) => {
          if (isExternal(args.path)) {
            return { path: args.path, external: true };
          }
          const importerDir = args.importer === '<stdin>'
            ? (entryPath.includes('/') ? entryPath.slice(0, entryPath.lastIndexOf('/')) : '')
            : (args.importer.includes('/') ? args.importer.slice(0, args.importer.lastIndexOf('/')) : '');
          const candidate = normalize(importerDir, args.path);
          const resolved = resolveCandidate(candidate);
          if (resolved) {
            return { path: resolved, namespace: 'codb-vfs' };
          }
          return { errors: [{ text: `Cannot resolve "${args.path}" from "${args.importer}"` }] };
        });

        build.onLoad({ filter: /.*/, namespace: 'codb-vfs' }, (args: { path: string }) => {
          const contents = fileMap.get(args.path);
          if (contents === undefined) {
            return { errors: [{ text: `File not found: ${args.path}` }] };
          }
          return { contents, loader: args.path.endsWith('.json') ? 'json' : 'js' };
        });
      }
    };

    const result = await mod.build({
      entryPoints: [entryPath],
      bundle: options.bundle ?? true,
      format: options.format || 'iife',
      minify: options.minify ?? false,
      sourcemap: options.sourceMap ?? false,
      target: options.target || 'es2022',
      platform: options.platform || 'browser',
      write: false,
      plugins: [plugin]
    });

    const out = result.outputFiles?.[0];
    if (!out) {
      return { ok: false, error: 'esbuild produced no output' };
    }
    const text = typeof out.text === 'string' ? out.text : new TextDecoder().decode(out.contents);
    return { ok: true, code: text };
  } catch (err) {
    return { ok: false, error: err instanceof Error ? err.message : String(err) };
  }
}
