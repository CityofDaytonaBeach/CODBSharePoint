// ============================================================================
// Esbuild Runtime Wrapper
// Lazily loads esbuild-wasm (offline-friendly) to perform real JSX / TypeScript
// transforms and module bundling. Falls back to a no-op when esbuild-wasm is
// not available (e.g. standalone CDN builds) so the rest of the SDK still works.
// ============================================================================

export interface TransformOptions {
  loader: 'js' | 'jsx' | 'ts' | 'tsx';
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

// CDN ESM fallback so pure browser/ESM runs can resolve esbuild-wasm without a
// bundler or node_modules (the bare `import('esbuild-wasm')` would 404).
const ESBUILD_WASM_VERSION = '0.28.2';
const CDN_BASE = 'https://cdn.jsdelivr.net/npm';
let esbuildCDNURL: string | undefined;

async function resolveModule(): Promise<EsbuildModule | undefined> {
  if (cachedModule) return cachedModule;

  // 1. Try the bundled/native resolution (works when a bundler inlined it or
  //    when importing the npm package in a Node/test environment).
  if (availabilityError === undefined) {
    try {
      const mod = (await import(/* webpackIgnore: true */ 'esbuild-wasm')) as EsbuildModule;
      if (mod && typeof mod.transform === 'function') {
        cachedModule = mod;
        return mod;
      }
    } catch {
      // fall through to CDN
    }
  }

  // 2. Browser/ESM fallback: import the esbuild-wasm ESM build from a CDN.
  if (typeof globalThis !== 'undefined' && !globalThis.process) {
    esbuildCDNURL = esbuildCDNURL || `${CDN_BASE}/esbuild-wasm@${ESBUILD_WASM_VERSION}/lib/browser.js`;
    try {
      // eslint-disable-next-line @typescript-eslint/no-var-requires
      const mod = (await import(/* webpackIgnore: true */ /* @vite-ignore */ esbuildCDNURL)) as unknown as EsbuildModule;
      if (mod && typeof mod.transform === 'function') {
        cachedModule = mod;
        return mod;
      }
      availabilityError = 'esbuild-wasm CDN module has no transform function';
    } catch (err) {
      availabilityError = err instanceof Error ? err.message : String(err);
      return undefined;
    }
  }

  availabilityError = availabilityError || 'esbuild-wasm is not available in this environment';
  return undefined;
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
        loader: /\.tsx$/.test(entry.path) ? 'tsx' : /\.ts$/.test(entry.path) ? 'ts' : /\.jsx$/.test(entry.path) ? 'jsx' : 'js',
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
// marking SPFx/React runtime modules external while resolving project CSS.
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
  /** Structured esbuild diagnostics (file, line, column, text) — preserved so callers can surface precise errors. */
  warnings?: BundleNote[];
  errors?: BundleNote[];
}

export interface BundleNote {
  pluginName?: string;
  text: string;
  location?: {
    file?: string;
    line?: number;
    column?: number;
    length?: number;
    lineText?: string;
  };
  detail?: string;
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
      return false;
    };

    const resolveCandidate = (base: string): string | undefined => {
      if (fileMap.has(base)) return base;
      if (!base.endsWith('.js') && fileMap.has(`${base}.js`)) return `${base}.js`;
      if (base.endsWith('.scss') && fileMap.has(base.replace(/\.scss$/, '.css'))) return base.replace(/\.scss$/, '.css');
      if (!base.endsWith('.json') && fileMap.has(`${base}.json`)) return `${base}.json`;
      return undefined;
    };

    const cssModuleObject = (path: string, contents: string): string => {
      const classes = new Set<string>();
      const classPattern = /\.([_a-zA-Z]+[_a-zA-Z0-9-]*)/g;
      let match: RegExpExecArray | null;

      while ((match = classPattern.exec(contents)) !== null) {
        classes.add(match[1]);
      }

      const mappings = Array.from(classes).sort().map(name => {
        return `  ${JSON.stringify(name)}: ${JSON.stringify(name)}`;
      }).join(',\n');

      return `const css = ${JSON.stringify(contents)};
if (typeof document !== 'undefined') {
  const style = document.createElement('style');
  style.setAttribute('data-codb-source', ${JSON.stringify(path)});
  style.textContent = css;
  document.head.appendChild(style);
}
export default {\n${mappings}\n};`;
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
          if (args.path.endsWith('.css')) {
            return { contents: cssModuleObject(args.path, contents), loader: 'js' };
          }
          return { contents, loader: args.path.endsWith('.json') ? 'json' : args.path.endsWith('.jsx') ? 'jsx' : 'js' };
        });
      }
    };

    let result;
    try {
      result = await mod.build({
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
    } catch (err) {
      // esbuild throws with structured .errors and .warnings arrays (member / location / text).
      const buildError = err as { message?: string; errors?: BundleNote[]; warnings?: BundleNote[] };
      const errors = normalizeStructuredMessages(buildError.errors);
      const warnings = normalizeStructuredMessages(buildError.warnings);
      const summary = errors.length > 0
        ? errors.map(e => formatNote(e)).join(' | ')
        : (buildError.message || (err instanceof Error ? err.message : String(err)));
      return { ok: false, error: summary, errors, warnings };
    }

    // Surface non-fatal structured esbuild errors (e.g. syntax errors that still
    // produced a partial output, or strict-mode violations) so callers see them.
    const structuredErrors = normalizeStructuredMessages((result as any)?.errors);
    const structuredWarnings = normalizeStructuredMessages((result as any)?.warnings);

    if (structuredErrors.length > 0) {
      return {
        ok: false,
        error: structuredErrors.map(e => formatNote(e)).join(' | '),
        errors: structuredErrors,
        warnings: structuredWarnings
      };
    }

    const out = result.outputFiles?.[0];
    if (!out) {
      return { ok: false, error: 'esbuild produced no output', warnings: structuredWarnings };
    }
    const text = typeof out.text === 'string' ? out.text : new TextDecoder().decode(out.contents);
    return { ok: true, code: text, warnings: structuredWarnings };
  } catch (err) {
    // Outer guard: never let an unexpected runtime error silently pass as success.
    return { ok: false, error: err instanceof Error ? err.message : String(err) };
  }
}

// ---------------------------------------------------------------------------
// Structured message normalization
// ---------------------------------------------------------------------------

function normalizeStructuredMessages(notes?: Array<{
  text?: string;
  detail?: string;
  location?: { file?: string; line?: number; column?: number; length?: number; lineText?: string };
  pluginName?: string;
} | string>): BundleNote[] {
  if (!Array.isArray(notes)) return [];
  const out: BundleNote[] = [];
  for (const note of notes) {
    if (typeof note === 'string') {
      out.push({ text: note });
    } else if (note && typeof note.text === 'string') {
      const n: BundleNote = { text: note.text };
      if (note.detail) n.detail = note.detail;
      if (note.pluginName) n.pluginName = note.pluginName;
      if (note.location) {
        n.location = {
          file: note.location.file,
          line: note.location.line,
          column: note.location.column,
          length: note.location.length,
          lineText: note.location.lineText
        };
      }
      out.push(n);
    }
  }
  return out;
}

function formatNote(note: BundleNote): string {
  const loc = note.location;
  const pos = loc && (loc.file || loc.line != null)
    ? ` (${loc.file || '<entry>'}${loc.line != null ? `:${loc.line}${loc.column != null ? `:${loc.column}` : ''}` : ''})`
    : '';
  const detail = note.detail ? `\n  ${note.detail}` : '';
  return `[esbuild]${pos} ${note.text}${detail}`;
}
