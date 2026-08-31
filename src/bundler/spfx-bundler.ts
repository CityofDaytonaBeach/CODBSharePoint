// ============================================================================
// Bundler - Combines compiled files into production bundles
// ============================================================================

import type { CODBIR, VFSFile, SPFxVersion } from '../types/index.js';
import { createVFS, type VFS } from '../core/vfs.js';
import { bundleFromVFS } from './esbuild-runtime.js';

export interface BundleOptions {
  minify?: boolean;
  sourceMaps?: boolean;
  treeshake?: boolean;
  external?: string[];
  format?: 'amd' | 'system' | 'esm';
}

export interface BundleResult {
  success: boolean;
  chunks: BundleChunk[];
  externals: string[];
  totalSize: number;
  files: Map<string, string | Uint8Array>;
  errors: string[];
}

export interface BundleChunk {
  name: string;
  content: string;
  size: number;
  isEntry: boolean;
  modules: string[];
}

export class SPFxBundle {
  private vfs: VFS;

  constructor() {
    this.vfs = createVFS();
  }

  async bundle(ir: CODBIR, compiledFiles: VFSFile[], options: BundleOptions = {}): Promise<BundleResult> {
    const chunks: BundleChunk[] = [];
    const files = new Map<string, string | Uint8Array>();
    const errors: string[] = [];
    let totalSize = 0;

    const externals = [
      'react',
      'react-dom',
      '@microsoft/sp-core-library',
      '@microsoft/sp-lodash-subset',
      '@microsoft/sp-property-pane',
      '@microsoft/sp-http',
      '@microsoft/sp-webpart-base',
      '@microsoft/sp-application-base',
      '@microsoft/sp-listview-extensibility',
      ...(options.external || [])
    ];

    // Process each component
    for (const component of ir.components) {
      const componentFiles = compiledFiles.filter(f => f.path.includes(component.name));

      if (componentFiles.length > 0) {
        const name = component.name;
        const candidates = this.findEntryCandidates(componentFiles, name);
        const bundle = await this.bundleEntry(componentFiles, candidates, externals, options);

        if (!bundle.success || !bundle.content) {
          errors.push(bundle.error || `Failed to bundle component ${name}`);
          continue;
        }

        const bundleContent = this.wrapAsAmdModule(`${name}.bundle`, bundle.content, externals);
        const chunk: BundleChunk = {
          name: `${name}.bundle.js`,
          content: bundleContent,
          size: bundleContent.length,
          isEntry: true,
          modules: componentFiles.map(f => f.path)
        };

        chunks.push(chunk);
        files.set(chunk.name, bundleContent);
        totalSize += chunk.size;
      }
    }

    // Process extensions
    for (const ext of ir.extensions) {
      const extFiles = compiledFiles.filter(f => f.path.includes(ext.name));

      if (extFiles.length > 0) {
        const name = ext.name;
        const candidates = this.findEntryCandidates(extFiles, name);
        const bundle = await this.bundleEntry(extFiles, candidates, externals, options);

        if (!bundle.success || !bundle.content) {
          errors.push(bundle.error || `Failed to bundle extension ${name}`);
          continue;
        }

        const bundleContent = this.wrapAsAmdModule(`${name}.bundle`, bundle.content, externals);
        const chunk: BundleChunk = {
          name: `${name}.bundle.js`,
          content: bundleContent,
          size: bundleContent.length,
          isEntry: true,
          modules: extFiles.map(f => f.path)
        };

        chunks.push(chunk);
        files.set(chunk.name, bundleContent);
        totalSize += chunk.size;
      }
    }

    // Process styles
    const styleFiles = compiledFiles.filter(f => f.path.endsWith('.css'));
    for (const styleFile of styleFiles) {
      const cssContent = typeof styleFile.content === 'string' ? styleFile.content : new TextDecoder().decode(styleFile.content);
      files.set(`${styleFile.path.split('/').pop()}`, cssContent);
      totalSize += cssContent.length;
    }

    return {
      success: errors.length === 0 && chunks.length > 0,
      chunks,
      externals,
      totalSize,
      files,
      errors
    };
  }

  private findEntryCandidates(files: VFSFile[], name: string): string[] {
    const candidates = files
      .map(f => f.path)
      .filter(p => p.endsWith('.js'))
      .filter(p => !p.endsWith('.map'))
      .sort((a, b) => this.entryScore(a, name) - this.entryScore(b, name));
    return candidates;
  }

  private entryScore(path: string, name: string): number {
    let score = 1000;
    if (/WebPart|ApplicationCustomizer|FieldCustomizer|CommandSet/.test(path)) score -= 300;
    if (path.includes(name)) score -= 100;
    if (path.includes(`${name}WebPart`)) score -= 100;
    return score;
  }

  private async bundleEntry(
    files: VFSFile[],
    candidates: string[],
    externals: string[],
    options: BundleOptions
  ): Promise<{ success: boolean; content?: string; error?: string }> {
    const entry = candidates[0];
    if (!entry) {
      return { success: false, error: 'No JavaScript entry point found for bundle' };
    }

    const fileMap = files.map(f => ({
      path: f.path,
      content: typeof f.content === 'string' ? f.content : new TextDecoder().decode(f.content)
    }));

    const result = await bundleFromVFS(entry, fileMap, {
      bundle: true,
      format: 'cjs',
      minify: options.minify ?? false,
      sourceMap: options.sourceMaps ?? false,
      external: externals,
      target: 'es2022',
      platform: 'browser'
    });

    if (result.ok && result.code) {
      return { success: true, content: result.code };
    }
    return { success: false, error: result.error || `Failed to bundle entry ${entry}` };
  }

  private wrapAsAmdModule(name: string, commonJsBundle: string, externals: string[]): string {
    const deps = Array.from(new Set(externals));
    const depArgs = deps.map((_, index) => `dep${index}`);
    const moduleMap = deps.map((dep, index) => `    ${JSON.stringify(dep)}: dep${index}`).join(',\n');

    return `(function (root, factory) {
  if (typeof define === 'function' && define.amd) {
    define(${JSON.stringify(name)}, ${JSON.stringify(deps)}, factory);
  } else {
    root[${JSON.stringify(name)}] = factory(${depArgs.map(arg => `root.${arg}`).join(', ')});
  }
})(typeof self !== 'undefined' ? self : this, function (${depArgs.join(', ')}) {
  var __codbModules = {
${moduleMap}
  };
  var require = function (id) {
    if (Object.prototype.hasOwnProperty.call(__codbModules, id)) return __codbModules[id];
    throw new Error('SPFx external not provided: ' + id);
  };
  var module = { exports: {} };
  var exports = module.exports;
${commonJsBundle.split('\n').map(line => `  ${line}`).join('\n')}
  return module.exports && module.exports.__esModule && module.exports.default ? module.exports.default : module.exports;
});
`;
  }

  getVFS(): VFS {
    return this.vfs;
  }
}
