// ============================================================================
// Bundler - Combines compiled files into production bundles
// ============================================================================

import type { CODBIR, VFSFile, SPFxVersion } from '../types/index.js';
import { createVFS, type VFS } from '../core/vfs.js';

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
        // Create bundle for this component
        const bundleContent = this.createBundle(componentFiles, externals);
        const chunk: BundleChunk = {
          name: `${component.name}.bundle.js`,
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
        const bundleContent = this.createBundle(extFiles, externals);
        const chunk: BundleChunk = {
          name: `${ext.name}.bundle.js`,
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

    // Generate vendor bundle
    const vendorBundle = this.createVendorBundle(ir, externals);
    if (vendorBundle) {
      files.set('vendor.bundle.js', vendorBundle);
      totalSize += vendorBundle.length;
    }

    return {
      success: true,
      chunks,
      externals,
      totalSize,
      files
    };
  }

  private createBundle(files: VFSFile[], externals: string[]): string {
    let bundle = '';
    const moduleMap = new Map<string, string>();

    // Wrap each file as a module
    for (const file of files) {
      let content = typeof file.content === 'string' ? file.content : new TextDecoder().decode(file.content);

      // Replace external requires
      for (const external of externals) {
        const regex = new RegExp(`require\\(['"]${external.replace('/', '\\/')}['"]\\)`, 'g');
        content = content.replace(regex, `window['${external.replace('/', '.')}']`);
      }

      moduleMap.set(file.path, content);
    }

    // Create AMD-compatible bundle
    bundle += `(function() {\n`;
    bundle += `  var modules = {};\n`;
    bundle += `  var cache = {};\n\n`;

    bundle += `  function require(moduleId) {\n`;
    bundle += `    if (cache[moduleId]) return cache[moduleId].exports;\n`;
    bundle += `    var module = cache[moduleId] = { exports: {} };\n`;
    bundle += `    modules[moduleId](module, module.exports, require);\n`;
    bundle += `    return module.exports;\n`;
    bundle += `  }\n\n`;

    let moduleIndex = 0;
    for (const [path, content] of moduleMap) {
      const moduleId = path.replace(/\.js$/, '').replace(/[^a-zA-Z0-9]/g, '_');
      bundle += `  modules['${moduleId}'] = function(module, exports, require) {\n`;
      bundle += `    // ${path}\n`;
      bundle += `    ${content}\n`;
      bundle += `  };\n\n`;
      moduleIndex++;
    }

    // Entry point
    if (moduleMap.size > 0) {
      const firstModule = Array.from(moduleMap.keys())[0];
      const entryId = firstModule.replace(/\.js$/, '').replace(/[^a-zA-Z0-9]/g, '_');
      bundle += `  require('${entryId}');\n`;
    }

    bundle += `})();\n`;

    return bundle;
  }

  private createVendorBundle(ir: CODBIR, externals: string[]): string | null {
    // Create a stub for external dependencies
    let vendor = `(function() {\n`;
    vendor += `  // Vendor bundle - external dependencies\n`;

    for (const external of externals) {
      const namespace = external.replace('/', '.');
      vendor += `  window['${namespace}'] = window['${namespace}'] || {};\n`;
    }

    vendor += `})();\n`;

    return vendor;
  }

  getVFS(): VFS {
    return this.vfs;
  }
}
