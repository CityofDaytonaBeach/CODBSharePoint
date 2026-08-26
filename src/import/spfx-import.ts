// ============================================================================
// Import/Export - Round-trip support for existing SPFx projects
// ============================================================================

import type { CODBIR, ImportSource, ImportResult, SPFxVersion } from '../types/index.js';
import { createIR, addWebPart, addList, addColumn, addGraphPermission } from '../core/ir.js';
import { safeJsonParse } from '../utils/helpers.js';

export class SPFxImporter {

  async import(data: File | Blob | ArrayBuffer | string): Promise<ImportResult> {
    try {
      // Determine the source type
      const source = this.detectSource(data);

      switch (source) {
        case 'sppkg':
          return await this.importSPPKG(data);
        case 'spfx-zip':
          return await this.importSPFxZip(data);
        case 'source-directory':
          return await this.importSourceDirectory(data);
        case 'codbsharepoint-json':
          return await this.importCODBJson(data);
        default:
          return {
            success: false,
            ir: createIR(),
            source: 'auto',
            warnings: ['Unable to detect source format']
          };
      }
    } catch (error) {
      return {
        success: false,
        ir: createIR(),
        source: 'auto',
        warnings: [error instanceof Error ? error.message : 'Import failed']
      };
    }
  }

  private detectSource(data: File | Blob | ArrayBuffer | string): ImportSource {
    if (typeof data === 'string') {
      try {
        const json = JSON.parse(data);
        if (json.$schema === 'codbsharepoint/ir/1.0') return 'codbsharepoint-json';
        if (json.solution) return 'codbsharepoint-json';
      } catch {}
      return 'source-directory';
    }

    // For binary data, check magic bytes
    const bytes = data instanceof ArrayBuffer
      ? new Uint8Array(data.slice(0, 4))
      : new Uint8Array(0);

    // ZIP magic bytes: PK
    if (bytes[0] === 0x50 && bytes[1] === 0x4B) {
      // Could be SPPKG or SPFx ZIP
      return 'sppkg';
    }

    return 'auto';
  }

  private async importSPPKG(data: File | Blob | ArrayBuffer): Promise<ImportResult> {
    const { unzipSync } = await import('fflate');

    let zipData: Uint8Array;
    if (data instanceof ArrayBuffer) {
      zipData = new Uint8Array(data);
    } else if (data instanceof Blob) {
      zipData = new Uint8Array(await data.arrayBuffer());
    } else {
      zipData = new Uint8Array(await data.arrayBuffer());
    }

    const extracted = unzipSync(zipData);
    return this.parseExtractedFiles(extracted, 'sppkg');
  }

  private async importSPFxZip(data: File | Blob | ArrayBuffer): Promise<ImportResult> {
    const { unzipSync } = await import('fflate');

    let zipData: Uint8Array;
    if (data instanceof ArrayBuffer) {
      zipData = new Uint8Array(data);
    } else if (data instanceof Blob) {
      zipData = new Uint8Array(await data.arrayBuffer());
    } else {
      zipData = new Uint8Array(await data.arrayBuffer());
    }

    const extracted = unzipSync(zipData);
    return this.parseExtractedFiles(extracted, 'spfx-zip');
  }

  private async importSourceDirectory(data: string): Promise<ImportResult> {
    const ir = createIR();
    const warnings: string[] = [];

    // Parse source directory structure
    try {
      const files = JSON.parse(data);
      if (typeof files === 'object') {
        // Try to find package-solution.json
        for (const [path, content] of Object.entries(files)) {
          if (path.includes('package-solution.json')) {
            const solution = JSON.parse(content as string);
            if (solution.solution) {
              ir.solution.name = solution.solution.name || ir.solution.name;
              ir.solution.version = solution.solution.version || ir.solution.version;
            }
          }
        }
      }
    } catch {
      warnings.push('Unable to parse source directory');
    }

    return {
      success: true,
      ir,
      source: 'source-directory',
      warnings
    };
  }

  private async importCODBJson(data: string | File | Blob | ArrayBuffer): Promise<ImportResult> {
    let jsonStr: string;

    if (typeof data === 'string') {
      jsonStr = data;
    } else if (data instanceof File || data instanceof Blob) {
      jsonStr = await data.text();
    } else {
      jsonStr = new TextDecoder().decode(data);
    }

    try {
      const ir = JSON.parse(jsonStr) as CODBIR;
      return {
        success: true,
        ir,
        source: 'codbsharepoint-json',
        warnings: []
      };
    } catch (error) {
      return {
        success: false,
        ir: createIR(),
        source: 'codbsharepoint-json',
        warnings: ['Invalid JSON format']
      };
    }
  }

  private async parseExtractedFiles(
    extracted: Record<string, Uint8Array>,
    source: ImportSource
  ): Promise<ImportResult> {
    const ir = createIR();
    const warnings: string[] = [];
    const decoder = new TextDecoder();

    // Look for key files
    for (const [path, content] of Object.entries(extracted)) {
      const contentStr = decoder.decode(content);

      // package-solution.json
      if (path.endsWith('package-solution.json')) {
        try {
          const solution = JSON.parse(contentStr);
          if (solution.solution) {
            ir.solution.name = solution.solution.name || ir.solution.name;
            ir.solution.version = solution.solution.version || ir.solution.version;
            ir.solution.description = solution.solution.description || '';
            ir.solution.id = solution.solution.id || ir.solution.id;
          }
        } catch {
          warnings.push('Unable to parse package-solution.json');
        }
      }

      // Manifest files
      if (path.endsWith('.manifest.json')) {
        try {
          const manifest = JSON.parse(contentStr);
          if (manifest.componentType === 'WebPart') {
            addWebPart(ir, {
              id: manifest.id?.replace(/[{}]/g, ''),
              name: manifest.alias?.split('-').pop() || 'ImportedWebPart',
              displayName: manifest.title?.default || '',
              description: manifest.description?.default || '',
              version: manifest.version || '1.0.0',
              framework: 'react'
            });
          }
        } catch {
          warnings.push(`Unable to parse manifest: ${path}`);
        }
      }
    }

    return {
      success: true,
      ir,
      source,
      detectedVersion: '1.22.0' as SPFxVersion,
      warnings
    };
  }
}
