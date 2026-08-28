// ============================================================================
// Import/Export - Round-trip support for existing SPFx projects
// ============================================================================

import type { CODBIR, ImportSource, ImportResult, SPFxVersion } from '../types/index.js';
import { createIR, addWebPart, addExtension, addList, addColumn, addGraphPermission } from '../core/ir.js';
import { safeJsonParse } from '../utils/helpers.js';

export class SPFxImporter {

  async import(data: File | Blob | ArrayBuffer | Uint8Array | string): Promise<ImportResult> {
    try {
      // Determine the source type
      const source = this.detectSource(data);

      switch (source) {
        case 'sppkg':
          return await this.importSPPKG(data as File | Blob | ArrayBuffer | Uint8Array);
        case 'spfx-zip':
          return await this.importSPFxZip(data as File | Blob | ArrayBuffer | Uint8Array);
        case 'source-directory':
          return await this.importSourceDirectory(data as string);
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

  private detectSource(data: File | Blob | ArrayBuffer | Uint8Array | string): ImportSource {
    if (typeof data === 'string') {
      try {
        const json = JSON.parse(data);
        if (json.$schema === 'codbsharepoint/ir/1.0') return 'codbsharepoint-json';
        if (json.solution) return 'codbsharepoint-json';
      } catch {}
      return 'source-directory';
    }

    // For binary data, check magic bytes
    let bytes: Uint8Array;
    if (data instanceof ArrayBuffer) {
      bytes = new Uint8Array(data.slice(0, 4));
    } else if (data instanceof Uint8Array) {
      bytes = data.slice(0, 4);
    } else {
      bytes = new Uint8Array(0);
    }

    // ZIP magic bytes: PK
    if (bytes[0] === 0x50 && bytes[1] === 0x4B) {
      // Could be SPPKG or SPFx ZIP
      return 'sppkg';
    }

    return 'auto';
  }

  private async importSPPKG(data: File | Blob | ArrayBuffer | Uint8Array): Promise<ImportResult> {
    const { unzipSync } = await import('fflate');

    let zipData: Uint8Array;
    if (data instanceof ArrayBuffer) {
      zipData = new Uint8Array(data);
    } else if (data instanceof Uint8Array) {
      zipData = data;
    } else {
      zipData = new Uint8Array(await data.arrayBuffer());
    }

    const extracted = unzipSync(zipData);
    return this.parseExtractedFiles(extracted, 'sppkg');
  }

  private async importSPFxZip(data: File | Blob | ArrayBuffer | Uint8Array): Promise<ImportResult> {
    const { unzipSync } = await import('fflate');

    let zipData: Uint8Array;
    if (data instanceof ArrayBuffer) {
      zipData = new Uint8Array(data);
    } else if (data instanceof Uint8Array) {
      zipData = data;
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

  private async importCODBJson(data: string | File | Blob | ArrayBuffer | Uint8Array): Promise<ImportResult> {
    let jsonStr: string;

    if (typeof data === 'string') {
      jsonStr = data;
    } else if (data instanceof File || data instanceof Blob) {
      jsonStr = await data.text();
    } else if (data instanceof Uint8Array) {
      jsonStr = new TextDecoder().decode(data);
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
          } else if (manifest.componentType === 'Extension') {
            const type = mapExtensionType(manifest.extensionType);
            if (type) {
              addExtension(ir, {
                id: manifest.id?.replace(/[{}]/g, ''),
                name: manifest.alias?.split('-').pop() || 'ImportedExtension',
                displayName: manifest.title?.default || '',
                description: manifest.description?.default || 'Extension',
                type
              });
            }
          }
        } catch {
          warnings.push(`Unable to parse manifest: ${path}`);
        }
      }

      // Graph permission requests in Elements.xml
      if (path.endsWith('Elements.xml') || path.endsWith('Feature.xml')) {
        const permMatches = contentStr.match(/<Permission>([^<]+)<\/Permission>/g) || [];
        for (const raw of permMatches) {
          const scope = raw.replace(/<\/?Permission>/g, '').trim();
          if (scope) addGraphPermission(ir, scope);
        }
      }

      // List/library templates in Elements.xml
      if (path.endsWith('Elements.xml')) {
        const listMatches = contentStr.match(/<ListInstance[^>]*Title="([^"]+)"[^>]*TemplateType="(\d+)"/g) || [];
        for (const raw of listMatches) {
          const title = /Title="([^"]+)"/.exec(raw)?.[1];
          const tpl = /TemplateType="(\d+)"/.exec(raw)?.[1];
          if (title && tpl) {
            addList(ir, {
              title,
              description: 'Imported list',
              template: Number(tpl) || 100
            });
          }
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

function mapExtensionType(extensionType?: string): 'ApplicationCustomizer' | 'FieldCustomizer' | 'ListViewCommandSet' | 'FormCustomizer' | undefined {
  switch ((extensionType || '').toLowerCase()) {
    case 'applicationcustomizer':
      return 'ApplicationCustomizer';
    case 'fieldcustomizer':
      return 'FieldCustomizer';
    case 'listviewcommandset':
      return 'ListViewCommandSet';
    case 'formcustomizer':
      return 'FormCustomizer';
    default:
      return undefined;
  }
}
