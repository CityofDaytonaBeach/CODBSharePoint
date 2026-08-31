// ============================================================================
// OPC (Open Packaging Conventions) & SPPKG Generator
// Generates .sppkg packages from SPFx project files
// ============================================================================

import type { CODBIR, VFSFile, VFS, ComponentDefinition } from '../types/index.js';
import { createVFS } from '../core/vfs.js';
import {
  generatePackageSolution,
  generateComponentManifest,
  generateExtensionManifest,
  generateFeatureXml,
  generateElementsXml
} from '../manifest/generator.js';

// OPC Content Types
const OPC_CONTENT_TYPES = {
  xml: 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml',
  rels: 'application/vnd.openxmlformats-package.relationships+xml',
  manifest: 'application/vnd.microsoft.sharepoint.client-side-solution+xml',
  json: 'application/json',
  js: 'application/javascript',
  css: 'text/css',
  default: 'application/octet-stream'
};

// SPPKG specific content types
const SPPKG_CONTENT_TYPES = {
  solution: 'application/vnd.microsoft.sharepoint.client-side-solution',
  manifest: 'application/vnd.microsoft.sharepoint.client-side-solution',
  feature: 'application/xml',
  elements: 'application/xml'
};

// OPC Relationships
interface OPCRelationship {
  type: string;
  target: string;
  id: string;
}

// ---------------------------------------------------------------------------
// Generate OPC Package Structure
// ---------------------------------------------------------------------------

export function generateSPPKG(ir: CODBIR, bundleFiles: Map<string, string | Uint8Array>): Uint8Array {
  const vfs = createVFS();
  const solutionName = ir.solution.name;

  // 1. Root [Content_Types].xml
  const contentTypes = generateContentTypesXml(bundleFiles);
  vfs.addFile('[Content_Types].xml', contentTypes);

  // 2. Root .rels
  const rootRels = generateRootRels(solutionName);
  vfs.addFile('_rels/.rels', rootRels);

  // 2a. OPC package properties referenced by root relationships
  vfs.addFile('docProps/core.xml', generateCorePropertiesXml(ir));
  vfs.addFile('docProps/app.xml', generateAppPropertiesXml(ir));

  // 3. Solution folder
  const solutionFolder = `${solutionName}/`;

  // 4. package-solution.json
  const packageSolution = generatePackageSolution(ir);
  vfs.addFile(`${solutionFolder}package-solution.json`, JSON.stringify(packageSolution, null, 2));

  // 5. package.rels (relationships within the solution)
  const packageRels = generatePackageRels(ir, bundleFiles);
  vfs.addFile(`${solutionFolder}_rels/.rels`, packageRels);

  // 6. Client-side manifests
  for (const component of ir.components) {
    if (component.type === 'webpart') {
      const manifest = generateComponentManifest(component, ir.solution.namespace);
      vfs.addFile(`${solutionFolder}${component.name}.manifest.json`, JSON.stringify(manifest, null, 2));
    }
  }

  for (const ext of ir.extensions) {
    const manifest = generateExtensionManifest(ext, ir.solution.namespace);
    vfs.addFile(`${solutionFolder}${ext.name}.manifest.json`, JSON.stringify(manifest, null, 2));
  }

  // 7. Bundle files
  for (const [path, content] of bundleFiles) {
    const fullPath = `${solutionFolder}${path}`;
    vfs.addFile(fullPath, content);
  }

  // 8. Feature XML
  const featureXml = generateFeatureXml(ir);
  vfs.addFile(`${solutionFolder}Feature.xml`, featureXml);

  // 9. Elements.xml
  const elementsXml = generateElementsXml(ir);
  vfs.addFile(`${solutionFolder}Elements.xml`, elementsXml);

  // 10. Additional metadata
  vfs.addFile(`${solutionFolder}_version.txt`, ir.solution.version);

  // Convert to ZIP
  return vfs.toZip();
}

// ---------------------------------------------------------------------------
// Content Types XML
// ---------------------------------------------------------------------------

function generateContentTypesXml(bundleFiles: Map<string, string | Uint8Array>): string {
  const extensions = new Map<string, string>();

  // Default content types
  extensions.set('rels', OPC_CONTENT_TYPES.rels);
  extensions.set('xml', OPC_CONTENT_TYPES.xml);
  extensions.set('json', OPC_CONTENT_TYPES.json);
  extensions.set('js', OPC_CONTENT_TYPES.js);
  extensions.set('css', OPC_CONTENT_TYPES.css);

  // Detect from bundle files
  for (const [path] of bundleFiles) {
    const ext = path.split('.').pop()?.toLowerCase();
    if (ext && !extensions.has(ext)) {
      extensions.set(ext, OPC_CONTENT_TYPES[ext as keyof typeof OPC_CONTENT_TYPES] || OPC_CONTENT_TYPES.default);
    }
  }

  let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`;

  for (const [ext, contentType] of extensions) {
    xml += `
  <Default Extension="${ext}" ContentType="${contentType}" />`;
  }

  xml += `
</Types>`;

  return xml;
}

// ---------------------------------------------------------------------------
// Root .rels
// ---------------------------------------------------------------------------

function generateRootRels(solutionName: string): string {
  const relationships: OPCRelationship[] = [
    {
      type: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
      target: 'docProps/core.xml',
      id: 'rId1'
    },
    {
      type: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/extended-properties',
      target: 'docProps/app.xml',
      id: 'rId2'
    },
    {
      type: 'http://schemas.microsoft.com/sharepoint/2010/03/containers/container',
      target: `${solutionName}`,
      id: 'rId3'
    }
  ];

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${relationships.map(r => `  <Relationship Id="${r.id}" Type="${r.type}" Target="${r.target}" />`).join('\n')}
</Relationships>`;
}

function generateCorePropertiesXml(ir: CODBIR): string {
  const now = new Date(0).toISOString();
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>${escapeXmlText(ir.solution.name)}</dc:title>
  <dc:creator>${escapeXmlText(ir.solution.author || ir.solution.company)}</dc:creator>
  <cp:lastModifiedBy>CODBSharePoint</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
</cp:coreProperties>`;
}

function generateAppPropertiesXml(ir: CODBIR): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>CODBSharePoint</Application>
  <AppVersion>${escapeXmlText(ir.solution.version)}</AppVersion>
</Properties>`;
}

function escapeXmlText(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// ---------------------------------------------------------------------------
// Package .rels (inside solution folder)
// ---------------------------------------------------------------------------

function generatePackageRels(ir: CODBIR, bundleFiles: Map<string, string | Uint8Array>): string {
  const relationships: OPCRelationship[] = [];
  let rid = 1;

  for (const component of ir.components) {
    relationships.push({
      type: 'http://schemas.microsoft.com/sharepoint/2010/03/containers/manifest',
      target: `${component.name}.manifest.json`,
      id: `rId${rid++}`
    });
  }

  for (const ext of ir.extensions) {
    relationships.push({
      type: 'http://schemas.microsoft.com/sharepoint/2010/03/containers/manifest',
      target: `${ext.name}.manifest.json`,
      id: `rId${rid++}`
    });
  }

  // Add bundle relationships
  for (const [path] of bundleFiles) {
    if (path.endsWith('.js')) {
      relationships.push({
        type: 'http://schemas.microsoft.com/sharepoint/2010/03/containers/entry',
        target: path,
        id: `rId${rid++}`
      });
    }
  }

  // Add feature relationship
  relationships.push({
    type: 'http://schemas.microsoft.com/sharepoint/2010/03/containers/feature',
    target: 'Feature.xml',
    id: `rId${rid++}`
  });

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${relationships.map(r => `  <Relationship Id="${r.id}" Type="${r.type}" Target="${r.target}" />`).join('\n')}
</Relationships>`;
}

// ---------------------------------------------------------------------------
// Helper: Extract manifests from bundle files
// ---------------------------------------------------------------------------

export function extractManifestsFromBundle(bundleFiles: Map<string, string | Uint8Array>): Record<string, unknown>[] {
  const manifests: Record<string, unknown>[] = [];

  for (const [path, content] of bundleFiles) {
    if (path.endsWith('.manifest.json')) {
      try {
        const json = typeof content === 'string' ? content : new TextDecoder().decode(content);
        manifests.push(JSON.parse(json));
      } catch {
        // Skip invalid manifests
      }
    }
  }

  return manifests;
}

// ---------------------------------------------------------------------------
// SPPKG Validation (structural)
// ---------------------------------------------------------------------------

export function validateSPPKGStructure(ir: CODBIR, bundleFiles: Map<string, string | Uint8Array>): string[] {
  const errors: string[] = [];
  const solutionName = ir.solution.name;

  // Check for manifest files. Manifests may be generated directly by the SPPKG
  // generator from IR rather than passed through bundleFiles.
  let hasManifest = ir.components.length > 0 || ir.extensions.length > 0;
  for (const [path] of bundleFiles) {
    if (path.endsWith('.manifest.json')) {
      hasManifest = true;
      break;
    }
  }

  if (!hasManifest) {
    errors.push('No manifest files found in bundle');
  }

  // Check for JS bundles
  let hasBundle = false;
  for (const [path] of bundleFiles) {
    if (path.endsWith('.js') || path.endsWith('.js.map')) {
      hasBundle = true;
      break;
    }
  }

  if (!hasBundle) {
    errors.push('No JavaScript bundle files found');
  }

  // Check component IDs
  const componentIds = ir.components.map(c => c.id);
  const uniqueIds = new Set(componentIds);
  if (componentIds.length !== uniqueIds.size) {
    errors.push('Duplicate component IDs detected');
  }

  // Check manifest IDs match component IDs
  for (const [path, content] of bundleFiles) {
    if (path.endsWith('.manifest.json')) {
      try {
        const json = typeof content === 'string' ? content : new TextDecoder().decode(content);
        const manifest = JSON.parse(json);
        const manifestId = manifest.id?.replace(/[{}]/g, '');
        if (manifestId && !componentIds.includes(manifestId) && !ir.extensions.some(e => e.clientSideComponentId === manifestId)) {
          errors.push(`Manifest ${path} contains unknown component ID: ${manifestId}`);
        }
      } catch {
        errors.push(`Invalid JSON in manifest: ${path}`);
      }
    }
  }

  return errors;
}
