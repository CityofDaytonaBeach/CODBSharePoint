// ============================================================================
// SharePoint Artifact Generator
// Generates standalone deliverables: themes, formatting, site designs,
// provisioning, and pages (JSON artifacts deployable to SharePoint)
// ============================================================================

import type {
  CODBIR,
  ThemeDefinition,
  FormattingDefinition,
  ProvisioningDefinition,
  PageDefinition,
  VFSFile
} from '../types/index.js';

// ---------------------------------------------------------------------------
// SharePoint Theme
// ---------------------------------------------------------------------------

export function generateThemeJson(theme: ThemeDefinition): Record<string, unknown> {
  const palette = theme.primary || {};

  return {
    name: theme.name,
    isInverted: theme.isInverted ?? false,
    ...palette
  };
}

// ---------------------------------------------------------------------------
// Column / List / Form formatting
// ---------------------------------------------------------------------------

function formattingSchema(type: FormattingDefinition['type']): string {
  switch (type) {
    case 'columnFormatting':
      return 'https://developer.microsoft.com/json-schemas/sp/column-formatting.schema.json';
    case 'listFormatting':
    case 'viewFormatting':
      return 'https://developer.microsoft.com/json-schemas/sp/view-formatting.schema.json';
    case 'formFormatting':
      return 'https://developer.microsoft.com/json-schemas/sp/form-formatting.schema.json';
    default:
      return 'https://developer.microsoft.com/json-schemas/sp/column-formatting.schema.json';
  }
}

export function generateFormattingJson(formatting: FormattingDefinition): Record<string, unknown> {
  return {
    $schema: formattingSchema(formatting.type),
    ...(formatting.json || {})
  };
}

// ---------------------------------------------------------------------------
// Site design / site script / provisioning
// ---------------------------------------------------------------------------

export function generateSiteScript(prov: ProvisioningDefinition): Record<string, unknown> {
  const data = prov.data || {};

  return {
    $schema: 'https://developer.microsoft.com/json-schemas/sp/site-design-script-actions.schema.json',
    actions: data.actions || [],
    ...data
  };
}

export function generateProvisioningJson(prov: ProvisioningDefinition): Record<string, unknown> {
  const data = prov.data || {};

  if (prov.type === 'siteScript') {
    return generateSiteScript(prov);
  }

  if (prov.type === 'siteDesign') {
    return {
      $schema: 'https://developer.microsoft.com/json-schemas/sp/site-design.schema.json',
      title: prov.name,
      description: prov.description || '',
      ...(data.siteScriptIds ? { siteScriptIds: data.siteScriptIds } : {}),
      ...(data.webTemplate ? { webTemplate: data.webTemplate } : {}),
      ...(data.previewImageUrl ? { previewImageUrl: data.previewImageUrl } : {})
    };
  }

  // list / library / column / contentType
  return {
    displayName: prov.name,
    description: prov.description || '',
    ...data
  };
}

// ---------------------------------------------------------------------------
// Pages
// ---------------------------------------------------------------------------

export function generatePageJson(page: PageDefinition): Record<string, unknown> {
  return {
    name: page.name,
    title: page.title,
    layout: page.layout || 'TitleAndDescription',
    promotedState: page.promotedState ?? 0,
    content: page.content
  };
}

// ---------------------------------------------------------------------------
// Top-level generator: produce files for all non-component artifacts
// ---------------------------------------------------------------------------

export function generateSharePointArtifacts(ir: CODBIR): VFSFile[] {
  const files: VFSFile[] = [];

  for (const theme of ir.themes) {
    files.push({
      path: `sharepoint/themes/${theme.name}.json`,
      content: JSON.stringify(generateThemeJson(theme), null, 2),
      encoding: 'utf-8'
    });
  }

  for (const formatting of ir.formatting) {
    files.push({
      path: `sharepoint/formatting/${formatting.name}.json`,
      content: JSON.stringify(generateFormattingJson(formatting), null, 2),
      encoding: 'utf-8'
    });
  }

  for (const prov of ir.provisioning) {
    const dir = prov.type === 'siteDesign' ? 'sharepoint/site-designs' : 'sharepoint/provisioning';
    files.push({
      path: `${dir}/${prov.name}.json`,
      content: JSON.stringify(generateProvisioningJson(prov), null, 2),
      encoding: 'utf-8'
    });
  }

  for (const page of ir.pages) {
    files.push({
      path: `sharepoint/pages/${page.name}.json`,
      content: JSON.stringify(generatePageJson(page), null, 2),
      encoding: 'utf-8'
    });
  }

  return files;
}
