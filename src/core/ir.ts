// ============================================================================
// IR Factory - Creates and manipulates the CODBSharePoint IR
// ============================================================================

import { randomUUID } from '../utils/crypto.js';
import type {
  CODBIR,
  SolutionConfig,
  ComponentDefinition,
  ExtensionDefinition,
  ACEDefinition,
  ListDefinition,
  LibraryDefinition,
  FieldDefinition,
  ContentTypeDefinition,
  PermissionDefinition,
  GraphPermissionDefinition,
  ThemeDefinition,
  FormattingDefinition,
  ProvisioningDefinition,
  PageDefinition,
  SPFxVersion,
  Framework,
  ComponentType,
  ExtensionType,
  PropertyPaneDefinition,
  PropertyPaneField,
  BuildTarget
} from '../types/index.js';
import { SPFx_COMPATIBILITY, GRAPH_PERMISSIONS } from '../types/index.js';

export function createIR(config?: Partial<SolutionConfig>): CODBIR {
  const now = new Date().toISOString();
  const id = randomUUID();

  return {
    $schema: 'codbsharepoint/ir/1.0',
    solution: {
      name: config?.name || 'NewSolution',
      id: config?.id || id,
      version: config?.version || '1.0.0',
      description: config?.description || '',
      author: config?.author || '',
      company: config?.company || 'City of Daytona Beach',
      namespace: config?.namespace || 'NewSolution',
      environment: config?.environment || 'spo',
      includeClientSideAssets: config?.includeClientSideAssets ?? true,
      skipFeatureDeployment: config?.skipFeatureDeployment ?? false,
      isDomainIsolated: config?.isDomainIsolated ?? false,
      developer: config?.developer || {
        name: config?.author || '',
        websiteUrl: ''
      },
      metadata: config?.metadata || {
        screenshot: ''
      },
      features: config?.features || []
    },
    components: [],
    lists: [],
    libraries: [],
    fields: [],
    contentTypes: [],
    pages: [],
    extensions: [],
    permissions: [],
    graph: [],
    themes: [],
    formatting: [],
    provisioning: [],
    localization: {
      defaultLanguage: 'en-us',
      languages: []
    },
    metadata: {
      generator: 'codbsharepoint',
      version: __VERSION__ || '1.0.0',
      createdAt: now,
      modifiedAt: now,
      spfxVersion: '1.22.0',
      buildTool: 'heft'
    }
  };
}

export function addWebPart(ir: CODBIR, config: Partial<ComponentDefinition>): ComponentDefinition {
  const id = config.id || randomUUID();
  const name = config.name || 'NewWebPart';

  const webPart: ComponentDefinition = {
    type: 'webpart',
    id,
    name,
    displayName: config.displayName || name,
    description: config.description || '',
    officeFabricIconFontName: config.officeFabricIconFontName,
    iconUrl: config.iconUrl,
    group: config.group || { id: '5c03119e-3074-46fd-976b-c60198311f70', name: 'Other' },
    version: config.version || '1.0.0',
    entry: config.entry || `src/webparts/${name}/${name}WebPart.ts`,
    entryModule: config.entryModule,
    framework: config.framework || 'react',
    properties: config.properties || [],
    preconfiguredEntries: config.preconfiguredEntries || [],
    assets: config.assets,
    contextSpecific: config.contextSpecific,
    supportedHosts: config.supportedHosts || ['SharePointWebPart']
  };

  ir.components.push(webPart);
  ir.metadata.modifiedAt = new Date().toISOString();
  return webPart;
}

export function addExtension(ir: CODBIR, config: Partial<ExtensionDefinition>): ExtensionDefinition {
  const id = config.id || randomUUID();
  const name = config.name || 'NewExtension';

  const ext: ExtensionDefinition = {
    type: config.type || 'ApplicationCustomizer',
    id,
    name,
    displayName: config.displayName || name,
    description: config.description || '',
    entry: config.entry || `src/extensions/${name}/${name}ApplicationCustomizer.ts`,
    clientSideComponentId: id,
    registeredClientSideExtensions: config.registeredClientSideExtensions,
    topNavigationZone: config.topNavigationZone,
    bottomNavigationZone: config.bottomNavigationZone,
    pageAction: config.pageAction,
    pageHeader: config.pageHeader,
    ribbon: config.ribbon
  };

  ir.extensions.push(ext);
  ir.metadata.modifiedAt = new Date().toISOString();
  return ext;
}

export function addACE(ir: CODBIR, config: Partial<ACEDefinition>): ACEDefinition {
  const id = config.id || randomUUID();
  const name = config.name || 'NewACE';

  const ace: ACEDefinition = {
    id,
    name,
    description: config.description || '',
    type: config.type || 'Primary',
    iconProperty: config.iconProperty,
    cardComponents: config.cardComponents || [],
    quickViews: config.quickViews,
    properties: config.properties || []
  };

  // ACEs are stored as components with type 'ace'
  ir.components.push({
    type: 'ace',
    id,
    name,
    displayName: name,
    description: config.description || '',
    group: { id: '', name: '' },
    version: '1.0.0',
    entry: `src/ace/${name}/${name} ACE.ts`,
    framework: 'react',
    properties: config.properties || [],
    preconfiguredEntries: []
  });

  ir.metadata.modifiedAt = new Date().toISOString();
  return ace;
}

export function addList(ir: CODBIR, config: Partial<ListDefinition>): ListDefinition {
  const list: ListDefinition = {
    title: config.title || 'NewList',
    description: config.description,
    template: config.template ?? 100,
    hidden: config.hidden,
    contentTypes: config.contentTypes,
    fields: config.fields || [],
    rows: config.rows,
    folderCreation: config.folderCreation,
    versioning: config.versioning,
    majorVersionLimit: config.majorVersionLimit,
    minorVersionLimit: config.minorVersionLimit
  };

  ir.lists.push(list);
  ir.metadata.modifiedAt = new Date().toISOString();
  return list;
}

export function addLibrary(ir: CODBIR, config: Partial<LibraryDefinition>): LibraryDefinition {
  const library: LibraryDefinition = {
    title: config.title || 'NewLibrary',
    description: config.description,
    template: config.template,
    hidden: config.hidden,
    contentTypes: config.contentTypes,
    fields: config.fields,
    versioningEnabled: config.versioningEnabled,
    majorVersionLimit: config.majorVersionLimit,
    minorVersionLimit: config.minorVersionLimit,
    templateFeatureId: config.templateFeatureId
  };

  ir.libraries.push(library);
  ir.metadata.modifiedAt = new Date().toISOString();
  return library;
}

export function addColumn(ir: CODBIR, config: Partial<FieldDefinition>): FieldDefinition {
  const field: FieldDefinition = {
    name: config.name || 'NewField',
    displayName: config.displayName || config.name || 'NewField',
    type: config.type || 'Text',
    group: config.group,
    description: config.description,
    required: config.required,
    hidden: config.hidden,
    indexed: config.indexed,
    unique: config.unique,
    minLength: config.minLength,
    maxLength: config.maxLength,
    minValue: config.minValue,
    maxValue: config.maxValue,
    choices: config.choices,
    defaultValue: config.defaultValue,
    lookup: config.lookup,
    taxonomy: config.taxonomy,
    richText: config.richText,
    multiline: config.multiline,
    appendChanges: config.appendChanges,
    datetime: config.datetime
  };

  ir.fields.push(field);
  ir.metadata.modifiedAt = new Date().toISOString();
  return field;
}

export function addContentType(ir: CODBIR, config: Partial<ContentTypeDefinition>): ContentTypeDefinition {
  const ct: ContentTypeDefinition = {
    name: config.name || 'NewContentType',
    description: config.description,
    group: config.group,
    parentContentType: config.parentContentType,
    hidden: config.hidden,
    sealed: config.sealed,
    fields: config.fields || [],
    documentTemplate: config.documentTemplate
  };

  ir.contentTypes.push(ct);
  ir.metadata.modifiedAt = new Date().toISOString();
  return ct;
}

export function addGraphPermission(ir: CODBIR, scope: string, type: 'Delegated' | 'Application' = 'Delegated'): CODBIR {
  const exists = ir.graph.some(p => p.scope === scope && p.type === type);
  if (!exists) {
    const permInfo = (GRAPH_PERMISSIONS as Record<string, { requiresAdminApproval: boolean; description: string }>)[scope];
    ir.graph.push({
      resource: 'Microsoft Graph',
      scope,
      type,
      requiresAdminApproval: permInfo?.requiresAdminApproval ?? true,
      description: permInfo?.description || ''
    });
  }
  ir.metadata.modifiedAt = new Date().toISOString();
  return ir;
}

export function addPropertyPane(component: ComponentDefinition, pane: PropertyPaneDefinition): ComponentDefinition {
  component.properties.push(pane);
  return component;
}

export function addPermission(ir: CODBIR, permission: PermissionDefinition): CODBIR {
  ir.permissions.push(permission);
  ir.metadata.modifiedAt = new Date().toISOString();
  return ir;
}

export function addTheme(ir: CODBIR, theme: ThemeDefinition): CODBIR {
  ir.themes.push(theme);
  ir.metadata.modifiedAt = new Date().toISOString();
  return ir;
}

export function addFormatting(ir: CODBIR, formatting: FormattingDefinition): CODBIR {
  ir.formatting.push(formatting);
  ir.metadata.modifiedAt = new Date().toISOString();
  return ir;
}

export function addProvisioning(ir: CODBIR, provisioning: ProvisioningDefinition): CODBIR {
  ir.provisioning.push(provisioning);
  ir.metadata.modifiedAt = new Date().toISOString();
  return ir;
}

export function addPage(ir: CODBIR, page: PageDefinition): CODBIR {
  ir.pages.push(page);
  ir.metadata.modifiedAt = new Date().toISOString();
  return ir;
}

// Validate IR
export function validateIR(ir: CODBIR): string[] {
  const errors: string[] = [];

  if (ir.$schema !== 'codbsharepoint/ir/1.0') {
    errors.push('Invalid IR schema version');
  }

  if (!ir.solution.name) {
    errors.push('Solution name is required');
  }

  if (!ir.solution.id) {
    errors.push('Solution ID is required');
  }

  if (!ir.solution.version) {
    errors.push('Solution version is required');
  }

  const componentTypes = ir.components.map(c => c.type);
  const hasWebParts = componentTypes.includes('webpart');
  const hasExtensions = componentTypes.some(t => t !== 'webpart' && t !== 'ace');

  if (hasExtensions && !ir.extensions.length) {
    errors.push('Extensions declared in components but no extension definitions found');
  }

  // Check for duplicate component IDs
  const ids = ir.components.map(c => c.id);
  const uniqueIds = new Set(ids);
  if (ids.length !== uniqueIds.size) {
    errors.push('Duplicate component IDs detected');
  }

  return errors;
}

// Transform IR to different targets
export function getTargetIR(ir: CODBIR, target: BuildTarget): CODBIR {
  // Clone the IR
  const targetIR = JSON.parse(JSON.stringify(ir)) as CODBIR;

  switch (target) {
    case 'sharepoint.spfx.webpart':
      targetIR.components = targetIR.components.filter(c => c.type === 'webpart');
      break;
    case 'sharepoint.spfx.extension':
    case 'sharepoint.spfx.applicationCustomizer':
      targetIR.extensions = targetIR.extensions.filter(e => e.type === 'ApplicationCustomizer');
      break;
    case 'sharepoint.spfx.fieldCustomizer':
      targetIR.extensions = targetIR.extensions.filter(e => e.type === 'FieldCustomizer');
      break;
    case 'sharepoint.spfx.commandSet':
      targetIR.extensions = targetIR.extensions.filter(e => e.type === 'ListViewCommandSet');
      break;
    case 'sharepoint.spfx.formCustomizer':
      targetIR.extensions = targetIR.extensions.filter(e => e.type === 'FormCustomizer');
      break;
    case 'sharepoint.ace':
      targetIR.components = targetIR.components.filter(c => c.type === 'ace');
      break;
    case 'sharepoint.solution':
    case 'sharepoint.spfx.library':
    default:
      // Keep everything
      break;
  }

  return targetIR;
}

// Serialize IR to JSON
export function serializeIR(ir: CODBIR): string {
  return JSON.stringify(ir, null, 2);
}

// Deserialize IR from JSON
export function deserializeIR(json: string): CODBIR {
  const ir = JSON.parse(json) as CODBIR;
  const errors = validateIR(ir);
  if (errors.length > 0) {
    throw new Error(`Invalid IR: ${errors.join(', ')}`);
  }
  return ir;
}

// Helper to generate namespace from solution name
export function generateNamespace(name: string): string {
  return name
    .replace(/[^a-zA-Z0-9]/g, '')
    .split(/(?=[A-Z])/)
    .join('')
    .substring(0, 40);
}
