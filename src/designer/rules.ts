// ============================================================================
// Designer Rules Engine
// Extracts patterns from real apps, stores them as reusable rules, and
// creates new projects from rules. Integrates with the knowledge catalog
// to auto-infer permissions, fields, and dependencies.
// ============================================================================

import type {
  CODBIR,
  ComponentDefinition,
  ExtensionDefinition,
  FieldDefinition,
  ListDefinition,
  LibraryDefinition,
  ContentTypeDefinition,
  GraphPermissionDefinition
} from '../types/index.js';
import {
  createIR,
  addWebPart,
  addExtension,
  addList,
  addLibrary,
  addColumn,
  addContentType,
  addGraphPermission
} from '../core/ir.js';
import { getKnowledgeCatalog, type PatternEntry } from '../knowledge/index.js';

// ---------------------------------------------------------------------------
// Rule Schema
// ---------------------------------------------------------------------------

export type RuleDataSourceType = 'graph' | 'sharepoint-list' | 'sharepoint-library' | 'static' | 'external-api';

export interface RuleField {
  name: string;
  displayName: string;
  type: 'text' | 'note' | 'choice' | 'number' | 'datetime' | 'boolean' | 'lookup' | 'taxonomy' | 'url' | 'person';
  required?: boolean;
  choices?: string[];
  defaultValue?: string | number | boolean;
  lookupList?: string;
  description?: string;
}

export interface RulePropertyPane {
  name: string;
  displayName: string;
  type: 'text' | 'number' | 'boolean' | 'dropdown' | 'checkbox' | 'slider';
  default?: unknown;
  description?: string;
  options?: string[];
  min?: number;
  max?: number;
}

export interface RuleDataSource {
  type: RuleDataSourceType;
  name: string;
  permissions?: string[];
  listName?: string;
  endpoint?: string;
  fields?: string[];
}

export interface RuleValidation {
  field: string;
  rule: 'required' | 'minLength' | 'maxLength' | 'pattern' | 'min' | 'max' | 'unique';
  value?: string | number;
  message: string;
}

export interface RuleFormatting {
  target: 'column' | 'view' | 'form';
  name: string;
  json: Record<string, unknown>;
}

export interface DesignerRule {
  id: string;
  name: string;
  description: string;
  version: string;
  author: string;
  tags: string[];

  // Component structure
  components: Array<{
    name: string;
    displayName: string;
    type: 'webpart' | 'extension' | 'ace';
    framework: 'react' | 'vanilla';
    description?: string;
    propertyPane?: RulePropertyPane[];
  }>;

  // Data layer
  dataSources: RuleDataSource[];
  fields?: RuleField[];
  lists?: Array<{
    name: string;
    displayName: string;
    template: number;
    fields: string[];
    contentTypes?: string[];
  }>;
  libraries?: Array<{
    name: string;
    displayName: string;
    fields?: string[];
  }>;
  contentTypes?: Array<{
    name: string;
    displayName: string;
    fields: string[];
    parentContentType?: string;
  }>;

  // Validation
  validations?: RuleValidation[];
  formatting?: RuleFormatting[];

  // Graph permissions (auto-inferred from dataSources)
  graphPermissions?: string[];

  // Source files (template code)
  sourceFiles?: Record<string, string>;

  // Metadata
  spfxVersion?: string;
  localization?: {
    defaultLanguage: string;
    languages?: string[];
  };

  // Provenance
  createdAt?: string;
  updatedAt?: string;
  basedOn?: string;
}

// ---------------------------------------------------------------------------
// Rules Engine
// ---------------------------------------------------------------------------

export class DesignerRulesEngine {
  private rules = new Map<string, DesignerRule>();

  constructor(builtInRules: DesignerRule[] = BUILT_IN_RULES) {
    for (const rule of builtInRules) {
      this.rules.set(rule.id, rule);
    }
  }

  // ---------------------------------------------------------------------------
  // CRUD
  // ---------------------------------------------------------------------------

  register(rule: DesignerRule): void {
    this.rules.set(rule.id, rule);
  }

  get(id: string): DesignerRule | undefined {
    return this.rules.get(id);
  }

  remove(id: string): boolean {
    return this.rules.delete(id);
  }

  list(): DesignerRule[] {
    return Array.from(this.rules.values());
  }

  search(query: string): DesignerRule[] {
    const q = query.toLowerCase();
    return this.list().filter(rule =>
      rule.name.toLowerCase().includes(q) ||
      rule.description.toLowerCase().includes(q) ||
      rule.tags.some(t => t.toLowerCase().includes(q))
    );
  }

  // ---------------------------------------------------------------------------
  // Export: Extract rule from IR + assets
  // ---------------------------------------------------------------------------

  exportFromProject(
    ir: CODBIR,
    assets: Record<string, string>,
    metadata: { name: string; description?: string; author?: string; tags?: string[] } = { name: 'ExportedRule' }
  ): DesignerRule {
    const components: DesignerRule['components'] = ir.components.map(c => ({
      name: c.name,
      displayName: c.displayName,
      type: c.type as 'webpart',
      framework: c.framework as 'react' | 'vanilla',
      description: c.description,
      propertyPane: this.extractPropertyPane(assets, c)
    }));

    const dataSources: RuleDataSource[] = this.inferDataSources(ir, assets);
    const fields = this.inferFields(ir, assets);
    const lists = ir.lists?.map(l => ({
      name: l.title,
      displayName: l.title,
      template: l.template || 100,
      fields: l.fields || []
    }));
    const libraries = ir.libraries?.map(l => ({
      name: l.title,
      displayName: l.title,
      fields: l.fields || []
    }));
    const graphPermissions = ir.graph.map(p => p.scope);
    const sourceFiles = this.extractSourceFiles(assets);

    const rule: DesignerRule = {
      id: `exported-${metadata.name.toLowerCase().replace(/[^a-z0-9]/g, '-')}`,
      name: metadata.name,
      description: metadata.description || `Exported from ${ir.solution.name}`,
      version: '1.0.0',
      author: metadata.author || 'CODBSharePoint',
      tags: metadata.tags || ['exported'],
      components,
      dataSources,
      fields: fields.length > 0 ? fields : undefined,
      lists: lists && lists.length > 0 ? lists : undefined,
      libraries: libraries && libraries.length > 0 ? libraries : undefined,
      graphPermissions: graphPermissions.length > 0 ? graphPermissions : undefined,
      sourceFiles: Object.keys(sourceFiles).length > 0 ? sourceFiles : undefined,
      spfxVersion: ir.metadata.spfxVersion,
      createdAt: new Date().toISOString()
    };

    return rule;
  }

  // ---------------------------------------------------------------------------
  // Create: Generate project from rule
  // ---------------------------------------------------------------------------

  createProject(
    rule: DesignerRule,
    overrides: { name?: string; description?: string } = {}
  ): { ir: CODBIR; files: Map<string, string>; permissions: string[] } {
    const solutionName = overrides.name || rule.name;
    const ir = createIR({
      name: solutionName,
      description: overrides.description || rule.description,
      version: rule.version
    });

    // Set SPFx version
    if (rule.spfxVersion) {
      ir.metadata.spfxVersion = rule.spfxVersion as any;
    }

    // Add components
    for (const comp of rule.components) {
      if (comp.type === 'webpart') {
        addWebPart(ir, {
          name: comp.name,
          displayName: comp.displayName,
          description: comp.description,
          framework: comp.framework
        });
      } else {
        // Extensions use addExtension with type
        addExtension(ir, {
          name: comp.name,
          displayName: comp.displayName,
          description: comp.description,
          type: comp.type as any
        });
      }
    }

    // Add lists
    if (rule.lists) {
      for (const list of rule.lists) {
        addList(ir, {
          title: list.displayName || list.name,
          template: list.template,
          fields: list.fields
        });
      }
    }

    // Add libraries
    if (rule.libraries) {
      for (const lib of rule.libraries) {
        addLibrary(ir, {
          title: lib.displayName || lib.name,
          fields: lib.fields
        });
      }
    }

    // Add content types
    if (rule.contentTypes) {
      for (const ct of rule.contentTypes) {
        addContentType(ir, {
          name: ct.name,
          description: ct.displayName,
          fields: ct.fields
        });
      }
    }

    // Add Graph permissions
    const permissions = rule.graphPermissions || this.inferPermissions(rule);
    for (const scope of permissions) {
      addGraphPermission(ir, scope);
    }

    // Build source files map
    const files = new Map<string, string>();
    if (rule.sourceFiles) {
      for (const [path, content] of Object.entries(rule.sourceFiles)) {
        files.set(path, content);
      }
    }

    return { ir, files, permissions };
  }

  // ---------------------------------------------------------------------------
  // Inference Helpers
  // ---------------------------------------------------------------------------

  inferPermissions(rule: DesignerRule): string[] {
    const perms = new Set<string>();
    for (const ds of rule.dataSources) {
      if (ds.permissions) {
        ds.permissions.forEach(p => perms.add(p));
      }
    }
    // Cross-reference with knowledge catalog
    const catalog = getKnowledgeCatalog();
    for (const ds of rule.dataSources) {
      if (ds.type === 'graph') {
        const matching = catalog.graph.filter(g => ds.endpoint && g.endpoint.includes(ds.endpoint));
        for (const entry of matching) {
          entry.permissions?.forEach(p => perms.add(p));
        }
      }
    }
    return Array.from(perms).sort();
  }

  validateRule(rule: DesignerRule): { valid: boolean; errors: string[]; warnings: string[] } {
    const errors: string[] = [];
    const warnings: string[] = [];

    if (!rule.id) errors.push('Rule must have an id');
    if (!rule.name) errors.push('Rule must have a name');
    if (rule.components.length === 0) errors.push('Rule must have at least one component');

    for (const comp of rule.components) {
      if (!comp.name) errors.push('Component must have a name');
      if (!comp.displayName) warnings.push(`Component "${comp.name}" has no displayName`);
    }

    for (const ds of rule.dataSources) {
      if (ds.type === 'graph' && !ds.endpoint) warnings.push(`DataSource "${ds.name}" is graph type but has no endpoint`);
      if (ds.type === 'sharepoint-list' && !ds.listName) warnings.push(`DataSource "${ds.name}" is list type but has no listName`);
    }

    // Knowledge-aware validation
    const catalog = getKnowledgeCatalog();
    for (const ds of rule.dataSources) {
      if (ds.type === 'graph' && ds.endpoint) {
        const supported = catalog.graph.some(g => ds.endpoint!.startsWith(g.endpoint.replace(/\{.*?\}/g, '')));
        if (!supported) warnings.push(`DataSource "${ds.name}" endpoint "${ds.endpoint}" is not in the knowledge catalog`);
      }
    }

    return { valid: errors.length === 0, errors, warnings };
  }

  // ---------------------------------------------------------------------------
  // Private Helpers
  // ---------------------------------------------------------------------------

  private extractPropertyPane(assets: Record<string, string>, component: ComponentDefinition): RulePropertyPane[] {
    const props: RulePropertyPane[] = [];
    const webpartPath = Object.keys(assets).find(p => p.includes(component.name) && p.endsWith('WebPart.ts'));
    if (!webpartPath) return props;

    const content = assets[webpartPath];
    // Extract PropertyPaneTextField calls
    const textFieldPattern = /PropertyPaneTextField\(['"](\w+)['"]\s*,\s*\{\s*label:\s*['"]([^'"]+)['"]/g;
    let match: RegExpExecArray | null;
    while ((match = textFieldPattern.exec(content)) !== null) {
      props.push({ name: match[1], displayName: match[2], type: 'text' });
    }

    const numberFieldPattern = /PropertyPaneSlider\(['"](\w+)['"]\s*,\s*\{\s*label:\s*['"]([^'"]+)['"]\s*,\s*min:\s*(\d+)\s*,\s*max:\s*(\d+)/g;
    while ((match = numberFieldPattern.exec(content)) !== null) {
      props.push({ name: match[1], displayName: match[2], type: 'slider', min: parseInt(match[3]), max: parseInt(match[4]) });
    }

    const togglePattern = /PropertyPaneToggle\(['"](\w+)['"]\s*,\s*\{\s*label:\s*['"]([^'"]+)['"]/g;
    while ((match = togglePattern.exec(content)) !== null) {
      props.push({ name: match[1], displayName: match[2], type: 'boolean' });
    }

    return props;
  }

  private inferDataSources(ir: CODBIR, assets: Record<string, string>): RuleDataSource[] {
    const sources: RuleDataSource[] = [];
    const allContent = Object.values(assets).join('\n');

    // Detect Graph usage
    if (allContent.includes('MSGraphClient') || allContent.includes('msGraphClientFactory')) {
      const graphEndpoints = new Set<string>();
      const endpointPattern = /\.api\(\s*['"`]([^'"`]+)['"`]/g;
      let match: RegExpExecArray | null;
      while ((match = endpointPattern.exec(allContent)) !== null) {
        graphEndpoints.add(match[1]);
      }

      for (const endpoint of graphEndpoints) {
        sources.push({
          type: 'graph',
          name: `graph-${endpoint.replace(/\//g, '-').replace(/^-/, '')}`,
          endpoint,
          permissions: this.inferGraphPermissions(endpoint)
        });
      }
    }

    // Detect SharePoint REST usage
    if (allContent.includes('SPHttpClient') || allContent.includes('_api/')) {
      sources.push({
        type: 'sharepoint-list',
        name: 'sharepoint-rest',
        endpoint: '/_api/web'
      });
    }

    // Detect list/library usage from IR
    for (const list of ir.lists || []) {
      sources.push({
        type: 'sharepoint-list',
        name: list.title,
        listName: list.title
      });
    }
    for (const lib of ir.libraries || []) {
      sources.push({
        type: 'sharepoint-library',
        name: lib.title,
        listName: lib.title
      });
    }

    return sources;
  }

  private inferFields(ir: CODBIR, assets: Record<string, string>): RuleField[] {
    const fields: RuleField[] = [];
    // IR lists only store field names as strings, not full FieldDefinition objects
    // Fields come from the rule's own fields array, not from IR
    return fields;
  }

  private inferGraphPermissions(endpoint: string): string[] {
    const perms: string[] = [];
    if (endpoint.includes('/users')) perms.push('User.Read.All');
    if (endpoint.includes('/groups')) perms.push('Group.Read.All');
    if (endpoint.includes('/sites')) perms.push('Sites.Read.All');
    if (endpoint.includes('/drives') || endpoint.includes('/drive')) perms.push('Files.Read.All');
    if (endpoint.includes('/me')) perms.push('User.Read');
    return perms;
  }

  private extractSourceFiles(assets: Record<string, string>): Record<string, string> {
    const sources: Record<string, string> = {};
    for (const [path, content] of Object.entries(assets)) {
      if (path.endsWith('.ts') || path.endsWith('.tsx') || path.endsWith('.js') || path.endsWith('.jsx') || path.endsWith('.scss')) {
        sources[path] = content;
      }
    }
    return sources;
  }
}

// ---------------------------------------------------------------------------
// Built-in Rules
// ---------------------------------------------------------------------------

const BUILT_IN_RULES: DesignerRule[] = [
  {
    id: 'employee-directory',
    name: 'Employee Directory',
    description: 'Search and display employee profiles from Microsoft Graph',
    version: '1.0.0',
    author: 'CODBSharePoint',
    tags: ['graph', 'users', 'search', 'people'],
    components: [{
      name: 'EmployeeDirectory',
      displayName: 'Employee Directory',
      type: 'webpart',
      framework: 'react',
      description: 'Search and browse employee profiles',
      propertyPane: [
        { name: 'pageSize', displayName: 'Results per page', type: 'number', default: 10 },
        { name: 'showPhoto', displayName: 'Show profile photo', type: 'boolean', default: true },
        { name: 'defaultSearch', displayName: 'Default search term', type: 'text', default: '' }
      ]
    }],
    dataSources: [{
      type: 'graph',
      name: 'graph-users',
      endpoint: '/users',
      permissions: ['User.Read.All']
    }],
    graphPermissions: ['User.Read.All'],
    sourceFiles: {}
  },
  {
    id: 'quick-links',
    name: 'Quick Links',
    description: 'Display a grid of configurable navigation links',
    version: '1.0.0',
    author: 'CODBSharePoint',
    tags: ['navigation', 'links', 'static'],
    components: [{
      name: 'QuickLinks',
      displayName: 'Quick Links',
      type: 'webpart',
      framework: 'react',
      description: 'Configurable grid of navigation links',
      propertyPane: [
        { name: 'linksPerRow', displayName: 'Links per row', type: 'number', default: 4 },
        { name: 'openInNewTab', displayName: 'Open in new tab', type: 'boolean', default: true }
      ]
    }],
    dataSources: [{ type: 'static', name: 'static-links' }],
    sourceFiles: {}
  },
  {
    id: 'data-table',
    name: 'Data Table',
    description: 'Display SharePoint list data in a sortable, filterable table',
    version: '1.0.0',
    author: 'CODBSharePoint',
    tags: ['list', 'table', 'data'],
    components: [{
      name: 'DataTable',
      displayName: 'Data Table',
      type: 'webpart',
      framework: 'react',
      description: 'Sortable, filterable SharePoint list viewer',
      propertyPane: [
        { name: 'listName', displayName: 'List name', type: 'text', default: '' },
        { name: 'pageSize', displayName: 'Page size', type: 'number', default: 25 },
        { name: 'showFilters', displayName: 'Show filters', type: 'boolean', default: true }
      ]
    }],
    dataSources: [{ type: 'sharepoint-list', name: 'sp-list', listName: '' }],
    sourceFiles: {}
  },
  {
    id: 'faq-accordion',
    name: 'FAQ Accordion',
    description: 'Expandable FAQ sections from a SharePoint list or static content',
    version: '1.0.0',
    author: 'CODBSharePoint',
    tags: ['faq', 'accordion', 'content'],
    components: [{
      name: 'FaqAccordion',
      displayName: 'FAQ Accordion',
      type: 'webpart',
      framework: 'react',
      description: 'Expandable FAQ sections',
      propertyPane: [
        { name: 'openMultiple', displayName: 'Allow multiple open', type: 'boolean', default: false },
        { name: 'showSearch', displayName: 'Show search box', type: 'boolean', default: true }
      ]
    }],
    dataSources: [{ type: 'static', name: 'static-faq' }],
    sourceFiles: {}
  },
  {
    id: 'calendar-view',
    name: 'Calendar View',
    description: 'Display events from SharePoint list or Graph calendar in a month/week view',
    version: '1.0.0',
    author: 'CODBSharePoint',
    tags: ['calendar', 'events', 'graph'],
    components: [{
      name: 'CalendarView',
      displayName: 'Calendar View',
      type: 'webpart',
      framework: 'react',
      description: 'Month/week calendar view of events',
      propertyPane: [
        { name: 'defaultView', displayName: 'Default view', type: 'dropdown', options: ['month', 'week', 'day'], default: 'month' },
        { name: 'listName', displayName: 'Events list name', type: 'text', default: 'Events' }
      ]
    }],
    dataSources: [{
      type: 'sharepoint-list',
      name: 'events-list',
      listName: 'Events'
    }],
    fields: [
      { name: 'Title', displayName: 'Title', type: 'text', required: true },
      { name: 'EventDate', displayName: 'Start Date', type: 'datetime', required: true },
      { name: 'EndDate', displayName: 'End Date', type: 'datetime', required: true },
      { name: 'Location', displayName: 'Location', type: 'text' },
      { name: 'Description', displayName: 'Description', type: 'note' }
    ],
    lists: [{
      name: 'Events',
      displayName: 'Events',
      template: 100,
      fields: ['Title', 'EventDate', 'EndDate', 'Location', 'Description']
    }],
    sourceFiles: {}
  },
  {
    id: 'org-chart',
    name: 'Organization Chart',
    description: 'Hierarchical org chart from Microsoft Graph user relationships',
    version: '1.0.0',
    author: 'CODBSharePoint',
    tags: ['graph', 'org', 'hierarchy', 'users'],
    components: [{
      name: 'OrgChart',
      displayName: 'Organization Chart',
      type: 'webpart',
      framework: 'react',
      description: 'Hierarchical org chart from Graph',
      propertyPane: [
        { name: 'rootUser', displayName: 'Root user email', type: 'text', default: '' },
        { name: 'depth', displayName: 'Depth levels', type: 'number', default: 3 }
      ]
    }],
    dataSources: [{
      type: 'graph',
      name: 'graph-users-manager',
      endpoint: '/users',
      permissions: ['User.Read.All']
    }],
    graphPermissions: ['User.Read.All'],
    sourceFiles: {}
  },
  {
    id: 'document-explorer',
    name: 'Document Explorer',
    description: 'Browse and preview documents from SharePoint libraries',
    version: '1.0.0',
    author: 'CODBSharePoint',
    tags: ['documents', 'library', 'files'],
    components: [{
      name: 'DocumentExplorer',
      displayName: 'Document Explorer',
      type: 'webpart',
      framework: 'react',
      description: 'Browse and preview library documents',
      propertyPane: [
        { name: 'libraryName', displayName: 'Library name', type: 'text', default: 'Documents' },
        { name: 'showThumbnails', displayName: 'Show thumbnails', type: 'boolean', default: true },
        { name: 'itemsPerPage', displayName: 'Items per page', type: 'number', default: 20 }
      ]
    }],
    dataSources: [{
      type: 'sharepoint-library',
      name: 'documents-library',
      listName: 'Documents'
    }],
    sourceFiles: {}
  },
  {
    id: 'approval-dashboard',
    name: 'Approval Dashboard',
    description: 'Track and manage approval workflows from a SharePoint list',
    version: '1.0.0',
    author: 'CODBSharePoint',
    tags: ['approvals', 'workflow', 'list'],
    components: [{
      name: 'ApprovalDashboard',
      displayName: 'Approval Dashboard',
      type: 'webpart',
      framework: 'react',
      description: 'Track approval requests and status',
      propertyPane: [
        { name: 'listName', displayName: 'Approvals list', type: 'text', default: 'Approvals' },
        { name: 'showMyItems', displayName: 'Show my items only', type: 'boolean', default: true }
      ]
    }],
    dataSources: [{
      type: 'sharepoint-list',
      name: 'approvals-list',
      listName: 'Approvals'
    }],
    fields: [
      { name: 'Title', displayName: 'Title', type: 'text', required: true },
      { name: 'Status', displayName: 'Status', type: 'choice', choices: ['Pending', 'Approved', 'Rejected'], defaultValue: 'Pending' },
      { name: 'Requester', displayName: 'Requester', type: 'person', required: true },
      { name: 'Approver', displayName: 'Approver', type: 'person' },
      { name: 'RequestDate', displayName: 'Request Date', type: 'datetime', required: true },
      { name: 'DecisionDate', displayName: 'Decision Date', type: 'datetime' },
      { name: 'Comments', displayName: 'Comments', type: 'note' }
    ],
    lists: [{
      name: 'Approvals',
      displayName: 'Approvals',
      template: 100,
      fields: ['Title', 'Status', 'Requester', 'Approver', 'RequestDate', 'DecisionDate', 'Comments']
    }],
    sourceFiles: {}
  }
];

export { BUILT_IN_RULES };
