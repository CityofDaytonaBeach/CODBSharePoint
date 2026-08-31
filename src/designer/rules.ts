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
  addACE,
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

export interface RuleDiffChange {
  path: string;
  type: 'added' | 'removed' | 'modified';
  before?: string;
  after?: string;
}

export interface RuleDiff {
  ruleA: string;
  ruleB: string;
  identical: boolean;
  changes: RuleDiffChange[];
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
      } else if (comp.type === 'ace') {
        addACE(ir, {
          name: comp.name,
          description: comp.displayName || comp.description
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
  // Inheritance: Create a new rule that extends an existing one
  // ---------------------------------------------------------------------------

  inherit(baseRuleId: string, overrides: Partial<DesignerRule>): DesignerRule {
    const base = this.rules.get(baseRuleId);
    if (!base) throw new Error(`Base rule "${baseRuleId}" not found`);

    const inherited: DesignerRule = {
      ...base,
      ...overrides,
      id: overrides.id || `${base.id}-derived`,
      version: overrides.version || incrementVersion(base.version),
      basedOn: base.id,
      createdAt: new Date().toISOString()
    };

    // Deep merge arrays instead of overwriting
    if (overrides.components) {
      inherited.components = [...base.components, ...overrides.components];
    }
    if (overrides.dataSources) {
      inherited.dataSources = [...base.dataSources, ...overrides.dataSources];
    }
    if (overrides.fields) {
      inherited.fields = [...(base.fields || []), ...overrides.fields];
    }
    if (overrides.lists) {
      inherited.lists = [...(base.lists || []), ...overrides.lists];
    }
    if (overrides.graphPermissions) {
      const merged = new Set([...(base.graphPermissions || []), ...overrides.graphPermissions]);
      inherited.graphPermissions = Array.from(merged);
    }
    if (overrides.tags) {
      const merged = new Set([...base.tags, ...overrides.tags]);
      inherited.tags = Array.from(merged);
    }

    return inherited;
  }

  // ---------------------------------------------------------------------------
  // Diff: Compare two rules and return what changed
  // ---------------------------------------------------------------------------

  diff(ruleIdA: string, ruleIdB: string): RuleDiff {
    const a = this.rules.get(ruleIdA);
    const b = this.rules.get(ruleIdB);
    if (!a) throw new Error(`Rule "${ruleIdA}" not found`);
    if (!b) throw new Error(`Rule "${ruleIdB}" not found`);

    const changes: RuleDiffChange[] = [];

    // Metadata
    if (a.name !== b.name) changes.push({ path: 'name', type: 'modified', before: a.name, after: b.name });
    if (a.description !== b.description) changes.push({ path: 'description', type: 'modified', before: a.description, after: b.description });

    // Components
    const aCompNames = a.components.map(c => c.name);
    const bCompNames = b.components.map(c => c.name);
    for (const name of aCompNames) {
      if (!bCompNames.includes(name)) changes.push({ path: `components.${name}`, type: 'removed' });
    }
    for (const name of bCompNames) {
      if (!aCompNames.includes(name)) changes.push({ path: `components.${name}`, type: 'added' });
    }
    for (const name of aCompNames) {
      if (bCompNames.includes(name)) {
        const ac = a.components.find(c => c.name === name)!;
        const bc = b.components.find(c => c.name === name)!;
        if (JSON.stringify(ac.propertyPane) !== JSON.stringify(bc.propertyPane)) {
          changes.push({ path: `components.${name}.propertyPane`, type: 'modified', before: JSON.stringify(ac.propertyPane), after: JSON.stringify(bc.propertyPane) });
        }
      }
    }

    // Data Sources
    const aDsNames = a.dataSources.map(d => d.name);
    const bDsNames = b.dataSources.map(d => d.name);
    for (const name of aDsNames) {
      if (!bDsNames.includes(name)) changes.push({ path: `dataSources.${name}`, type: 'removed' });
    }
    for (const name of bDsNames) {
      if (!aDsNames.includes(name)) changes.push({ path: `dataSources.${name}`, type: 'added' });
    }

    // Permissions
    const aPerms = new Set(a.graphPermissions || []);
    const bPerms = new Set(b.graphPermissions || []);
    for (const p of aPerms) {
      if (!bPerms.has(p)) changes.push({ path: `graphPermissions.${p}`, type: 'removed' });
    }
    for (const p of bPerms) {
      if (!aPerms.has(p)) changes.push({ path: `graphPermissions.${p}`, type: 'added' });
    }

    // Fields
    const aFieldNames = (a.fields || []).map(f => f.name);
    const bFieldNames = (b.fields || []).map(f => f.name);
    for (const name of aFieldNames) {
      if (!bFieldNames.includes(name)) changes.push({ path: `fields.${name}`, type: 'removed' });
    }
    for (const name of bFieldNames) {
      if (!aFieldNames.includes(name)) changes.push({ path: `fields.${name}`, type: 'added' });
    }

    // Lists
    const aListNames = (a.lists || []).map(l => l.name);
    const bListNames = (b.lists || []).map(l => l.name);
    for (const name of aListNames) {
      if (!bListNames.includes(name)) changes.push({ path: `lists.${name}`, type: 'removed' });
    }
    for (const name of bListNames) {
      if (!aListNames.includes(name)) changes.push({ path: `lists.${name}`, type: 'added' });
    }

    return {
      ruleA: ruleIdA,
      ruleB: ruleIdB,
      identical: changes.length === 0,
      changes
    };
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
  },
  {
    id: 'ace-announcements',
    name: 'ACE Announcements Card',
    description: 'Viva Connections adaptive card showing latest announcements',
    version: '1.0.0',
    author: 'CODBSharePoint',
    tags: ['ace', 'viva', 'cards', 'announcements'],
    components: [{
      name: 'AnnouncementsACE',
      displayName: 'Announcements Card',
      type: 'ace',
      framework: 'react',
      description: 'Viva Connections card for announcements',
      propertyPane: [
        { name: 'listName', displayName: 'Announcements list', type: 'text', default: 'Announcements' },
        { name: 'maxItems', displayName: 'Max items', type: 'number', default: 3 }
      ]
    }],
    dataSources: [{
      type: 'sharepoint-list',
      name: 'announcements-list',
      listName: 'Announcements'
    }],
    fields: [
      { name: 'Title', displayName: 'Title', type: 'text', required: true },
      { name: 'Announcement', displayName: 'Announcement', type: 'note' },
      { name: 'Expires', displayName: 'Expires', type: 'datetime' }
    ],
    lists: [{
      name: 'Announcements',
      displayName: 'Announcements',
      template: 100,
      fields: ['Title', 'Announcement', 'Expires']
    }],
    sourceFiles: {}
  },
  {
    id: 'ace-task-card',
    name: 'ACE Task Card',
    description: 'Viva Connections card showing tasks from Graph Planner',
    version: '1.0.0',
    author: 'CODBSharePoint',
    tags: ['ace', 'viva', 'tasks', 'planner', 'graph'],
    components: [{
      name: 'TaskCardACE',
      displayName: 'My Tasks Card',
      type: 'ace',
      framework: 'react',
      description: 'Viva Connections card for my Planner tasks',
      propertyPane: [
        { name: 'planId', displayName: 'Plan ID', type: 'text', default: '' },
        { name: 'maxTasks', displayName: 'Max tasks', type: 'number', default: 5 }
      ]
    }],
    dataSources: [{
      type: 'graph',
      name: 'graph-planner-tasks',
      endpoint: '/planner/plans/{planId}/tasks',
      permissions: ['Tasks.Read.All']
    }],
    graphPermissions: ['Tasks.Read.All'],
    sourceFiles: {}
  },
  {
    id: 'command-set-approvals',
    name: 'Approvals Command Set',
    description: 'ListView Command Set to approve/reject from the command bar',
    version: '1.0.0',
    author: 'CODBSharePoint',
    tags: ['commandset', 'listview', 'approvals', 'actions'],
    components: [{
      name: 'ApprovalsCommandSet',
      displayName: 'Approvals Actions',
      type: 'extension',
      framework: 'react',
      description: 'Approve/reject list items from command bar',
      propertyPane: []
    }],
    dataSources: [{
      type: 'sharepoint-list',
      name: 'approvals-list',
      listName: 'Approvals'
    }],
    fields: [
      { name: 'Title', displayName: 'Title', type: 'text', required: true },
      { name: 'Status', displayName: 'Status', type: 'choice', choices: ['Pending', 'Approved', 'Rejected'], defaultValue: 'Pending' }
    ],
    lists: [{
      name: 'Approvals',
      displayName: 'Approvals',
      template: 100,
      fields: ['Title', 'Status']
    }],
    sourceFiles: {}
  },
  {
    id: 'field-customizer-status',
    name: 'Status Field Customizer',
    description: 'Field Customizer rendering status as colored badges',
    version: '1.0.0',
    author: 'CODBSharePoint',
    tags: ['fieldcustomizer', 'status', 'badges', 'formatting'],
    components: [{
      name: 'StatusFieldCustomizer',
      displayName: 'Status Badge',
      type: 'extension',
      framework: 'vanilla',
      description: 'Render status field as colored badges',
      propertyPane: []
    }],
    dataSources: [{
      type: 'sharepoint-list',
      name: 'target-list',
      listName: ''
    }],
    sourceFiles: {}
  },
  {
    id: 'news-feed',
    name: 'News Feed',
    description: 'Display recent news from a communications site page library',
    version: '1.0.0',
    author: 'CODBSharePoint',
    tags: ['news', 'pages', 'communications'],
    components: [{
      name: 'NewsFeed',
      displayName: 'News Feed',
      type: 'webpart',
      framework: 'react',
      description: 'Recent news articles from Site Pages',
      propertyPane: [
        { name: 'maxItems', displayName: 'Max articles', type: 'number', default: 6 },
        { name: 'showThumbnails', displayName: 'Show thumbnails', type: 'boolean', default: true },
        { name: 'siteUrl', displayName: 'Site URL (empty = current site)', type: 'text', default: '' }
      ]
    }],
    dataSources: [{
      type: 'sharepoint-list',
      name: 'site-pages',
      listName: 'Site Pages'
    }],
    sourceFiles: {}
  },
  {
    id: 'task-board',
    name: 'Task Board',
    description: 'Kanban task board from Graph Planner buckets',
    version: '1.0.0',
    author: 'CODBSharePoint',
    tags: ['tasks', 'kanban', 'planner', 'graph', 'board'],
    components: [{
      name: 'TaskBoard',
      displayName: 'Task Board',
      type: 'webpart',
      framework: 'react',
      description: 'Kanban board of Planner tasks by bucket',
      propertyPane: [
        { name: 'planId', displayName: 'Plan ID', type: 'text', default: '' },
        { name: 'showBuckets', displayName: 'Show bucket columns', type: 'boolean', default: true },
        { name: 'refreshInterval', displayName: 'Auto-refresh (minutes)', type: 'number', default: 0 }
      ]
    }],
    dataSources: [{
      type: 'graph',
      name: 'graph-planner',
      endpoint: '/planner/plans/{planId}',
      permissions: ['Tasks.Read.All']
    }, {
      type: 'graph',
      name: 'graph-planner-buckets',
      endpoint: '/planner/plans/{planId}/buckets',
      permissions: ['Tasks.Read.All']
    }],
    graphPermissions: ['Tasks.Read.All'],
    sourceFiles: {}
  },
  {
    id: 'mailbox-viewer',
    name: 'Mailbox Viewer',
    description: 'Read and display recent emails from the current user',
    version: '1.0.0',
    author: 'CODBSharePoint',
    tags: ['mail', 'outlook', 'graph', 'messages'],
    components: [{
      name: 'MailboxViewer',
      displayName: 'Mailbox Viewer',
      type: 'webpart',
      framework: 'react',
      description: 'Recent emails from Outlook inbox',
      propertyPane: [
        { name: 'folder', displayName: 'Folder', type: 'dropdown', options: ['inbox', 'sentItems', 'drafts'], default: 'inbox' },
        { name: 'maxEmails', displayName: 'Max emails', type: 'number', default: 10 }
      ]
    }],
    dataSources: [{
      type: 'graph',
      name: 'graph-mails',
      endpoint: '/me/mailFolders/inbox/messages',
      permissions: ['Mail.Read']
    }],
    graphPermissions: ['Mail.Read'],
    sourceFiles: {}
  },
  {
    id: 'file-manager',
    name: 'File Manager',
    description: 'Browse and manage OneDrive/SharePoint files from Graph',
    version: '1.0.0',
    author: 'CODBSharePoint',
    tags: ['files', 'onedrive', 'documents', 'graph', 'manager'],
    components: [{
      name: 'FileManager',
      displayName: 'File Manager',
      type: 'webpart',
      framework: 'react',
      description: 'Browse OneDrive or SharePoint files',
      propertyPane: [
        { name: 'driveType', displayName: 'Drive', type: 'dropdown', options: ['me', 'site'], default: 'me' },
        { name: 'showPreview', displayName: 'Show preview panel', type: 'boolean', default: true }
      ]
    }],
    dataSources: [{
      type: 'graph',
      name: 'graph-drive',
      endpoint: '/me/drive/root/children',
      permissions: ['Files.Read.All']
    }],
    graphPermissions: ['Files.Read.All'],
    sourceFiles: {}
  },
  {
    id: 'event-registration',
    name: 'Event Registration',
    description: 'Event list with registration form and attendee management',
    version: '1.0.0',
    author: 'CODBSharePoint',
    tags: ['events', 'registration', 'forms', 'list'],
    components: [{
      name: 'EventRegistration',
      displayName: 'Event Registration',
      type: 'webpart',
      framework: 'react',
      description: 'Register for events and manage attendees',
      propertyPane: [
        { name: 'eventList', displayName: 'Events list', type: 'text', default: 'Events' },
        { name: 'regList', displayName: 'Registrations list', type: 'text', default: 'Registrations' }
      ]
    }],
    dataSources: [{
      type: 'sharepoint-list',
      name: 'events-list',
      listName: 'Events'
    }, {
      type: 'sharepoint-list',
      name: 'registrations-list',
      listName: 'Registrations'
    }],
    fields: [
      { name: 'Title', displayName: 'Title', type: 'text', required: true },
      { name: 'EventDate', displayName: 'Event Date', type: 'datetime', required: true },
      { name: 'Capacity', displayName: 'Capacity', type: 'number', defaultValue: 50 },
      { name: 'EventLink', displayName: 'Event (lookup)', type: 'lookup', lookupList: 'Events' },
      { name: 'AttendeeName', displayName: 'Attendee Name', type: 'person', required: true },
      { name: 'Registered', displayName: 'Registered At', type: 'datetime' }
    ],
    lists: [{
      name: 'Events',
      displayName: 'Events',
      template: 100,
      fields: ['Title', 'EventDate', 'Capacity']
    }, {
      name: 'Registrations',
      displayName: 'Registrations',
      template: 100,
      fields: ['Title', 'EventLink', 'AttendeeName', 'Registered']
    }],
    sourceFiles: {}
  },
  {
    id: 'ticketing-system',
    name: 'Ticketing System',
    description: 'Support ticket list with status workflow and assignment',
    version: '1.0.0',
    author: 'CODBSharePoint',
    tags: ['tickets', 'support', 'workflow', 'status', 'list'],
    components: [{
      name: 'TicketingSystem',
      displayName: 'Support Tickets',
      type: 'webpart',
      framework: 'react',
      description: 'Create, track, and manage support tickets',
      propertyPane: [
        { name: 'listName', displayName: 'Tickets list', type: 'text', default: 'Tickets' },
        { name: 'showAssignedToMe', displayName: 'Default filter to assigned to me', type: 'boolean', default: true }
      ]
    }],
    dataSources: [{
      type: 'sharepoint-list',
      name: 'tickets-list',
      listName: 'Tickets'
    }],
    fields: [
      { name: 'Title', displayName: 'Title', type: 'text', required: true },
      { name: 'Description', displayName: 'Description', type: 'note', required: true },
      { name: 'Priority', displayName: 'Priority', type: 'choice', choices: ['Low', 'Medium', 'High', 'Critical'], defaultValue: 'Medium' },
      { name: 'Status', displayName: 'Status', type: 'choice', choices: ['New', 'In Progress', 'Waiting on Customer', 'Resolved', 'Closed'], defaultValue: 'New' },
      { name: 'AssignedTo', displayName: 'Assigned To', type: 'person' },
      { name: 'SubmittedBy', displayName: 'Submitted By', type: 'person', required: true },
      { name: 'CreatedDate', displayName: 'Created Date', type: 'datetime', required: true },
      { name: 'ResolvedDate', displayName: 'Resolved Date', type: 'datetime' }
    ],
    lists: [{
      name: 'Tickets',
      displayName: 'Tickets',
      template: 100,
      fields: ['Title', 'Description', 'Priority', 'Status', 'AssignedTo', 'SubmittedBy', 'CreatedDate', 'ResolvedDate']
    }],
    sourceFiles: {}
  },
  {
    id: 'leave-request',
    name: 'Leave Request',
    description: 'Employee leave request form with Power Automate approval',
    version: '1.0.0',
    author: 'CODBSharePoint',
    tags: ['leave', 'hr', 'requests', 'approval', 'powerautomate'],
    components: [{
      name: 'LeaveRequest',
      displayName: 'Leave Request',
      type: 'webpart',
      framework: 'react',
      description: 'Submit and track leave requests',
      propertyPane: [
        { name: 'listName', displayName: 'Requests list', type: 'text', default: 'LeaveRequests' },
        { name: 'flowUrl', displayName: 'Power Automate flow URL', type: 'text', default: '' }
      ]
    }],
    dataSources: [{
      type: 'sharepoint-list',
      name: 'leave-list',
      listName: 'LeaveRequests'
    }],
    fields: [
      { name: 'Title', displayName: 'Title', type: 'text', required: true },
      { name: 'LeaveType', displayName: 'Leave Type', type: 'choice', choices: ['Annual', 'Sick', 'Personal', 'Maternity', 'Paternity', 'Bereavement'], required: true },
      { name: 'StartDate', displayName: 'Start Date', type: 'datetime', required: true },
      { name: 'EndDate', displayName: 'End Date', type: 'datetime', required: true },
      { name: 'Days', displayName: 'Days', type: 'number', required: true },
      { name: 'Reason', displayName: 'Reason', type: 'note' },
      { name: 'Status', displayName: 'Status', type: 'choice', choices: ['Pending', 'Approved', 'Rejected', 'Cancelled'], defaultValue: 'Pending' },
      { name: 'Approver', displayName: 'Approver', type: 'person' }
    ],
    lists: [{
      name: 'LeaveRequests',
      displayName: 'Leave Requests',
      template: 100,
      fields: ['Title', 'LeaveType', 'StartDate', 'EndDate', 'Days', 'Reason', 'Status', 'Approver']
    }],
    sourceFiles: {}
  },
  {
    id: 'image-gallery',
    name: 'Image Gallery',
    description: 'Responsive image gallery from a picture library',
    version: '1.0.0',
    author: 'CODBSharePoint',
    tags: ['gallery', 'images', 'library', 'photos'],
    components: [{
      name: 'ImageGallery',
      displayName: 'Image Gallery',
      type: 'webpart',
      framework: 'react',
      description: 'Responsive image gallery with lightbox',
      propertyPane: [
        { name: 'libraryName', displayName: 'Picture library', type: 'text', default: 'Pictures' },
        { name: 'columns', displayName: 'Columns', type: 'number', default: 4 },
        { name: 'showCaptions', displayName: 'Show captions', type: 'boolean', default: true }
      ]
    }],
    dataSources: [{
      type: 'sharepoint-library',
      name: 'pictures-library',
      listName: 'Pictures'
    }],
    sourceFiles: {}
  },
  {
    id: 'inventory-tracker',
    name: 'Inventory Tracker',
    description: 'Track inventory items with quantities and reorder alerts',
    version: '1.0.0',
    author: 'CODBSharePoint',
    tags: ['inventory', 'stock', 'tracking', 'list'],
    components: [{
      name: 'InventoryTracker',
      displayName: 'Inventory',
      type: 'webpart',
      framework: 'react',
      description: 'Track stock levels and reorder alerts',
      propertyPane: [
        { name: 'listName', displayName: 'Inventory list', type: 'text', default: 'Inventory' },
        { name: 'reorderThreshold', displayName: 'Reorder threshold', type: 'number', default: 10 }
      ]
    }],
    dataSources: [{
      type: 'sharepoint-list',
      name: 'inventory-list',
      listName: 'Inventory'
    }],
    fields: [
      { name: 'Title', displayName: 'Item Name', type: 'text', required: true },
      { name: 'SKU', displayName: 'SKU', type: 'text' },
      { name: 'Quantity', displayName: 'Quantity', type: 'number', required: true, defaultValue: 0 },
      { name: 'MinQuantity', displayName: 'Minimum Quantity', type: 'number', defaultValue: 10 },
      { name: 'UnitCost', displayName: 'Unit Cost', type: 'number' },
      { name: 'Location', displayName: 'Storage Location', type: 'text' },
      { name: 'Category', displayName: 'Category', type: 'choice', choices: ['Supplies', 'Equipment', 'Raw Material', 'Finished Goods'] }
    ],
    lists: [{
      name: 'Inventory',
      displayName: 'Inventory',
      template: 100,
      fields: ['Title', 'SKU', 'Quantity', 'MinQuantity', 'UnitCost', 'Location', 'Category']
    }],
    sourceFiles: {}
  }
];

export { BUILT_IN_RULES };

function incrementVersion(version: string): string {
  const parts = version.split('.');
  const patch = parseInt(parts[2] || '0') + 1;
  return `${parts[0]}.${parts[1]}.${patch}`;
}
