// ============================================================================
// Designer (Authoring-First Project Service)
// Holds a full SharePoint project (IR + settings + assets), scaffolds via the
// template registry, persists through a StorageAdapter, builds, and publishes
// static artifacts. The primary entry point for an authoring UI or AI agent.
// ============================================================================

import type {
  CODBIR,
  ComponentDefinition,
  ExtensionDefinition,
  ACEDefinition,
  SPFxVersion,
  BuildResult,
  FieldDefinition,
  ListDefinition,
  LibraryDefinition,
  ContentTypeDefinition,
  ThemeDefinition,
  FormattingDefinition,
  ProvisioningDefinition,
  PageDefinition
} from '../types/index.js';
import { createIR } from '../core/ir.js';
import {
  addWebPart as irAddWebPart,
  addExtension as irAddExtension,
  addACE as irAddACE,
  addList as irAddList,
  addLibrary as irAddLibrary,
  addColumn as irAddColumn,
  addContentType as irAddContentType,
  addTheme as irAddTheme,
  addFormatting as irAddFormatting,
  addProvisioning as irAddProvisioning,
  addPage as irAddPage,
  addGraphPermission as irAddGraphPermission
} from '../core/ir.js';
import { createStorage, type StorageAdapter, type StorageKind } from '../storage/index.js';
import { TemplateRegistry, type ComponentTemplate } from '../templates/index.js';
import { generateStaticPublish, type StaticPublishOptions, type StaticPublishResult } from '../publish/index.js';
import { DesignerRulesEngine, type DesignerRule, BUILT_IN_RULES } from './rules.js';

export interface ProjectSettings {
  name?: string;
  version?: string;
  description?: string;
  spfxVersion?: SPFxVersion;
  framework?: 'react' | 'vanilla';
  storageKey?: string;
}

export interface ProjectManifest {
  $schema: string;
  settings: ProjectSettings;
  ir: CODBIR;
  assets: Record<string, string>;
}

export interface DesignerBuilder {
  buildFromIR(ir: CODBIR, sourceFiles?: Map<string, string>, options?: unknown): Promise<BuildResult>;
}

const PROJECT_SCHEMA = 'codbsharepoint/project/1.0';

export class Designer {
  private builder: DesignerBuilder;
  private storage: StorageAdapter;
  readonly templates: TemplateRegistry;
  readonly rules: DesignerRulesEngine;
  private manifest: ProjectManifest;
  private assets: Record<string, string>;

  constructor(builder: DesignerBuilder, options: { storage?: StorageAdapter | StorageKind; templates?: ComponentTemplate[]; rules?: DesignerRule[] } = {}) {
    this.builder = builder;
    this.storage = options.storage
      ? (typeof options.storage === 'string' ? createStorage(options.storage) : options.storage)
      : createStorage('memory');
    this.templates = new TemplateRegistry(options.templates);
    this.rules = new DesignerRulesEngine(options.rules || BUILT_IN_RULES);

    this.manifest = this.emptyManifest('NewSolution');
    this.assets = {};
  }

  // ---------------------------------------------------------------------------

  private emptyManifest(name: string): ProjectManifest {
    return {
      $schema: PROJECT_SCHEMA,
      settings: { name },
      ir: createIR({ name }),
      assets: {}
    };
  }

  getIR(): CODBIR {
    return this.manifest.ir;
  }

  getSettings(): ProjectSettings {
    return this.manifest.settings;
  }

  getManifest(): ProjectManifest {
    return this.manifest;
  }

  getAssets(): Record<string, string> {
    return this.assets;
  }

  setAsset(path: string, content: string): void {
    this.assets[path] = content;
  }

  // ---------------------------------------------------------------------------
  // Project lifecycle
  // ---------------------------------------------------------------------------

  async create(name: string, settings: ProjectSettings = {}): Promise<void> {
    this.manifest = this.emptyManifest(name);
    this.manifest.settings = {
      ...this.manifest.settings,
      ...settings,
      name
    };
    this.assets = {};
  }

  async save(): Promise<void> {
    const key = this.manifest.settings.storageKey || `codb-project:${this.manifest.settings.name}`;
    const payload: ProjectManifest = {
      $schema: PROJECT_SCHEMA,
      settings: this.manifest.settings,
      ir: this.manifest.ir,
      assets: this.assets
    };
    await this.storage.setItem(key, payload as unknown as Record<string, unknown>);
  }

  async load(name: string): Promise<boolean> {
    const key = `codb-project:${name}`;
    const value = await this.storage.getItem(key);
    if (value === undefined || typeof value !== 'object') return false;
    const manifest = value as unknown as ProjectManifest;
    if (!manifest || !manifest.ir || manifest.$schema !== PROJECT_SCHEMA) return false;
    this.manifest = manifest;
    this.assets = manifest.assets || {};
    return true;
  }

  async delete(name: string): Promise<void> {
    const key = `codb-project:${name}`;
    await this.storage.removeItem(key);
  }

  serialize(): string {
    return JSON.stringify(this.getManifest(), null, 2);
  }

  // ---------------------------------------------------------------------------
  // Authoring (template-driven)
  // ---------------------------------------------------------------------------

  async addWebPart(config: Partial<ComponentDefinition> & { name: string }): Promise<ComponentDefinition> {
    const component = irAddWebPart(this.manifest.ir, config);
    const template = this.templates.get('react-webpart');
    if (template) {
      const files = template.render({ component, namespace: this.manifest.ir.solution.namespace });
      for (const [path, content] of files) this.assets[path] = content;
    }
    return component;
  }

  async addVanillaWebPart(config: Partial<ComponentDefinition> & { name: string }): Promise<ComponentDefinition> {
    const component = irAddWebPart(this.manifest.ir, { ...config, framework: 'vanilla' });
    const template = this.templates.get('vanilla-webpart');
    if (template) {
      const files = template.render({ component, namespace: this.manifest.ir.solution.namespace });
      for (const [path, content] of files) this.assets[path] = content;
    }
    return component;
  }

  async addExtension(config: Partial<ExtensionDefinition> & { name: string; type: ExtensionDefinition['type'] }): Promise<ExtensionDefinition> {
    const extension = irAddExtension(this.manifest.ir, config);
    const template = this.templates.resolveForComponent(
      { type: 'extension' } as ComponentDefinition,
      extension
    );
    if (template) {
      const files = template.render({ component: {} as ComponentDefinition, extension, namespace: this.manifest.ir.solution.namespace });
      for (const [path, content] of files) this.assets[path] = content;
    }
    return extension;
  }

  async addACE(config: Partial<ACEDefinition> & { name: string }): Promise<ACEDefinition> {
    const ace = irAddACE(this.manifest.ir, config);
    const template = this.templates.get('ace');
    if (template) {
      const files = template.render({ component: {} as ComponentDefinition, ace, namespace: this.manifest.ir.solution.namespace });
      for (const [path, content] of files) this.assets[path] = content;
    }
    return ace;
  }

  addList(config: Partial<ListDefinition>): ListDefinition {
    return irAddList(this.manifest.ir, config);
  }

  addLibrary(config: Partial<LibraryDefinition>): LibraryDefinition {
    return irAddLibrary(this.manifest.ir, config);
  }

  addColumn(config: Partial<FieldDefinition>): FieldDefinition {
    return irAddColumn(this.manifest.ir, config);
  }

  addContentType(config: Partial<ContentTypeDefinition>): ContentTypeDefinition {
    return irAddContentType(this.manifest.ir, config);
  }

  addTheme(config: ThemeDefinition): void {
    irAddTheme(this.manifest.ir, config);
  }

  addFormatting(config: FormattingDefinition): void {
    irAddFormatting(this.manifest.ir, config);
  }

  addProvisioning(config: ProvisioningDefinition): void {
    irAddProvisioning(this.manifest.ir, config);
  }

  addPage(config: PageDefinition): void {
    irAddPage(this.manifest.ir, config);
  }

  addGraphPermission(scope: string, type: 'Delegated' | 'Application' = 'Delegated'): void {
    irAddGraphPermission(this.manifest.ir, scope, type);
  }

  // ---------------------------------------------------------------------------
  // Rule-based authoring
  // ---------------------------------------------------------------------------

  /**
   * Create a new project from a rule, applying all components, data sources,
   * fields, lists, permissions, and source files defined in the rule.
   */
  async createFromRule(ruleId: string, overrides: { name?: string; description?: string } = {}): Promise<DesignerRule | undefined> {
    const rule = this.rules.get(ruleId);
    if (!rule) return undefined;

    const { ir, files, permissions } = this.rules.createProject(rule, overrides);
    this.manifest.ir = ir;
    this.manifest.settings = {
      ...this.manifest.settings,
      name: overrides.name || rule.name,
      description: overrides.description || rule.description,
      spfxVersion: rule.spfxVersion as any
    };

    // Apply source files from rule
    this.assets = {};
    for (const [path, content] of files) {
      this.assets[path] = content;
    }

    return rule;
  }

  /**
   * Export the current project as a reusable rule.
   */
  exportAsRule(metadata: { name: string; description?: string; author?: string; tags?: string[] } = { name: this.manifest.settings.name || 'ExportedRule' }): DesignerRule {
    return this.rules.exportFromProject(this.manifest.ir, this.assets, metadata);
  }

  /**
   * Validate the current project against a specific rule.
   */
  validateAgainstRule(ruleId: string): { valid: boolean; errors: string[]; warnings: string[] } {
    const rule = this.rules.get(ruleId);
    if (!rule) return { valid: false, errors: [`Rule "${ruleId}" not found`], warnings: [] };
    return this.rules.validateRule(rule);
  }

  /**
   * Search built-in and registered rules.
   */
  searchRules(query: string): DesignerRule[] {
    return this.rules.search(query);
  }

  /**
   * List all available rules.
   */
  listRules(): DesignerRule[] {
    return this.rules.list();
  }

  setLocalization(config: { defaultLanguage: string; languages?: string[]; strings?: Record<string, Record<string, string>> }): void {
    this.manifest.ir.localization = {
      defaultLanguage: config.defaultLanguage,
      languages: config.languages || [],
      ...(config.strings ? { strings: config.strings } : {})
    };
  }

  // ---------------------------------------------------------------------------
  // Build & Publish
  // ---------------------------------------------------------------------------

  async build(): Promise<BuildResult> {
    const sourceFiles = new Map<string, string>();
    for (const [path, content] of Object.entries(this.assets)) {
      sourceFiles.set(path, content);
    }
    return this.builder.buildFromIR(this.getIR(), sourceFiles);
  }

  async publish(options: StaticPublishOptions = {}): Promise<StaticPublishResult> {
    const sourceFiles = new Map<string, string>();
    for (const [path, content] of Object.entries(this.assets)) {
      sourceFiles.set(path, content);
    }
    const buildResult = await this.builder.buildFromIR(this.getIR(), sourceFiles);
    return generateStaticPublish(this.getIR(), buildResult, options);
  }
}
