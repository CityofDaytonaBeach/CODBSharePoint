// ============================================================================
// CODBSharePoint - Main Entry Point
// Browser-native SharePoint compiler, validator, and packaging SDK
// ============================================================================

import type {
  CODBIR,
  BuildOptions,
  BuildResult,
  ValidationResult,
  SecurityReport,
  CompatibilityReport,
  BundleAnalysis,
  AnalysisResult,
  DeploymentManifest,
  CODBEvent,
  CODBEventHandler,
  CODBSharePointConfig,
  ImportSource,
  ImportResult,
  ExportOptions,
  SimulatorConfig,
  ToolAPI,
  SolutionConfig,
  ComponentDefinition,
  ExtensionDefinition,
  ACEDefinition,
  ListDefinition,
  LibraryDefinition,
  FieldDefinition,
  ContentTypeDefinition,
  PropertyPaneDefinition,
  SPFxVersion,
  Framework,
  GraphPermissionDefinition,
  VFSFile
} from './types/index.js';
import { SPFx_COMPATIBILITY } from './types/index.js';

// Core
import { createVFS } from './core/vfs.js';
import {
  createIR,
  addWebPart,
  addExtension,
  addACE,
  addList,
  addLibrary,
  addColumn,
  addContentType,
  addGraphPermission,
  addPropertyPane,
  addPermission,
  addTheme,
  addFormatting,
  addProvisioning,
  addPage,
  validateIR,
  serializeIR,
  deserializeIR,
  generateNamespace
} from './core/ir.js';

// Compiler
import { SPFxCompiler } from './compiler/spfx-compiler.js';

// Bundler
import { SPFxBundle } from './bundler/spfx-bundler.js';

// Manifest
import {
  generatePackageSolution,
  generateComponentManifest,
  generateExtensionManifest,
  generateFeatureXml,
  generateElementsXml,
  generateConfigJson,
  generateTsConfig,
  generatePackageJson
} from './manifest/generator.js';

// OPC
import { generateSPPKG, validateSPPKGStructure, validateSPPKGPackage } from './opc/sppkg.js';

// Provisioning artifacts
import {
  generateSharePointArtifacts,
  generateThemeJson,
  generateFormattingJson,
  generateSiteScript,
  generateProvisioningJson,
  generatePageJson
} from './provisioning/generator.js';

// Localization
import {
  generateResx,
  generateStringsModule,
  generateLocalizationFiles,
  resolveStrings
} from './localization/generator.js';

// Esbuild runtime (offline bundler control)
import {
  setWasmURL,
  isAvailable as isEsbuildAvailable,
  transformContent as esbuildTransform,
  bundleFromVFS as esbuildBundle
} from './bundler/esbuild-runtime.js';

// Serverless storage
import { createStorage, MemoryStorage, LocalStorageAdapter, IndexedDBStorage } from './storage/index.js';
import type { StorageAdapter, StorageValue, StorageKind } from './storage/index.js';

// Templates
import { TemplateRegistry } from './templates/index.js';
import type { ComponentTemplate, TemplateKind, TemplateRenderContext } from './templates/index.js';

// Designer
import { Designer } from './designer/index.js';
import { DesignerRulesEngine, BUILT_IN_RULES } from './designer/rules.js';
import type { DesignerRule, RuleField, RulePropertyPane, RuleDataSource } from './designer/rules.js';

// Static publish
import { generateStaticPublish } from './publish/index.js';

// Validator
import { SPFxValidator } from './validator/spfx-validator.js';

// Analyzer
import { SPFxAnalyzer } from './analyzer/spfx-analyzer.js';

// Security
import { SecurityScanner } from './security/scanner.js';

// Tools
import { ToolAPIService } from './tools/index.js';
import { AIService } from './ai/index.js';
import { findKnowledgeEntries, getKnowledgeCatalog, summarizeKnowledge } from './knowledge/index.js';

// Browser runtime
import { initBrowser, initBrowserCustom, downloadFile, downloadSPPKG, browserBuildProof } from './browser/index.js';

// Deploy
import { uploadSPPKG, DEFAULT_LIBRARY } from './deploy/index.js';

// Simulator
import { SharePointSimulator } from './simulator/index.js';

// Import/Export
import { SPFxImporter } from './import/spfx-import.js';
import { SPFxExporter } from './export/spfx-export.js';

// Utils
import { EventEmitter } from './utils/events.js';
import { randomUUID, generateSPFxId } from './utils/crypto.js';

// ---------------------------------------------------------------------------
// CODBSharePoint Class
// ---------------------------------------------------------------------------

export class CODBSharePoint {
  private config: CODBSharePointConfig;
  private emitter: EventEmitter;
  private compiler: SPFxCompiler;
  private bundler: SPFxBundle;
  private validator: SPFxValidator;
  private analyzer: SPFxAnalyzer;
  private securityScanner: SecurityScanner;
  private importer: SPFxImporter;
  private exporter: SPFxExporter;

  constructor(config: CODBSharePointConfig = {}) {
    this.config = {
      version: config.version || '1.0.0',
      spfxVersion: config.spfxVersion || '1.22.0',
      environment: config.environment || 'production',
      ...config
    };

    this.emitter = new EventEmitter();
    this.compiler = new SPFxCompiler({
      target: this.config.spfxVersion
    });
    this.bundler = new SPFxBundle();
    this.validator = new SPFxValidator();
    this.analyzer = new SPFxAnalyzer();
    this.securityScanner = new SecurityScanner();
    this.importer = new SPFxImporter();
    this.exporter = new SPFxExporter();
  }

  // ---------------------------------------------------------------------------
  // Main Build API
  // ---------------------------------------------------------------------------

  /**
   * Build a SharePoint solution from specification
   */
  async build(spec: {
    target?: string;
    type?: string;
    name?: string;
    framework?: Framework;
    solution?: Partial<SolutionConfig>;
    components?: Partial<ComponentDefinition>[];
    extensions?: Partial<ExtensionDefinition>[];
    lists?: Partial<ListDefinition>[];
    libraries?: Partial<LibraryDefinition>[];
    fields?: Partial<FieldDefinition>[];
    contentTypes?: Partial<ContentTypeDefinition>[];
    permissions?: string[];
    graph?: string[];
    files?: Map<string, string>;
    spfxVersion?: SPFxVersion;
    options?: Partial<BuildOptions>;
  }): Promise<BuildResult> {
    const startTime = Date.now();
    this.emit({ type: 'build:start', timestamp: startTime });

    try {
      // 1. Create IR from specification
      let ir = createIR(spec.solution);

      // 2. Add components
      if (spec.components) {
        for (const comp of spec.components) {
          addWebPart(ir, comp);
        }
      }

      if (spec.extensions) {
        for (const ext of spec.extensions) {
          addExtension(ir, ext);
        }
      }

      // 3. Add SharePoint resources
      if (spec.lists) {
        for (const list of spec.lists) {
          addList(ir, list);
        }
      }

      if (spec.libraries) {
        for (const lib of spec.libraries) {
          addLibrary(ir, lib);
        }
      }

      if (spec.fields) {
        for (const field of spec.fields) {
          addColumn(ir, field);
        }
      }

      if (spec.contentTypes) {
        for (const ct of spec.contentTypes) {
          addContentType(ir, ct);
        }
      }

      // 4. Add permissions
      if (spec.permissions) {
        for (const perm of spec.permissions) {
          addGraphPermission(ir, perm);
        }
      }

      if (spec.graph) {
        for (const scope of spec.graph) {
          addGraphPermission(ir, scope);
        }
      }

      // 5. Set SPFx version
      if (spec.spfxVersion) {
        ir.metadata.spfxVersion = spec.spfxVersion;
      }

      // 6. Build the solution
      return await this.buildFromIR(ir, spec.files, spec.options);
    } catch (error) {
      const duration = Date.now() - startTime;
      this.emit({ type: 'build:error', error: error as Error });

      return {
        success: false,
        files: [],
        deployment: {
          status: 'errors',
          artifact: '',
          destination: '',
          requiresAdmin: false,
          permissions: [],
          provisioning: [],
          warnings: [],
          instructions: [],
          metadata: {
            generator: 'codbsharepoint',
            version: __VERSION__ || '1.0.0',
            buildTime: new Date().toISOString(),
            spfxVersion: '1.22.0'
          }
        },
        validation: {
          valid: false,
          errors: [{ code: 'BUILD001', message: (error as Error).message, severity: 'error', category: 'build' }],
          warnings: [],
          info: [],
          summary: { total: 1, errors: 1, warnings: 0, info: 0, categories: { build: { errors: 1, warnings: 0 } } }
        },
        security: { passed: false, score: 0, findings: [], secrets: [], externalUrls: [], permissions: [], recommendations: [] },
        compatibility: { compatible: false, targetVersion: '1.22.0', issues: [], recommendations: [], deprecated: [] },
        bundle: { totalSize: 0, chunks: [], dependencies: [], duplicates: [], externals: [], recommendations: [] },
        errors: [{ code: 'BUILD001', message: (error as Error).message, severity: 'error', category: 'build', file: '' }],
        warnings: [],
        duration
      };
    }
  }

  /**
   * Build from an existing IR
   */
  async buildFromIR(
    ir: CODBIR,
    sourceFiles?: Map<string, string>,
    options?: Partial<BuildOptions>
  ): Promise<BuildResult> {
    const startTime = Date.now();
    const buildOptions: BuildOptions = {
      target: options?.target || 'sharepoint.solution',
      spfxVersion: options?.spfxVersion || ir.metadata.spfxVersion,
      includeSource: options?.includeSource ?? true,
      skipValidation: options?.skipValidation ?? false,
      minify: options?.minify ?? true,
      sourceMaps: options?.sourceMaps ?? false,
      ...options
    };

    const errors: any[] = [];
    const warnings: any[] = [];

    // Emit progress
    this.emit({ type: 'build:progress', stage: 'validation', progress: 0 });

    // 1. Validate IR
    let validation: ValidationResult;
    if (!buildOptions.skipValidation) {
      validation = await this.validator.validate(ir);
      if (!validation.valid) {
        return {
          success: false,
          files: [],
          deployment: this.exporter.generateDeploymentManifest(ir, {
            success: false,
            files: [],
            deployment: {
              status: 'errors',
              artifact: '',
              destination: '',
              requiresAdmin: false,
              permissions: [],
              provisioning: [],
              warnings: [],
              instructions: [],
              metadata: {
                generator: 'codbsharepoint',
                version: __VERSION__ || '1.0.0',
                buildTime: new Date().toISOString(),
                spfxVersion: ir.metadata.spfxVersion
              }
            },
            validation,
            security: {
              passed: false,
              score: 0,
              findings: [],
              secrets: [],
              externalUrls: [],
              permissions: [],
              recommendations: []
            },
            compatibility: {
              compatible: false,
              targetVersion: ir.metadata.spfxVersion,
              issues: [],
              recommendations: [],
              deprecated: []
            },
            bundle: {
              totalSize: 0,
              chunks: [],
              dependencies: [],
              duplicates: [],
              externals: [],
              recommendations: []
            },
            errors: validation.errors,
            warnings: validation.warnings,
            duration: Date.now() - startTime
          }),
          validation,
          security: {
            passed: false,
            score: 0,
            findings: [],
            secrets: [],
            externalUrls: [],
            permissions: [],
            recommendations: []
          },
          compatibility: {
            compatible: false,
            targetVersion: ir.metadata.spfxVersion,
            issues: [],
            recommendations: [],
            deprecated: []
          },
          bundle: {
            totalSize: 0,
            chunks: [],
            dependencies: [],
            duplicates: [],
            externals: [],
            recommendations: []
          },
          errors: validation.errors,
          warnings: validation.warnings,
          duration: Date.now() - startTime
        };
      }
    } else {
      validation = {
        valid: true,
        errors: [],
        warnings: [],
        info: [],
        summary: { total: 0, errors: 0, warnings: 0, info: 0, categories: {} }
      };
    }

    this.emit({ type: 'build:progress', stage: 'compilation', progress: 25 });

    // 2. Compile source files
    const compileResult = await this.compiler.compile(ir, sourceFiles);

    if (!compileResult.success) {
      errors.push(...compileResult.errors.map(error => ({
        code: 'TS001',
        message: error.message,
        severity: error.severity,
        category: 'compilation',
        file: error.file,
        line: error.line,
        column: error.column
      })));
    }

    this.emit({ type: 'build:progress', stage: 'bundling', progress: 50 });

    // 3. Bundle
    const bundleResult = await this.bundler.bundle(ir, compileResult.files, {
      minify: buildOptions.minify,
      sourceMaps: buildOptions.sourceMaps
    });

    if (!bundleResult.success) {
      errors.push(...bundleResult.errors.map(message => ({
        code: 'BUNDLE001',
        message,
        severity: 'error',
        category: 'bundling'
      })));
    }

    this.emit({ type: 'build:progress', stage: 'packaging', progress: 75 });

    // 4. Generate SPPKG
    let sppkg: Uint8Array | undefined;
    if (compileResult.success && bundleResult.success) {
      try {
        sppkg = generateSPPKG(ir, bundleResult.files);
        const packageErrors = validateSPPKGPackage(sppkg);
        if (packageErrors.length > 0) {
          errors.push(...packageErrors.map(message => ({
            code: 'SPPKG002',
            message,
            severity: 'error',
            category: 'packaging'
          })));
        }
      } catch (error) {
        errors.push({
          code: 'PKG001',
          message: `SPPKG generation failed: ${(error as Error).message}`,
          severity: 'error',
          category: 'packaging'
        });
      }
    }

    const packageStructureErrors = validateSPPKGStructure(ir, bundleResult.files);
    if (packageStructureErrors.length > 0) {
      errors.push(...packageStructureErrors.map(message => ({
        code: 'SPPKG001',
        message,
        severity: 'error',
        category: 'packaging'
      })));
    }

    this.emit({ type: 'build:progress', stage: 'security', progress: 85 });

    // 5. Security scan
    const security = await this.securityScanner.scan(ir, sourceFiles);

    // 6. Compatibility check
    const compatibility = await this.checkCompatibility(ir);

    // 7. Bundle analysis
    const bundleAnalysis = this.analyzeBundle(bundleResult);
    const success =
      errors.length === 0 &&
      sppkg !== undefined &&
      validation.valid &&
      compileResult.success &&
      bundleResult.success &&
      security.passed &&
      compatibility.compatible;

    // 8. Generate deployment manifest
    const buildResult: BuildResult = {
      success,
      sppkg,
      files: compileResult.files,
      deployment: this.exporter.generateDeploymentManifest(ir, {
        success,
        sppkg,
        files: compileResult.files,
        validation,
        security,
        compatibility,
        bundle: bundleAnalysis,
        errors,
        warnings: compileResult.warnings as any,
        duration: Date.now() - startTime
      }),
      validation,
      security,
      compatibility,
      bundle: bundleAnalysis,
      errors,
      warnings: compileResult.warnings as any,
      duration: Date.now() - startTime
    };

    this.emit({ type: 'build:progress', stage: 'complete', progress: 100 });
    this.emit({ type: 'build:complete', result: buildResult, duration: buildResult.duration });

    return buildResult;
  }

  // ---------------------------------------------------------------------------
  // Analyze API
  // ---------------------------------------------------------------------------

  /**
   * Analyze a project for framework, permissions, and complexity
   */
  async analyze(ir: CODBIR): Promise<AnalysisResult> {
    return this.analyzer.analyze(ir);
  }

  // ---------------------------------------------------------------------------
  // Validate API
  // ---------------------------------------------------------------------------

  /**
   * Validate a project
   */
  async validate(ir: CODBIR): Promise<ValidationResult> {
    return this.validator.validate(ir);
  }

  // ---------------------------------------------------------------------------
  // Import API
  // ---------------------------------------------------------------------------

  /**
   * Import an existing SPFx project or SPPKG
   */
  async import(data: File | Blob | ArrayBuffer | string): Promise<ImportResult> {
    return this.importer.import(data);
  }

  // ---------------------------------------------------------------------------
  // Export API
  // ---------------------------------------------------------------------------

  /**
   * Export a project with deployment artifacts
   */
  async export(ir: CODBIR, buildResult: BuildResult, options: ExportOptions = { format: 'all' }): Promise<Uint8Array> {
    return this.exporter.generateDeploymentZip(ir, buildResult);
  }

  // ---------------------------------------------------------------------------
  // Compatibility API
  // ---------------------------------------------------------------------------

  /**
   * Check compatibility with target SPFx version
   */
  async compatibility(ir: CODBIR): Promise<CompatibilityReport> {
    return this.checkCompatibility(ir);
  }

  capabilities(): Record<string, boolean> {
    return {
      browserCompiler: true,
      esbuildWasm: true,
      typescriptChecking: false,
      tsx: true,
      jsx: true,
      react: true,
      sass: true,
      cssModules: true,
      productionBundling: false,
      spfxExternals: false,
      spfx122: false,
      sppkg: false,
      webParts: true,
      applicationCustomizers: false,
      fieldCustomizers: false,
      commandSets: false,
      ace: false,
      graphPermissions: true,
      featureFramework: false,
      offline: false
    };
  }

  async browserProductionSmoke(): Promise<{
    success: boolean;
    diagnostics: string[];
    sppkgBytes: number;
    bundleCount: number;
  }> {
    const diagnostics: string[] = [];
    const result = await this.build({
      solution: {
        name: 'BrowserSmoke',
        version: '1.0.0',
        description: 'Browser-only production smoke test',
        author: 'CODBSharePoint'
      },
      components: [{
        name: 'BrowserSmokeWebPart',
        displayName: 'Browser Smoke Web Part',
        description: 'Validates browser-native compile, bundle, and package output',
        framework: 'react'
      }],
      options: {
        runtime: 'browser',
        minify: true,
        sourceMaps: false
      }
    });

    if (!result.success) {
      diagnostics.push(...result.errors.map(error => `${error.code}: ${error.message}`));
    }

    if (!result.sppkg) {
      diagnostics.push('No SPPKG bytes were produced.');
    } else {
      diagnostics.push(...validateSPPKGPackage(result.sppkg));
    }

    for (const file of result.files) {
      if (typeof file.content !== 'string') continue;
      if (/\b(process|Buffer|__dirname|__filename)\b/.test(file.content)) {
        diagnostics.push(`Generated file contains Node global reference: ${file.path}`);
      }
    }

    for (const chunk of result.bundle.chunks as Array<{ name: string; content?: string }>) {
      if (!chunk.content) continue;
      if (!chunk.content.includes('define(')) {
        diagnostics.push(`Bundle is not AMD-wrapped for SPFx loader: ${chunk.name}`);
      }
      if (/\b(process|Buffer|__dirname|__filename)\b/.test(chunk.content)) {
        diagnostics.push(`Bundle contains Node global reference: ${chunk.name}`);
      }
    }

    return {
      success: result.success && diagnostics.length === 0,
      diagnostics,
      sppkgBytes: result.sppkg?.length || 0,
      bundleCount: result.bundle.chunks.length
    };
  }

  // ---------------------------------------------------------------------------
  // Tool API (for AI agents)
  // ---------------------------------------------------------------------------

  /**
   * Get the Tool API for programmatic access
   */
  tools(config?: Partial<SolutionConfig>): ToolAPI {
    return new ToolAPIService(config);
  }

  ai(): AIService {
    return new AIService(this);
  }

  // ---------------------------------------------------------------------------
  // Simulator API
  // ---------------------------------------------------------------------------

  /**
   * Create a SharePoint simulator for previewing solutions
   */
  simulator(config?: SimulatorConfig): SharePointSimulator {
    return new SharePointSimulator(config);
  }

  designer(options: { storage?: StorageAdapter | StorageKind; templates?: ComponentTemplate[] } = {}): Designer {
    return new Designer(this, options);
  }

  get designerAPI() {
    return {
      DesignerRulesEngine,
      BUILT_IN_RULES
    };
  }

  get deployAPI() {
    return {
      upload: uploadSPPKG,
      DEFAULT_LIBRARY
    };
  }

  // ---------------------------------------------------------------------------
  // Lower-level APIs
  // ---------------------------------------------------------------------------

  get compilerAPI() { return this.compiler; }
  get bundlerAPI() { return this.bundler; }
  get storageAPI() { return { createStorage, MemoryStorage, LocalStorageAdapter, IndexedDBStorage }; }
  get templatesAPI() { return { TemplateRegistry }; }
  get publishAPI() { return { generateStaticPublish }; }
  get manifestAPI() {
    return {
      generatePackageSolution,
      generateComponentManifest,
      generateExtensionManifest,
      generateFeatureXml,
      generateElementsXml,
      generateConfigJson,
      generateTsConfig,
      generatePackageJson
    };
  }
  get opcAPI() { return { generateSPPKG, validateSPPKGStructure, validateSPPKGPackage }; }
  get validatorAPI() { return this.validator; }
  get analyzerAPI() { return this.analyzer; }
  get securityAPI() { return this.securityScanner; }
  get provisioningAPI() {
    return {
      generateSharePointArtifacts,
      generateThemeJson,
      generateFormattingJson,
      generateSiteScript,
      generateProvisioningJson,
      generatePageJson
    };
  }
  get localizationAPI() {
    return {
      generateResx,
      generateStringsModule,
      generateLocalizationFiles,
      resolveStrings
    };
  }

  get bundleAPI() {
    return {
      setWasmURL,
      isEsbuildAvailable,
      transform: esbuildTransform,
      bundle: esbuildBundle
    };
  }

  get knowledgeAPI() {
    return {
      catalog: getKnowledgeCatalog,
      find: findKnowledgeEntries,
      summary: summarizeKnowledge
    };
  }

  get browserAPI() {
    return {
      init: initBrowser,
      initCustom: initBrowserCustom,
      download: downloadFile,
      downloadSPPKG,
      buildProof: browserBuildProof
    };
  }

  // ---------------------------------------------------------------------------
  // Static Methods
  // ---------------------------------------------------------------------------

  static createIR = createIR;
  static addWebPart = addWebPart;
  static addExtension = addExtension;
  static addACE = addACE;
  static addList = addList;
  static addLibrary = addLibrary;
  static addColumn = addColumn;
  static addContentType = addContentType;
  static addGraphPermission = addGraphPermission;
  static addPropertyPane = addPropertyPane;
  static addPermission = addPermission;
  static addTheme = addTheme;
  static addFormatting = addFormatting;
  static addProvisioning = addProvisioning;
  static addPage = addPage;
  static validateIR = validateIR;
  static serializeIR = serializeIR;
  static deserializeIR = deserializeIR;
  static generateNamespace = generateNamespace;
  static generateSPFxId = generateSPFxId;
  static randomUUID = randomUUID;
  static createVFS = createVFS;

  static fromAI = async (input: {
    description?: string;
    files?: Map<string, string>;
    spec?: Record<string, unknown>;
  }): Promise<BuildResult> => {
    const sdk = new CODBSharePoint();
    const ir = createIR({ name: input.spec?.name as string || 'AIGenerated' });

    // Auto-detect from files
    if (input.files) {
      // Detect framework
      let framework: Framework = 'none';
      for (const [path, content] of input.files) {
        if (content.includes('React') || content.includes('useState')) {
          framework = 'react';
          break;
        }
      }

      // Add a web part
      addWebPart(ir, {
        name: input.spec?.name as string || 'AIWebPart',
        framework,
        displayName: input.description || 'AI Generated Web Part'
      });

      // Detect Graph usage
      const content = Array.from(input.files.values()).join('\n');
      if (content.includes('graph') || content.includes('Graph')) {
        addGraphPermission(ir, 'User.Read.All');
      }
    }

    return sdk.buildFromIR(ir, input.files);
  };

  // ---------------------------------------------------------------------------
  // Event System
  // ---------------------------------------------------------------------------

  on(handler: CODBEventHandler): () => void {
    return this.emitter.on(handler);
  }

  private emit(event: CODBEvent): void {
    this.emitter.emit(event);
  }

  // ---------------------------------------------------------------------------
  // Private Helpers
  // ---------------------------------------------------------------------------

  private async checkCompatibility(ir: CODBIR): Promise<CompatibilityReport> {
    const targetVersion = ir.metadata.spfxVersion;
    const compatibility = SPFx_COMPATIBILITY[targetVersion as keyof typeof SPFx_COMPATIBILITY];

    const issues: any[] = [];
    const recommendations: string[] = [];
    const deprecated: any[] = [];

    if (!compatibility) {
      issues.push({
        severity: 'warning',
        component: 'SPFx Version',
        message: `Unknown SPFx version: ${targetVersion}`
      });
    }

    for (const component of ir.components) {
      if (component.type !== 'webpart') {
        issues.push({
          severity: 'error',
          component: component.name,
          message: `Component type "${component.type}" is not production-proven for ${targetVersion}`
        });
      }

      if (component.framework !== 'react' && component.framework !== 'vanilla') {
        issues.push({
          severity: 'error',
          component: component.name,
          message: `Framework "${component.framework}" is not supported by the production build path`
        });
      }
    }

    for (const extension of ir.extensions) {
      issues.push({
        severity: 'error',
        component: extension.name,
        message: `Extension type "${extension.type}" is not production-proven yet`
      });
    }

    return {
      compatible: issues.filter(i => i.severity === 'error').length === 0,
      targetVersion,
      issues,
      recommendations,
      deprecated
    };
  }

  private analyzeBundle(bundleResult: any): BundleAnalysis {
    return {
      totalSize: bundleResult.totalSize || 0,
      chunks: bundleResult.chunks || [],
      dependencies: bundleResult.dependencies || [],
      duplicates: [],
      externals: bundleResult.externals || [],
      recommendations: []
    };
  }

  // ---------------------------------------------------------------------------
  // Version & Info
  // ---------------------------------------------------------------------------

  get version(): string {
    return __VERSION__ || '1.0.0';
  }

  getInfo(): Record<string, unknown> {
    return {
      name: 'CODBSharePoint',
      version: this.version,
      description: 'Browser-native SharePoint compiler, validator, and packaging SDK',
      capabilities: this.capabilities(),
      spfxVersions: Object.keys(SPFx_COMPATIBILITY),
      buildTargets: [
        'sharepoint.spfx.webpart',
        'sharepoint.spfx.extension',
        'sharepoint.spfx.applicationCustomizer',
        'sharepoint.spfx.fieldCustomizer',
        'sharepoint.spfx.commandSet',
        'sharepoint.spfx.ace',
        'sharepoint.solution'
      ]
    };
  }
}

// ---------------------------------------------------------------------------
// Export default and named exports
// ---------------------------------------------------------------------------

export default CODBSharePoint;

// Re-export all types
export * from './types/index.js';

// Re-export core functions
export {
  createIR,
  addWebPart,
  addExtension,
  addACE,
  addList,
  addLibrary,
  addColumn,
  addContentType,
  addGraphPermission,
  addPropertyPane,
  addPermission,
  addTheme,
  addFormatting,
  addProvisioning,
  addPage,
  validateIR,
  serializeIR,
  deserializeIR,
  generateNamespace
} from './core/ir.js';

export { createVFS } from './core/vfs.js';
export { randomUUID, generateSPFxId } from './utils/crypto.js';
export { SPFxCompiler } from './compiler/spfx-compiler.js';
export { SPFxBundle } from './bundler/spfx-bundler.js';
export { SPFxValidator } from './validator/spfx-validator.js';
export { SPFxAnalyzer } from './analyzer/spfx-analyzer.js';
export { SecurityScanner } from './security/scanner.js';
export { ToolAPIService } from './tools/index.js';
export { AIService } from './ai/index.js';
export type { AIDiagnostic, AIFix, AIPlan, AIProfile, AIProject, AIProjectInput, AISourceCompatibility } from './ai/index.js';
export { findKnowledgeEntries, getKnowledgeCatalog, summarizeKnowledge } from './knowledge/index.js';
export type { ApiEntry, DependencyEntry, KnowledgeEntry, KnowledgeStatus, PatternEntry, SchemaEntry, SharePointKnowledgeCatalog } from './knowledge/index.js';
export { SharePointSimulator } from './simulator/index.js';
export { SPFxImporter } from './import/spfx-import.js';
export { SPFxExporter } from './export/spfx-export.js';
export {
  generatePackageSolution,
  generateComponentManifest,
  generateExtensionManifest,
  generateFeatureXml,
  generateElementsXml,
  generateConfigJson,
  generateTsConfig,
  generatePackageJson,
} from './manifest/generator.js';

export {
  generateSharePointArtifacts,
  generateThemeJson,
  generateFormattingJson,
  generateSiteScript,
  generateProvisioningJson,
  generatePageJson
} from './provisioning/generator.js';
export {
  generateResx,
  generateStringsModule,
  generateLocalizationFiles,
  resolveStrings
} from './localization/generator.js';
export {
  setWasmURL,
  isAvailable as isEsbuildAvailable,
  transformContent as esbuildTransform,
  bundleFromVFS as esbuildBundle
} from './bundler/esbuild-runtime.js';
export { generateSPPKG, validateSPPKGStructure, validateSPPKGPackage } from './opc/sppkg.js';

// Serverless authoring support
export { Designer } from './designer/index.js';
export type { ProjectManifest, ProjectSettings, DesignerBuilder } from './designer/index.js';
export { DesignerRulesEngine, BUILT_IN_RULES } from './designer/rules.js';
export type { DesignerRule, RuleField, RulePropertyPane, RuleDataSource, RuleValidation, RuleFormatting, RuleDiff, RuleDiffChange } from './designer/rules.js';
export { TemplateRegistry } from './templates/index.js';
export type { ComponentTemplate, TemplateKind, TemplateRenderContext } from './templates/index.js';
export { createStorage, MemoryStorage, LocalStorageAdapter, IndexedDBStorage } from './storage/index.js';
export type { StorageAdapter, StorageValue, StorageKind } from './storage/index.js';
export { generateStaticPublish } from './publish/index.js';
export type { StaticPublishOptions, StaticPublishResult } from './publish/index.js';

// Browser runtime utilities
export {
  initBrowser,
  initBrowserCustom,
  downloadFile,
  downloadSPPKG,
  browserBuildProof
} from './browser/index.js';
export type { BrowserInitResult, BrowserBuildProof, DownloadOptions } from './browser/index.js';

// Deployment
export { uploadSPPKG, DEFAULT_LIBRARY } from './deploy/index.js';
export type { UploadSPPKGOptions, UploadSPPKGResult, AppCatalogType, SPHttpClientLike } from './deploy/index.js';
