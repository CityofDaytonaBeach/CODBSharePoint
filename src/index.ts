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
import { generateSPPKG, validateSPPKGStructure } from './opc/sppkg.js';

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

    this.emit({ type: 'build:progress', stage: 'bundling', progress: 50 });

    // 3. Bundle
    const bundleResult = await this.bundler.bundle(ir, compileResult.files, {
      minify: buildOptions.minify,
      sourceMaps: buildOptions.sourceMaps
    });

    this.emit({ type: 'build:progress', stage: 'packaging', progress: 75 });

    // 4. Generate SPPKG
    let sppkg: Uint8Array | undefined;
    try {
      sppkg = generateSPPKG(ir, bundleResult.files);
    } catch (error) {
      errors.push({
        code: 'PKG001',
        message: `SPPKG generation failed: ${(error as Error).message}`,
        severity: 'error',
        category: 'packaging'
      });
    }

    this.emit({ type: 'build:progress', stage: 'security', progress: 85 });

    // 5. Security scan
    const security = await this.securityScanner.scan(ir, sourceFiles);

    // 6. Compatibility check
    const compatibility = await this.checkCompatibility(ir);

    // 7. Bundle analysis
    const bundleAnalysis = this.analyzeBundle(bundleResult);

    // 8. Generate deployment manifest
    const buildResult: BuildResult = {
      success: sppkg !== undefined && validation.valid,
      sppkg,
      files: compileResult.files,
      deployment: this.exporter.generateDeploymentManifest(ir, {
        success: sppkg !== undefined,
        sppkg,
        files: compileResult.files,
        validation,
        security,
        compatibility,
        bundle: bundleAnalysis,
        errors: compileResult.errors as any,
        warnings: compileResult.warnings as any,
        duration: Date.now() - startTime
      }),
      validation,
      security,
      compatibility,
      bundle: bundleAnalysis,
      errors: compileResult.errors as any,
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

  // ---------------------------------------------------------------------------
  // Tool API (for AI agents)
  // ---------------------------------------------------------------------------

  /**
   * Get the Tool API for programmatic access
   */
  tools(config?: Partial<SolutionConfig>): ToolAPI {
    return new ToolAPIService(config);
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
  get opcAPI() { return { generateSPPKG, validateSPPKGStructure }; }
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
      features: [
        'SPFx Web Part compilation',
        'Extension support (Application Customizer, Field Customizer, Command Set)',
        'Adaptive Card Extensions',
        'SPPKG generation',
        'Source project generation',
        'Graph permission analysis',
        'Security scanning',
        'Compatibility checking',
        'Bundle analysis',
        'SharePoint simulator',
        'Import/Export',
        'Tool API for AI agents'
      ],
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
export { SharePointSimulator } from './simulator/index.js';
export { SPFxImporter } from './import/spfx-import.js';
export { SPFxExporter } from './export/spfx-export.js';
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

// Serverless authoring support
export { Designer } from './designer/index.js';
export type { ProjectManifest, ProjectSettings, DesignerBuilder } from './designer/index.js';
export { TemplateRegistry } from './templates/index.js';
export type { ComponentTemplate, TemplateKind, TemplateRenderContext } from './templates/index.js';
export { createStorage, MemoryStorage, LocalStorageAdapter, IndexedDBStorage } from './storage/index.js';
export type { StorageAdapter, StorageValue, StorageKind } from './storage/index.js';
export { generateStaticPublish } from './publish/index.js';
export type { StaticPublishOptions, StaticPublishResult } from './publish/index.js';
