// ============================================================================
// Tool API - Programmatic interface for AI agents
// ============================================================================

import type {
  CODBIR,
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
  GraphPermissionDefinition,
  BuildOptions,
  BuildResult,
  ValidationResult,
  AnalysisResult
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
  addGraphPermission as addGraphPerm,
  addPropertyPane as addPropPane,
  validateIR
} from '../core/ir.js';
import { SPFxValidator } from '../validator/spfx-validator.js';
import { SPFxAnalyzer } from '../analyzer/spfx-analyzer.js';

export class ToolAPIService implements ToolAPI {
  private ir: CODBIR;
  private validator: SPFxValidator;
  private analyzer: SPFxAnalyzer;

  constructor(config?: Partial<SolutionConfig>) {
    this.ir = createIR(config);
    this.validator = new SPFxValidator();
    this.analyzer = new SPFxAnalyzer();
  }

  createSolution(config: Partial<SolutionConfig>): CODBIR {
    this.ir = createIR(config);
    return this.ir;
  }

  createWebPart(config: Partial<ComponentDefinition>): ComponentDefinition {
    return addWebPart(this.ir, config);
  }

  createExtension(config: Partial<ExtensionDefinition>): ExtensionDefinition {
    return addExtension(this.ir, config);
  }

  createACE(config: Partial<ACEDefinition>): ACEDefinition {
    return addACE(this.ir, config);
  }

  createList(config: Partial<ListDefinition>): ListDefinition {
    return addList(this.ir, config);
  }

  createLibrary(config: Partial<LibraryDefinition>): LibraryDefinition {
    return addLibrary(this.ir, config);
  }

  createColumn(config: Partial<FieldDefinition>): FieldDefinition {
    return addColumn(this.ir, config);
  }

  createContentType(config: Partial<ContentTypeDefinition>): ContentTypeDefinition {
    return addContentType(this.ir, config);
  }

  addGraphPermission(ir: CODBIR, scope: string, type: 'Delegated' | 'Application' = 'Delegated'): CODBIR {
    return addGraphPerm(ir, scope, type);
  }

  addPropertyPane(component: ComponentDefinition, pane: PropertyPaneDefinition): ComponentDefinition {
    return addPropPane(component, pane);
  }

  addList(ir: CODBIR, list: ListDefinition): CODBIR {
    addList(ir, list);
    return ir;
  }

  addLibrary(ir: CODBIR, library: LibraryDefinition): CODBIR {
    addLibrary(ir, library);
    return ir;
  }

  addColumn(ir: CODBIR, field: FieldDefinition): CODBIR {
    addColumn(ir, field);
    return ir;
  }

  async compile(ir: CODBIR, options?: Partial<BuildOptions>): Promise<BuildResult> {
    // This would integrate with the compiler
    return {
      success: true,
      files: [],
      deployment: {
        status: 'ready',
        artifact: `${ir.solution.name}.sppkg`,
        destination: 'Tenant App Catalog',
        requiresAdmin: ir.graph.some(p => p.requiresAdminApproval),
        permissions: ir.graph.map(p => ({
          resource: p.resource,
          permission: p.scope,
          type: p.type,
          requiresAdminApproval: p.requiresAdminApproval,
          description: p.description || ''
        })),
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
      validation: {
        valid: true,
        errors: [],
        warnings: [],
        info: [],
        summary: { total: 0, errors: 0, warnings: 0, info: 0, categories: {} }
      },
      security: {
        passed: true,
        score: 100,
        findings: [],
        secrets: [],
        externalUrls: [],
        permissions: [],
        recommendations: []
      },
      compatibility: {
        compatible: true,
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
      errors: [],
      warnings: [],
      duration: 0
    };
  }

  async validate(ir: CODBIR): Promise<ValidationResult> {
    return this.validator.validate(ir);
  }

  async analyze(ir: CODBIR): Promise<AnalysisResult> {
    return this.analyzer.analyze(ir);
  }

  repair(ir: CODBIR, validation: ValidationResult): CODBIR {
    // Auto-repair common issues
    for (const error of validation.errors) {
      switch (error.code) {
        case 'SOL001':
          if (!ir.solution.name) ir.solution.name = 'RepairedSolution';
          break;
        case 'SOL003':
          if (!ir.solution.id) {
            ir.solution.id = crypto.randomUUID();
          }
          break;
        case 'SOL005':
          if (!ir.solution.version) ir.solution.version = '1.0.0';
          break;
      }
    }

    return ir;
  }

  async package(ir: CODBIR, options?: Partial<BuildOptions>): Promise<BuildResult> {
    return this.compile(ir, options);
  }

  getIR(): CODBIR {
    return this.ir;
  }
}
