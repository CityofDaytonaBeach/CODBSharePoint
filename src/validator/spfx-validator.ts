// ============================================================================
// Validator - Comprehensive SPFx project validation
// ============================================================================

import type {
  CODBIR,
  ValidationResult,
  ValidationError,
  ValidationWarning,
  ValidationInfo,
  ValidationSummary,
  SPFxVersion,
  ComponentDefinition
} from '../types/index.js';
import { SPFx_COMPATIBILITY, GRAPH_PERMISSIONS } from '../types/index.js';

export class SPFxValidator {

  async validate(ir: CODBIR, bundleFiles?: Map<string, string | Uint8Array>): Promise<ValidationResult> {
    const errors: ValidationError[] = [];
    const warnings: ValidationWarning[] = [];
    const info: ValidationInfo[] = [];

    // Run all validation checks
    this.validateSolution(ir, errors, warnings, info);
    this.validateComponents(ir, errors, warnings, info);
    this.validateExtensions(ir, errors, warnings, info);
    this.validatePermissions(ir, errors, warnings, info);
    this.validateManifests(ir, errors, warnings, info);
    this.validateCompatibility(ir, errors, warnings, info);
    this.validateSecurity(ir, warnings, info);
    this.validateStructure(ir, errors, warnings, info);

    if (bundleFiles) {
      this.validateBundle(ir, bundleFiles, errors, warnings, info);
    }

    const summary = this.generateSummary(errors, warnings, info);

    return {
      valid: errors.length === 0,
      errors,
      warnings,
      info,
      summary
    };
  }

  private validateSolution(ir: CODBIR, errors: ValidationError[], warnings: ValidationWarning[], info: ValidationInfo[]): void {
    const solution = ir.solution;

    // Required fields
    if (!solution.name) {
      errors.push({
        code: 'SOL001',
        message: 'Solution name is required',
        severity: 'error',
        category: 'solution'
      });
    } else if (!/^[a-zA-Z][a-zA-Z0-9]*$/.test(solution.name)) {
      errors.push({
        code: 'SOL002',
        message: 'Solution name must start with a letter and contain only alphanumeric characters',
        severity: 'error',
        category: 'solution'
      });
    }

    if (!solution.id) {
      errors.push({
        code: 'SOL003',
        message: 'Solution ID is required',
        severity: 'error',
        category: 'solution'
      });
    } else if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(solution.id)) {
      errors.push({
        code: 'SOL004',
        message: 'Solution ID must be a valid GUID',
        severity: 'error',
        category: 'solution'
      });
    }

    if (!solution.version) {
      errors.push({
        code: 'SOL005',
        message: 'Solution version is required',
        severity: 'error',
        category: 'solution'
      });
    } else if (!/^\d+\.\d+\.\d+$/.test(solution.version)) {
      warnings.push({
        code: 'SOL006',
        message: 'Version should follow semantic versioning (x.y.z)',
        category: 'solution'
      });
    }

    if (!solution.author) {
      warnings.push({
        code: 'SOL007',
        message: 'Author is recommended for distribution',
        category: 'solution'
      });
    }

    if (!solution.description) {
      info.push({
        code: 'SOL008',
        message: 'Adding a description improves discoverability',
        category: 'solution'
      });
    }
  }

  private validateComponents(ir: CODBIR, errors: ValidationError[], warnings: ValidationWarning[], info: ValidationInfo[]): void {
    if (ir.components.length === 0 && ir.extensions.length === 0) {
      warnings.push({
        code: 'CMP001',
        message: 'No components defined in the solution',
        category: 'components'
      });
    }

    const ids = new Set<string>();

    for (const component of ir.components) {
      // ID validation
      if (!component.id) {
        errors.push({
          code: 'CMP002',
          message: `Component "${component.name}" is missing an ID`,
          severity: 'error',
          category: 'components'
        });
      } else if (ids.has(component.id)) {
        errors.push({
          code: 'CMP003',
          message: `Duplicate component ID: ${component.id}`,
          severity: 'error',
          category: 'components'
        });
      } else {
        ids.add(component.id);
      }

      // Name validation
      if (!component.name) {
        errors.push({
          code: 'CMP004',
          message: 'Component name is required',
          severity: 'error',
          category: 'components'
        });
      } else if (!/^[a-zA-Z][a-zA-Z0-9]*$/.test(component.name)) {
        errors.push({
          code: 'CMP005',
          message: `Component name "${component.name}" must start with a letter and contain only alphanumeric characters`,
          severity: 'error',
          category: 'components'
        });
      }

      // Entry point
      if (!component.entry) {
        warnings.push({
          code: 'CMP006',
          message: `Component "${component.name}" is missing an entry point`,
          category: 'components'
        });
      }

      // Version
      if (!component.version) {
        warnings.push({
          code: 'CMP007',
          message: `Component "${component.name}" is missing a version`,
          category: 'components'
        });
      }

      // Group
      if (!component.group?.id) {
        info.push({
          code: 'CMP008',
          message: `Component "${component.name}" will appear in the default group`,
          category: 'components'
        });
      }

      // Preconfigured entries
      if (component.preconfiguredEntries.length === 0) {
        info.push({
          code: 'CMP009',
          message: `Component "${component.name}" has no preconfigured entries`,
          category: 'components'
        });
      }
    }
  }

  private validateExtensions(ir: CODBIR, errors: ValidationError[], warnings: ValidationWarning[], info: ValidationInfo[]): void {
    for (const ext of ir.extensions) {
      // Extension type
      if (!ext.type) {
        errors.push({
          code: 'EXT001',
          message: `Extension "${ext.name}" is missing a type`,
          severity: 'error',
          category: 'extensions'
        });
      }

      // Client side component ID
      if (!ext.clientSideComponentId) {
        errors.push({
          code: 'EXT002',
          message: `Extension "${ext.name}" is missing a client side component ID`,
          severity: 'error',
          category: 'extensions'
        });
      }

      // Entry
      if (!ext.entry) {
        warnings.push({
          code: 'EXT003',
          message: `Extension "${ext.name}" is missing an entry point`,
          category: 'extensions'
        });
      }
    }
  }

  private validatePermissions(ir: CODBIR, errors: ValidationError[], warnings: ValidationWarning[], info: ValidationInfo[]): void {
    // Check for declared permissions vs used permissions
    for (const perm of ir.graph) {
      const knownPerm = GRAPH_PERMISSIONS[perm.scope as keyof typeof GRAPH_PERMISSIONS];

      if (!knownPerm) {
        warnings.push({
          code: 'PERM001',
          message: `Unknown Graph permission: ${perm.scope}`,
          category: 'permissions'
        });
      } else if (knownPerm.requiresAdminApproval && !perm.requiresAdminApproval) {
        warnings.push({
          code: 'PERM002',
          message: `Permission "${perm.scope}" requires admin approval`,
          category: 'permissions'
        });
      }
    }

    // Check for excessive permissions
    const adminPermissions = ir.graph.filter(p => {
      const known = GRAPH_PERMISSIONS[p.scope as keyof typeof GRAPH_PERMISSIONS];
      return known?.requiresAdminApproval;
    });

    if (adminPermissions.length > 3) {
      warnings.push({
        code: 'PERM003',
        message: `Solution requests ${adminPermissions.length} admin-approved permissions. Consider reducing.`,
        category: 'permissions'
      });
    }

    // Info about permissions
    for (const perm of ir.graph) {
      info.push({
        code: 'PERM004',
        message: `Graph permission declared: ${perm.scope} (${perm.type})`,
        category: 'permissions'
      });
    }
  }

  private validateManifests(ir: CODBIR, errors: ValidationError[], warnings: ValidationWarning[], info: ValidationInfo[]): void {
    // Check manifest IDs match component IDs
    const componentIds = new Set(ir.components.map(c => c.id));

    for (const component of ir.components) {
      if (!component.displayName) {
        warnings.push({
          code: 'MAN001',
          message: `Component "${component.name}" has no display name`,
          category: 'manifests'
        });
      }

      if (!component.description) {
        info.push({
          code: 'MAN002',
          message: `Component "${component.name}" has no description`,
          category: 'manifests'
        });
      }
    }
  }

  private validateCompatibility(ir: CODBIR, errors: ValidationError[], warnings: ValidationWarning[], info: ValidationInfo[]): void {
    const spfxVersion = ir.metadata.spfxVersion;
    const compatibility = SPFx_COMPATIBILITY[spfxVersion as keyof typeof SPFx_COMPATIBILITY];

    if (!compatibility) {
      warnings.push({
        code: 'COMP001',
        message: `Unknown SPFx version: ${spfxVersion}`,
        category: 'compatibility'
      });
    } else {
      info.push({
        code: 'COMP002',
        message: `Targeting SPFx ${spfxVersion} with ${compatibility.buildTool}`,
        category: 'compatibility'
      });

      // Check React compatibility
      const hasReact = ir.components.some(c => c.framework === 'react');
      if (hasReact) {
        info.push({
          code: 'COMP003',
          message: `React ${compatibility.react} required for SPFx ${spfxVersion}`,
          category: 'compatibility'
        });
      }
    }
  }

  private validateSecurity(ir: CODBIR, warnings: ValidationWarning[], info: ValidationInfo[]): void {
    // Check for sensitive data in components
    const jsonIr = JSON.stringify(ir);

    const sensitivePatterns = [
      { pattern: /password/gi, message: 'Password-like content detected' },
      { pattern: /api[_-]?key/gi, message: 'API key-like content detected' },
      { pattern: /secret/gi, message: 'Secret-like content detected' },
      { pattern: /token/gi, message: 'Token-like content detected' }
    ];

    for (const { pattern, message } of sensitivePatterns) {
      if (pattern.test(jsonIr)) {
        warnings.push({
          code: 'SEC001',
          message,
          category: 'security'
        });
      }
    }
  }

  private validateStructure(ir: CODBIR, errors: ValidationError[], warnings: ValidationWarning[], info: ValidationInfo[]): void {
    // Check for required files
    const requiredFiles = [
      'package-solution.json',
      'config/package-solution.json'
    ];

    // Check IR schema
    if (ir.$schema !== 'codbsharepoint/ir/1.0') {
      errors.push({
        code: 'STR001',
        message: 'Invalid IR schema version',
        severity: 'error',
        category: 'structure'
      });
    }

    // Check metadata
    if (!ir.metadata.generator) {
      warnings.push({
        code: 'STR002',
        message: 'IR metadata is missing generator information',
        category: 'structure'
      });
    }
  }

  private validateBundle(ir: CODBIR, bundleFiles: Map<string, string | Uint8Array>, errors: ValidationError[], warnings: ValidationWarning[], info: ValidationInfo[]): void {
    let hasManifest = false;
    let hasBundle = false;

    for (const [path] of bundleFiles) {
      if (path.endsWith('.manifest.json')) {
        hasManifest = true;
      }
      if (path.endsWith('.js')) {
        hasBundle = true;
      }
    }

    if (!hasManifest) {
      errors.push({
        code: 'BND001',
        message: 'No manifest files found in bundle',
        severity: 'error',
        category: 'bundle'
      });
    }

    if (!hasBundle) {
      errors.push({
        code: 'BND002',
        message: 'No JavaScript bundles found',
        severity: 'error',
        category: 'bundle'
      });
    }

    // Check bundle size
    let totalSize = 0;
    for (const [, content] of bundleFiles) {
      const size = typeof content === 'string' ? content.length : content.length;
      totalSize += size;
    }

    if (totalSize > 5 * 1024 * 1024) { // 5MB
      warnings.push({
        code: 'BND003',
        message: `Bundle size is ${(totalSize / 1024 / 1024).toFixed(2)}MB. Consider code splitting.`,
        category: 'bundle'
      });
    }

    info.push({
      code: 'BND004',
      message: `Bundle contains ${bundleFiles.size} files, total size: ${(totalSize / 1024).toFixed(2)}KB`,
      category: 'bundle'
    });
  }

  private generateSummary(errors: ValidationError[], warnings: ValidationWarning[], info: ValidationInfo[]): ValidationSummary {
    const categories: Record<string, { errors: number; warnings: number }> = {};

    for (const error of errors) {
      if (!categories[error.category]) {
        categories[error.category] = { errors: 0, warnings: 0 };
      }
      categories[error.category].errors++;
    }

    for (const warning of warnings) {
      if (!categories[warning.category]) {
        categories[warning.category] = { errors: 0, warnings: 0 };
      }
      categories[warning.category].warnings++;
    }

    return {
      total: errors.length + warnings.length + info.length,
      errors: errors.length,
      warnings: warnings.length,
      info: info.length,
      categories
    };
  }
}
