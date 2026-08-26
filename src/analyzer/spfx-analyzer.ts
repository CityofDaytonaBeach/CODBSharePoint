// ============================================================================
// Analyzer - Analyzes code for permissions, dependencies, and complexity
// ============================================================================

import type {
  CODBIR,
  AnalysisResult,
  GraphPermissionDefinition,
  Framework,
  ComponentType
} from '../types/index.js';

export class SPFxAnalyzer {

  async analyze(ir: CODBIR): Promise<AnalysisResult> {
    const framework = this.detectFramework(ir);
    const componentTypes = this.detectComponentTypes(ir);
    const requiredPermissions = this.analyzePermissions(ir);
    const complexity = this.assessComplexity(ir);
    const estimatedBundleSize = this.estimateBundleSize(ir);
    const dependencies = this.analyzeDependencies(ir);
    const recommendations = this.generateRecommendations(ir);

    return {
      framework,
      componentTypes,
      requiredPermissions,
      complexity,
      estimatedBundleSize,
      dependencies,
      recommendations
    };
  }

  private detectFramework(ir: CODBIR): Framework {
    const frameworks = ir.components.map(c => c.framework);
    if (frameworks.includes('react')) return 'react';
    if (frameworks.includes('knockout')) return 'knockout';
    if (frameworks.includes('vue')) return 'vue';
    return 'none';
  }

  private detectComponentTypes(ir: CODBIR): ComponentType[] {
    const types = new Set<ComponentType>();
    for (const comp of ir.components) {
      types.add(comp.type);
    }
    for (const ext of ir.extensions) {
      types.add('extension');
    }
    return Array.from(types);
  }

  private analyzePermissions(ir: CODBIR): GraphPermissionDefinition[] {
    const permissions: GraphPermissionDefinition[] = [];

    // Check declared permissions
    for (const perm of ir.graph) {
      permissions.push(perm);
    }

    // Analyze code for implicit permissions (simplified)
    const irJson = JSON.stringify(ir);

    // Graph API patterns that indicate specific permissions
    const patterns: Array<{ regex: RegExp; scope: string }> = [
      { regex: /\/users/g, scope: 'User.Read.All' },
      { regex: /\/me\/photo/g, scope: 'User.Read' },
      { regex: /\/groups/g, scope: 'Group.Read.All' },
      { regex: /\/sites/g, scope: 'Sites.Read.All' },
      { regex: /\/drive/g, scope: 'Files.Read.All' },
      { regex: /\/messages/g, scope: 'Mail.Read' },
      { regex: /\/events/g, scope: 'Calendars.ReadWrite' },
      { regex: /\/contacts/g, scope: 'Contacts.Read' }
    ];

    for (const { regex, scope } of patterns) {
      if (regex.test(irJson)) {
        const exists = permissions.some(p => p.scope === scope);
        if (!exists) {
          permissions.push({
            resource: 'Microsoft Graph',
            scope,
            type: 'Delegated',
            requiresAdminApproval: true,
            description: `Detected from code usage pattern`
          });
        }
      }
    }

    return permissions;
  }

  private assessComplexity(ir: CODBIR): 'simple' | 'moderate' | 'complex' {
    let score = 0;

    // Component count
    score += ir.components.length * 2;
    score += ir.extensions.length * 3;

    // Lists and libraries
    score += ir.lists.length;
    score += ir.libraries.length;

    // Fields and content types
    score += ir.fields.length * 0.5;
    score += ir.contentTypes.length;

    // Permissions
    score += ir.graph.length * 2;

    // Formatting
    score += ir.formatting.length;

    if (score <= 5) return 'simple';
    if (score <= 15) return 'moderate';
    return 'complex';
  }

  private estimateBundleSize(ir: CODBIR): number {
    let estimated = 0;

    // Base SPFx runtime
    estimated += 50 * 1024; // 50KB base

    // React
    if (ir.components.some(c => c.framework === 'react')) {
      estimated += 130 * 1024; // 130KB React
    }

    // Per component
    estimated += ir.components.length * 20 * 1024; // 20KB per component
    estimated += ir.extensions.length * 15 * 1024; // 15KB per extension

    return estimated;
  }

  private analyzeDependencies(ir: CODBIR): string[] {
    const deps = new Set<string>();

    // SPFx core dependencies
    deps.add('@microsoft/sp-core-library');
    deps.add('@microsoft/sp-lodash-subset');
    deps.add('@microsoft/sp-property-pane');
    deps.add('@microsoft/sp-http');

    // Framework dependencies
    if (ir.components.some(c => c.framework === 'react')) {
      deps.add('react');
      deps.add('react-dom');
      deps.add('@types/react');
      deps.add('@types/react-dom');
      deps.add('@microsoft/sp-component-base');
    }

    // Graph usage
    if (ir.graph.length > 0) {
      deps.add('@microsoft/microsoft-graph-client');
    }

    return Array.from(deps);
  }

  private generateRecommendations(ir: CODBIR): string[] {
    const recommendations: string[] = [];

    // Check for large number of components
    if (ir.components.length > 5) {
      recommendations.push('Consider splitting into multiple solutions for better maintainability');
    }

    // Check for permissions
    if (ir.graph.length > 3) {
      recommendations.push('Multiple Graph permissions detected. Ensure each is necessary.');
    }

    // Check for deprecated patterns
    if (ir.metadata.spfxVersion && parseInt(ir.metadata.spfxVersion) < 18) {
      recommendations.push('Consider upgrading to SPFx 1.18+ for Heft-based tooling');
    }

    // Check for missing descriptions
    const missingDescriptions = ir.components.filter(c => !c.description);
    if (missingDescriptions.length > 0) {
      recommendations.push('Add descriptions to components for better App Catalog discoverability');
    }

    // Check for source maps
    if (!ir.metadata.modifiedAt) {
      recommendations.push('Enable source maps for easier debugging');
    }

    // Check for localization
    if (!ir.localization?.languages || ir.localization.languages.length <= 1) {
      recommendations.push('Consider adding localization support for broader deployment');
    }

    return recommendations;
  }
}
