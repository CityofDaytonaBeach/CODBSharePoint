// ============================================================================
// AI Contract Layer
// Structured, deterministic APIs for AI agents to create, validate, repair,
// and build CODBSharePoint projects without guessing SPFx internals.
// ============================================================================

import type {
  BuildOptions,
  BuildResult,
  CODBIR,
  ComponentDefinition,
  ComponentType,
  FieldDefinition,
  Framework,
  ListDefinition,
  SPFxVersion,
  ValidationResult,
  VFSFile
} from '../types/index.js';
import {
  addColumn,
  addContentType,
  addGraphPermission,
  addList,
  addPage,
  addProvisioning,
  addTheme,
  addWebPart,
  createIR,
  deserializeIR,
  serializeIR
} from '../core/ir.js';
import { SPFx_COMPATIBILITY } from '../types/index.js';
import { getKnowledgeCatalog, summarizeKnowledge } from '../knowledge/index.js';

export interface AIHost {
  buildFromIR(ir: CODBIR, sourceFiles?: Map<string, string>, options?: Partial<BuildOptions>): Promise<BuildResult>;
  validate(ir: CODBIR): Promise<ValidationResult>;
  capabilities(): Record<string, boolean>;
}

export interface AIProjectInput {
  schemaVersion?: 'codbsharepoint.ai/1.0';
  intent?: string;
  target?: {
    spfxVersion?: SPFxVersion;
    runtime?: 'browser' | 'webworker';
    componentTypes?: ComponentType[];
  };
  solution?: {
    name: string;
    version?: string;
    description?: string;
    author?: string;
  };
  components?: AIComponentInput[];
  data?: {
    graphPermissions?: string[];
    lists?: Array<Partial<ListDefinition>>;
    fields?: Array<Partial<FieldDefinition>>;
    contentTypes?: Array<{ name: string; fields?: string[]; description?: string }>;
  };
  artifacts?: {
    themes?: Array<{ name: string; primary: Record<string, string> }>;
    pages?: Array<{ name: string; title: string; content: string }>;
    provisioning?: Array<{ type: 'siteScript' | 'siteDesign'; name: string; description?: string; data: Record<string, unknown> }>;
  };
  files?: Record<string, string>;
}

export interface AIComponentInput {
  type?: 'webpart';
  name: string;
  displayName?: string;
  description?: string;
  framework?: Framework;
  features?: string[];
}

export interface AIProject {
  schemaVersion: 'codbsharepoint.ai/1.0';
  ir: CODBIR;
  files: Map<string, string>;
  intent: string;
}

export interface AIDiagnostic {
  code: string;
  severity: 'error' | 'warning' | 'info';
  stage: string;
  message: string;
  file?: string;
  line?: number;
  column?: number;
  fix?: AIFix;
}

export type AIFix =
  | { type: 'addGraphPermission'; scope: string }
  | { type: 'setSolutionName'; name: string }
  | { type: 'setVersion'; version: string }
  | { type: 'replaceFile'; path: string; contentHint: string }
  | { type: 'addWebPart'; name: string; framework: Framework };

export interface AIPlan {
  intent: string;
  profile: string;
  steps: string[];
  warnings: string[];
}

export interface AIProfile {
  name: string;
  spfxVersion: SPFxVersion;
  supportedComponents: ComponentType[];
  frameworks: Framework[];
  requiredBuildStages: string[];
  allowedImports: string[];
}

export interface AISourceCompatibility {
  compatible: boolean;
  supportedImports: string[];
  unsupportedImports: string[];
  graphPermissions: string[];
  diagnostics: AIDiagnostic[];
}

const DEFAULT_PROFILE: AIProfile = {
  name: 'spfx-1.22-react-webpart',
  spfxVersion: '1.22.0',
  supportedComponents: ['webpart'],
  frameworks: ['react', 'vanilla'],
  requiredBuildStages: ['validate', 'typecheck', 'compile', 'sass', 'bundle', 'manifest', 'sppkg', 'package-validation'],
  allowedImports: ['react', 'react-dom', '@microsoft/sp-core-library', '@microsoft/sp-webpart-base', '@microsoft/sp-property-pane', '@microsoft/sp-http']
};

export class AIService {
  constructor(private readonly host: AIHost) {}

  capabilities(): Record<string, unknown> {
    return {
      schemaVersion: 'codbsharepoint.ai/1.0',
      ...this.host.capabilities(),
      aiContract: true,
      deterministicPaths: true,
      machineRepairHints: true,
      promptContextExport: true,
      sourceCompatibilityAssessment: true,
      sharePointKnowledge: summarizeKnowledge(),
      profiles: ['spfx-1.22-react-webpart', 'spfx-1.22-graph-webpart', 'spfx-1.22-list-provisioning']
    };
  }

  profile(name = DEFAULT_PROFILE.name): AIProfile {
    if (name === 'spfx-1.22-graph-webpart') {
      return { ...DEFAULT_PROFILE, name, allowedImports: [...DEFAULT_PROFILE.allowedImports, '@microsoft/sp-http'] };
    }
    if (name === 'spfx-1.22-list-provisioning') {
      return { ...DEFAULT_PROFILE, name, requiredBuildStages: [...DEFAULT_PROFILE.requiredBuildStages, 'feature-framework'] };
    }
    return DEFAULT_PROFILE;
  }

  plan(input: AIProjectInput): AIPlan {
    const graph = input.data?.graphPermissions || [];
    const lists = input.data?.lists || [];
    return {
      intent: input.intent || `Build ${input.solution?.name || 'a SharePoint web part'}`,
      profile: graph.length > 0 ? 'spfx-1.22-graph-webpart' : lists.length > 0 ? 'spfx-1.22-list-provisioning' : DEFAULT_PROFILE.name,
      steps: ['normalize-input', 'create-ir', 'write-deterministic-source-files', 'validate', 'build', 'validate-sppkg'],
      warnings: input.target?.componentTypes?.some(type => type !== 'webpart')
        ? ['Only web parts are production-proven in the current AI contract.']
        : []
    };
  }

  generate(input: AIProjectInput): AIProject {
    const ir = createIR({
      name: input.solution?.name || safeName(input.intent || 'AIProject'),
      version: input.solution?.version || '1.0.0',
      description: input.solution?.description || input.intent || '',
      author: input.solution?.author || 'AI'
    });
    ir.metadata.spfxVersion = input.target?.spfxVersion || '1.22.0';

    const files = new Map<string, string>(Object.entries(input.files || {}));
    for (const component of input.components || [{ name: 'MainWebPart', framework: 'react' as Framework }]) {
      const webPart = addWebPart(ir, {
        name: component.name,
        displayName: component.displayName || component.name,
        description: component.description || input.intent || '',
        framework: component.framework || 'react'
      });
      if (!hasComponentFiles(files, webPart.name)) {
        this.writeComponentSource({ ir, files, componentName: webPart.name, fileRole: 'component', content: defaultComponentSource(webPart) });
        this.writeComponentSource({ ir, files, componentName: webPart.name, fileRole: 'style', content: defaultStyleSource(webPart) });
        this.writeComponentSource({ ir, files, componentName: webPart.name, fileRole: 'webpart', content: defaultWebPartSource(webPart) });
      }
    }

    for (const scope of input.data?.graphPermissions || []) addGraphPermission(ir, scope);
    for (const field of input.data?.fields || []) addColumn(ir, field);
    for (const list of input.data?.lists || []) addList(ir, list);
    for (const contentType of input.data?.contentTypes || []) addContentType(ir, { fields: [], ...contentType });
    for (const theme of input.artifacts?.themes || []) addTheme(ir, { name: theme.name, primary: theme.primary });
    for (const page of input.artifacts?.pages || []) addPage(ir, page);
    for (const provisioning of input.artifacts?.provisioning || []) addProvisioning(ir, provisioning);

    return {
      schemaVersion: 'codbsharepoint.ai/1.0',
      ir,
      files,
      intent: input.intent || ''
    };
  }

  writeComponentSource(args: {
    ir: CODBIR;
    files: Map<string, string>;
    componentName: string;
    fileRole: 'webpart' | 'component' | 'style' | 'props';
    content: string;
  }): string {
    const path = componentFilePath(args.componentName, args.fileRole);
    args.files.set(path, args.content);
    return path;
  }

  async validate(project: AIProject): Promise<{ valid: boolean; diagnostics: AIDiagnostic[]; validation: ValidationResult }> {
    const validation = await this.host.validate(project.ir);
    return {
      valid: validation.valid,
      diagnostics: validationToDiagnostics(validation),
      validation
    };
  }

  async build(project: AIProject, options: Partial<BuildOptions> = {}): Promise<BuildResult & { aiDiagnostics: AIDiagnostic[] }> {
    const compatibility = this.assessSourceCompatibility(project);
    if (!compatibility.compatible) {
      const result = await this.host.buildFromIR(project.ir, project.files, { runtime: 'browser', ...options });
      return Object.assign(result, { aiDiagnostics: [...compatibility.diagnostics, ...buildToDiagnostics(result)] });
    }

    const result = await this.host.buildFromIR(project.ir, project.files, { runtime: 'browser', ...options });
    return Object.assign(result, { aiDiagnostics: buildToDiagnostics(result) });
  }

  assessSourceCompatibility(project: AIProject): AISourceCompatibility {
    const supportedImports = new Set<string>();
    const unsupportedImports = new Set<string>();
    const graphPermissions = new Set<string>();
    const diagnostics: AIDiagnostic[] = [];

    for (const [path, content] of project.files) {
      for (const specifier of extractImports(content)) {
        if (isRelativeImport(specifier) || isSupportedRuntimeImport(specifier)) {
          supportedImports.add(specifier);
          continue;
        }

        unsupportedImports.add(specifier);
        diagnostics.push({
          code: 'AI_UNSUPPORTED_IMPORT',
          severity: 'error',
          stage: 'source-compatibility',
          file: path,
          message: `Import "${specifier}" requires a browser package registry/prebundle before this SDK can rebuild the full PnP sample.`,
          fix: {
            type: 'replaceFile',
            path,
            contentHint: `Remove or replace "${specifier}", or add it to a future browser package registry.`
          }
        });
      }

      for (const permission of inferGraphPermissions(content)) {
        graphPermissions.add(permission);
        if (!project.ir.graph.some(existing => existing.scope === permission)) {
          diagnostics.push({
            code: 'AI_MISSING_GRAPH_PERMISSION',
            severity: 'error',
            stage: 'permissions',
            file: path,
            message: `Source appears to require Microsoft Graph permission ${permission}.`,
            fix: { type: 'addGraphPermission', scope: permission }
          });
        }
      }
    }

    return {
      compatible: diagnostics.filter(diagnostic => diagnostic.severity === 'error').length === 0,
      supportedImports: Array.from(supportedImports).sort(),
      unsupportedImports: Array.from(unsupportedImports).sort(),
      graphPermissions: Array.from(graphPermissions).sort(),
      diagnostics
    };
  }

  repair(project: AIProject, diagnostics: AIDiagnostic[]): AIProject {
    const repaired: AIProject = {
      ...project,
      ir: deserializeIR(serializeIR(project.ir)),
      files: new Map(project.files)
    };

    for (const diagnostic of diagnostics) {
      if (!diagnostic.fix) continue;
      if (diagnostic.fix.type === 'addGraphPermission') addGraphPermission(repaired.ir, diagnostic.fix.scope);
      if (diagnostic.fix.type === 'setSolutionName') repaired.ir.solution.name = diagnostic.fix.name;
      if (diagnostic.fix.type === 'setVersion') repaired.ir.solution.version = diagnostic.fix.version;
      if (diagnostic.fix.type === 'addWebPart') addWebPart(repaired.ir, { name: diagnostic.fix.name, framework: diagnostic.fix.framework });
    }

    return repaired;
  }

  exportPromptContext(project?: AIProject): string {
    const profile = this.profile();
    return JSON.stringify({
      schemaVersion: 'codbsharepoint.ai/1.0',
      capabilities: this.capabilities(),
      profile,
      spfxCompatibility: SPFx_COMPATIBILITY[profile.spfxVersion],
      knowledge: getKnowledgeCatalog(),
      rules: [
        'AI supplies intent and source; SDK owns paths, manifests, IDs, package structure, and validation.',
        'Use writeComponentSource file roles instead of inventing SPFx paths.',
        'Do not claim success unless BuildResult.success is true and aiDiagnostics is empty.',
        'Use Graph permissions declared in data.graphPermissions for Microsoft Graph calls.',
        'Run assessSourceCompatibility before build when importing PnP/gallery source.'
      ],
      project: project ? {
        solution: project.ir.solution,
        components: project.ir.components.map(component => ({ name: component.name, type: component.type, framework: component.framework })),
        files: Array.from(project.files.keys()).sort()
      } : undefined
    }, null, 2);
  }

  import(input: AIProjectInput | CODBIR | string): AIProject {
    if (typeof input === 'string') {
      const parsed = JSON.parse(input) as AIProjectInput | CODBIR;
      return this.import(parsed);
    }
    if (isCODBIR(input)) {
      return { schemaVersion: 'codbsharepoint.ai/1.0', ir: input, files: new Map(), intent: '' };
    }
    return this.generate(input);
  }
}

function extractImports(content: string): string[] {
  const imports = new Set<string>();
  const patterns = [
    /import\s+(?:type\s+)?(?:[^'";]+\s+from\s+)?['"]([^'"]+)['"]/g,
    /export\s+(?:[^'";]+\s+from\s+)?['"]([^'"]+)['"]/g,
    /import\(\s*['"]([^'"]+)['"]\s*\)/g
  ];

  for (const pattern of patterns) {
    let match: RegExpExecArray | null;
    while ((match = pattern.exec(content)) !== null) {
      imports.add(match[1]);
    }
  }

  return Array.from(imports);
}

function isRelativeImport(specifier: string): boolean {
  return specifier.startsWith('./') || specifier.startsWith('../') || specifier.startsWith('/');
}

function isSupportedRuntimeImport(specifier: string): boolean {
  if (specifier === 'react' || specifier === 'react-dom') return true;
  if (specifier.startsWith('@microsoft/')) return true;
  return false;
}

function inferGraphPermissions(content: string): string[] {
  const permissions = new Set<string>();
  if (/\.api\(\s*['"]\/users\b|https:\/\/graph\.microsoft\.com\/[^'"`]*\/users\b/i.test(content)) permissions.add('User.Read.All');
  if (/\.api\(\s*['"]\/groups\b|\/groups\/[^'"`]+\/members\b|https:\/\/graph\.microsoft\.com\/[^'"`]*\/groups\b/i.test(content)) permissions.add('Group.Read.All');
  if (/\.api\(\s*['"]\/sites\b|https:\/\/graph\.microsoft\.com\/[^'"`]*\/sites\b/i.test(content)) permissions.add('Sites.Read.All');
  if (/\.api\(\s*['"]\/me\b|https:\/\/graph\.microsoft\.com\/[^'"`]*\/me\b/i.test(content)) permissions.add('User.Read');
  return Array.from(permissions);
}

function isCODBIR(input: AIProjectInput | CODBIR): input is CODBIR {
  return '$schema' in input && input.$schema === 'codbsharepoint/ir/1.0';
}

function componentFilePath(componentName: string, role: 'webpart' | 'component' | 'style' | 'props'): string {
  if (role === 'webpart') return `src/webparts/${componentName}/${componentName}WebPart.ts`;
  if (role === 'component') return `src/webparts/${componentName}/components/${componentName}.tsx`;
  if (role === 'style') return `src/webparts/${componentName}/components/${componentName}.module.scss`;
  return `src/webparts/${componentName}/components/${componentName}Props.ts`;
}

function hasComponentFiles(files: Map<string, string>, componentName: string): boolean {
  return files.has(componentFilePath(componentName, 'webpart')) && files.has(componentFilePath(componentName, 'component'));
}

function validationToDiagnostics(validation: ValidationResult): AIDiagnostic[] {
  return [
    ...validation.errors.map(error => ({
      code: error.code,
      severity: 'error' as const,
      stage: error.category,
      message: error.message,
      file: error.file,
      line: error.line,
      column: error.column,
      fix: fixFor(error.code)
    })),
    ...validation.warnings.map(warning => ({
      code: warning.code,
      severity: 'warning' as const,
      stage: warning.category,
      message: warning.message,
      file: warning.file
    }))
  ];
}

function buildToDiagnostics(result: BuildResult): AIDiagnostic[] {
  return result.errors.map(error => ({
    code: error.code,
    severity: error.severity === 'info' ? 'info' : error.severity === 'warning' ? 'warning' : 'error',
    stage: error.category,
    message: error.message,
    file: error.file,
    line: error.line,
    column: error.column,
    fix: error.file ? { type: 'replaceFile', path: error.file, contentHint: 'Return valid browser-compatible TypeScript/TSX/SCSS for this file.' } : undefined
  }));
}

function fixFor(code: string): AIFix | undefined {
  if (code === 'SOL001') return { type: 'setSolutionName', name: 'AIRepairedSolution' };
  if (code === 'SOL005') return { type: 'setVersion', version: '1.0.0' };
  return undefined;
}

function safeName(value: string): string {
  const name = value.replace(/[^a-zA-Z0-9]/g, '');
  return /^[a-zA-Z]/.test(name) ? name.slice(0, 40) : `AI${name.slice(0, 38)}`;
}

function defaultWebPartSource(component: ComponentDefinition): string {
  return `import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import ${component.name} from './components/${component.name}';

export default class ${component.name}WebPart extends BaseClientSideWebPart<any> {
  public render(): void {
    ReactDOM.render(React.createElement(${component.name}, { context: this.context }), this.domElement);
  }
  protected onDispose(): void { ReactDOM.unmountComponentAtNode(this.domElement); }
  protected get dataVersion(): Version { return Version.parse('1.0'); }
}
`;
}

function defaultComponentSource(component: ComponentDefinition): string {
  return `import * as React from 'react';
import styles from './${component.name}.module.scss';

export default class ${component.name} extends React.Component<any, any> {
  public render(): React.ReactElement<any> {
    return <section className={styles.root}><h2>${component.displayName}</h2><p>${component.description}</p></section>;
  }
}
`;
}

function defaultStyleSource(component: ComponentDefinition): string {
  return `$accent: #005a9e;
.root {
  border-left: 4px solid $accent;
  display: grid;
  gap: 8px;
  padding: 16px;
}
`;
}
