// ============================================================================
// Compiler - Transforms TypeScript/React source code
// Uses esbuild-wasm for browser-based compilation
// ============================================================================

import type { CODBIR, ComponentDefinition, Framework, SPFxVersion, BuildOptions, VFSFile } from '../types/index.js';
import { createVFS, type VFS } from '../core/vfs.js';
import { generateSharePointArtifacts } from '../provisioning/generator.js';
import { transformContent } from '../bundler/esbuild-runtime.js';
import { generateLocalizationFiles } from '../localization/generator.js';
import { checkTypeScriptFiles } from './typescript-checker.js';

export interface CompileOptions {
  framework: Framework;
  target: SPFxVersion;
  minify?: boolean;
  sourceMaps?: boolean;
  externals?: string[];
  define?: Record<string, string>;
}

export interface CompileResult {
  success: boolean;
  files: VFSFile[];
  errors: CompileError[];
  warnings: string[];
  entryPoint: string;
  bundleSize: number;
}

export interface CompileError {
  message: string;
  file?: string;
  line?: number;
  column?: number;
  severity: 'error' | 'warning';
}

// ---------------------------------------------------------------------------
// Main Compiler Class
// ---------------------------------------------------------------------------

export class SPFxCompiler {
  private options: CompileOptions;
  private vfs: VFS;

  constructor(options: Partial<CompileOptions> = {}) {
    this.options = {
      framework: options.framework || 'react',
      target: options.target || '1.22.0',
      minify: options.minify ?? false,
      sourceMaps: options.sourceMaps ?? true,
      externals: options.externals || [],
      define: options.define || {}
    };
    this.vfs = createVFS();
  }

  // Compile IR to bundle files
  async compile(ir: CODBIR, sourceFiles?: Map<string, string>): Promise<CompileResult> {
    const errors: CompileError[] = [];
    const warnings: string[] = [];
    const startTime = Date.now();

    try {
      // Generate scaffold if no source files provided
      if (!sourceFiles || sourceFiles.size === 0) {
        sourceFiles = this.generateScaffold(ir);
      }

      const typeErrors = await checkTypeScriptFiles(sourceFiles);
      if (typeErrors.length > 0) {
        return {
          success: false,
          files: [],
          errors: typeErrors,
          warnings,
          entryPoint: ir.components[0]?.entry || '',
          bundleSize: 0
        };
      }

      // Compile each source file
      const compiledFiles: VFSFile[] = [];

      for (const [path, content] of sourceFiles) {
        const result = await this.compileFile(path, content);
        if (result.errors.length > 0) {
          errors.push(...result.errors);
        }
        compiledFiles.push(...result.files);
        warnings.push(...result.warnings);
      }

      // Generate bundle entry points
      const entryFiles = this.generateEntryPoints(ir);
      compiledFiles.push(...entryFiles);

      // Generate manifest files
      const manifestFiles = this.generateManifestFiles(ir);
      compiledFiles.push(...manifestFiles);

      // Generate SharePoint artifacts (themes, formatting, provisioning, pages)
      compiledFiles.push(...generateSharePointArtifacts(ir));

      // Generate localization files (.resx + localized strings)
      compiledFiles.push(...generateLocalizationFiles(ir));

      // Add to VFS
      for (const file of compiledFiles) {
        this.vfs.addFile(file.path, file.content);
      }

      return {
        success: errors.length === 0,
        files: compiledFiles,
        errors,
        warnings,
        entryPoint: ir.components[0]?.entry || '',
        bundleSize: compiledFiles.reduce((acc, f) => acc + (typeof f.content === 'string' ? f.content.length : f.content.length), 0)
      };
    } catch (err) {
      errors.push({
        message: err instanceof Error ? err.message : 'Unknown compile error',
        severity: 'error'
      });

      return {
        success: false,
        files: [],
        errors,
        warnings,
        entryPoint: '',
        bundleSize: 0
      };
    }
  }

  // Compile a single file
  private async compileFile(path: string, content: string): Promise<CompileResult> {
    const errors: CompileError[] = [];
    const warnings: string[] = [];
    const files: VFSFile[] = [];

    // Determine file type
    const ext = path.split('.').pop()?.toLowerCase();

    if (ext === 'ts' || ext === 'tsx') {
      // TypeScript compilation
      const result = await this.compileTypeScript(path, content);
      files.push(...result.files);
      errors.push(...result.errors);
      warnings.push(...result.warnings);
    } else if (ext === 'jsx') {
      const result = await this.compileJavaScript(path, content);
      files.push(...result.files);
      errors.push(...result.errors);
      warnings.push(...result.warnings);
    } else if (ext === 'scss' || ext === 'css') {
      // CSS/SCSS compilation
      const result = await this.compileStyles(path, content);
      files.push(...result.files);
      errors.push(...result.errors);
    } else if (ext === 'json') {
      // JSON files pass through
      files.push({
        path: path.replace(/\.json$/, '.json'),
        content,
        encoding: 'utf-8'
      });
    } else {
      // Other files pass through
      files.push({
        path,
        content,
        encoding: 'utf-8'
      });
    }

    return {
      success: errors.length === 0,
      files,
      errors,
      warnings,
      entryPoint: path,
      bundleSize: content.length
    };
  }

  private async compileJavaScript(path: string, content: string): Promise<CompileResult> {
    const errors: CompileError[] = [];
    const warnings: string[] = [];
    const files: VFSFile[] = [];

    try {
      const transformResult = await transformContent(content, {
        loader: 'jsx',
        minify: this.options.minify,
        sourceMap: false,
        target: 'es2022'
      });

      if (!transformResult.ok || !transformResult.code) {
        errors.push({
          message: transformResult.error || 'esbuild produced no output',
          file: path,
          severity: 'error'
        });
      } else {
        files.push({
          path: path.replace(/\.jsx$/, '.js').replace(/^src\//, 'lib/'),
          content: this.rewriteStyleImports(transformResult.code),
          encoding: 'utf-8'
        });
      }
    } catch (err) {
      errors.push({
        message: err instanceof Error ? err.message : 'JavaScript compilation failed',
        file: path,
        severity: 'error'
      });
    }

    return {
      success: errors.length === 0,
      files,
      errors,
      warnings,
      entryPoint: path,
      bundleSize: content.length
    };
  }

  // TypeScript compilation
  private async compileTypeScript(path: string, content: string): Promise<CompileResult> {
    const errors: CompileError[] = [];
    const warnings: string[] = [];
    const files: VFSFile[] = [];

    // Use esbuild-wasm for real TypeScript/TSX transformation. Do not report
    // production success when the compiler is unavailable or failed.
    try {
      const transformResult = await transformContent(content, {
        loader: path.endsWith('.tsx') ? 'tsx' : 'ts',
        minify: this.options.minify,
        sourceMap: false,
        target: 'es2022'
      });

      if (!transformResult.ok || !transformResult.code) {
        errors.push({
          message: transformResult.error || 'esbuild produced no output',
          file: path,
          severity: 'error'
        });

        return {
          success: false,
          files,
          errors,
          warnings,
          entryPoint: path,
          bundleSize: content.length
        };
      }

      const compiled = this.rewriteStyleImports(transformResult.code);

      // Output compiled JS
      const outputPath = path
        .replace(/\.tsx?$/, '.js')
        .replace(/^src\//, 'lib/');

      files.push({
        path: outputPath,
        content: compiled,
        encoding: 'utf-8'
      });

      // Generate declaration file if needed
      if (path.endsWith('.ts')) {
        const dtsPath = path.replace(/\.ts$/, '.d.ts').replace(/^src\//, 'lib/');
        files.push({
          path: dtsPath,
          content: this.generateDeclaration(content),
          encoding: 'utf-8'
        });
      }

      // Source maps
      if (this.options.sourceMaps) {
        const mapPath = `${outputPath}.map`;
        files.push({
          path: mapPath,
          content: this.generateSourceMap(path, content, compiled),
          encoding: 'utf-8'
        });
      }
    } catch (err) {
      errors.push({
        message: err instanceof Error ? err.message : 'TypeScript compilation failed',
        file: path,
        severity: 'error'
      });
    }

    return {
      success: errors.length === 0,
      files,
      errors,
      warnings,
      entryPoint: path,
      bundleSize: content.length
    };
  }

  // Style compilation (simplified)
  private async compileStyles(path: string, content: string): Promise<CompileResult> {
    const errors: CompileError[] = [];
    const files: VFSFile[] = [];

    try {
      let compiled = content;

      // SCSS to CSS (simplified - remove nesting)
      if (path.endsWith('.scss')) {
        const unsupported = this.findUnsupportedScssFeatures(content);
        if (unsupported.length > 0) {
          errors.push({
            message: `Unsupported SCSS syntax: ${unsupported.join(', ')}. A browser Sass compiler is required for this file.`,
            file: path,
            severity: 'error'
          });

          return {
            success: false,
            files,
            errors,
            warnings: [],
            entryPoint: path,
            bundleSize: content.length
          };
        }

        compiled = this.compileSCSS(content);
      }

      const outputPath = path
        .replace(/\.scss$/, '.css')
        .replace(/^src\//, 'lib/');

      files.push({
        path: outputPath,
        content: compiled,
        encoding: 'utf-8'
      });
    } catch (err) {
      errors.push({
        message: err instanceof Error ? err.message : 'Style compilation failed',
        file: path,
        severity: 'error'
      });
    }

    return {
      success: errors.length === 0,
      files,
      errors,
      warnings: [],
      entryPoint: path,
      bundleSize: content.length
    };
  }

  // Generate scaffold files from IR
  private generateScaffold(ir: CODBIR): Map<string, string> {
    const files = new Map<string, string>();
    const namespace = ir.solution.namespace;

    for (const component of ir.components) {
      if (component.type === 'webpart') {
        const dir = `src/webparts/${component.name}`;

        // Main web part file
        if (component.framework === 'react') {
          files.set(`${dir}/${component.name}WebPart.ts`, this.generateReactWebPart(component, namespace));
          files.set(`${dir}/components/${component.name}.tsx`, this.generateReactComponent(component));
          files.set(`${dir}/components/${component.name}.module.scss`, this.generateStyles(component));
          files.set(`${dir}/components/${component.name}Props.ts`, this.generatePropsInterface(component));
        } else {
          files.set(`${dir}/${component.name}WebPart.ts`, this.generateBasicWebPart(component, namespace));
        }
      }
    }

    for (const ext of ir.extensions) {
      const dir = `src/extensions/${ext.name}`;
      if (ext.type === 'ApplicationCustomizer') {
        files.set(`${dir}/${ext.name}ApplicationCustomizer.ts`, this.generateApplicationCustomizer(ext));
        files.set(`${dir}/${ext.name}ApplicationCustomizer.module.scss`, this.generateExtensionStyles(ext));
      } else if (ext.type === 'FieldCustomizer') {
        files.set(`${dir}/${ext.name}FieldCustomizer.ts`, this.generateFieldCustomizer(ext));
      } else if (ext.type === 'ListViewCommandSet') {
        files.set(`${dir}/${ext.name}CommandSet.ts`, this.generateCommandSet(ext));
      }
    }

    return files;
  }

  // Generate React WebPart file
  private generateReactWebPart(component: ComponentDefinition, namespace: string): string {
    return `import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import ${component.name}Component from './components/${component.name}';
import type { I${component.name}Props } from './components/${component.name}Props';

export interface I${component.name}WebPartProps {
  description: string;
}

export default class ${component.name}WebPart extends BaseClientSideWebPart<I${component.name}WebPartProps> {

  public render(): void {
    const element: React.ReactElement<I${component.name}Props> = React.createElement(
      ${component.name}Component,
      {
        description: this.properties.description,
        context: this.context,
        siteUrl: this.context.pageContext.web.absoluteUrl
      }
    );

    ReactDOM.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDOM.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: '${component.displayName} Settings'
          },
          groups: [
            {
              groupName: 'Basic Settings',
              groupFields: [
                PropertyPaneTextField('description', {
                  label: 'Description'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
`;
  }

  // Generate React Component
  private generateReactComponent(component: ComponentDefinition): string {
    return `import * as React from 'react';
import styles from './${component.name}.module.scss';

export interface I${component.name}Props {
  description: string;
  context: any;
  siteUrl: string;
}

export interface I${component.name}State {
  loading: boolean;
  data: any[];
}

export default class ${component.name}Component extends React.Component<I${component.name}Props, I${component.name}State> {
  constructor(props: I${component.name}Props) {
    super(props);
    this.state = {
      loading: true,
      data: []
    };
  }

  public componentDidMount(): void {
    this.loadData();
  }

  private async loadData(): Promise<void> {
    try {
      // Load data from SharePoint
      this.setState({ loading: false });
    } catch (error) {
      console.error('Error loading data:', error);
      this.setState({ loading: false });
    }
  }

  public render(): React.ReactElement<I${component.name}Props> {
    const { loading, data } = this.state;

    return (
      <div className={styles.${component.name}}>
        <div className={styles.container}>
          <h2>${component.displayName}</h2>
          {loading ? (
            <div className={styles.loading}>Loading...</div>
          ) : (
            <div className={styles.content}>
              {data.map((item, index) => (
                <div key={index} className={styles.item}>
                  {JSON.stringify(item)}
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    );
  }
}
`;
  }

  // Generate Styles
  private generateStyles(component: ComponentDefinition): string {
    return `@import '~@microsoft/sp-tslint-theme/vars.scss';

.${component.name} {
  .container {
    max-width: 700px;
    margin: 0px auto;
    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.2);
    padding: 20px;
  }

  .loading {
    text-align: center;
    padding: 20px;
  }

  .content {
    margin-top: 10px;
  }

  .item {
    padding: 10px;
    border-bottom: 1px solid #eee;
  }
}
`;
  }

  // Generate Props Interface
  private generatePropsInterface(component: ComponentDefinition): string {
    return `import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface I${component.name}Props {
  description: string;
  context: WebPartContext;
  siteUrl: string;
}
`;
  }

  // Generate Basic WebPart (non-React)
  private generateBasicWebPart(component: ComponentDefinition, namespace: string): string {
    return `import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

export interface I${component.name}WebPartProps {
  description: string;
}

export default class ${component.name}WebPart extends BaseClientSideWebPart<I${component.name}WebPartProps> {

  public render(): void {
    this.domElement.innerHTML = \`
      <div class="${namespace}-${component.name}">
        <h2>${component.displayName}</h2>
        <p>\${this.properties.description}</p>
      </div>
    \`;
  }

  protected onDispose(): void {
    // Cleanup
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: '${component.displayName}' },
          groups: [
            {
              groupName: 'Settings',
              groupFields: [
                PropertyPaneTextField('description', {
                  label: 'Description'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
`;
  }

  // Generate Application Customizer
  private generateApplicationCustomizer(ext: any): string {
    return `import { Log } from '@microsoft/sp-core-library';
import {
  type IApplicationCustomizerProperties,
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';

export interface I${ext.name}Properties extends IApplicationCustomizerProperties {
  testMessage: string;
}

export default class ${ext.name}ApplicationCustomizer extends BaseApplicationCustomizer<I${ext.name}Properties> {

  public async onInit(): Promise<void> {
    Log.info('${ext.name}', 'Initialized');
    return Promise.resolve();
  }

  public onDispose(): void {
    // Cleanup
  }
}
`;
  }

  // Generate Field Customizer
  private generateFieldCustomizer(ext: any): string {
    return `import { Log } from '@microsoft/sp-core-library';
import {
  type IFieldCustomizerCellProperties,
  BaseFieldCustomizer
} from '@microsoft/sp-listview-extensibility';

export default class ${ext.name}FieldCustomizer extends BaseFieldCustomizer<IFieldCustomizerCellProperties> {

  public onInit(): Promise<void> {
    return Promise.resolve();
  }

  public onRenderCell(event: any): void {
    // Render cell
  }

  public onDispose(): void {
    // Cleanup
  }
}
`;
  }

  // Generate Command Set
  private generateCommandSet(ext: any): string {
    return `import { Log } from '@microsoft/sp-core-library';
import {
  type IListViewCommandSetProperties,
  BaseListViewCommandSet,
  RowClickedEvent
} from '@microsoft/sp-listview-extensibility';

export default class ${ext.name}CommandSet extends BaseListViewCommandSet<IListViewCommandSetProperties> {

  public onInit(): Promise<void> {
    return Promise.resolve();
  }

  public onListViewSelectedChanged(event: RowClickedEvent): void {
    // Handle selection
  }
}
`;
  }

  // Generate Extension Styles
  private generateExtensionStyles(ext: any): string {
    return `.${ext.name} {
  // Extension styles
}
`;
  }

  // Generate entry points
  private generateEntryPoints(ir: CODBIR): VFSFile[] {
    const files: VFSFile[] = [];

    files.push({
      path: 'lib/index.js',
      content: `'use strict';
// Auto-generated entry point for ${ir.solution.name}
module.exports = require('./webparts/${ir.components[0]?.name || 'index'}/${ir.components[0]?.name || 'index'}WebPart');
`,
      encoding: 'utf-8'
    });

    return files;
  }

  // Generate manifest files
  private generateManifestFiles(ir: CODBIR): VFSFile[] {
    const files: VFSFile[] = [];

    for (const component of ir.components) {
      if (component.type === 'webpart') {
        const manifest = {
          id: `{${component.id}}`,
          alias: `${ir.solution.namespace}-${component.name}WebPart`,
          componentType: 'WebPart',
          version: component.version,
          manifestVersion: 2,
          preconfiguredEntries: component.preconfiguredEntries.map(entry => ({
            groupId: component.group.id,
            group: { default: component.group.name },
            title: { default: entry.defaultTitle },
            description: { default: entry.description },
            officeFabricIconFontName: entry.officeFabricIconFontName,
            properties: entry.properties
          }))
        };

        files.push({
          path: `sharepoint/assets/${component.name}.manifest.json`,
          content: JSON.stringify(manifest, null, 2),
          encoding: 'utf-8'
        });
      }
    }

    return files;
  }

  private rewriteStyleImports(content: string): string {
    return content.replace(/(from\s+['"][^'"]+)\.scss(['"])/g, '$1.css$2');
  }

  // Generate declaration file (simplified)
  private generateDeclaration(content: string): string {
    // Extract exports and interfaces
    let declaration = '';
    const exportMatches = content.match(/export\s+(?:default\s+)?(?:class|interface|type|function|const|let|var)\s+(\w+)/g);
    if (exportMatches) {
      for (const match of exportMatches) {
        declaration += `${match};\n`;
      }
    }
    return declaration;
  }

  // Generate source map (simplified)
  private generateSourceMap(sourcePath: string, source: string, compiled: string): string {
    return JSON.stringify({
      version: 3,
      file: sourcePath.replace(/\.tsx?$/, '.js'),
      sources: [sourcePath],
      sourcesContent: [source],
      names: [],
      mappings: 'AAAA'
    });
  }

  // SCSS compilation for the generated Phase 1 scaffold. Full Sass syntax still
  // needs a browser Sass compiler before claiming complete SCSS support.
  private compileSCSS(content: string): string {
    const source = content.replace(/@import\s+['"][^'"]+['"]\s*;?/g, '');
    const output: string[] = [];
    const stack: string[] = [];
    let declarations: string[] = [];

    const flush = () => {
      if (stack.length > 0 && declarations.length > 0) {
        output.push(`${stack.join(' ')} {`);
        for (const declaration of declarations) {
          output.push(`  ${declaration}`);
        }
        output.push('}');
        declarations = [];
      }
    };

    for (const rawLine of source.split(/\r?\n/)) {
      const line = rawLine.trim();
      if (!line) continue;

      if (line.endsWith('{')) {
        flush();
        stack.push(line.slice(0, -1).trim());
      } else if (line === '}') {
        flush();
        stack.pop();
      } else {
        declarations.push(line);
      }
    }

    flush();
    return `${output.join('\n')}\n`;
  }

  private findUnsupportedScssFeatures(content: string): string[] {
    const unsupported: string[] = [];
    const checks: Array<[RegExp, string]> = [
      [/\$[a-zA-Z_][\w-]*\s*:/, 'variables'],
      [/@mixin\b/, 'mixins'],
      [/@include\b/, 'includes'],
      [/@use\b/, 'use rules'],
      [/@forward\b/, 'forward rules'],
      [/&[.:#\[]/, 'parent selectors'],
      [/@if\b|@for\b|@each\b|@while\b/, 'control flow']
    ];

    for (const [pattern, label] of checks) {
      if (pattern.test(content)) unsupported.push(label);
    }

    return unsupported;
  }

  getVFS(): VFS {
    return this.vfs;
  }
}
