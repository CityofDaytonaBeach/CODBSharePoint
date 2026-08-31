// ============================================================================
// TypeScript semantic diagnostics for VFS-backed projects
// ============================================================================

import type { CompileError } from './spfx-compiler.js';

type TypeScriptModule = typeof import('typescript');

let cachedTypeScript: TypeScriptModule | undefined;
let availabilityError: string | undefined;

async function resolveTypeScript(): Promise<TypeScriptModule | undefined> {
  if (cachedTypeScript) return cachedTypeScript;
  if (availabilityError !== undefined) return undefined;

  try {
    cachedTypeScript = (await import(/* webpackIgnore: true */ 'typescript')) as TypeScriptModule;
    return cachedTypeScript;
  } catch (err) {
    availabilityError = err instanceof Error ? err.message : String(err);
    return undefined;
  }
}

const ambientTypes = `
declare const console: any;
declare const window: any;
declare const document: any;
declare const URL: any;
declare const Blob: any;
declare class Error { constructor(message?: string); message: string; }
declare class Promise<T> { constructor(executor: any); static resolve(): Promise<void>; static resolve<T>(value: T): Promise<T>; }
declare interface Array<T> { length: number; map<U>(callback: (value: T, index: number) => U): U[]; }
declare interface ReadonlyArray<T> { length: number; }
declare interface String { length: number; }
declare interface Number {}
declare interface Boolean {}
declare interface Object {}
declare interface Function {}
declare interface RegExp {}
declare interface IArguments {}
declare type Record<K extends string | number | symbol, T> = { [P in K]: T };
declare type Partial<T> = { [P in keyof T]?: T[P] };
declare const JSON: { stringify(value: any): string };

declare namespace JSX {
  interface IntrinsicElements { [elementName: string]: any; }
}

declare module 'react' {
  export type ReactElement<T = any> = any;
  export class Component<P = any, S = any> {
    constructor(props: P);
    props: P;
    state: S;
    setState(state: Partial<S>): void;
  }
  export function createElement(type: any, props?: any, ...children: any[]): any;
}

declare module 'react-dom' {
  export function render(element: any, container: any): void;
  export function unmountComponentAtNode(container: any): void;
}

declare module '@microsoft/sp-core-library' { export class Version { static parse(value: string): Version; } export const Log: any; }
declare module '@microsoft/sp-property-pane' { export type IPropertyPaneConfiguration = any; export const PropertyPaneTextField: any; }
declare module '@microsoft/sp-webpart-base' { export class BaseClientSideWebPart<T = any> { properties: T; context: any; domElement: any; } export type WebPartContext = any; }
declare module '@microsoft/sp-application-base' { export type IApplicationCustomizerProperties = any; export class BaseApplicationCustomizer<T = any> {} }
declare module '@microsoft/sp-listview-extensibility' { export type IFieldCustomizerCellProperties = any; export type IListViewCommandSetProperties = any; export type RowClickedEvent = any; export class BaseFieldCustomizer<T = any> {} export class BaseListViewCommandSet<T = any> {} }
declare module '*.module.css' { const styles: Record<string, string>; export default styles; }
declare module '*.module.scss' { const styles: Record<string, string>; export default styles; }
`;

export interface TypeScriptCheckResult {
  /** Real semantic type diagnostics (only when the compiler was available). */
  errors: CompileError[];
  /** True when the TypeScript compiler could not be loaded (browser without TS). */
  unavailable: boolean;
  /** Human-readable reason for unavailability, if any. */
  unavailableReason?: string;
}

export async function checkTypeScriptFiles(files: Map<string, string>): Promise<TypeScriptCheckResult> {
  const ts = await resolveTypeScript();
  if (!ts) {
    // TypeScript is optional for transpilation (esbuild handles it). Report
    // unavailability separately so callers can downgrade to a warning instead
    // of aborting the whole build in browser-only environments.
    return {
      errors: [],
      unavailable: true,
      unavailableReason: availabilityError || 'TypeScript compiler is not available'
    };
  }

  const virtualFiles = new Map(files);
  virtualFiles.set('/codb/ambient.d.ts', ambientTypes);

  const rootNames = Array.from(virtualFiles.keys()).filter(path => /\.tsx?$|\.d\.ts$/.test(path));
  if (rootNames.length === 0) return { errors: [], unavailable: false };

  const options: import('typescript').CompilerOptions = {
    target: ts.ScriptTarget.ES2022,
    module: ts.ModuleKind.ESNext,
    moduleResolution: ts.ModuleResolutionKind.Bundler,
    jsx: ts.JsxEmit.React,
    noEmit: true,
    noLib: true,
    types: [],
    strict: true,
    skipLibCheck: true,
    allowSyntheticDefaultImports: true,
    esModuleInterop: true
  };

  const host = ts.createCompilerHost(options, true);
  host.getCurrentDirectory = () => '';
  host.fileExists = fileName => virtualFiles.has(normalizePath(fileName));
  host.readFile = fileName => virtualFiles.get(normalizePath(fileName));
  host.getSourceFile = (fileName, languageVersion) => {
    const normalized = normalizePath(fileName);
    const content = virtualFiles.get(normalized);
    return content === undefined ? undefined : ts.createSourceFile(normalized, content, languageVersion, true);
  };
  host.resolveModuleNames = (moduleNames, containingFile) => moduleNames.map(moduleName => {
    if (!moduleName.startsWith('.')) return undefined;

    const containingDir = containingFile.includes('/')
      ? containingFile.slice(0, containingFile.lastIndexOf('/'))
      : '';
    const base = normalizeRelativePath(`${containingDir}/${moduleName}`);
    const resolved = resolveVirtualModule(base, virtualFiles);
    if (!resolved) return undefined;

    return {
      resolvedFileName: resolved.path,
      extension: resolved.extension
    };
  });
  host.writeFile = () => undefined;

  const program = ts.createProgram(rootNames, options, host);
  const diagnostics = ts.getPreEmitDiagnostics(program).filter(diagnostic => !isSuppressedDiagnostic(diagnostic.code));

  return {
    errors: diagnostics.map(diagnostic => {
      const message = ts.flattenDiagnosticMessageText(diagnostic.messageText, '\n');
      const file = diagnostic.file;
      const position = file && diagnostic.start !== undefined ? file.getLineAndCharacterOfPosition(diagnostic.start) : undefined;

      return {
        message,
        file: file?.fileName === '/codb/ambient.d.ts' ? undefined : file?.fileName,
        line: position ? position.line + 1 : undefined,
        column: position ? position.character + 1 : undefined,
        severity: 'error'
      };
    }),
    unavailable: false
  };
}

function normalizePath(path: string): string {
  return path.replace(/\\/g, '/');
}

function normalizeRelativePath(path: string): string {
  const parts: string[] = [];
  for (const part of normalizePath(path).split('/')) {
    if (!part || part === '.') continue;
    if (part === '..') {
      parts.pop();
      continue;
    }
    parts.push(part);
  }
  return parts.join('/');
}

function resolveVirtualModule(
  base: string,
  files: Map<string, string>
): { path: string; extension: import('typescript').Extension } | undefined {
  if (/\.(s?css|json|png|jpe?g|gif|svg|woff2?|ttf|eot)$/i.test(base)) {
    return undefined;
  }

  const candidates: Array<[string, import('typescript').Extension]> = [
    [`${base}.ts`, '.ts' as import('typescript').Extension],
    [`${base}.tsx`, '.tsx' as import('typescript').Extension],
    [`${base}.d.ts`, '.d.ts' as import('typescript').Extension],
    [`${base}/index.ts`, '.ts' as import('typescript').Extension],
    [`${base}/index.tsx`, '.tsx' as import('typescript').Extension]
  ];

  for (const [path, extension] of candidates) {
    if (files.has(path)) return { path, extension };
  }

  return undefined;
}

function isSuppressedDiagnostic(code: number): boolean {
  return [2318, 6053].includes(code);
}
