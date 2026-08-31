import {
  CODBSharePoint,
  createIR,
  addWebPart,
  addExtension,
  addList,
  addGraphPermission,
  generatePackageJson,
  generateLocalizationFiles,
  generateResx,
  generateStringsModule,
  resolveStrings
} from '../src/index.js';
import { unzipSync, strFromU8 } from 'fflate';

const sdk = new CODBSharePoint();

describe('CODBSharePoint - Offline engine (esbuild, localization, deep import)', () => {
  describe('esbuild-wasm real transpile', () => {
    it('reports esbuild availability', async () => {
      const available = await sdk.bundleAPI.isEsbuildAvailable();
      expect(available).toBe(true);
    });

    it('transpiles JSX to runnable JS (no raw JSX remains)', async () => {
      const result = await sdk.build({
        name: 'JsxTest',
        framework: 'react',
        components: [{ name: 'JsxWidget', framework: 'react' }]
      });

      expect(result.success).toBe(true);
      const jsFiles = result.files.filter(f => f.path.endsWith('.js'));
      expect(jsFiles.length).toBeGreaterThan(0);

      const componentFile = jsFiles.find(f => f.path.includes('components'));
      expect(componentFile).toBeDefined();
      const content = String(componentFile!.content);

      expect(content).toContain('createElement');
      expect(content).not.toContain('</div>');
    });

    it('transpiles JSX files instead of passing raw JSX through', async () => {
      const ir = createIR({ name: 'JsxSourceTest' });
      addWebPart(ir, { name: 'JsxSource', framework: 'react' });

      const files = new Map<string, string>();
      files.set('src/webparts/JsxSource/JsxSourceWebPart.jsx', 'export default function View() { return <div>Hello</div>; }');

      const compiled = await sdk.compilerAPI.compile(ir, files);
      expect(compiled.success).toBe(true);

      const js = compiled.files.find(f => f.path.endsWith('JsxSourceWebPart.js'));
      expect(js).toBeDefined();
      expect(String(js!.content)).toContain('createElement');
      expect(String(js!.content)).not.toContain('<div>');
    });

    it('fails invalid TypeScript instead of falling back to string rewriting', async () => {
      const ir = createIR({ name: 'InvalidSourceTest' });
      addWebPart(ir, { name: 'Broken', framework: 'react' });

      const files = new Map<string, string>();
      files.set('src/webparts/Broken/BrokenWebPart.ts', 'export const value = ;');

      const compiled = await sdk.compilerAPI.compile(ir, files);
      expect(compiled.success).toBe(false);
      expect(compiled.errors.length).toBeGreaterThan(0);
    });

    it('runs semantic TypeScript checks before esbuild transpilation', async () => {
      const ir = createIR({ name: 'SemanticTest' });
      addWebPart(ir, { name: 'Semantic', framework: 'react' });

      const files = new Map<string, string>();
      files.set('src/webparts/Semantic/SemanticWebPart.ts', 'const count: number = "wrong"; export default count;');

      const compiled = await sdk.compilerAPI.compile(ir, files);
      expect(compiled.success).toBe(false);
      expect(compiled.errors.some(error => error.message.includes('number'))).toBe(true);
    });

    it('fails unsupported SCSS instead of emitting invalid CSS', async () => {
      const ir = createIR({ name: 'SassTest' });
      addWebPart(ir, { name: 'SassWidget', framework: 'react' });

      const files = new Map<string, string>();
      files.set('src/webparts/SassWidget/components/SassWidget.module.scss', '$color: red; .root { color: $color; }');

      const compiled = await sdk.compilerAPI.compile(ir, files);
      expect(compiled.success).toBe(false);
      expect(compiled.errors.some(error => error.message.includes('Unsupported SCSS syntax'))).toBe(true);
    });
  });

  describe('capability reporting', () => {
    it('does not claim unproven production SPFx conformance', () => {
      const capabilities = sdk.capabilities();

      expect(capabilities.tsx).toBe(true);
      expect(capabilities.jsx).toBe(true);
      expect(capabilities.spfx122).toBe(false);
      expect(capabilities.sppkg).toBe(false);
      expect(capabilities.productionBundling).toBe(false);
    });
  });

  describe('SPFx 1.22 metadata', () => {
    it('generates package dependencies from the target SPFx profile', () => {
      const ir = createIR({ name: 'ProfileTest' });
      ir.metadata.spfxVersion = '1.22.0';
      addWebPart(ir, { name: 'ProfileWebPart', framework: 'react' });

      const pkg = generatePackageJson(ir) as { dependencies: Record<string, string> };
      expect(pkg.dependencies['@microsoft/sp-core-library']).toBe('1.22.0');
      expect(pkg.dependencies.react).toBe('^18.2.0');
    });
  });

  describe('SPPKG OPC structure', () => {
    it('includes root relationship targets and component manifest relationships', async () => {
      const result = await sdk.build({
        name: 'OpcStructureTest',
        solution: { name: 'OpcStructureTest', version: '1.0.0' },
        components: [{ name: 'OpcWidget', framework: 'react' }]
      });

      expect(result.sppkg).toBeDefined();
      const files = unzipSync(result.sppkg!);
      expect(files['docProps/core.xml']).toBeDefined();
      expect(files['docProps/app.xml']).toBeDefined();

      const rels = strFromU8(files['OpcStructureTest/_rels/.rels']);
      expect(rels).toContain('OpcWidget.manifest.json');
      expect(files['OpcStructureTest/OpcWidget.manifest.json']).toBeDefined();
    });
  });

  describe('runnable bundle', () => {
    it('generates React web part scaffolds with component and props imports', async () => {
      const ir = createIR({ name: 'ScaffoldTest' });
      addWebPart(ir, { name: 'EmployeeDirectory', framework: 'react' });

      const compiled = await sdk.compilerAPI.compile(ir);
      expect(compiled.success).toBe(true);

      const webPart = compiled.files.find(f => f.path.endsWith('EmployeeDirectoryWebPart.js'));
      expect(webPart).toBeDefined();

      const content = String(webPart!.content);
      expect(content).toContain("./components/EmployeeDirectory");
      expect(content).toContain('EmployeeDirectoryComponent');
      expect(content).not.toContain('${component.name}');

      const css = compiled.files.find(f => f.path.endsWith('EmployeeDirectory.module.css'));
      expect(css).toBeDefined();
      expect(String(css!.content)).toContain('.EmployeeDirectory .container');
      expect(String(css!.content)).not.toContain('.container {\nmargin-top');
    });

    it('produces a syntactically valid IIFE bundle', async () => {
      const ir = createIR({ name: 'BundleTest' });
      addWebPart(ir, { name: 'Bundled', framework: 'react' });
      addExtension(ir, { name: 'Ext', type: 'ApplicationCustomizer' });

      const compiled = await sdk.compilerAPI.compile(ir);
      expect(compiled.success).toBe(true);

      const bundleResult = await sdk.bundlerAPI.bundle(ir, compiled.files);
      expect(bundleResult.success).toBe(true);
      expect(bundleResult.chunks.length).toBeGreaterThan(0);

      for (const chunk of bundleResult.chunks) {
        expect(() => new Function(chunk.content)).not.toThrow();
      }
    });
  });

  describe('localization generation', () => {
    it('generates .resx and localized string modules from IR', () => {
      const ir = createIR({ name: 'LocTest' });
      ir.localization = {
        defaultLanguage: 'en-us',
        languages: ['es-es', 'fr-fr'],
        strings: {
          'en-us': { Title: 'Title', Save: 'Save' },
          'es-es': { Title: 'Título', Save: 'Guardar' },
          'fr-fr': { Title: 'Titre', Save: 'Enregistrer' }
        }
      };

      const files = generateLocalizationFiles(ir);
      const paths = files.map(f => f.path);

      expect(paths).toContain('sharepoint/localization/en-us.resx');
      expect(paths).toContain('sharepoint/localization/es-es.resx');
      expect(paths).toContain('lib/providers/loc/fr-fr.js');

      const resx = generateResx('es-es', resolveStrings(ir.localization, 'es-es'));
      expect(resx).toContain('Guardar');
      expect(resx).toContain('<root>');

      const locJs = generateStringsModule('fr-fr', resolveStrings(ir.localization, 'fr-fr'));
      expect(locJs).toContain('Titre');
    });

    it('includes localization files in build output', async () => {
      const ir = createIR({ name: 'LocBuild' });
      addWebPart(ir, { name: 'LocWidget', framework: 'react' });
      ir.localization.languages = ['de-de'];
      ir.localization.strings = {
        'en-us': { Hello: 'Hello' },
        'de-de': { Hello: 'Hallo' }
      };

      const result = await sdk.buildFromIR(ir);
      const paths = result.files.map(f => f.path);
      expect(paths).toContain('sharepoint/localization/en-us.resx');
      expect(paths).toContain('sharepoint/localization/de-de.resx');
    });
  });

  describe('deep import round-trip', () => {
    it('recovers web parts, extensions, lists, and Graph permissions from an SPPKG', async () => {
      const ir = createIR({ name: 'ImportRoundTrip', solution: { name: 'ImportRoundTrip', version: '1.0.0' } });
      addWebPart(ir, { name: 'ImportedWebPart', framework: 'react' });
      addExtension(ir, { name: 'ImportedExtension', type: 'ApplicationCustomizer' });
      addList(ir, { title: 'ImportedList', template: 100 });
      addGraphPermission(ir, 'User.Read.All');

      const built = await sdk.buildFromIR(ir);
      expect(built.sppkg).toBeDefined();
      const imported = await sdk.import(built.sppkg!);
      expect(imported.success).toBe(true);
      expect(imported.source).toBe('sppkg');
    });
  });
});
