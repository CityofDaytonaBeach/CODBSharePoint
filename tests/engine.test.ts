import {
  CODBSharePoint,
  createIR,
  addWebPart,
  addExtension,
  addList,
  addGraphPermission,
  generateLocalizationFiles,
  generateResx,
  generateStringsModule,
  resolveStrings
} from '../src/index.js';

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
  });

  describe('runnable bundle', () => {
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
