import {
  CODBSharePoint,
  createIR,
  addTheme,
  addFormatting,
  addProvisioning,
  addPage,
  addWebPart,
  generateThemeJson,
  generateFormattingJson,
  generateSiteScript,
  generateProvisioningJson,
  generatePageJson,
  generateSharePointArtifacts
} from '../src/index.js';

const sdk = new CODBSharePoint();

describe('CODBSharePoint - Themes, Formatting, Site Designs, Provisioning, Pages', () => {
  describe('IR-level creation', () => {
    it('creates a theme', () => {
      const ir = createIR({ name: 'Branding' });
      addTheme(ir, {
        name: 'CityBlue',
        primary: { themePrimary: '#0078d4', themeDarker: '#004578' }
      });

      expect(ir.themes).toHaveLength(1);
      expect(ir.themes[0].name).toBe('CityBlue');
      expect(ir.themes[0].primary.themePrimary).toBe('#0078d4');
    });

    it('creates column formatting and list formatting', () => {
      const ir = createIR({ name: 'Formatting' });
      addFormatting(ir, {
        type: 'columnFormatting',
        name: 'StatusColumn',
        target: 'Status',
        json: { elmType: 'div', txtContent: '@currentField' }
      });
      addFormatting(ir, {
        type: 'listFormatting',
        name: 'TaskList',
        target: 'Tasks',
        json: { elmType: 'div', children: [] }
      });

      expect(ir.formatting).toHaveLength(2);
      expect(ir.formatting[0].type).toBe('columnFormatting');
      expect(ir.formatting[1].type).toBe('listFormatting');
    });

    it('creates a site design and a site script', () => {
      const ir = createIR({ name: 'Provisioning' });
      addProvisioning(ir, {
        type: 'siteScript',
        name: 'CreateSite',
        data: { actions: [{ verb: 'applyTheme' }] }
      });
      addProvisioning(ir, {
        type: 'siteDesign',
        name: 'ProjectHub',
        description: 'Project hub site design',
        data: { siteScriptIds: ['11111111-1111-1111-1111-111111111111'] }
      });

      expect(ir.provisioning).toHaveLength(2);
      expect(ir.provisioning[0].type).toBe('siteScript');
      expect(ir.provisioning[1].type).toBe('siteDesign');
    });

    it('creates a page', () => {
      const ir = createIR({ name: 'Pages' });
      addPage(ir, {
        name: 'home',
        title: 'Home',
        content: '<h1>Welcome</h1>'
      });

      expect(ir.pages).toHaveLength(1);
      expect(ir.pages[0].name).toBe('home');
    });
  });

  describe('JSON generators', () => {
    it('generates a theme JSON with palette merged at top level', () => {
      const json = generateThemeJson({
        name: 'CityBlue',
        primary: { themePrimary: '#0078d4' }
      });
      expect(json.name).toBe('CityBlue');
      expect(json.themePrimary).toBe('#0078d4');
    });

    it('generates column formatting JSON with the correct schema', () => {
      const json = generateFormattingJson({
        type: 'columnFormatting',
        name: 'StatusColumn',
        target: 'Status',
        json: { elmType: 'div' }
      });
      expect(json.$schema).toContain('column-formatting');
      expect(json.elmType).toBe('div');
    });

    it('generates a site script JSON', () => {
      const json = generateSiteScript({
        type: 'siteScript',
        name: 'CreateSite',
        data: { actions: [{ verb: 'applyTheme' }] }
      });
      expect(json.$schema).toContain('site-design-script-actions');
      expect((json.actions as unknown[]).length).toBe(1);
    });

    it('generates a site design JSON', () => {
      const json = generateProvisioningJson({
        type: 'siteDesign',
        name: 'ProjectHub',
        data: { siteScriptIds: ['a'] }
      });
      expect(json.title).toBe('ProjectHub');
      expect(json.siteScriptIds).toEqual(['a']);
    });

    it('generates a page JSON', () => {
      const json = generatePageJson({
        name: 'home',
        title: 'Home',
        content: '<h1>Welcome</h1>'
      });
      expect(json.name).toBe('home');
      expect(json.promotedState).toBe(0);
    });
  });

  describe('SharePoint artifact file generation', () => {
    it('generates files for all artifact types', () => {
      const ir = createIR({ name: 'AllArtifacts' });
      addTheme(ir, { name: 'CityBlue', primary: { themePrimary: '#0078d4' } });
      addFormatting(ir, {
        type: 'columnFormatting', name: 'StatusColumn', target: 'Status',
        json: { elmType: 'div' }
      });
      addProvisioning(ir, {
        type: 'siteDesign', name: 'ProjectHub', data: { siteScriptIds: [] }
      });
      addPage(ir, { name: 'home', title: 'Home', content: '' });

      const files = generateSharePointArtifacts(ir);
      const paths = files.map(f => f.path);

      expect(paths).toContain('sharepoint/themes/CityBlue.json');
      expect(paths).toContain('sharepoint/formatting/StatusColumn.json');
      expect(paths).toContain('sharepoint/site-designs/ProjectHub.json');
      expect(paths).toContain('sharepoint/pages/home.json');
    });
  });

  describe('End-to-end build integration', () => {
    it('includes artifact files and provisioning in the build output', async () => {
      const ir = createIR({ name: 'ArtifactSolution' });
      addWebPart(ir, { name: 'MyWebPart', framework: 'react' });
      addTheme(ir, { name: 'CityBlue', primary: { themePrimary: '#0078d4' } });
      addFormatting(ir, {
        type: 'columnFormatting', name: 'StatusColumn', target: 'Status',
        json: { elmType: 'div' }
      });
      addProvisioning(ir, {
        type: 'siteDesign', name: 'ProjectHub', data: { siteScriptIds: [] }
      });
      addPage(ir, { name: 'home', title: 'Home', content: '' });

      const result = await sdk.buildFromIR(ir);

      expect(result.success).toBe(true);
      const filePaths = result.files.map(f => f.path);
      expect(filePaths).toContain('sharepoint/themes/CityBlue.json');
      expect(filePaths).toContain('sharepoint/formatting/StatusColumn.json');
      expect(filePaths).toContain('sharepoint/site-designs/ProjectHub.json');
      expect(filePaths).toContain('sharepoint/pages/home.json');

      const provTypes = result.deployment.provisioning.map(p => p.type);
      expect(provTypes).toContain('theme');
      expect(provTypes).toContain('columnFormatting');
      expect(provTypes).toContain('siteDesign');
      expect(provTypes).toContain('page');
    });

    it('exposes the provisioning API on the SDK', () => {
      expect(typeof sdk.provisioningAPI.generateSharePointArtifacts).toBe('function');
      expect(typeof sdk.provisioningAPI.generateThemeJson).toBe('function');
      expect(typeof sdk.provisioningAPI.generateFormattingJson).toBe('function');
    });
  });
});
