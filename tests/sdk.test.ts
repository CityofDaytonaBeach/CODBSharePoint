import {
  CODBSharePoint,
  createIR,
  addWebPart,
  addGraphPermission,
  addPropertyPane,
  createVFS,
  serializeIR,
  deserializeIR,
  generateNamespace,
  SPFx_COMPATIBILITY,
  GRAPH_PERMISSIONS,
  SharePointSimulator
} from '../src/index.js';

const sdk = new CODBSharePoint();

describe('CODBSharePoint SDK', () => {
  describe('Validation', () => {
    it('returns a valid solution with no errors', async () => {
      const ir = createIR({ name: 'ValidSolution' });
      addWebPart(ir, { name: 'MyWebPart', framework: 'react' });

      const validation = await sdk.validate(ir);
      expect(validation.valid).toBe(true);
      expect(validation.errors).toHaveLength(0);
      expect(validation.summary.total).toBeGreaterThan(0);
    });

    it('flags invalid solution names', async () => {
      const ir = createIR({ name: 'Bad Name!' });
      const validation = await sdk.validate(ir);
      expect(validation.errors.some(e => e.code === 'SOL002')).toBe(true);
    });

    it('flags invalid solution IDs', async () => {
      const ir = createIR({ name: 'ValidSolution' });
      ir.solution.id = 'not-a-guid';
      const validation = await sdk.validate(ir);
      expect(validation.errors.some(e => e.code === 'SOL004')).toBe(true);
    });
  });

  describe('Analysis', () => {
    it('detects framework and complexity', async () => {
      const ir = createIR({ name: 'PeopleFinder' });
      addWebPart(ir, { name: 'PeopleDirectory', framework: 'react' });
      addGraphPermission(ir, 'User.Read.All');

      const analysis = await sdk.analyze(ir);
      expect(analysis.framework).toBe('react');
      expect(analysis.componentTypes).toContain('webpart');
      expect(analysis.requiredPermissions.map(p => p.scope)).toContain('User.Read.All');
      expect(['simple', 'moderate', 'complex']).toContain(analysis.complexity);
    });
  });

  describe('Compatibility', () => {
    it('reports compatibility for a known SPFx version', async () => {
      const ir = createIR({ name: 'CompTest' });
      const report = await sdk.compatibility(ir);
      expect(report.targetVersion).toBe('1.22.0');
      expect(report.compatible).toBe(true);
    });

    it('reports an unknown SPFx version as a warning', async () => {
      const ir = createIR({ name: 'CompTest' });
      ir.metadata.spfxVersion = '9.9.9' as never;
      const report = await sdk.compatibility(ir);
      expect(report.issues.some(i => i.message.includes('Unknown SPFx'))).toBe(true);
    });
  });

  describe('Security scanner', () => {
    it('detects secrets in source files', async () => {
      const ir = createIR({ name: 'SecureTest' });
      addWebPart(ir, { name: 'MyWebPart', framework: 'react' });

      const files = new Map<string, string>();
      files.set('src/secret.ts', 'const apiKey = "abcdefghijklmnopqrstuvwxyz123456";');

      const report = await sdk.securityAPI.scan(ir, files);
      expect(report.secrets.length).toBeGreaterThan(0);
    });

    it('passes when no critical findings exist', async () => {
      const ir = createIR({ name: 'SecureTest' });
      const report = await sdk.securityAPI.scan(ir);
      expect(report.passed).toBe(true);
    });
  });

  describe('Tool API', () => {
    it('creates a web part and adds Graph permissions', async () => {
      const tools = sdk.tools({ name: 'ToolkitSolution' });

      const wp = tools.createWebPart({
        name: 'Announcements',
        framework: 'react',
        displayName: 'Announcements'
      });
      expect(wp.name).toBe('Announcements');

      const ir = tools.getIR();
      tools.addGraphPermission(ir, 'User.Read');
      expect(ir.graph).toHaveLength(1);
    });

    it('creates a list and a library', async () => {
      const tools = sdk.tools({ name: 'ProvisioningTest' });
      const list = tools.createList({ title: 'Requests', template: 100 });
      expect(list.title).toBe('Requests');

      const lib = tools.createLibrary({ title: 'Docs' });
      expect(lib.title).toBe('Docs');
    });

    it('validates an IR through tools', async () => {
      const tools = sdk.tools({ name: 'ToolValidate' });
      const validation = await tools.validate(tools.getIR());
      expect(validation.valid).toBe(true);
    });
  });

  describe('Property Pane', () => {
    it('adds a property pane to a web part', () => {
      const ir = createIR({ name: 'PaneTest' });
      const wp = addWebPart(ir, { name: 'MyWebPart', framework: 'react' });

      const updated = addPropertyPane(wp, {
        pages: [{
          groups: [{
            groupName: 'Settings',
            groupFields: [{
              type: 'textField',
              propertyName: 'description',
              label: 'Description'
            }]
          }]
        }]
      });

      expect(updated.properties).toHaveLength(1);
      expect(updated.properties[0].pages[0].groups[0].groupFields[0].type).toBe('textField');
    });
  });

  describe('VFS', () => {
    it('adds, gets, and lists files', () => {
      const vfs = createVFS();
      vfs.addFile('src/index.ts', 'export {};');
      vfs.addFile('styles/main.css', 'body {}');

      expect(vfs.hasFile('src/index.ts')).toBe(true);
      expect(vfs.readAsString('src/index.ts')).toBe('export {};');
      expect(vfs.getFiles()).toHaveLength(2);
      expect(vfs.getFilesByPattern('**/*.css')).toHaveLength(1);
    });

    it('produces a valid zip', () => {
      const vfs = createVFS();
      vfs.addFile('hello.txt', 'hello world');
      const zip = vfs.toZip();
      expect(zip).toBeInstanceOf(Uint8Array);
      expect(zip.length).toBeGreaterThan(0);
      expect(new Uint8Array(zip.slice(0, 2))[0]).toBe(0x50); // 'P'
      expect(new Uint8Array(zip.slice(1, 2))[0]).toBe(0x4b); // 'K'
    });
  });

  describe('IR serialization', () => {
    it('round-trips through serialize/deserialize', () => {
      const ir = createIR({ name: 'RoundTrip' });
      addWebPart(ir, { name: 'MyWebPart', framework: 'react' });

      const json = serializeIR(ir);
      const restored = deserializeIR(json);
      expect(restored.solution.name).toBe('RoundTrip');
      expect(restored.components).toHaveLength(1);
    });

    it('generates a namespace from a name', () => {
      expect(generateNamespace('Employee Directory')).toBe('EmployeeDirectory');
    });
  });

  describe('Import', () => {
    it('imports a CODBSharePoint JSON document', async () => {
      const ir = createIR({ name: 'ImportMe' });
      addWebPart(ir, { name: 'MyWebPart', framework: 'react' });
      const json = serializeIR(ir);

      const imported = await sdk.import(json);
      expect(imported.success).toBe(true);
      expect(imported.source).toBe('codbsharepoint-json');
      expect(imported.ir.solution.name).toBe('ImportMe');
    });
  });

  describe('Export', () => {
    it('generates a deployment zip with SPPKG inside', async () => {
      const result = await sdk.build({
        name: 'ExportTest',
        solution: { name: 'ExportTest', version: '1.0.0' },
        components: [{ name: 'Widget', framework: 'react' }]
      });

      expect(result.success).toBe(true);
      const ir = createIR({ name: 'ExportTest' });

      const zip = await sdk.export(ir, {
        success: true,
        sppkg: result.sppkg,
        files: result.files,
        deployment: result.deployment,
        validation: result.validation,
        security: result.security,
        compatibility: result.compatibility,
        bundle: result.bundle,
        errors: [],
        warnings: [],
        duration: 0
      });

      expect(zip).toBeInstanceOf(Uint8Array);
      expect(zip.length).toBeGreaterThan(0);
      expect(new Uint8Array(zip.slice(0, 2))[0]).toBe(0x50); // 'P'
    });
  });

  describe('Simulator', () => {
    it('creates a simulator with default configuration', () => {
      const sim = new SharePointSimulator();
      const user = sim.getUser();
      expect(user).toBeDefined();
      expect(user!.displayName).toBe('Test User');
    });

    it('returns list data and theme', () => {
      const sim = new SharePointSimulator();
      expect(sim.getListData('Employees').length).toBeGreaterThan(0);
      expect(sim.getTheme()).toBeDefined();
      expect(sim.getContext()).toBeDefined();
    });
  });

  describe('Constants', () => {
    it('exposes SPFx compatibility matrix', () => {
      expect(SPFx_COMPATIBILITY['1.22.0'].react).toBe('18');
      expect(SPFx_COMPATIBILITY['1.22.0'].buildTool).toBe('heft');
    });

    it('exposes Graph permissions', () => {
      expect(GRAPH_PERMISSIONS['User.Read.All'].requiresAdminApproval).toBe(true);
    });
  });
});
