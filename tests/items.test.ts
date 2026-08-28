import {
  CODBSharePoint,
  createIR,
  addWebPart,
  addACE,
  addExtension,
  addList,
  addLibrary,
  addColumn,
  addContentType,
  addGraphPermission
} from '../src/index.js';

const sdk = new CODBSharePoint();

describe('CODBSharePoint - Build each SharePoint item type', () => {
  describe('Web Part', () => {
    it('builds a React web part solution successfully', async () => {
      const result = await sdk.build({
        name: 'EmployeeDirectory',
        framework: 'react',
        solution: {
          name: 'EmployeeDirectory',
          version: '1.0.0',
          description: 'Employee directory web part',
          author: 'City of Daytona Beach'
        },
        components: [
          {
            name: 'EmployeeDirectory',
            displayName: 'Employee Directory',
            description: 'Search and browse employees',
            framework: 'react'
          }
        ]
      });

      expect(result.success).toBe(true);
      expect(result.validation.valid).toBe(true);
      expect(result.validation.errors).toHaveLength(0);
      expect(result.sppkg).toBeDefined();
      expect(result.sppkg!.length).toBeGreaterThan(0);
      expect(result.files.length).toBeGreaterThan(0);
      expect(result.deployment.status).toBe('ready');
      expect(result.deployment.artifact).toBe('EmployeeDirectory.sppkg');
    });

    it('builds a vanilla JS web part', async () => {
      const result = await sdk.build({
        name: 'Announcements',
        framework: 'vanilla',
        components: [{ name: 'Announcements', framework: 'vanilla' }]
      });

      expect(result.success).toBe(true);
      expect(result.validation.valid).toBe(true);
    });
  });

  describe('Adaptive Card Extension (ACE)', () => {
    it('builds a solution with an ACE component', async () => {
      const result = await sdk.build({
        name: 'QuickLinksACE',
        solution: { name: 'QuickLinksACE', version: '1.0.0' }
      });

      // Build an IR with an ACE and validate explicitly
      const ir = createIR({ name: 'QuickLinksACE' });
      const ace = addACE(ir, {
        name: 'QuickLinks',
        description: 'Quick links card',
        iconProperty: 'QuickLinks',
        type: 'Primary'
      });

      expect(ace.id).toBeDefined();
      expect(ace.name).toBe('QuickLinks');
      expect(ir.components[0].type).toBe('ace');

      const validation = await sdk.validate(ir);
      expect(validation.errors).toHaveLength(0);
      expect(result.success).toBe(true);
    });
  });

  describe('Extensions', () => {
    it('builds an Application Customizer', async () => {
      const ir = createIR({ name: 'HeaderCustomizer' });
      const ext = addExtension(ir, {
        name: 'HeaderCustomizer',
        type: 'ApplicationCustomizer',
        displayName: 'Header Customizer',
        topNavigationZone: true
      });

      expect(ext.type).toBe('ApplicationCustomizer');
      expect(ext.clientSideComponentId).toBe(ext.id);
      expect(ir.extensions).toHaveLength(1);

      const result = await sdk.build({
        name: 'HeaderCustomizer',
        extensions: [{ name: 'HeaderCustomizer', type: 'ApplicationCustomizer' }]
      });
      expect(result.validation.valid).toBe(true);
    });

    it('builds a Field Customizer', async () => {
      const ir = createIR({ name: 'ScoreField' });
      const ext = addExtension(ir, {
        name: 'ScoreField',
        type: 'FieldCustomizer'
      });
      expect(ext.type).toBe('FieldCustomizer');

      const result = await sdk.build({
        name: 'ScoreField',
        extensions: [{ name: 'ScoreField', type: 'FieldCustomizer' }]
      });
      expect(result.validation.valid).toBe(true);
    });

    it('builds a ListView Command Set', async () => {
      const ir = createIR({ name: 'ApprovalsCommands' });
      const ext = addExtension(ir, {
        name: 'ApprovalsCommands',
        type: 'ListViewCommandSet'
      });
      expect(ext.type).toBe('ListViewCommandSet');

      const result = await sdk.build({
        name: 'ApprovalsCommands',
        extensions: [{ name: 'ApprovalsCommands', type: 'ListViewCommandSet' }]
      });
      expect(result.validation.valid).toBe(true);
    });

    it('builds a Form Customizer', async () => {
      const ir = createIR({ name: 'ProjectForm' });
      const ext = addExtension(ir, {
        name: 'ProjectForm',
        type: 'FormCustomizer'
      });
      expect(ext.type).toBe('FormCustomizer');
    });
  });

  describe('Provisioning items', () => {
    it('builds a List', async () => {
      const ir = createIR({ name: 'Requests' });
      const list = addList(ir, {
        title: 'Requests',
        description: 'Employee requests',
        template: 100
      });

      expect(list.title).toBe('Requests');
      expect(list.template).toBe(100);
      expect(ir.lists).toHaveLength(1);
    });

    it('builds a Document Library', async () => {
      const ir = createIR({ name: 'Contracts' });
      const lib = addLibrary(ir, {
        title: 'Contracts',
        versioningEnabled: true,
        majorVersionLimit: 5
      });

      expect(lib.title).toBe('Contracts');
      expect(lib.versioningEnabled).toBe(true);
      expect(ir.libraries).toHaveLength(1);
    });

    it('builds a Column (Field)', async () => {
      const ir = createIR({ name: 'EmployeeData' });
      const field = addColumn(ir, {
        name: 'Department',
        displayName: 'Department',
        type: 'Choice',
        choices: ['IT', 'HR', 'Finance']
      });

      expect(field.name).toBe('Department');
      expect(field.type).toBe('Choice');
      expect(field.choices).toEqual(['IT', 'HR', 'Finance']);
      expect(ir.fields).toHaveLength(1);
    });

    it('builds a Content Type', async () => {
      const ir = createIR({ name: 'CaseManagement' });
      const ct = addContentType(ir, {
        name: 'Case',
        description: 'Case management content type',
        group: 'Custom',
        fields: ['Department', 'Status']
      });

      expect(ct.name).toBe('Case');
      expect(ct.fields).toEqual(['Department', 'Status']);
      expect(ir.contentTypes).toHaveLength(1);
    });
  });

  describe('Graph permissions', () => {
    it('adds a known Graph permission', async () => {
      const ir = createIR({ name: 'PeopleFinder' });
      addGraphPermission(ir, 'User.Read.All');

      expect(ir.graph).toHaveLength(1);
      expect(ir.graph[0].scope).toBe('User.Read.All');
      expect(ir.graph[0].requiresAdminApproval).toBe(true);
    });

    it('does not duplicate an existing permission', async () => {
      const ir = createIR({ name: 'PeopleFinder' });
      addGraphPermission(ir, 'User.Read');
      addGraphPermission(ir, 'User.Read');

      expect(ir.graph).toHaveLength(1);
    });

    it('builds a full solution with Graph permissions flagged for admin approval', async () => {
      const result = await sdk.build({
        name: 'PeoplePortal',
        solution: { name: 'PeoplePortal', version: '1.0.0' },
        components: [{ name: 'PeopleDirectory', framework: 'react' }],
        graph: ['User.Read.All']
      });

      expect(result.deployment.requiresAdmin).toBe(true);
      expect(result.deployment.permissions.map(p => p.permission)).toContain('User.Read.All');
    });
  });
});
