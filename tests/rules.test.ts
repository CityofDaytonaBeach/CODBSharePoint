import { CODBSharePoint, DesignerRulesEngine, BUILT_IN_RULES, Designer } from '../src/index.js';
import { createIR, addWebPart, addList, addGraphPermission } from '../src/core/ir.js';

describe('CODBSharePoint - Designer Rules Engine', () => {
  it('has 21 built-in rules', () => {
    expect(BUILT_IN_RULES.length).toBeGreaterThanOrEqual(20);
    expect(BUILT_IN_RULES.map(r => r.id)).toContain('employee-directory');
    expect(BUILT_IN_RULES.map(r => r.id)).toContain('quick-links');
    expect(BUILT_IN_RULES.map(r => r.id)).toContain('data-table');
    expect(BUILT_IN_RULES.map(r => r.id)).toContain('faq-accordion');
    expect(BUILT_IN_RULES.map(r => r.id)).toContain('calendar-view');
    expect(BUILT_IN_RULES.map(r => r.id)).toContain('org-chart');
    expect(BUILT_IN_RULES.map(r => r.id)).toContain('document-explorer');
    expect(BUILT_IN_RULES.map(r => r.id)).toContain('approval-dashboard');
    expect(BUILT_IN_RULES.map(r => r.id)).toContain('ace-announcements');
    expect(BUILT_IN_RULES.map(r => r.id)).toContain('ace-task-card');
    expect(BUILT_IN_RULES.map(r => r.id)).toContain('command-set-approvals');
    expect(BUILT_IN_RULES.map(r => r.id)).toContain('field-customizer-status');
    expect(BUILT_IN_RULES.map(r => r.id)).toContain('news-feed');
    expect(BUILT_IN_RULES.map(r => r.id)).toContain('task-board');
    expect(BUILT_IN_RULES.map(r => r.id)).toContain('mailbox-viewer');
    expect(BUILT_IN_RULES.map(r => r.id)).toContain('file-manager');
    expect(BUILT_IN_RULES.map(r => r.id)).toContain('event-registration');
    expect(BUILT_IN_RULES.map(r => r.id)).toContain('ticketing-system');
    expect(BUILT_IN_RULES.map(r => r.id)).toContain('leave-request');
    expect(BUILT_IN_RULES.map(r => r.id)).toContain('image-gallery');
    expect(BUILT_IN_RULES.map(r => r.id)).toContain('inventory-tracker');
  });

  it('searches rules by name and tags', () => {
    const engine = new DesignerRulesEngine();
    const graphRules = engine.search('graph');
    expect(graphRules.length).toBeGreaterThanOrEqual(2);
    expect(graphRules.some(r => r.id === 'employee-directory')).toBe(true);

    const approvalRules = engine.search('approval');
    expect(approvalRules.some(r => r.id === 'approval-dashboard')).toBe(true);
  });

  it('registers and retrieves custom rules', () => {
    const engine = new DesignerRulesEngine([]);
    expect(engine.list().length).toBe(0);

    const customRule: any = {
      id: 'custom-test',
      name: 'Custom Test',
      description: 'A test rule',
      version: '1.0.0',
      author: 'test',
      tags: ['test'],
      components: [{ name: 'TestPart', displayName: 'Test', type: 'webpart', framework: 'react' }],
      dataSources: []
    };

    engine.register(customRule);
    expect(engine.list().length).toBe(1);
    expect(engine.get('custom-test')?.name).toBe('Custom Test');
    expect(engine.remove('custom-test')).toBe(true);
    expect(engine.list().length).toBe(0);
  });

  it('validates rule structure', () => {
    const engine = new DesignerRulesEngine();
    const validRule = engine.get('employee-directory')!;
    const result = engine.validateRule(validRule);
    expect(result.valid).toBe(true);
    expect(result.errors.length).toBe(0);
  });

  it('detects invalid rules', () => {
    const engine = new DesignerRulesEngine();
    const invalidRule: any = {
      id: '',
      name: '',
      description: '',
      version: '1.0.0',
      author: 'test',
      tags: [],
      components: [],
      dataSources: []
    };
    const result = engine.validateRule(invalidRule);
    expect(result.valid).toBe(false);
    expect(result.errors.length).toBeGreaterThan(0);
  });

  it('creates project from rule with IR, permissions, and lists', () => {
    const engine = new DesignerRulesEngine();
    const { ir, files, permissions } = engine.createProject(
      engine.get('employee-directory')!,
      { name: 'MyDirectory' }
    );

    expect(ir.solution.name).toBe('MyDirectory');
    expect(ir.components.length).toBe(1);
    expect(ir.components[0].name).toBe('EmployeeDirectory');
    expect(permissions).toContain('User.Read.All');
    expect(ir.graph.some(p => p.scope === 'User.Read.All')).toBe(true);
  });

  it('creates project from rule with lists and fields', () => {
    const engine = new DesignerRulesEngine();
    const { ir } = engine.createProject(engine.get('approval-dashboard')!);

    expect(ir.lists.length).toBe(1);
    expect(ir.lists[0].title).toBe('Approvals');
    expect(ir.lists[0].fields!.length).toBe(7);
    expect(ir.lists[0].fields!.includes('Status')).toBe(true);
  });

  it('exports rule from IR and assets', () => {
    const engine = new DesignerRulesEngine();
    const ir = createIR({ name: 'TestExport' });
    addWebPart(ir, { name: 'ExportPart', framework: 'react' });
    addGraphPermission(ir, 'User.Read.All');

    const assets: Record<string, string> = {
      'src/webparts/ExportPart/ExportPartWebPart.ts': `import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { PropertyPaneTextField } from '@microsoft/sp-property-pane';
export default class ExportPartWebPart extends BaseClientSideWebPart<any> {
  protected getPropertyPaneConfiguration() {
    return { pages: [{ groups: [{ groupFields: [PropertyPaneTextField('title', { label: 'Title' })] }] }] };
  }
}`
    };

    const rule = engine.exportFromProject(ir, assets, { name: 'MyExport', author: 'test' });

    expect(rule.id).toContain('exported-myexport');
    expect(rule.name).toBe('MyExport');
    expect(rule.components.length).toBe(1);
    expect(rule.components[0].propertyPane!.length).toBe(1);
    expect(rule.components[0].propertyPane![0].name).toBe('title');
    expect(rule.graphPermissions).toContain('User.Read.All');
    expect(rule.createdAt).toBeDefined();
  });

  it('infers Graph permissions from data sources', () => {
    const engine = new DesignerRulesEngine();
    const rule = engine.get('employee-directory')!;
    const perms = engine.inferPermissions(rule);
    expect(perms).toContain('User.Read.All');
  });

  it('validates rule against knowledge catalog', () => {
    const engine = new DesignerRulesEngine();
    const rule = engine.get('employee-directory')!;
    const validation = engine.validateRule(rule);
    expect(validation.valid).toBe(true);
    expect(validation.warnings.length).toBe(0);
  });
});

describe('CODBSharePoint - Designer + Rules Integration', () => {
  it('designer exposes rules engine', () => {
    const sdk = new CODBSharePoint();
    const designer = sdk.designer();

    expect(designer.rules).toBeDefined();
    expect(designer.listRules().length).toBeGreaterThanOrEqual(20);
    expect(designer.searchRules('graph').length).toBeGreaterThanOrEqual(2);
    expect(designer.searchRules('ace').length).toBeGreaterThanOrEqual(1);
    expect(designer.searchRules('commandset').length).toBeGreaterThanOrEqual(1);
  });

  it('supports rule inheritance', () => {
    const engine = new DesignerRulesEngine();
    const base = engine.get('employee-directory')!;

    const derived = engine.inherit('employee-directory', {
      id: 'my-custom-directory',
      name: 'My Custom Directory',
      tags: ['custom']
    });

    expect(derived.basedOn).toBe('employee-directory');
    expect(derived.id).toBe('my-custom-directory');
    expect(derived.version).not.toBe(base.version);
    expect(derived.graphPermissions).toContain('User.Read.All');
    expect(derived.tags).toContain('custom');
    expect(derived.tags).toContain('graph');
  });

  it('inheritance deep-merges components and data sources', () => {
    const engine = new DesignerRulesEngine();
    const derived = engine.inherit('employee-directory', {
      components: [{
        name: 'ExtraPart',
        displayName: 'Extra',
        type: 'webpart',
        framework: 'react'
      }],
      dataSources: [{
        type: 'graph',
        name: 'extra-graph',
        endpoint: '/me'
      }]
    });

    expect(derived.components.length).toBe(2);
    expect(derived.dataSources.length).toBe(2);
  });

  it('diffs two rules to find changes', () => {
    const engine = new DesignerRulesEngine();
    const derived = engine.inherit('employee-directory', {
      id: 'diff-target',
      name: 'Diff Target',
      graphPermissions: ['User.Read.All', 'Mail.Read']
    });
    engine.register(derived);

    const diff = engine.diff('employee-directory', 'diff-target');

    expect(diff.identical).toBe(false);
    expect(diff.changes.some(c => c.path === 'name')).toBe(true);
    expect(diff.changes.some(c => c.path === 'graphPermissions.Mail.Read' && c.type === 'added')).toBe(true);
  });

  it('reports identical rules as identical', () => {
    const engine = new DesignerRulesEngine();
    engine.register(engine.get('quick-links')!);
    const diff = engine.diff('quick-links', 'quick-links');
    expect(diff.identical).toBe(true);
    expect(diff.changes.length).toBe(0);
  });

  it('creates ACE-type project from rule', () => {
    const engine = new DesignerRulesEngine();
    const { ir } = engine.createProject(engine.get('ace-announcements')!);
    expect(ir.components.length).toBe(1);
    expect(ir.components[0].type).toBe('ace');
  });

  it('designer creates project from rule', async () => {
    const sdk = new CODBSharePoint();
    const designer = sdk.designer();

    const rule = await designer.createFromRule('employee-directory', { name: 'MyDir' });
    expect(rule).toBeDefined();
    expect(rule!.name).toBe('Employee Directory');

    const ir = designer.getIR();
    expect(ir.solution.name).toBe('MyDir');
    expect(ir.components.length).toBe(1);
    expect(ir.graph.some(p => p.scope === 'User.Read.All')).toBe(true);
  });

  it('designer exports current project as rule', async () => {
    const sdk = new CODBSharePoint();
    const designer = sdk.designer();

    await designer.create('MyProject');
    designer.addWebPart({ name: 'MyPart', displayName: 'My Part' });
    designer.addGraphPermission('User.Read.All');

    const exported = designer.exportAsRule({ name: 'MyPartRule', author: 'test' });
    expect(exported.name).toBe('MyPartRule');
    expect(exported.components.length).toBe(1);
    expect(exported.graphPermissions).toContain('User.Read.All');
  });

  it('designer validates project against rule', async () => {
    const sdk = new CODBSharePoint();
    const designer = sdk.designer();

    const validation = designer.validateAgainstRule('employee-directory');
    expect(validation).toHaveProperty('valid');
    expect(validation).toHaveProperty('errors');
    expect(validation).toHaveProperty('warnings');
  });

  it('designer validates against non-existent rule', async () => {
    const sdk = new CODBSharePoint();
    const designer = sdk.designer();

    const validation = designer.validateAgainstRule('non-existent');
    expect(validation.valid).toBe(false);
    expect(validation.errors).toContain('Rule "non-existent" not found');
  });

  it('sdk.designerAPI exposes rules engine', () => {
    const sdk = new CODBSharePoint();
    expect(sdk.designerAPI.DesignerRulesEngine).toBeDefined();
    expect(sdk.designerAPI.BUILT_IN_RULES.length).toBeGreaterThanOrEqual(20);
  });
});
