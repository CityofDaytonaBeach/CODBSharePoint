import { CODBSharePoint, DesignerRulesEngine, BUILT_IN_RULES, Designer } from '../src/index.js';
import { createIR, addWebPart, addList, addGraphPermission } from '../src/core/ir.js';

describe('CODBSharePoint - Designer Rules Engine', () => {
  it('has 8 built-in rules', () => {
    expect(BUILT_IN_RULES.length).toBe(8);
    expect(BUILT_IN_RULES.map(r => r.id)).toContain('employee-directory');
    expect(BUILT_IN_RULES.map(r => r.id)).toContain('quick-links');
    expect(BUILT_IN_RULES.map(r => r.id)).toContain('data-table');
    expect(BUILT_IN_RULES.map(r => r.id)).toContain('faq-accordion');
    expect(BUILT_IN_RULES.map(r => r.id)).toContain('calendar-view');
    expect(BUILT_IN_RULES.map(r => r.id)).toContain('org-chart');
    expect(BUILT_IN_RULES.map(r => r.id)).toContain('document-explorer');
    expect(BUILT_IN_RULES.map(r => r.id)).toContain('approval-dashboard');
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
    expect(designer.listRules().length).toBe(8);
    expect(designer.searchRules('graph').length).toBeGreaterThanOrEqual(2);
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
    expect(sdk.designerAPI.BUILT_IN_RULES.length).toBe(8);
  });
});
