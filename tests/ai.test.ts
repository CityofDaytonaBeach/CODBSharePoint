import { CODBSharePoint, type AIDiagnostic } from '../src/index.js';

const sdk = new CODBSharePoint();

describe('CODBSharePoint - AI contract layer', () => {
  it('generates deterministic project structure from AI intent and builds a package', async () => {
    const ai = sdk.ai();
    const project = ai.generate({
      schemaVersion: 'codbsharepoint.ai/1.0',
      intent: 'Build a service request dashboard with Graph people lookup',
      target: { spfxVersion: '1.22.0', runtime: 'browser', componentTypes: ['webpart'] },
      solution: { name: 'AIServiceRequests', version: '1.0.0', description: 'AI generated service request dashboard', author: 'AI' },
      components: [{ name: 'ServiceRequests', displayName: 'Service Requests', framework: 'react' }],
      data: {
        graphPermissions: ['User.Read.All'],
        fields: [{ name: 'Status', displayName: 'Status', type: 'Choice', choices: ['Open', 'Closed'] }],
        lists: [{ title: 'Requests', template: 100, fields: ['Status'] }]
      }
    });

    expect(project.schemaVersion).toBe('codbsharepoint.ai/1.0');
    expect(project.files.has('src/webparts/ServiceRequests/ServiceRequestsWebPart.ts')).toBe(true);
    expect(project.files.has('src/webparts/ServiceRequests/components/ServiceRequests.tsx')).toBe(true);
    expect(project.files.has('src/webparts/ServiceRequests/components/ServiceRequests.module.scss')).toBe(true);

    const result = await ai.build(project);
    expect(result.success).toBe(true);
    expect(result.aiDiagnostics).toEqual([]);
    expect(result.sppkg).toBeDefined();
    expect(result.bundle.chunks).toHaveLength(1);
    expect(result.deployment.permissions.map(permission => permission.permission)).toEqual(['User.Read.All']);
  });

  it('writes component source by role instead of requiring AI to guess file paths', () => {
    const ai = sdk.ai();
    const project = ai.generate({
      solution: { name: 'AIRoleWriter' },
      components: [{ name: 'RoleWriter', framework: 'react' }]
    });

    const path = ai.writeComponentSource({
      ir: project.ir,
      files: project.files,
      componentName: 'RoleWriter',
      fileRole: 'component',
      content: 'export default function RoleWriter() { return null; }'
    });

    expect(path).toBe('src/webparts/RoleWriter/components/RoleWriter.tsx');
    expect(project.files.get(path)).toContain('RoleWriter');
  });

  it('returns prompt context with capability and rule guidance for AI agents', () => {
    const ai = sdk.ai();
    const project = ai.generate({ solution: { name: 'AIContext' }, components: [{ name: 'ContextPart' }] });
    const context = JSON.parse(ai.exportPromptContext(project));

    expect(context.schemaVersion).toBe('codbsharepoint.ai/1.0');
    expect(context.capabilities.aiContract).toBe(true);
    expect(context.rules.join('\n')).toContain('SDK owns paths');
    expect(context.project.files).toContain('src/webparts/ContextPart/components/ContextPart.tsx');
  });

  it('normalizes build errors into machine-repairable diagnostics', async () => {
    const ai = sdk.ai();
    const project = ai.generate({ solution: { name: 'AIBroken' }, components: [{ name: 'BrokenPart' }] });
    project.files.set('src/webparts/BrokenPart/components/BrokenPart.tsx', 'export const value = ;');

    const result = await ai.build(project);

    expect(result.success).toBe(false);
    expect(result.aiDiagnostics.length).toBeGreaterThan(0);
    expect(result.aiDiagnostics.some((diagnostic: AIDiagnostic) => diagnostic.fix?.type === 'replaceFile')).toBe(true);
  });

  it('plans with Graph and list provisioning profiles from structured intent', () => {
    const plan = sdk.ai().plan({
      intent: 'Build a people finder backed by Graph',
      solution: { name: 'AIPlanner' },
      components: [{ name: 'PeopleFinder' }],
      data: { graphPermissions: ['User.Read.All'], lists: [{ title: 'People Cache', template: 100 }] }
    });

    expect(plan.profile).toBe('spfx-1.22-graph-webpart');
    expect(plan.steps).toContain('validate-sppkg');
  });

  it('assesses PnP gallery source dependencies before rebuilding', () => {
    const ai = sdk.ai();
    const project = ai.generate({
      solution: { name: 'AIPnPImport' },
      components: [{ name: 'ImportedSample' }],
      files: {
        'src/webparts/ImportedSample/ImportedSampleWebPart.ts': `import * as React from 'react';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { spfi } from '@pnp/sp';
import { PrimaryButton } from '@fluentui/react';
export default class ImportedSampleWebPart extends BaseClientSideWebPart<any> {}`,
        'src/webparts/ImportedSample/components/ImportedSample.tsx': `import * as React from 'react';
export default class ImportedSample extends React.Component<any, any> {
  public async componentDidMount(): Promise<void> {
    const client = await this.props.context.msGraphClientFactory.getClient('3');
    await client.api('/users').get();
    await client.api('/groups/group-id/members').get();
  }
  public render(): React.ReactElement<any> { return <section>Imported</section>; }
}`
      }
    });

    const compatibility = ai.assessSourceCompatibility(project);

    expect(compatibility.compatible).toBe(false);
    expect(compatibility.supportedImports).toContain('react');
    expect(compatibility.supportedImports).toContain('@microsoft/sp-webpart-base');
    expect(compatibility.unsupportedImports).toEqual(['@fluentui/react', '@pnp/sp']);
    expect(compatibility.graphPermissions).toEqual(['Group.Read.All', 'User.Read.All']);
    expect(compatibility.diagnostics.some(diagnostic => diagnostic.code === 'AI_UNSUPPORTED_IMPORT')).toBe(true);
    expect(compatibility.diagnostics.some(diagnostic => diagnostic.fix?.type === 'addGraphPermission')).toBe(true);
  });

  it('can repair inferred Graph permissions from compatibility diagnostics', () => {
    const ai = sdk.ai();
    const project = ai.generate({
      solution: { name: 'AIGraphRepair' },
      components: [{ name: 'GraphRepair' }],
      files: {
        'src/webparts/GraphRepair/GraphRepairWebPart.ts': 'export default class GraphRepairWebPart {}',
        'src/webparts/GraphRepair/components/GraphRepair.tsx': "export async function load(context: any) { const c = await context.msGraphClientFactory.getClient('3'); return c.api('/sites/root').get(); }"
      }
    });

    const compatibility = ai.assessSourceCompatibility(project);
    const repaired = ai.repair(project, compatibility.diagnostics);

    expect(repaired.ir.graph.map(permission => permission.scope)).toEqual(['Sites.Read.All']);
  });
});
