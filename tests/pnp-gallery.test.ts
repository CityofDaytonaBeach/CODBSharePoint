import {
  CODBSharePoint,
  addGraphPermission,
  addWebPart,
  createIR,
  validateSPPKGPackage
} from '../src/index.js';
import { strFromU8, unzipSync } from 'fflate';

const sdk = new CODBSharePoint();

describe('CODBSharePoint - PnP gallery recreation smoke tests', () => {
  it('recreates and packages a Quick Links style web part', async () => {
    const ir = createIR({
      name: 'PnpQuickLinksRecreated',
      version: '1.0.0',
      description: 'Recreated from the PnP Quick Links web part pattern',
      author: 'CODBSharePoint'
    });
    addWebPart(ir, {
      name: 'QuickLinks',
      displayName: 'Quick Links',
      description: 'Displays a responsive set of promoted links',
      framework: 'react'
    });

    const result = await sdk.buildFromIR(ir, quickLinksFiles());

    expect(result.success).toBe(true);
    expect(result.sppkg).toBeDefined();
    expect(validateSPPKGPackage(result.sppkg!)).toEqual([]);

    const packageFiles = unzipSync(result.sppkg!);
    const manifest = JSON.parse(strFromU8(packageFiles['PnpQuickLinksRecreated/QuickLinks.manifest.json']));
    expect(manifest.loaderConfig.entryModuleId).toBe('QuickLinks.bundle');
    expect(packageFiles['PnpQuickLinksRecreated/QuickLinks.bundle.js']).toBeDefined();

    const bundle = strFromU8(packageFiles['PnpQuickLinksRecreated/QuickLinks.bundle.js']);
    expect(bundle).toContain('define(');
    expect(bundle).toContain('City Services');
    expect(bundle).not.toContain('<section');
  });

  it('recreates and packages a Graph Group Viewer style web part with Graph permissions', async () => {
    const ir = createIR({
      name: 'PnpGraphGroupViewerRecreated',
      version: '1.0.0',
      description: 'Recreated from the PnP Graph Group Viewer web part pattern',
      author: 'CODBSharePoint'
    });
    addWebPart(ir, {
      name: 'GraphGroupViewer',
      displayName: 'Graph Group Viewer',
      description: 'Searches groups and previews members',
      framework: 'react'
    });
    addGraphPermission(ir, 'Group.Read.All');
    addGraphPermission(ir, 'User.Read.All');

    const result = await sdk.buildFromIR(ir, graphGroupViewerFiles());

    expect(result.success).toBe(true);
    expect(result.deployment.requiresAdmin).toBe(true);
    expect(result.sppkg).toBeDefined();
    expect(validateSPPKGPackage(result.sppkg!)).toEqual([]);

    const packageFiles = unzipSync(result.sppkg!);
    const packageSolution = JSON.parse(strFromU8(packageFiles['PnpGraphGroupViewerRecreated/package-solution.json']));
    expect(packageSolution.solution.webApiPermissionRequests).toEqual([
      { resource: 'Microsoft Graph', scope: 'Group.Read.All' },
      { resource: 'Microsoft Graph', scope: 'User.Read.All' }
    ]);

    const bundle = strFromU8(packageFiles['PnpGraphGroupViewerRecreated/GraphGroupViewer.bundle.js']);
    expect(bundle).toContain('/groups');
    expect(bundle).toContain('/members');
  });

  it('recreates and packages a Data Table style web part with real Sass', async () => {
    const ir = createIR({
      name: 'PnpDataTableRecreated',
      version: '1.0.0',
      description: 'Recreated from the PnP Data Table web part pattern',
      author: 'CODBSharePoint'
    });
    addWebPart(ir, {
      name: 'DataTable',
      displayName: 'Data Table',
      description: 'Renders SharePoint list rows in a table layout',
      framework: 'react'
    });

    const result = await sdk.buildFromIR(ir, dataTableFiles());

    expect(result.success).toBe(true);
    expect(result.sppkg).toBeDefined();
    expect(validateSPPKGPackage(result.sppkg!)).toEqual([]);

    const css = result.files.find(file => file.path.endsWith('DataTable.module.css'));
    expect(css).toBeDefined();
    expect(String(css!.content)).toContain('grid-template-columns: 2fr 1fr 1fr');
    expect(String(css!.content)).not.toContain('$header');

    const packageFiles = unzipSync(result.sppkg!);
    expect(packageFiles['PnpDataTableRecreated/DataTable.bundle.js']).toBeDefined();
  });
});

function quickLinksFiles(): Map<string, string> {
  const files = new Map<string, string>();
  files.set('src/webparts/QuickLinks/QuickLinksWebPart.ts', webPartSource('QuickLinks', 'QuickLinks'));
  files.set('src/webparts/QuickLinks/components/QuickLinks.tsx', `import * as React from 'react';
import styles from './QuickLinks.module.scss';

const links = [
  { title: 'City Services', url: 'https://www.codb.us', icon: 'CityNext' },
  { title: 'Permits', url: 'https://www.codb.us/permits', icon: 'ClipboardList' },
  { title: 'Meetings', url: 'https://www.codb.us/meetings', icon: 'Calendar' }
];

export default class QuickLinks extends React.Component<any, any> {
  public render(): React.ReactElement<any> {
    return (
      <section className={styles.quickLinks}>
        <h2>Quick Links</h2>
        <div className={styles.grid}>
          {links.map(link => (
            <a className={styles.card} href={link.url} key={link.title}>
              <span className={styles.icon}>{link.icon}</span>
              <span>{link.title}</span>
            </a>
          ))}
        </div>
      </section>
    );
  }
}
`);
  files.set('src/webparts/QuickLinks/components/QuickLinks.module.scss', `.quickLinks {
  padding: 16px;

  .grid {
    display: grid;
    grid-template-columns: repeat(3, minmax(0, 1fr));
    gap: 12px;
  }

  .card {
    border: 1px solid #d0d0d0;
    border-radius: 8px;
    color: #323130;
    padding: 14px;
    text-decoration: none;
  }

  .icon {
    display: block;
    font-weight: 600;
  }
}
`);
  return files;
}

function graphGroupViewerFiles(): Map<string, string> {
  const files = new Map<string, string>();
  files.set('src/webparts/GraphGroupViewer/GraphGroupViewerWebPart.ts', webPartSource('GraphGroupViewer', 'GraphGroupViewer'));
  files.set('src/webparts/GraphGroupViewer/components/GraphGroupViewer.tsx', `import * as React from 'react';
import styles from './GraphGroupViewer.module.scss';

export default class GraphGroupViewer extends React.Component<any, any> {
  public state = {
    groups: [{ id: 'group-1', displayName: 'Finance Team', mail: 'finance@example.com' }],
    members: [{ displayName: 'Ada Lovelace', mail: 'ada@example.com' }]
  };

  public async componentDidMount(): Promise<void> {
    const factory = this.props.context?.msGraphClientFactory;
    if (!factory) return;
    const client = await factory.getClient('3');
    await client.api('/groups').version('v1.0').select('id,displayName,mail').get();
    await client.api('/groups/group-1/members').version('v1.0').select('displayName,mail').get();
  }

  public render(): React.ReactElement<any> {
    return (
      <section className={styles.viewer}>
        <h2>Graph Group Viewer</h2>
        <ul>{this.state.groups.map((group: any) => <li key={group.id}>{group.displayName}</li>)}</ul>
        <h3>Members</h3>
        <ul>{this.state.members.map((member: any) => <li key={member.mail}>{member.displayName}</li>)}</ul>
      </section>
    );
  }
}
`);
  files.set('src/webparts/GraphGroupViewer/components/GraphGroupViewer.module.scss', `.viewer {
  border-left: 4px solid #0078d4;
  padding: 16px;

  ul {
    margin: 0;
    padding-left: 18px;
  }
}
`);
  return files;
}

function dataTableFiles(): Map<string, string> {
  const files = new Map<string, string>();
  files.set('src/webparts/DataTable/DataTableWebPart.ts', webPartSource('DataTable', 'DataTable'));
  files.set('src/webparts/DataTable/components/DataTable.tsx', `import * as React from 'react';
import styles from './DataTable.module.scss';

const rows = [
  { title: 'Capital Project', status: 'Active', owner: 'Public Works' },
  { title: 'Permit Review', status: 'Pending', owner: 'Permits' },
  { title: 'Agenda Item', status: 'Complete', owner: 'Clerk' }
];

export default class DataTable extends React.Component<any, any> {
  public render(): React.ReactElement<any> {
    return (
      <section className={styles.table}>
        <div className={styles.header}><span>Title</span><span>Status</span><span>Owner</span></div>
        {rows.map(row => (
          <div className={styles.row} key={row.title}>
            <span>{row.title}</span><span>{row.status}</span><span>{row.owner}</span>
          </div>
        ))}
      </section>
    );
  }
}
`);
  files.set('src/webparts/DataTable/components/DataTable.module.scss', `$header: #f3f2f1;
$border: #edebe9;

.table {
  border: 1px solid $border;

  .header,
  .row {
    display: grid;
    grid-template-columns: 2fr 1fr 1fr;
    gap: 8px;
    padding: 10px;
  }

  .header {
    background: $header;
    font-weight: 600;
  }

  .row {
    border-top: 1px solid $border;
  }
}
`);
  return files;
}

function webPartSource(componentName: string, exportName: string): string {
  return `import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import ${exportName} from './components/${componentName}';

export default class ${componentName}WebPart extends BaseClientSideWebPart<any> {
  public render(): void {
    const element = React.createElement(${exportName}, { context: this.context });
    ReactDOM.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDOM.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
`;
}
