// ============================================================================
// SharePoint Simulator - Browser-side mock implementation
// ============================================================================

import type {
  SimulatorConfig,
  SimulatorUser,
  SimulatorList,
  SimulatorLibrary,
  ThemeColorPalette,
  CODBIR
} from '../types/index.js';

export class SharePointSimulator {
  private config: SimulatorConfig;
  private container: HTMLElement | null = null;

  constructor(config: SimulatorConfig = {}) {
    this.config = {
      siteUrl: config.siteUrl || 'https://tenant.sharepoint.com/sites/test',
      siteTitle: config.siteTitle || 'Test Site',
      user: config.user || this.getDefaultUser(),
      theme: config.theme || this.getDefaultTheme(),
      lists: config.lists || this.getDefaultLists(),
      libraries: config.libraries || this.getDefaultLibraries(),
      webPartSize: config.webPartSize || 'medium'
    };
  }

  private getDefaultUser(): SimulatorUser {
    return {
      displayName: 'Test User',
      email: 'testuser@tenant.onmicrosoft.com',
      loginName: 'i:0#.f|membership|testuser@tenant.onmicrosoft.com',
      id: '12345678-1234-1234-1234-123456789012',
      isSiteAdmin: true,
      isSiteOwner: false,
      title: 'Software Developer'
    };
  }

  private getDefaultTheme(): ThemeColorPalette {
    return {
      themePrimary: '#0078d4',
      themeLighterAlt: '#eff6fc',
      themeLighter: '#deecf9',
      themeLight: '#c7e0f4',
      themeTertiary: '#71afe5',
      themeSecondary: '#2b88d8',
      themeDarkAlt: '#106ebe',
      themeDark: '#005a9e',
      themeDarker: '#004578',
      neutralLighterAlt: '#f3f2f1',
      neutralLighter: '#edebe9',
      neutralLight: '#d2d0ce',
      neutralQuaternaryAlt: '#e1dfdd',
      neutralQuaternary: '#d0d0d0',
      neutralTertiaryAlt: '#c8c6c4',
      neutralTertiary: '#a19f9d',
      neutralSecondaryAlt: '#979593',
      neutralSecondary: '#605e5c',
      neutralPrimaryAlt: '#3b3a39',
      neutralPrimary: '#323130',
      neutralDark: '#201f1e',
      black: '#000000',
      white: '#ffffff'
    };
  }

  private getDefaultLists(): SimulatorList[] {
    return [
      {
        title: 'Documents',
        items: [
          { Title: 'Document1.docx', FileLeafRef: 'Document1.docx', Created: '2024-01-15' },
          { Title: 'Spreadsheet.xlsx', FileLeafRef: 'Spreadsheet.xlsx', Created: '2024-01-16' }
        ],
        fields: [
          { name: 'Title', title: 'Title', type: 'Text' },
          { name: 'FileLeafRef', title: 'File Name', type: 'File' }
        ]
      },
      {
        title: 'Employees',
        items: [
          { Title: 'John Smith', Email: 'john@tenant.com', Department: 'IT', JobTitle: 'Developer' },
          { Title: 'Jane Doe', Email: 'jane@tenant.com', Department: 'HR', JobTitle: 'Manager' },
          { Title: 'Bob Johnson', Email: 'bob@tenant.com', Department: 'Finance', JobTitle: 'Analyst' }
        ],
        fields: [
          { name: 'Title', title: 'Name', type: 'Text' },
          { name: 'Email', title: 'Email', type: 'Text' },
          { name: 'Department', title: 'Department', type: 'Choice' },
          { name: 'JobTitle', title: 'Job Title', type: 'Text' }
        ]
      }
    ];
  }

  private getDefaultLibraries(): SimulatorLibrary[] {
    return [
      {
        title: 'Site Assets',
        files: [
          { name: 'logo.png', serverRelativeUrl: '/sites/test/SiteAssets/logo.png', size: 15000, timeCreated: '2024-01-01', timeLastModified: '2024-01-01' }
        ]
      }
    ];
  }

  render(container: HTMLElement): void {
    this.container = container;
    container.innerHTML = this.generateHTML();
    this.applyTheme();
  }

  private generateHTML(): string {
    const { siteTitle, user, lists } = this.config;

    return `
      <div class="codb-simulator" style="font-family: 'Segoe UI', sans-serif; max-width: 1200px; margin: 0 auto;">
        <div style="background: ${this.config.theme?.themePrimary}; color: white; padding: 12px 24px; display: flex; justify-content: space-between; align-items: center;">
          <div style="display: flex; align-items: center; gap: 16px;">
            <span style="font-size: 20px; font-weight: 600;">${siteTitle}</span>
            <nav style="display: flex; gap: 12px;">
              <a href="#" style="color: white; text-decoration: none;">Home</a>
              <a href="#" style="color: white; text-decoration: none;">Documents</a>
              <a href="#" style="color: white; text-decoration: none;">Site Contents</a>
            </nav>
          </div>
          <div style="display: flex; align-items: center; gap: 12px;">
            <span style="font-size: 14px;">${user.displayName}</span>
            <div style="width: 32px; height: 32px; border-radius: 50%; background: ${this.config.theme?.themeLighter}; display: flex; align-items: center; justify-content: center; color: ${this.config.theme?.themePrimary}; font-weight: 600;">
              ${user.displayName.charAt(0)}
            </div>
          </div>
        </div>

        <div style="padding: 24px;">
          <div id="webpart-zone" style="min-height: 400px; border: 2px dashed #ccc; border-radius: 4px; padding: 20px;">
            <div style="text-align: center; color: #666; padding: 40px;">
              <p>Web Part Preview Zone</p>
              <p style="font-size: 12px;">Drop a web part here or use the simulator API</p>
            </div>
          </div>
        </div>

        <div style="background: #f3f2f1; padding: 12px 24px; border-top: 1px solid #edebe9; font-size: 12px; color: #605e5c;">
          SharePoint Simulator | ${user.email} | ${new Date().toLocaleDateString()}
        </div>
      </div>
    `;
  }

  private applyTheme(): void {
    if (!this.container) return;

    const theme = this.config.theme;
    if (!theme) return;

    // Apply theme CSS variables
    const style = document.createElement('style');
    style.textContent = `
      .codb-simulator {
        --sp-theme-primary: ${theme.themePrimary};
        --sp-theme-lighter: ${theme.themeLighter};
        --sp-theme-light: ${theme.themeLight};
        --sp-theme-tertiary: ${theme.themeTertiary};
        --sp-theme-secondary: ${theme.themeSecondary};
        --sp-theme-dark: ${theme.themeDark};
        --sp-neutral-primary: ${theme.neutralPrimary};
        --sp-neutral-secondary: ${theme.neutralSecondary};
        --sp-neutral-lighter: ${theme.neutralLighter};
      }
    `;
    this.container.appendChild(style);
  }

  // API methods for web parts to interact with
  getContext(): Record<string, unknown> {
    return {
      pageContext: {
        web: {
          absoluteUrl: this.config.siteUrl,
          title: this.config.siteTitle,
          id: 'test-web-id'
        },
        user: this.config.user,
        culturalInfo: {
          currentUICultureName: 'en-us',
          currentCultureName: 'en-us'
        }
      },
      spHttpClient: {
        get: async (url: string) => {
          return { json: async () => ({}) };
        }
      },
      msGraphClient: {
        api: (url: string) => ({
          get: async () => ({ value: [] })
        })
      }
    };
  }

  getListData(listTitle: string): unknown[] {
    const list = this.config.lists?.find(l => l.title === listTitle);
    return list?.items || [];
  }

  getTheme(): ThemeColorPalette | undefined {
    return this.config.theme;
  }

  getUser(): SimulatorUser | undefined {
    return this.config.user;
  }

  // Mount a web part component
  mountWebPart(component: React.ComponentType<unknown>, props: Record<string, unknown> = {}): void {
    if (!this.container) return;

    const webpartZone = this.container.querySelector('#webpart-zone');
    if (!webpartZone) return;

    // In a real implementation, this would use React.render
    const placeholder = document.createElement('div');
    placeholder.className = 'simulated-webpart';
    placeholder.innerHTML = `
      <div style="border: 1px solid #ddd; border-radius: 4px; padding: 16px; margin: 8px 0;">
        <div style="font-weight: 600; margin-bottom: 8px;">${component.displayName || 'Web Part'}</div>
        <div style="color: #666;">Simulated rendering</div>
      </div>
    `;

    webpartZone.appendChild(placeholder);
  }

  // Destroy simulator
  destroy(): void {
    if (this.container) {
      this.container.innerHTML = '';
    }
    this.container = null;
  }
}
