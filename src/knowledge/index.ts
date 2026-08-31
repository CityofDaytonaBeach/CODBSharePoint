// ============================================================================
// Microsoft 365 / SharePoint Knowledge Catalog
// Compact, versioned knowledge base used by AI workflows to choose supported
// APIs, schemas, dependencies, permissions, patterns, and deployment settings.
// ============================================================================

export type KnowledgeStatus = 'supported' | 'experimental' | 'known' | 'unsupported';

export interface KnowledgeEntry {
  id: string;
  name: string;
  status: KnowledgeStatus;
  description: string;
  permissions?: string[];
  dependencies?: string[];
  notes?: string[];
}

export interface ApiEntry extends KnowledgeEntry {
  endpoint: string;
  method: 'GET' | 'POST' | 'PATCH' | 'PUT' | 'DELETE';
  client: 'SPHttpClient' | 'MSGraphClient' | 'AadHttpClient';
}

export interface SchemaEntry extends KnowledgeEntry {
  kind: 'field' | 'listTemplate' | 'contentType' | 'view' | 'formatting';
  value: string | number;
}

export interface DependencyEntry extends KnowledgeEntry {
  packageName: string;
  handling: 'external' | 'browser-prebundle-required' | 'replace-with-native-api' | 'unsupported';
}

export interface PatternEntry extends KnowledgeEntry {
  components: string[];
  dataSources: string[];
  recommendedPermissions: string[];
}

export interface SharePointKnowledgeCatalog {
  version: 'codbsharepoint.knowledge/1.0';
  sharePointRest: ApiEntry[];
  graph: ApiEntry[];
  schema: SchemaEntry[];
  spfxRuntime: KnowledgeEntry[];
  componentTypes: KnowledgeEntry[];
  permissions: KnowledgeEntry[];
  m365Integrations: KnowledgeEntry[];
  pnpDependencies: DependencyEntry[];
  fluentUI: DependencyEntry[];
  deployment: KnowledgeEntry[];
  sourcePatterns: PatternEntry[];
}

export const SHAREPOINT_KNOWLEDGE: SharePointKnowledgeCatalog = {
  version: 'codbsharepoint.knowledge/1.0',
  sharePointRest: [
    api('sp.lists', 'Lists', 'GET', '/_api/web/lists', 'SPHttpClient', 'supported', 'List and inspect SharePoint lists.'),
    api('sp.listItems', 'List Items', 'GET', "/_api/web/lists/getbytitle('{list}')/items", 'SPHttpClient', 'supported', 'Read list items.'),
    api('sp.createItem', 'Create List Item', 'POST', "/_api/web/lists/getbytitle('{list}')/items", 'SPHttpClient', 'supported', 'Create list items with request digest/context.'),
    api('sp.files', 'Library Files', 'GET', "/_api/web/getfolderbyserverrelativeurl('{folder}')/files", 'SPHttpClient', 'supported', 'Read files from document libraries.'),
    api('sp.search', 'SharePoint Search', 'GET', "/_api/search/query?querytext='{query}'", 'SPHttpClient', 'experimental', 'Search API output varies by tenant settings.'),
    api('sp.userProfile', 'User Profile', 'GET', '/_api/SP.UserProfiles.PeopleManager/GetMyProperties', 'SPHttpClient', 'known', 'Classic user profile endpoint.'),
    api('sp.taxonomy', 'Taxonomy', 'GET', '/_api/v2.1/termStore', 'SPHttpClient', 'experimental', 'Term store requires tenant capabilities and permissions.')
  ],
  graph: [
    api('graph.me', 'Current User', 'GET', '/me', 'MSGraphClient', 'supported', 'Read current user profile.', ['User.Read']),
    api('graph.users', 'Users', 'GET', '/users', 'MSGraphClient', 'supported', 'Search or list users.', ['User.Read.All']),
    api('graph.groups', 'Groups', 'GET', '/groups', 'MSGraphClient', 'supported', 'Search or list groups.', ['Group.Read.All']),
    api('graph.groupMembers', 'Group Members', 'GET', '/groups/{id}/members', 'MSGraphClient', 'supported', 'Read group membership.', ['Group.Read.All', 'User.Read.All']),
    api('graph.sites', 'Sites', 'GET', '/sites', 'MSGraphClient', 'supported', 'Read SharePoint sites through Graph.', ['Sites.Read.All']),
    api('graph.drives', 'Drives', 'GET', '/drives', 'MSGraphClient', 'known', 'Read OneDrive or document library drives.', ['Files.Read.All']),
    api('graph.planner', 'Planner Tasks', 'GET', '/planner/plans/{id}/tasks', 'MSGraphClient', 'known', 'Planner has group-backed permission requirements.', ['Tasks.ReadWrite']),
    api('graph.calendar', 'Calendar Events', 'GET', '/me/events', 'MSGraphClient', 'known', 'Read current user calendar events.', ['Calendars.ReadWrite'])
  ],
  schema: [
    schema('field.text', 'Text Field', 'field', 'Text', 'supported', 'Single line of text.'),
    schema('field.note', 'Multiline Text Field', 'field', 'Note', 'supported', 'Multiple lines of text.'),
    schema('field.choice', 'Choice Field', 'field', 'Choice', 'supported', 'Choice field with fixed options.'),
    schema('field.number', 'Number Field', 'field', 'Number', 'supported', 'Numeric field.'),
    schema('field.datetime', 'Date Time Field', 'field', 'DateTime', 'supported', 'Date or date/time field.'),
    schema('field.lookup', 'Lookup Field', 'field', 'Lookup', 'experimental', 'Requires target list and indexed lookup constraints.'),
    schema('field.taxonomy', 'Managed Metadata Field', 'field', 'Taxonomy', 'experimental', 'Requires term set IDs and taxonomy hidden fields.'),
    schema('list.generic', 'Custom List', 'listTemplate', 100, 'supported', 'Generic SharePoint list.'),
    schema('list.documentLibrary', 'Document Library', 'listTemplate', 101, 'supported', 'Document library.'),
    schema('format.column', 'Column Formatting', 'formatting', 'columnFormatting', 'supported', 'SharePoint JSON column formatting.'),
    schema('format.view', 'View Formatting', 'formatting', 'viewFormatting', 'known', 'SharePoint JSON view formatting.')
  ],
  spfxRuntime: [
    entry('runtime.context', 'Web Part Context', 'supported', 'Provides page context, SPHttpClient, MSGraphClient factory, and service scope.'),
    entry('runtime.spHttpClient', 'SPHttpClient', 'supported', 'Native SPFx client for SharePoint REST calls.'),
    entry('runtime.msGraphClient', 'MSGraphClient', 'supported', 'Native SPFx client for Microsoft Graph calls.'),
    entry('runtime.aadHttpClient', 'AadHttpClient', 'known', 'Calls custom AAD-secured APIs.'),
    entry('runtime.themeProvider', 'Theme Provider', 'known', 'Consumes SharePoint theme variants.'),
    entry('runtime.dynamicData', 'Dynamic Data', 'known', 'Connects web parts through SPFx dynamic data.')
  ],
  componentTypes: [
    entry('component.webpart.react', 'React Web Part', 'supported', 'Primary production-proven component path.'),
    entry('component.webpart.vanilla', 'Vanilla Web Part', 'supported', 'No React dependency.'),
    entry('component.applicationCustomizer', 'Application Customizer', 'experimental', 'Builds but requires SharePoint deployment proof.'),
    entry('component.fieldCustomizer', 'Field Customizer', 'experimental', 'Builds but requires SharePoint deployment proof.'),
    entry('component.commandSet', 'ListView Command Set', 'experimental', 'Builds but requires SharePoint deployment proof.'),
    entry('component.ace', 'Adaptive Card Extension', 'known', 'Modeled, not production package proven.')
  ],
  permissions: [
    entry('permission.graph.userRead', 'User.Read', 'supported', 'Read current user profile.', ['User.Read']),
    entry('permission.graph.userReadAll', 'User.Read.All', 'supported', 'Read users in the tenant. Requires admin consent.', ['User.Read.All']),
    entry('permission.graph.groupReadAll', 'Group.Read.All', 'supported', 'Read groups and members. Requires admin consent.', ['Group.Read.All']),
    entry('permission.graph.sitesReadAll', 'Sites.Read.All', 'supported', 'Read sites through Graph. Requires admin consent.', ['Sites.Read.All']),
    entry('permission.sharepoint.site', 'SharePoint Site Permissions', 'known', 'SharePoint REST is bounded by current user and deployed app permissions.')
  ],
  m365Integrations: [
    entry('m365.teams', 'Microsoft Teams', 'known', 'SPFx web parts can be exposed as Teams tabs with manifest work.'),
    entry('m365.viva', 'Viva Connections', 'known', 'ACE and dashboard scenarios.'),
    entry('m365.planner', 'Planner', 'known', 'Graph Planner APIs and group permissions.'),
    entry('m365.onedrive', 'OneDrive', 'known', 'Graph files/drives APIs.'),
    entry('m365.powerAutomate', 'Power Automate', 'known', 'Usually via HTTP triggers/custom connectors.'),
    entry('m365.azureFunctions', 'Azure Functions', 'known', 'Use AadHttpClient for secured APIs.')
  ],
  pnpDependencies: [
    dep('@pnp/sp', 'PnPjs SharePoint', 'browser-prebundle-required', 'known', 'Common PnP sample dependency. Needs browser package registry/prebundle support.'),
    dep('@pnp/graph', 'PnPjs Graph', 'browser-prebundle-required', 'known', 'Can often be replaced by MSGraphClient helper generation.'),
    dep('@pnp/spfx-controls-react', 'PnP React Controls', 'browser-prebundle-required', 'known', 'Common UI controls for people picker/file picker.'),
    dep('@pnp/spfx-property-controls', 'PnP Property Controls', 'browser-prebundle-required', 'known', 'Property pane controls used by many samples.')
  ],
  fluentUI: [
    dep('@fluentui/react', 'Fluent UI React v8', 'browser-prebundle-required', 'known', 'Common SPFx UI library.'),
    dep('@fluentui/react-components', 'Fluent UI React v9', 'browser-prebundle-required', 'experimental', 'SPFx compatibility depends on sample and styling setup.'),
    dep('office-ui-fabric-react', 'Office UI Fabric React', 'browser-prebundle-required', 'known', 'Legacy SPFx UI dependency.')
  ],
  deployment: [
    entry('deploy.tenantAppCatalog', 'Tenant App Catalog', 'known', 'Upload .sppkg for tenant-wide availability.'),
    entry('deploy.siteCollectionAppCatalog', 'Site Collection App Catalog', 'known', 'Site-scoped app catalog deployment.'),
    entry('deploy.skipFeatureDeployment', 'Tenant-wide Deployment', 'supported', 'Controlled by skipFeatureDeployment in package-solution.'),
    entry('deploy.includeClientSideAssets', 'Include Client Side Assets', 'supported', 'Package client-side assets for SharePoint hosting.'),
    entry('deploy.apiApproval', 'API Permission Approval', 'supported', 'Graph permissions appear for admin approval in SharePoint admin center.'),
    entry('deploy.isolatedDomain', 'Isolated Domain', 'known', 'Requires tenant setup and isolated web part metadata.')
  ],
  sourcePatterns: [
    pattern('pattern.employeeDirectory', 'Employee Directory', 'supported', ['React Web Part'], ['Microsoft Graph /users'], ['User.Read.All']),
    pattern('pattern.quickLinks', 'Quick Links', 'supported', ['React Web Part'], ['Static links', 'SharePoint list optional'], []),
    pattern('pattern.dataTable', 'Data Table', 'supported', ['React Web Part'], ['SharePoint list items'], []),
    pattern('pattern.documentExplorer', 'Document Explorer', 'known', ['React Web Part'], ['SharePoint library files', 'Graph drives optional'], ['Files.Read.All']),
    pattern('pattern.approvals', 'Approvals', 'known', ['React Web Part'], ['SharePoint list', 'Power Automate optional'], []),
    pattern('pattern.orgChart', 'Organization Chart', 'known', ['React Web Part'], ['Microsoft Graph /users/manager'], ['User.Read.All']),
    pattern('pattern.calendar', 'Calendar', 'known', ['React Web Part'], ['Graph calendar or SharePoint events list'], ['Calendars.ReadWrite']),
    pattern('pattern.faq', 'FAQ Accordion', 'supported', ['React Web Part'], ['Static data', 'SharePoint list optional'], [])
  ]
};

export function getKnowledgeCatalog(): SharePointKnowledgeCatalog {
  return SHAREPOINT_KNOWLEDGE;
}

export function findKnowledgeEntries(query: string): KnowledgeEntry[] {
  const normalized = query.toLowerCase();
  return flattenKnowledgeEntries().filter(item => {
    return item.id.toLowerCase().includes(normalized) ||
      item.name.toLowerCase().includes(normalized) ||
      item.description.toLowerCase().includes(normalized);
  });
}

export function summarizeKnowledge(): Record<string, number> {
  return {
    sharePointRest: SHAREPOINT_KNOWLEDGE.sharePointRest.length,
    graph: SHAREPOINT_KNOWLEDGE.graph.length,
    schema: SHAREPOINT_KNOWLEDGE.schema.length,
    spfxRuntime: SHAREPOINT_KNOWLEDGE.spfxRuntime.length,
    componentTypes: SHAREPOINT_KNOWLEDGE.componentTypes.length,
    permissions: SHAREPOINT_KNOWLEDGE.permissions.length,
    m365Integrations: SHAREPOINT_KNOWLEDGE.m365Integrations.length,
    pnpDependencies: SHAREPOINT_KNOWLEDGE.pnpDependencies.length,
    fluentUI: SHAREPOINT_KNOWLEDGE.fluentUI.length,
    deployment: SHAREPOINT_KNOWLEDGE.deployment.length,
    sourcePatterns: SHAREPOINT_KNOWLEDGE.sourcePatterns.length
  };
}

function flattenKnowledgeEntries(): KnowledgeEntry[] {
  return [
    ...SHAREPOINT_KNOWLEDGE.sharePointRest,
    ...SHAREPOINT_KNOWLEDGE.graph,
    ...SHAREPOINT_KNOWLEDGE.schema,
    ...SHAREPOINT_KNOWLEDGE.spfxRuntime,
    ...SHAREPOINT_KNOWLEDGE.componentTypes,
    ...SHAREPOINT_KNOWLEDGE.permissions,
    ...SHAREPOINT_KNOWLEDGE.m365Integrations,
    ...SHAREPOINT_KNOWLEDGE.pnpDependencies,
    ...SHAREPOINT_KNOWLEDGE.fluentUI,
    ...SHAREPOINT_KNOWLEDGE.deployment,
    ...SHAREPOINT_KNOWLEDGE.sourcePatterns
  ];
}

function entry(id: string, name: string, status: KnowledgeStatus, description: string, permissions?: string[]): KnowledgeEntry {
  return { id, name, status, description, permissions };
}

function api(id: string, name: string, method: ApiEntry['method'], endpoint: string, client: ApiEntry['client'], status: KnowledgeStatus, description: string, permissions?: string[]): ApiEntry {
  return { id, name, method, endpoint, client, status, description, permissions };
}

function schema(id: string, name: string, kind: SchemaEntry['kind'], value: string | number, status: KnowledgeStatus, description: string): SchemaEntry {
  return { id, name, kind, value, status, description };
}

function dep(packageName: string, name: string, handling: DependencyEntry['handling'], status: KnowledgeStatus, description: string): DependencyEntry {
  return { id: `dependency.${packageName}`, packageName, name, handling, status, description, dependencies: [packageName] };
}

function pattern(id: string, name: string, status: KnowledgeStatus, components: string[], dataSources: string[], recommendedPermissions: string[]): PatternEntry {
  return { id, name, status, components, dataSources, recommendedPermissions, description: `${name} app pattern.` };
}
