// ============================================================================
// CODBSharePoint - Core Type Definitions
// Browser-native SharePoint compiler, validator, and packaging SDK
// ============================================================================

// ---------------------------------------------------------------------------
// Version & Compatibility
// ---------------------------------------------------------------------------

export type SPFxVersion = '1.1.0' | '1.2.0' | '1.3.0' | '1.4.0' | '1.5.0' | '1.6.0' | '1.7.0' | '1.8.0' | '1.9.0' | '1.10.0' | '1.11.0' | '1.12.0' | '1.13.0' | '1.14.0' | '1.15.0' | '1.16.0' | '1.17.0' | '1.18.0' | '1.19.0' | '1.20.0' | '1.21.0' | '1.22.0';
export type TypeScriptVersion = '4.7' | '4.8' | '4.9' | '5.0' | '5.1' | '5.2' | '5.3' | '5.4' | '5.5';
export type ReactVersion = '16' | '17' | '18';

// ---------------------------------------------------------------------------
// Component Types
// ---------------------------------------------------------------------------

export type ComponentType = 'webpart' | 'extension' | 'ace' | 'fieldCustomizer' | 'commandSet' | 'formCustomizer';
export type ExtensionType = 'ApplicationCustomizer' | 'FieldCustomizer' | 'ListViewCommandSet' | 'FormCustomizer';
export type Framework = 'react' | 'knockout' | 'none' | 'vue' | 'vanilla';
export type ModuleType = 'modern' | 'classic';

// ---------------------------------------------------------------------------
// SPFx Version Compatibility Matrix
// ---------------------------------------------------------------------------

export interface SPFxCompatibility {
  spfx: SPFxVersion;
  react: ReactVersion;
  typescript: TypeScriptVersion;
  node: string;
  scaffold: string;
  buildTool: 'gulp' | 'heft';
  bundler: 'webpack' | 'none';
}

export const SPFx_COMPATIBILITY: Record<SPFxVersion, SPFxCompatibility> = {
  '1.18.0': { spfx: '1.18.0', react: '17', typescript: '4.7', node: '16.13.x', scaffold: '@microsoft/sharepoint', buildTool: 'heft', bundler: 'webpack' },
  '1.19.0': { spfx: '1.19.0', react: '17', typescript: '4.7', node: '16.13.x', scaffold: '@microsoft/sharepoint', buildTool: 'heft', bundler: 'webpack' },
  '1.20.0': { spfx: '1.20.0', react: '17', typescript: '4.7', node: '18.17.x', scaffold: '@microsoft/sharepoint', buildTool: 'heft', bundler: 'webpack' },
  '1.21.0': { spfx: '1.21.0', react: '17', typescript: '5.3', node: '18.17.x', scaffold: '@microsoft/sharepoint', buildTool: 'heft', bundler: 'webpack' },
  '1.22.0': { spfx: '1.22.0', react: '18', typescript: '5.3', node: '18.17.x', scaffold: '@microsoft/sharepoint', buildTool: 'heft', bundler: 'webpack' },
} as const;

// ---------------------------------------------------------------------------
// File System
// ---------------------------------------------------------------------------

export interface VFSFile {
  path: string;
  content: string | Uint8Array;
  encoding: 'utf-8' | 'binary';
  timestamp?: number;
}

export interface VFS {
  files: Map<string, VFSFile>;
  addFile(path: string, content: string | Uint8Array, encoding?: 'utf-8' | 'binary'): void;
  getFile(path: string): VFSFile | undefined;
  removeFile(path: string): void;
  getFiles(): VFSFile[];
  getFilesByPattern(pattern: string | RegExp): VFSFile[];
  hasFile(path: string): boolean;
  readAsString(path: string): string | undefined;
  toZip(): Promise<Uint8Array>;
}

// ---------------------------------------------------------------------------
// Intermediate Representation (IR)
// ---------------------------------------------------------------------------

export interface CODBIR {
  $schema: 'codbsharepoint/ir/1.0';
  solution: SolutionConfig;
  components: ComponentDefinition[];
  lists: ListDefinition[];
  libraries: LibraryDefinition[];
  fields: FieldDefinition[];
  contentTypes: ContentTypeDefinition[];
  pages: PageDefinition[];
  extensions: ExtensionDefinition[];
  permissions: PermissionDefinition[];
  graph: GraphPermissionDefinition[];
  themes: ThemeDefinition[];
  formatting: FormattingDefinition[];
  provisioning: ProvisioningDefinition[];
  localization: LocalizationConfig;
  metadata: ProjectMetadata;
}

export interface SolutionConfig {
  name: string;
  id: string;
  version: string;
  description: string;
  author: string;
  company: string;
  namespace: string;
  environment: 'spo' | 'odsp';
  includeClientSideAssets: boolean;
  skipFeatureDeployment: boolean;
  isDomainIsolated: boolean;
  developer: DeveloperConfig;
  metadata: SolutionMetadata;
  features: FeatureDefinition[];
}

export interface DeveloperConfig {
  name: string;
  websiteUrl: string;
  privacyUrl?: string;
  termsOfUseUrl?: string;
}

export interface SolutionMetadata {
  'screenshot': string;
  'videoUrl'?: string;
  'categories'?: string[];
}

export interface FeatureDefinition {
  id: string;
  title: string;
  description: string;
  version: string;
  components: FeatureComponent[];
}

export interface FeatureComponent {
  type: ComponentType;
  id: string;
  name: string;
}

export interface ProjectMetadata {
  generator: string;
  version: string;
  createdAt: string;
  modifiedAt: string;
  spfxVersion: SPFxVersion;
  buildTool: 'gulp' | 'heft';
}

// ---------------------------------------------------------------------------
// Component Definitions
// ---------------------------------------------------------------------------

export interface ComponentDefinition {
  type: ComponentType;
  id: string;
  name: string;
  displayName: string;
  description: string;
  officeFabricIconFontName?: string;
  iconUrl?: string;
  group: ClientSideGroupDefinition;
  version: string;
  entry: string;
  entryModule?: string;
  framework: Framework;
  properties: PropertyPaneDefinition[];
  preconfiguredEntries: PreconfiguredEntry[];
  assets?: ComponentAsset[];
  contextSpecific?: boolean;
  supportedHosts?: string[];
}

export interface ClientSideGroupDefinition {
  id: string;
  name: string;
}

export interface PreconfiguredEntry {
  defaultTitle: string;
  description: string;
  officeFabricIconFontName?: string;
  properties: Record<string, unknown>;
}

// ---------------------------------------------------------------------------
// Extension Definitions
// ---------------------------------------------------------------------------

export interface ExtensionDefinition {
  type: ExtensionType;
  id: string;
  name: string;
  displayName: string;
  description: string;
  entry: string;
  clientSideComponentId: string;
  registeredClientSideExtensions?: string[];
  topNavigationZone?: boolean;
  bottomNavigationZone?: boolean;
  pageAction?: boolean;
  pageHeader?: boolean;
  ribbon?: boolean;
}

// ---------------------------------------------------------------------------
// Adaptive Card Extension (ACE)
// ---------------------------------------------------------------------------

export interface ACEDefinition {
  id: string;
  name: string;
  description: string;
  type: 'Primary' | 'Card' | 'QuickView';
  iconProperty?: string;
  cardComponents: ACECardComponent[];
  quickViews?: ACEQuickView[];
  properties: PropertyPaneDefinition[];
}

export interface ACECardComponent {
  id: string;
  cardView: string;
}

export interface ACEQuickView {
  id: string;
  title: string;
  template: string;
}

// ---------------------------------------------------------------------------
// Property Pane
// ---------------------------------------------------------------------------

export type PropertyPaneField = PropertyPaneTextField | PropertyPaneDropdown | PropertyPaneToggle | PropertyPaneSlider | PropertyPaneLink | PropertyPaneColorPicker | PropertyPaneMultiSelect | PropertyPaneChoiceGroup;

export interface PropertyPaneTextField {
  type: 'textField';
  propertyName: string;
  label: string;
  description?: string;
  placeholder?: string;
  multiline?: boolean;
  required?: boolean;
  value?: string;
  maxLength?: number;
}

export interface PropertyPaneDropdown {
  type: 'dropdown';
  propertyName: string;
  label: string;
  options: PropertyPaneDropdownOption[];
  selectedKey?: string;
  required?: boolean;
}

export interface PropertyPaneDropdownOption {
  key: string;
  text: string;
}

export interface PropertyPaneToggle {
  type: 'toggle';
  propertyName: string;
  label: string;
  onText?: string;
  offText?: string;
  checked?: boolean;
}

export interface PropertyPaneSlider {
  type: 'slider';
  propertyName: string;
  label: string;
  min: number;
  max: number;
  step?: number;
  value?: number;
}

export interface PropertyPaneLink {
  type: 'link';
  propertyName: string;
  label: string;
  href: string;
  target?: string;
}

export interface PropertyPaneColorPicker {
  type: 'colorPicker';
  propertyName: string;
  label: string;
  value?: string;
}

export interface PropertyPaneMultiSelect {
  type: 'multiSelect';
  propertyName: string;
  label: string;
  options: PropertyPaneDropdownOption[];
  selectedKeys?: string[];
}

export interface PropertyPaneChoiceGroup {
  type: 'choiceGroup';
  propertyName: string;
  label: string;
  options: PropertyPaneChoiceOption[];
  selectedKey?: string;
}

export interface PropertyPaneChoiceOption {
  key: string;
  text: string;
  iconProps?: { officeFabricIconFontName: string };
}

export interface PropertyPaneDefinition {
  pages: PropertyPanePage[];
}

export interface PropertyPanePage {
  header?: { description: string };
  groups: PropertyPaneGroup[];
}

export interface PropertyPaneGroup {
  groupName?: string;
  groupFields: PropertyPaneField[];
}

// ---------------------------------------------------------------------------
// SharePoint Data Definitions
// ---------------------------------------------------------------------------

export interface ListDefinition {
  title: string;
  description?: string;
  template: number;
  hidden?: boolean;
  contentTypes?: string[];
  fields: string[];
  rows?: Record<string, unknown>[];
  folderCreation?: boolean;
  versioning?: boolean;
  majorVersionLimit?: number;
  minorVersionLimit?: number;
}

export interface LibraryDefinition {
  title: string;
  description?: string;
  template?: number;
  hidden?: boolean;
  contentTypes?: string[];
  fields?: string[];
  versioningEnabled?: boolean;
  majorVersionLimit?: number;
  minorVersionLimit?: number;
  templateFeatureId?: string;
}

export interface FieldDefinition {
  name: string;
  displayName: string;
  type: FieldType;
  group?: string;
  description?: string;
  required?: boolean;
  hidden?: boolean;
  indexed?: boolean;
  unique?: boolean;
  minLength?: number;
  maxLength?: number;
  minValue?: number;
  maxValue?: number;
  choices?: string[];
  defaultValue?: string;
  lookup?: LookupFieldConfig;
  taxonomy?: TaxonomyFieldConfig;
  richText?: boolean;
  multiline?: boolean;
  appendChanges?: boolean;
  datetime?: DateTimeFieldConfig;
}

export type FieldType =
  | 'Text' | 'Note' | 'Number' | 'Currency' | 'DateTime'
  | 'Choice' | 'MultiChoice' | 'Boolean' | 'URL'
  | 'Lookup' | 'User' | 'UserMulti' | 'Taxonomy'
  | 'TaxonomyMulti' | 'Calculated' | 'Attachments'
  | 'ModStat' | 'Invalid';

export interface LookupFieldConfig {
  listName: string;
  fieldName: string;
  showField?: string;
  multiValue?: boolean;
}

export interface TaxonomyFieldConfig {
  termSetId: string;
  anchorId?: string;
  multiValue?: boolean;
  storeId?: string;
}

export interface DateTimeFieldConfig {
  displayFormat?: 'DateOnly' | 'DateTime';
  calendarType?: number;
  showWeekNumber?: boolean;
}

export interface ContentTypeDefinition {
  name: string;
  description?: string;
  group?: string;
  parentContentType?: string;
  hidden?: boolean;
  sealed?: boolean;
  fields: string[];
  documentTemplate?: string;
}

export interface PageDefinition {
  name: string;
  title: string;
  layout?: string;
  promotedState?: number;
  content: string;
}

// ---------------------------------------------------------------------------
// Permissions
// ---------------------------------------------------------------------------

export interface PermissionDefinition {
  resource: 'sharepoint' | 'graph' | 'azureAD';
  scope: string;
  level?: 'read' | 'write' | 'manage' | 'fullControl';
  description?: string;
}

export interface GraphPermissionDefinition {
  resource: string;
  scope: string;
  type: 'Delegated' | 'Application';
  requiresAdminApproval: boolean;
  description?: string;
}

// Supported Graph permissions
export const GRAPH_PERMISSIONS = {
  'User.Read': { requiresAdminApproval: false, description: 'Read user profile' },
  'User.Read.All': { requiresAdminApproval: true, description: 'Read all users' },
  'User.ReadWrite.All': { requiresAdminApproval: true, description: 'Read and write all users' },
  'Group.Read.All': { requiresAdminApproval: true, description: 'Read all groups' },
  'Group.ReadWrite.All': { requiresAdminApproval: true, description: 'Read and write all groups' },
  'Sites.Read.All': { requiresAdminApproval: true, description: 'Read all site collections' },
  'Sites.ReadWrite.All': { requiresAdminApproval: true, description: 'Read and write all site collections' },
  'Files.Read.All': { requiresAdminApproval: true, description: 'Read all files' },
  'Files.ReadWrite.All': { requiresAdminApproval: true, description: 'Read and write all files' },
  'Mail.Read': { requiresAdminApproval: true, description: 'Read mail' },
  'Mail.ReadWrite': { requiresAdminApproval: true, description: 'Read and write mail' },
  'Calendars.ReadWrite': { requiresAdminApproval: true, description: 'Read and write calendars' },
  'Directory.Read.All': { requiresAdminApproval: true, description: 'Read directory data' },
  'Domain.Read.All': { requiresAdminApproval: true, description: 'Read domain data' },
  'IdentityRiskEvent.Read.All': { requiresAdminApproval: true, description: 'Read identity risk events' },
  'Tasks.ReadWrite': { requiresAdminApproval: true, description: 'Read and write tasks' },
  'Teamwork.Migrate.All': { requiresAdminApproval: true, description: 'Teamwork migrate' },
  'AgPlaceholds.Read': { requiresAdminApproval: true, description: 'Read placeholds' },
  'ChannelMessage.Read.All': { requiresAdminApproval: true, description: 'Read channel messages' },
  'Chat.ReadWrite': { requiresAdminApproval: true, description: 'Read and write chat' },
} as const;

// ---------------------------------------------------------------------------
// Themes & Formatting
// ---------------------------------------------------------------------------

export interface ThemeDefinition {
  name: string;
  primary: ThemeColorPalette;
  isInverted?: boolean;
}

export interface ThemeColorPalette {
  themePrimary?: string;
  themeLighterAlt?: string;
  themeLighter?: string;
  themeLight?: string;
  themeTertiary?: string;
  themeSecondary?: string;
  themeDarkAlt?: string;
  themeDark?: string;
  themeDarker?: string;
  neutralLighterAlt?: string;
  neutralLighter?: string;
  neutralLight?: string;
  neutralQuaternaryAlt?: string;
  neutralQuaternary?: string;
  neutralTertiaryAlt?: string;
  neutralTertiary?: string;
  neutralSecondaryAlt?: string;
  neutralSecondary?: string;
  neutralPrimary?: string;
  neutralPrimaryAlt?: string;
  neutralDark?: string;
  black?: string;
  white?: string;
}

export interface FormattingDefinition {
  type: 'listFormatting' | 'columnFormatting' | 'formFormatting' | 'viewFormatting';
  name: string;
  target: string;
  json: Record<string, unknown>;
  description?: string;
}

// ---------------------------------------------------------------------------
// Provisioning
// ---------------------------------------------------------------------------

export interface ProvisioningDefinition {
  type: 'siteScript' | 'siteDesign' | 'list' | 'library' | 'column' | 'contentType';
  name: string;
  description?: string;
  data: Record<string, unknown>;
}

// ---------------------------------------------------------------------------
// Build Pipeline
// ---------------------------------------------------------------------------

export type BuildTarget =
  | 'sharepoint.spfx.webpart'
  | 'sharepoint.spfx.extension'
  | 'sharepoint.spfx.applicationCustomizer'
  | 'sharepoint.spfx.fieldCustomizer'
  | 'sharepoint.spfx.commandSet'
  | 'sharepoint.spfx.formCustomizer'
  | 'sharepoint.spfx.library'
  | 'sharepoint.ace'
  | 'sharepoint.theme'
  | 'sharepoint.siteScript'
  | 'sharepoint.siteDesign'
  | 'sharepoint.listFormatting'
  | 'sharepoint.columnFormatting'
  | 'sharepoint.formFormatting'
  | 'sharepoint.provisioning'
  | 'sharepoint.solution';

export interface BuildOptions {
  target: BuildTarget;
  spfxVersion?: SPFxVersion;
  moduleType?: ModuleType;
  environment?: 'spo' | 'odsp';
  includeSource?: boolean;
  sourceOnly?: boolean;
  minify?: boolean;
  sourceMaps?: boolean;
  skipValidation?: boolean;
  runtime?: 'browser' | 'node' | 'webworker';
}

export interface BuildResult {
  success: boolean;
  sppkg?: Uint8Array;
  sourceZip?: Uint8Array;
  files: VFSFile[];
  deployment: DeploymentManifest;
  validation: ValidationResult;
  security: SecurityReport;
  compatibility: CompatibilityReport;
  bundle: BundleAnalysis;
  errors: BuildError[];
  warnings: BuildWarning[];
  duration: number;
}

// ---------------------------------------------------------------------------
// Validation
// ---------------------------------------------------------------------------

export interface ValidationResult {
  valid: boolean;
  errors: ValidationError[];
  warnings: ValidationWarning[];
  info: ValidationInfo[];
  summary: ValidationSummary;
}

export interface ValidationError {
  code: string;
  message: string;
  file?: string;
  line?: number;
  column?: number;
  severity: 'error' | 'critical';
  category: string;
  fix?: string;
}

export interface ValidationWarning {
  code: string;
  message: string;
  file?: string;
  category: string;
  suggestion?: string;
}

export interface ValidationInfo {
  code: string;
  message: string;
  category: string;
}

export interface ValidationSummary {
  total: number;
  errors: number;
  warnings: number;
  info: number;
  categories: Record<string, { errors: number; warnings: number }>;
}

// ---------------------------------------------------------------------------
// Security
// ---------------------------------------------------------------------------

export interface SecurityReport {
  passed: boolean;
  score: number;
  findings: SecurityFinding[];
  secrets: SecretFinding[];
  externalUrls: string[];
  permissions: PermissionAudit[];
  recommendations: string[];
}

export interface SecurityFinding {
  severity: 'critical' | 'high' | 'medium' | 'low' | 'info';
  category: string;
  rule: string;
  message: string;
  file?: string;
  line?: number;
  evidence?: string;
  fix?: string;
}

export interface SecretFinding {
  type: string;
  file: string;
  line: number;
  masked: string;
}

export interface PermissionAudit {
  permission: string;
  required: boolean;
  declared: boolean;
  actuallyUsed: boolean;
  risk: 'low' | 'medium' | 'high';
}

// ---------------------------------------------------------------------------
// Compatibility
// ---------------------------------------------------------------------------

export interface CompatibilityReport {
  compatible: boolean;
  targetVersion: SPFxVersion;
  issues: CompatibilityIssue[];
  recommendations: string[];
  deprecated: DeprecatedFeature[];
}

export interface CompatibilityIssue {
  severity: 'error' | 'warning' | 'info';
  component: string;
  message: string;
  requirement?: string;
}

export interface DeprecatedFeature {
  feature: string;
  since: SPFxVersion;
  replacement?: string;
}

// ---------------------------------------------------------------------------
// Bundle Analysis
// ---------------------------------------------------------------------------

export interface BundleAnalysis {
  totalSize: number;
  chunks: BundleChunk[];
  dependencies: BundleDependency[];
  duplicates: DuplicateDependency[];
  externals: string[];
  recommendations: string[];
}

export interface BundleChunk {
  name: string;
  size: number;
  modules: string[];
  isEntry: boolean;
}

export interface BundleDependency {
  name: string;
  version: string;
  size: number;
  type: 'production' | 'development' | 'peer' | 'optional';
}

export interface DuplicateDependency {
  name: string;
  versions: string[];
  totalSize: number;
  savings: number;
}

// ---------------------------------------------------------------------------
// Deployment Manifest
// ---------------------------------------------------------------------------

export interface DeploymentManifest {
  status: 'ready' | 'readyWithWarnings' | 'errors';
  artifact: string;
  artifactSize?: number;
  destination: string;
  requiresAdmin: boolean;
  permissions: DeploymentPermission[];
  provisioning: DeploymentProvisioning[];
  warnings: string[];
  instructions: DeploymentInstruction[];
  metadata: {
    generator: string;
    version: string;
    buildTime: string;
    spfxVersion: SPFxVersion;
  };
}

export interface DeploymentPermission {
  resource: string;
  permission: string;
  type: 'Delegated' | 'Application';
  requiresAdminApproval: boolean;
  description: string;
}

export interface DeploymentProvisioning {
  type: string;
  name: string;
  description?: string;
}

export interface DeploymentInstruction {
  step: number;
  action: string;
  target: string;
  notes?: string;
}

// ---------------------------------------------------------------------------
// Import/Export
// ---------------------------------------------------------------------------

export type ImportSource = 'spfx-zip' | 'sppkg' | 'source-directory' | 'codbsharepoint-json' | 'auto';

export interface ImportResult {
  success: boolean;
  ir: CODBIR;
  source: ImportSource;
  detectedVersion?: SPFxVersion;
  warnings: string[];
}

export interface ExportOptions {
  format: 'sppkg' | 'source-zip' | 'json' | 'all';
  target?: BuildTarget;
  spfxVersion?: SPFxVersion;
}

// ---------------------------------------------------------------------------
// Simulator
// ---------------------------------------------------------------------------

export interface SimulatorConfig {
  siteUrl?: string;
  siteTitle?: string;
  user?: SimulatorUser;
  theme?: ThemeColorPalette;
  lists?: SimulatorList[];
  libraries?: SimulatorLibrary[];
  webPartSize?: 'small' | 'medium' | 'large';
}

export interface SimulatorUser {
  displayName: string;
  email: string;
  loginName: string;
  id: string;
  isSiteAdmin: boolean;
  isSiteOwner: boolean;
  title: string;
  picture?: string;
}

export interface SimulatorList {
  title: string;
  items: Record<string, unknown>[];
  fields: SimulatorField[];
}

export interface SimulatorLibrary {
  title: string;
  files: SimulatorFile[];
}

export interface SimulatorFile {
  name: string;
  serverRelativeUrl: string;
  size: number;
  timeCreated: string;
  timeLastModified: string;
}

export interface SimulatorField {
  name: string;
  title: string;
  type: string;
}

// ---------------------------------------------------------------------------
// Tool API (for AI agents)
// ---------------------------------------------------------------------------

export interface ToolAPI {
  createSolution(config: Partial<SolutionConfig>): CODBIR;
  createWebPart(config: Partial<ComponentDefinition>): ComponentDefinition;
  createExtension(config: Partial<ExtensionDefinition>): ExtensionDefinition;
  createACE(config: Partial<ACEDefinition>): ACEDefinition;
  createList(config: Partial<ListDefinition>): ListDefinition;
  createLibrary(config: Partial<LibraryDefinition>): LibraryDefinition;
  createColumn(config: Partial<FieldDefinition>): FieldDefinition;
  createContentType(config: Partial<ContentTypeDefinition>): ContentTypeDefinition;
  addGraphPermission(ir: CODBIR, scope: string, type?: 'Delegated' | 'Application'): CODBIR;
  addPropertyPane(component: ComponentDefinition, pane: PropertyPaneDefinition): ComponentDefinition;
  addList(ir: CODBIR, list: ListDefinition): CODBIR;
  addLibrary(ir: CODBIR, library: LibraryDefinition): CODBIR;
  addColumn(ir: CODBIR, field: FieldDefinition): CODBIR;
  compile(ir: CODBIR, options?: Partial<BuildOptions>): Promise<BuildResult>;
  validate(ir: CODBIR): Promise<ValidationResult>;
  analyze(ir: CODBIR): Promise<AnalysisResult>;
  repair(ir: CODBIR, validation: ValidationResult): CODBIR;
  package(ir: CODBIR, options?: Partial<BuildOptions>): Promise<BuildResult>;
}

export interface AnalysisResult {
  framework: Framework;
  componentTypes: ComponentType[];
  requiredPermissions: GraphPermissionDefinition[];
  complexity: 'simple' | 'moderate' | 'complex';
  estimatedBundleSize: number;
  dependencies: string[];
  recommendations: string[];
}

// ---------------------------------------------------------------------------
// CDN / Distribution
// ---------------------------------------------------------------------------

export interface CODBSharePointConfig {
  version?: string;
  cdnUrl?: string;
  spfxVersion?: SPFxVersion;
  environment?: 'production' | 'development';
  wasm?: boolean;
  pwa?: boolean;
}

// ---------------------------------------------------------------------------
// Events
// ---------------------------------------------------------------------------

export type CODBEvent =
  | { type: 'build:start'; timestamp: number }
  | { type: 'build:progress'; stage: string; progress: number }
  | { type: 'build:complete'; result: BuildResult; duration: number }
  | { type: 'build:error'; error: Error }
  | { type: 'validate:start'; timestamp: number }
  | { type: 'validate:complete'; result: ValidationResult }
  | { type: 'compile:start'; timestamp: number }
  | { type: 'compile:complete'; files: VFSFile[] }
  | { type: 'package:start'; timestamp: number }
  | { type: 'package:complete'; sppkg: Uint8Array };

export type CODBEventHandler = (event: CODBEvent) => void;
