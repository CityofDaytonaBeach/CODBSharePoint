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
    // Lists & Items
    api('sp.lists', 'Lists', 'GET', '/_api/web/lists', 'SPHttpClient', 'supported', 'Enumerate all lists in the current web.'),
    api('sp.listByTitle', 'Get List By Title', 'GET', "/_api/web/lists/getbytitle('{list}')", 'SPHttpClient', 'supported', 'Retrieve a single list metadata by title.'),
    api('sp.listItems', 'List Items', 'GET', "/_api/web/lists/getbytitle('{list}')/items", 'SPHttpClient', 'supported', 'Read list items with optional OData query.'),
    api('sp.createItem', 'Create List Item', 'POST', "/_api/web/lists/getbytitle('{list}')/items", 'SPHttpClient', 'supported', 'Create list items. Requires request digest.'),
    api('sp.updateItem', 'Update List Item', 'PATCH', "/_api/web/lists/getbytitle('{list}')/items({id})", 'SPHttpClient', 'supported', 'Update a single list item by ID.'),
    api('sp.deleteItem', 'Delete List Item', 'DELETE', "/_api/web/lists/getbytitle('{list}')/items({id})", 'SPHttpClient', 'supported', 'Delete a list item by ID.'),
    api('sp.listItemFields', 'List Item Fields', 'GET', "/_api/web/lists/getbytitle('{list}')/fields", 'SPHttpClient', 'supported', 'Read field definitions for a list.'),
    api('sp.listItemAttachments', 'List Item Attachments', 'GET', "/_api/web/lists/getbytitle('{list}')/items({id})/AttachmentFiles", 'SPHttpClient', 'supported', 'Read attachments for a list item.'),

    // Files & Libraries
    api('sp.files', 'Library Files', 'GET', "/_api/web/getfolderbyserverrelativeurl('{folder}')/files", 'SPHttpClient', 'supported', 'Read files from a document library folder.'),
    api('sp.file', 'Get File', 'GET', "/_api/web/getfilebyserverrelativeurl('{url}')", 'SPHttpClient', 'supported', 'Read metadata for a single file.'),
    api('sp.fileContent', 'File Content', 'GET', "/_api/web/getfilebyserverrelativeurl('{url}')/$value", 'SPHttpClient', 'supported', 'Download raw file content.'),
    api('sp.folders', 'Folders', 'GET', "/_api/web/getfolderbyserverrelativeurl('{folder}')/folders", 'SPHttpClient', 'supported', 'List subfolders in a library folder.'),
    api('sp.createFolder', 'Create Folder', 'POST', "/_api/web/folders", 'SPHttpClient', 'supported', 'Create a folder in a document library.'),
    api('sp.fileCheckIn', 'Check In File', 'POST', "/_api/web/getfilebyserverrelativeurl('{url}')/CheckIn(comment,checkInType)", 'SPHttpClient', 'experimental', 'Check in a file to a library.'),
    api('sp.fileCheckOut', 'Check Out File', 'POST', "/_api/web/getfilebyserverrelativeurl('{url}')/Checkout()", 'SPHttpClient', 'experimental', 'Check out a file from a library.'),
    api('sp.fileVersions', 'File Versions', 'GET', "/_api/web/getfilebyserverrelativeurl('{url}')/versions", 'SPHttpClient', 'supported', 'Read version history for a file.'),

    // Search
    api('sp.search', 'SharePoint Search', 'GET', "/_api/search/query?querytext='{query}'", 'SPHttpClient', 'experimental', 'Full-text search across SharePoint content.'),
    api('sp.searchPost', 'SharePoint Search (POST)', 'POST', '/_api/search/postquery', 'SPHttpClient', 'experimental', 'Advanced search with POST body for complex queries.'),

    // Site & Web
    api('sp.web', 'Current Web', 'GET', '/_api/web', 'SPHttpClient', 'supported', 'Read current web metadata.'),
    api('sp.site', 'Current Site', 'GET', '/_api/site', 'SPHttpClient', 'supported', 'Read current site collection metadata.'),
    api('sp.webs', 'Subsites', 'GET', '/_api/web/webs', 'SPHttpClient', 'supported', 'List immediate subsites.'),
    api('sp.navigation', 'Navigation', 'GET', '/_api/web/navigation/quicklaunch', 'SPHttpClient', 'known', 'Read quick launch navigation nodes.'),
    api('sp.webLists', 'Web Lists', 'GET', '/_api/web/lists', 'SPHttpClient', 'supported', 'Lists in the current web.'),

    // User Profiles & People
    api('sp.userProfile', 'User Profile', 'GET', '/_api/SP.UserProfiles.PeopleManager/GetMyProperties', 'SPHttpClient', 'known', 'Read current user profile properties.'),
    api('sp.userProfileByAccount', 'User Profile By Account', 'GET', "/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='{account}'", 'SPHttpClient', 'known', 'Read profile for a specific user by account name.'),
    api('sp.userProfilePhoto', 'User Profile Photo', 'GET', '/_api/SP.UserProfiles.PeopleManager/GetMyProfilePicture', 'SPHttpClient', 'experimental', 'Read current user profile picture.'),

    // Taxonomy
    api('sp.taxonomyTermStore', 'Term Store', 'GET', '/_api/v2.1/termStore', 'SPHttpClient', 'experimental', 'Read term store metadata.'),
    api('sp.taxonomySets', 'Term Sets', 'GET', '/_api/v2.1/termStore/sets', 'SPHttpClient', 'experimental', 'List term sets.'),
    api('sp.taxonomyTerms', 'Terms', 'GET', "/_api/v2.1/termStore/sets/{setId}/terms", 'SPHttpClient', 'experimental', 'List terms in a term set.'),

    // Content Types
    api('sp.contentTypes', 'Content Types', 'GET', '/_api/web/contenttypes', 'SPHttpClient', 'supported', 'List available content types.'),
    api('sp.contentTypeById', 'Content Type By ID', 'GET', "/_api/web/contenttypes('{id}')", 'SPHttpClient', 'supported', 'Get a specific content type.'),

    // Web Parts & Pages
    api('sp.sitePages', 'Site Pages', 'GET', "/_api/web/lists/getbytitle('Site Pages')/items", 'SPHttpClient', 'supported', 'Read site pages.'),
    api('sp.publishPage', 'Publish Page', 'POST', "/_api/web/lists/getbytitle('Site Pages')/items({id})/Publish", 'SPHttpClient', 'experimental', 'Publish a site page.'),

    // Recycle Bin
    api('sp.recycleBin', 'Recycle Bin', 'GET', '/_api/web/RecycleBin', 'SPHttpClient', 'supported', 'List items in the recycle bin.'),
    api('sp.recycleBinRestore', 'Restore Recycle Bin Item', 'POST', '/_api/web/RecycleBin/RestoreByIds', 'SPHttpClient', 'experimental', 'Restore items from recycle bin by IDs.'),

    // Groups & Permissions
    api('sp.siteGroups', 'Site Groups', 'GET', '/_api/web/sitegroups', 'SPHttpClient', 'supported', 'List site groups.'),
    api('sp.siteUsers', 'Site Users', 'GET', '/_api/web/siteusers', 'SPHttpClient', 'supported', 'List site users.'),
    api('sp.roleAssignments', 'Role Assignments', 'GET', "/_api/web/lists/getbytitle('{list}')/roleassignments", 'SPHttpClient', 'known', 'Read role assignments for a list.'),

    // Regional Settings
    api('sp.regionalSettings', 'Regional Settings', 'GET', '/_api/web/RegionalSettings', 'SPHttpClient', 'known', 'Read regional settings (timezone, locale).'),

    // Features
    api('sp.siteFeatures', 'Site Features', 'GET', '/_api/web/features', 'SPHttpClient', 'experimental', 'List activated features.'),

    // Workflows
    api('sp.workflowSubscriptions', 'Workflow Subscriptions', 'GET', '/_api/web/WorkflowSubscriptions', 'SPHttpClient', 'experimental', 'List workflow subscriptions.')
  ],
  graph: [
    // ── Users ──────────────────────────────────────────────────────────
    api('graph.me', 'Current User', 'GET', '/me', 'MSGraphClient', 'supported', 'Read current user profile.', ['User.Read']),
    api('graph.meProfile', 'Current User Profile', 'GET', '/me/profile', 'MSGraphClient', 'supported', 'Read current user profile details.', ['User.Read']),
    api('graph.users', 'List Users', 'GET', '/users', 'MSGraphClient', 'supported', 'List or search users in the tenant.', ['User.Read.All']),
    api('graph.userById', 'Get User By ID', 'GET', '/users/{id}', 'MSGraphClient', 'supported', 'Read a specific user by ID.', ['User.Read.All']),
    api('graph.userManager', 'User Manager', 'GET', '/users/{id}/manager', 'MSGraphClient', 'supported', 'Read the manager of a user.', ['User.Read.All']),
    api('graph.userDirectReports', 'User Direct Reports', 'GET', '/users/{id}/directReports', 'MSGraphClient', 'supported', 'Read direct reports for a user.', ['User.Read.All']),
    api('graph.userPhoto', 'User Photo', 'GET', '/users/{id}/photo/$value', 'MSGraphClient', 'supported', 'Download user profile photo.', ['User.Read.All']),
    api('graph.userCalendar', 'User Calendar Events', 'GET', '/users/{id}/calendar/events', 'MSGraphClient', 'known', 'Read calendar events for a user.', ['Calendars.Read']),
    api('graph.meCalendar', 'My Calendar Events', 'GET', '/me/calendar/events', 'MSGraphClient', 'known', 'Read current user calendar events.', ['Calendars.Read']),
    api('graph.userMailFolders', 'User Mail Folders', 'GET', '/users/{id}/mailFolders', 'MSGraphClient', 'known', 'List mail folders for a user.', ['Mail.Read']),
    api('graph.userSendMail', 'Send Mail', 'POST', '/users/{id}/sendMail', 'MSGraphClient', 'known', 'Send mail on behalf of a user.', ['Mail.Send']),
    api('graph.userTodos', 'User To-Do Tasks', 'GET', '/users/{id}/todo/lists', 'MSGraphClient', 'known', 'Read To-Do task lists for a user.', ['Tasks.ReadWrite']),

    // ── Groups ─────────────────────────────────────────────────────────
    api('graph.groups', 'List Groups', 'GET', '/groups', 'MSGraphClient', 'supported', 'List or search Microsoft 365 groups.', ['Group.Read.All']),
    api('graph.groupById', 'Get Group By ID', 'GET', '/groups/{id}', 'MSGraphClient', 'supported', 'Read a specific group by ID.', ['Group.Read.All']),
    api('graph.groupMembers', 'Group Members', 'GET', '/groups/{id}/members', 'MSGraphClient', 'supported', 'Read members of a group.', ['Group.Read.All']),
    api('graph.groupOwners', 'Group Owners', 'GET', '/groups/{id}/owners', 'MSGraphClient', 'supported', 'Read owners of a group.', ['Group.Read.All']),
    api('graph.groupMemberOf', 'Group Member Of', 'GET', '/groups/{id}/memberOf', 'MSGraphClient', 'known', 'Read groups this group belongs to.', ['Group.Read.All']),
    api('graph.groupEvents', 'Group Events', 'GET', '/groups/{id}/events', 'MSGraphClient', 'known', 'Read group calendar events.', ['Group.Read.All', 'Calendars.Read']),
    api('graph.groupConversations', 'Group Conversations', 'GET', '/groups/{id}/conversations', 'MSGraphClient', 'known', 'Read group conversations.', ['Group.Read.All']),
    api('graph.groupDrive', 'Group Drive', 'GET', '/groups/{id}/drive', 'MSGraphClient', 'known', 'Read the group OneDrive.', ['Group.Read.All', 'Files.Read.All']),
    api('graph.createGroup', 'Create Group', 'POST', '/groups', 'MSGraphClient', 'known', 'Create a new Microsoft 365 group.', ['Group.ReadWrite.All']),

    // ── Sites ──────────────────────────────────────────────────────────
    api('graph.sites', 'List Sites', 'GET', '/sites', 'MSGraphClient', 'supported', 'List SharePoint sites.', ['Sites.Read.All']),
    api('graph.siteById', 'Get Site By ID', 'GET', '/sites/{id}', 'MSGraphClient', 'supported', 'Get a specific site.', ['Sites.Read.All']),
    api('graph.siteRoot', 'Root Site', 'GET', '/sites/root', 'MSGraphClient', 'supported', 'Get the root SharePoint site.', ['Sites.Read.All']),
    api('graph.siteLists', 'Site Lists', 'GET', '/sites/{siteId}/lists', 'MSGraphClient', 'supported', 'List lists on a site.', ['Sites.Read.All']),
    api('graph.siteListItems', 'Site List Items', 'GET', '/sites/{siteId}/lists/{listId}/items', 'MSGraphClient', 'supported', 'Read items from a site list.', ['Sites.Read.All']),
    api('graph.sitePages', 'Site Pages', 'GET', "/sites/{siteId}/lists('Site Pages')/items", 'MSGraphClient', 'known', 'Read site pages through Graph.', ['Sites.Read.All']),
    api('graph.siteSearch', 'Search Sites', 'GET', "/search/q '{@query}'", 'MSGraphClient', 'known', 'Search SharePoint sites.', ['Sites.Read.All']),

    // ── Drives & Files ─────────────────────────────────────────────────
    api('graph.drives', 'List Drives', 'GET', '/drives', 'MSGraphClient', 'supported', 'List available drives.', ['Files.Read.All']),
    api('graph.driveById', 'Get Drive By ID', 'GET', '/drives/{id}', 'MSGraphClient', 'supported', 'Get a specific drive.', ['Files.Read.All']),
    api('graph.driveRoot', 'Drive Root', 'GET', '/drives/{driveId}/root', 'MSGraphClient', 'supported', 'Get root folder of a drive.', ['Files.Read.All']),
    api('graph.driveChildren', 'Drive Children', 'GET', '/drives/{driveId}/root/children', 'MSGraphClient', 'supported', 'List children in drive root.', ['Files.Read.All']),
    api('graph.driveItem', 'Get Drive Item', 'GET', '/drives/{driveId}/items/{itemId}', 'MSGraphClient', 'supported', 'Get a specific drive item.', ['Files.Read.All']),
    api('graph.driveItemContent', 'Download Drive Item', 'GET', '/drives/{driveId}/items/{itemId}/content', 'MSGraphClient', 'supported', 'Download drive item content.', ['Files.Read.All']),
    api('graph.driveRecent', 'Recent Files', 'GET', '/me/drive/recent', 'MSGraphClient', 'known', 'Get recently accessed files.', ['Files.Read']),
    api('graph.driveShared', 'Shared Files', 'GET', '/me/drive/shared', 'MSGraphClient', 'known', 'Get files shared with me.', ['Files.Read']),
    api('graph.driveSearch', 'Search Drive', 'GET', "/drives/{driveId}/root/search(q='{query}')", 'MSGraphClient', 'known', 'Search within a drive.', ['Files.Read.All']),
    api('graph.uploadDriveItem', 'Upload Drive Item', 'PUT', '/drives/{driveId}/root:/{path}:/content', 'MSGraphClient', 'known', 'Upload a file to a drive.', ['Files.ReadWrite.All']),
    api('graph.createUploadSession', 'Create Upload Session', 'POST', '/drives/{driveId}/root:/{path}:/createUploadSession', 'MSGraphClient', 'known', 'Create session for large file upload.', ['Files.ReadWrite.All']),

    // ── Planner ────────────────────────────────────────────────────────
    api('graph.plannerPlans', 'Planner Plans', 'GET', '/planner/plans', 'MSGraphClient', 'known', 'List Planner plans.', ['Tasks.Read.All']),
    api('graph.plannerPlanById', 'Get Plan By ID', 'GET', '/planner/plans/{planId}', 'MSGraphClient', 'known', 'Get a specific Planner plan.', ['Tasks.Read.All']),
    api('graph.plannerTasks', 'Planner Tasks', 'GET', '/planner/plans/{planId}/tasks', 'MSGraphClient', 'known', 'List tasks in a plan.', ['Tasks.Read.All']),
    api('graph.plannerBuckets', 'Planner Buckets', 'GET', '/planner/plans/{planId}/buckets', 'MSGraphClient', 'known', 'List buckets in a plan.', ['Tasks.Read.All']),
    api('graph.plannerTaskDetails', 'Task Details', 'GET', '/planner/tasks/{taskId}/details', 'MSGraphClient', 'known', 'Get details for a Planner task.', ['Tasks.Read.All']),
    api('graph.plannerCreateTask', 'Create Task', 'POST', '/planner/tasks', 'MSGraphClient', 'known', 'Create a Planner task.', ['Tasks.ReadWrite.All']),

    // ── Teams ──────────────────────────────────────────────────────────
    api('graph.teams', 'List Teams', 'GET', '/me/joinedTeams', 'MSGraphClient', 'known', 'List teams the user belongs to.', ['Team.ReadBasic.All']),
    api('graph.teamById', 'Get Team', 'GET', '/teams/{teamId}', 'MSGraphClient', 'known', 'Get a specific team.', ['Team.ReadBasic.All']),
    api('graph.teamChannels', 'Team Channels', 'GET', '/teams/{teamId}/channels', 'MSGraphClient', 'known', 'List channels in a team.', ['Channel.ReadBasic.All']),
    api('graph.teamChannelMessages', 'Channel Messages', 'GET', '/teams/{teamId}/channels/{channelId}/messages', 'MSGraphClient', 'known', 'Read messages in a channel.', ['ChannelMessage.Read.All']),
    api('graph.teamSendMessage', 'Send Channel Message', 'POST', '/teams/{teamId}/channels/{channelId}/messages', 'MSGraphClient', 'known', 'Send a message to a channel.', ['ChannelMessage.Send']),
    api('graph.teamTabs', 'Team Tabs', 'GET', '/teams/{teamId}/channels/{channelId}/tabs', 'MSGraphClient', 'known', 'List tabs in a channel.', ['Tab.Read.All']),

    // ── Chats & Messages ───────────────────────────────────────────────
    api('graph.chats', 'List Chats', 'GET', '/me/chats', 'MSGraphClient', 'known', 'List 1:1 and group chats.', ['Chat.Read']),
    api('graph.chatMessages', 'Chat Messages', 'GET', '/chats/{chatId}/messages', 'MSGraphClient', 'known', 'Read messages in a chat.', ['Chat.Read']),
    api('graph.sendChatMessage', 'Send Chat Message', 'POST', '/chats/{chatId}/messages', 'MSGraphClient', 'known', 'Send a message in a chat.', ['Chat.ReadWrite']),

    // ── OneDrive & SharePoint Files ────────────────────────────────────
    api('graph.meDrive', 'My Drive', 'GET', '/me/drive', 'MSGraphClient', 'supported', 'Get the current user OneDrive.', ['Files.Read']),
    api('graph.meDriveRoot', 'My Drive Root', 'GET', '/me/drive/root', 'MSGraphClient', 'supported', 'Get root of my OneDrive.', ['Files.Read']),
    api('graph.meDriveRecent', 'My Recent Files', 'GET', '/me/drive/recent', 'MSGraphClient', 'known', 'Files recently accessed.', ['Files.Read']),
    api('graph.meDriveShared', 'My Shared Files', 'GET', '/me/drive/shared', 'MSGraphClient', 'known', 'Files shared with me.', ['Files.Read']),
    api('graph.sitesRootDrive', 'Site Default Drive', 'GET', '/sites/{siteId}/drive', 'MSGraphClient', 'supported', 'Get the default drive for a site.', ['Sites.Read.All']),

    // ── Identity & Governance ──────────────────────────────────────────
    api('graph.organization', 'Organization', 'GET', '/organization', 'MSGraphClient', 'known', 'Read tenant organization info.', ['Organization.Read.All']),
    api('graph.subscribedSkus', 'Subscribed SKUs', 'GET', '/subscribedSkus', 'MSGraphClient', 'known', 'Read subscribed SKUs/licenses.', ['Directory.Read.All']),
    api('graph.directoryRoles', 'Directory Roles', 'GET', '/directoryRoles', 'MSGraphClient', 'known', 'Read directory roles.', ['Directory.Read.All']),
    api('graph.servicePrincipals', 'Service Principals', 'GET', '/servicePrincipals', 'MSGraphClient', 'known', 'List service principals.', ['Application.Read.All']),

    // ── Excel ──────────────────────────────────────────────────────────
    api('graph.excelSession', 'Open Excel Session', 'POST', '/me/drive/items/{itemId}/workbook/createSession', 'MSGraphClient', 'known', 'Create an Excel session for batch operations.', ['Sites.Read.All']),
    api('graph.excelRange', 'Read Excel Range', 'GET', "/me/drive/items/{itemId}/workbook/worksheets('{name}')/usedRange", 'MSGraphClient', 'known', 'Read used range from an Excel worksheet.', ['Sites.Read.All']),

    // ── To-Do ──────────────────────────────────────────────────────────
    api('graph.todoLists', 'To-Do Lists', 'GET', '/me/todo/lists', 'MSGraphClient', 'known', 'List To-Do task lists.', ['Tasks.ReadWrite']),
    api('graph.todoTasks', 'To-Do Tasks', 'GET', '/me/todo/lists/{listId}/tasks', 'MSGraphClient', 'known', 'List tasks in a To-Do list.', ['Tasks.ReadWrite']),

    // ── Bookings ───────────────────────────────────────────────────────
    api('graph.bookings', 'Bookings Businesses', 'GET', '/solutions/bookingBusinesses', 'MSGraphClient', 'experimental', 'List Bookings businesses.', ['Bookings.Read.All']),
    api('graph.bookingsAppointments', 'Bookings Appointments', 'GET', "/solutions/bookingBusinesses/{id}/appointments", 'MSGraphClient', 'experimental', 'List appointments in a Bookings business.', ['Bookings.Read.All']),

    // ── Education ──────────────────────────────────────────────────────
    api('graph.educationClasses', 'Education Classes', 'GET', '/education/classes', 'MSGraphClient', 'experimental', 'List education classes.', ['EduClass.Read.All']),
    api('graph.educationUsers', 'Education Users', 'GET', '/education/users', 'MSGraphClient', 'experimental', 'List education users.', ['EduUser.Read.All']),

    // ── Security & Compliance ──────────────────────────────────────────
    api('graph.securityAlerts', 'Security Alerts', 'GET', '/security/alerts', 'MSGraphClient', 'experimental', 'Read security alerts.', ['SecurityEvents.Read.All']),
    api('graph.identityRiskyUsers', 'Risky Users', 'GET', '/identityProtection/riskyUsers', 'MSGraphClient', 'experimental', 'Read risky users.', ['IdentityRiskyUser.Read.All'])
  ],
  schema: [
    // Fields
    schema('field.text', 'Text Field', 'field', 'Text', 'supported', 'Single line of text. Max 255 chars.'),
    schema('field.note', 'Multiline Text', 'field', 'Note', 'supported', 'Multiple lines of text (plain or rich text).'),
    schema('field.number', 'Number', 'field', 'Number', 'supported', 'Numeric field with optional min/max/decimals.'),
    schema('field.currency', 'Currency', 'field', 'Currency', 'supported', 'Currency field with locale-aware formatting.'),
    schema('field.datetime', 'Date Time', 'field', 'DateTime', 'supported', 'Date or date/time field.'),
    schema('field.choice', 'Choice', 'field', 'Choice', 'supported', 'Single-choice dropdown.'),
    schema('field.multiChoice', 'Multi-Choice', 'field', 'MultiChoice', 'supported', 'Multi-select choice field.'),
    schema('field.boolean', 'Yes/No', 'field', 'Boolean', 'supported', 'Boolean yes/no toggle.'),
    schema('field.url', 'Hyperlink', 'field', 'URL', 'supported', 'URL field (hyperlink or picture).'),
    schema('field.lookup', 'Lookup', 'field', 'Lookup', 'experimental', 'Lookup to another list. Requires indexed fields.'),
    schema('field.user', 'Person or Group', 'field', 'User', 'supported', 'Person or group picker field.'),
    schema('field.userMulti', 'People (Multi)', 'field', 'UserMulti', 'supported', 'Multi-select people picker.'),
    schema('field.taxonomy', 'Managed Metadata', 'field', 'Taxonomy', 'experimental', 'Requires term set ID and hidden taxonomy fields.'),
    schema('field.taxonomyMulti', 'Managed Metadata (Multi)', 'field', 'TaxonomyMulti', 'experimental', 'Multi-select managed metadata.'),
    schema('field.calculated', 'Calculated', 'field', 'Calculated', 'experimental', 'Calculated field from other columns.'),
    schema('field.attachments', 'Attachments', 'field', 'Attachments', 'supported', 'Attachment indicator field.'),
    schema('field.modStat', 'Moderation Status', 'field', 'ModStat', 'known', 'Content approval moderation status.'),

    // List Templates
    schema('list.generic', 'Custom List', 'listTemplate', 100, 'supported', 'Generic custom list.'),
    schema('list.documentLibrary', 'Document Library', 'listTemplate', 101, 'supported', 'Document library with versioning.'),
    schema('list wikiPageLibrary', 'Wiki Page Library', 'listTemplate', 1100, 'known', 'Wiki page library.'),
    schema('list.tasks', 'Tasks', 'listTemplate', 107, 'known', 'Task list with assignments and dates.'),
    schema('list.calendar', 'Calendar', 'listTemplate', 106, 'known', 'Calendar list (legacy).'),
    schema('list.contacts', 'Contacts', 'listTemplate', 102, 'known', 'Contact list.'),
    schema('list.issueTracking', 'Issue Tracking', 'listTemplate', 1100, 'known', 'Issue tracking list.'),
    schema('list.links', 'Links', 'listTemplate', 103, 'known', 'Links list.'),
    schema('list.announcements', 'Announcements', 'listTemplate', 104, 'known', 'Announcements list.'),
    schema('list.discussionBoard', 'Discussion Board', 'listTemplate', 108, 'known', 'Discussion board list.'),
    schema('list.survey', 'Survey', 'listTemplate', 109, 'known', 'Survey list.'),
    schema('list.xmlForm', 'XML Form', 'listTemplate', 115, 'known', 'XML form library.'),
    schema('list.dataConnectionLibrary', 'Data Connection Library', 'listTemplate', 150, 'known', 'Data connection library.'),
    schema('list.translationLibrary', 'Translation Library', 'listTemplate', 161, 'known', 'Translation library.'),
    schema('list.pictureLibrary', 'Picture Library', 'listTemplate', 109, 'known', 'Picture library.'),

    // Content Types
    schema('ct.item', 'Item Content Type', 'contentType', '0x01', 'supported', 'Base Item content type.'),
    schema('ct.document', 'Document Content Type', 'contentType', '0x0101', 'supported', 'Base Document content type.'),
    schema('ct.folder', 'Folder Content Type', 'contentType', '0x0120', 'supported', 'Base Folder content type.'),
    schema('ct.message', 'Message Content Type', 'contentType', '0x0107', 'known', 'Message content type.'),
    schema('ct.task', 'Task Content Type', 'contentType', '0x0108', 'known', 'Task content type.'),
    schema('ct.event', 'Event Content Type', 'contentType', '0x0102', 'known', 'Event content type.'),
    schema('ct.contact', 'Contact Content Type', 'contentType', '0x0106', 'known', 'Contact content type.'),

    // Formatting
    schema('format.column', 'Column Formatting', 'formatting', 'columnFormatting', 'supported', 'JSON column formatting for list views.'),
    schema('format.view', 'View Formatting', 'formatting', 'viewFormatting', 'supported', 'JSON view formatting for list views.'),
    schema('format.form', 'Form Formatting', 'formatting', 'formFormatting', 'experimental', 'JSON form layout customization.')
  ],
  spfxRuntime: [
    entry('runtime.context', 'Web Part Context', 'supported', 'Page context with web URL, user, spHttpClient, msGraphClientFactory.'),
    entry('runtime.spHttpClient', 'SPHttpClient', 'supported', 'Native HTTP client for SharePoint REST API calls.'),
    entry('runtime.msGraphClient', 'MSGraphClient', 'supported', 'Native HTTP client for Microsoft Graph via SPFx AadTokenProvider.'),
    entry('runtime.aadHttpClient', 'AadHttpClient', 'known', 'HTTP client for custom AAD-protected APIs.'),
    entry('runtime.serviceScope', 'Service Scope', 'supported', 'Dependency injection and service sharing between components.'),
    entry('runtime.themeProvider', 'Theme Provider', 'known', 'Read SharePoint theme colors and semantic slots.'),
    entry('runtime.dynamicData', 'Dynamic Data', 'known', 'Connect web parts as data source/consumer via dynamic data.'),
    entry('runtime.placeholder', 'Placeholder', 'known', 'Render content in designated page zones.'),
    entry('runtime.spListPicker', 'List Picker', 'experimental', 'Native list/library picker control.'),
    entry('runtime.propertyPane', 'Property Pane', 'supported', 'Web part property pane with typed controls.'),
    entry('runtime.baseComponent', 'BaseClientSideWebPart', 'supported', 'Base class for all client-side web parts.'),
    entry('runtime.baseExtension', 'BaseExtension', 'known', 'Base class for application customizers and other extensions.'),
    entry('runtime.environment', 'Environment Detection', 'known', 'Detect local/workbench/SharePoint environment.'),
    entry('runtime.openNavigate', 'Navigation', 'known', 'SPFx navigateTo, chromatNavigation, and page header.'),
    entry('runtime.graphHttpClient', 'GraphHttpClient', 'experimental', 'Legacy Graph HTTP client (prefer MSGraphClient).')
  ],
  componentTypes: [
    entry('component.webpart.react', 'React Web Part', 'supported', 'Class or function component with React rendering.'),
    entry('component.webpart.reactFunction', 'React Function Web Part', 'experimental', 'Function component with hooks. Works but class components more common.'),
    entry('component.webpart.vanilla', 'Vanilla Web Part', 'supported', 'No framework. Direct DOM manipulation.'),
    entry('component.applicationCustomizer', 'Application Customizer', 'experimental', 'Inject scripts into header/footer placeholders.'),
    entry('component.fieldCustomizer', 'Field Customizer', 'experimental', 'Customize list cell rendering with JavaScript.'),
    entry('component.commandSet', 'ListView Command Set', 'experimental', 'Add custom command bar/context menu buttons.'),
    entry('component.formCustomizer', 'Form Customizer', 'experimental', 'Customize new/edit/display forms.'),
    entry('component.ace', 'Adaptive Card Extension', 'known', 'Viva Connections card with quick view.')
  ],
  permissions: [
    // Graph Delegated
    entry('graphperm.userRead', 'User.Read', 'supported', 'Sign in and read user profile. Basic user context.', ['User.Read']),
    entry('graphperm.userReadWrite', 'User.ReadWrite', 'supported', 'Read and write user profile. Changes to own profile.', ['User.ReadWrite']),
    entry('graphperm.userReadAll', 'User.Read.All', 'supported', 'Read all user profiles. Requires admin consent.', ['User.Read.All']),
    entry('graphperm.userReadWriteAll', 'User.ReadWrite.All', 'supported', 'Read and write all user profiles. Requires admin consent.', ['User.ReadWrite.All']),
    entry('graphperm.groupReadAll', 'Group.Read.All', 'supported', 'Read all groups and memberships. Requires admin consent.', ['Group.Read.All']),
    entry('graphperm.groupReadWriteAll', 'Group.ReadWrite.All', 'supported', 'Create and manage groups. Requires admin consent.', ['Group.ReadWrite.All']),
    entry('graphperm.filesRead', 'Files.Read', 'supported', 'Read user files. Delegated to signed-in user.', ['Files.Read']),
    entry('graphperm.filesReadWrite', 'Files.ReadWrite', 'supported', 'Read and write user files. Delegated to signed-in user.', ['Files.ReadWrite']),
    entry('graphperm.filesReadAll', 'Files.Read.All', 'supported', 'Read all files across the tenant. Requires admin consent.', ['Files.Read.All']),
    entry('graphperm.filesReadWriteAll', 'Files.ReadWrite.All', 'supported', 'Read and write all files. Requires admin consent.', ['Files.ReadWrite.All']),
    entry('graphperm.sitesReadAll', 'Sites.Read.All', 'supported', 'Read SharePoint sites and lists via Graph. Requires admin consent.', ['Sites.Read.All']),
    entry('graphperm.sitesReadWriteAll', 'Sites.ReadWrite.All', 'supported', 'Read and write SharePoint sites via Graph. Requires admin consent.', ['Sites.ReadWrite.All']),
    entry('graphperm.mailRead', 'Mail.Read', 'supported', 'Read user mail. Delegated to signed-in user.', ['Mail.Read']),
    entry('graphperm.mailSend', 'Mail.Send', 'supported', 'Send mail on behalf of user. Requires admin consent for app-only.', ['Mail.Send']),
    entry('graphperm.calendarsRead', 'Calendars.Read', 'supported', 'Read user calendars. Delegated.', ['Calendars.Read']),
    entry('graphperm.calendarsReadWrite', 'Calendars.ReadWrite', 'supported', 'Read and write calendars. Delegated.', ['Calendars.ReadWrite']),
    entry('graphperm.tasksReadWrite', 'Tasks.ReadWrite', 'supported', 'Read and write To-Do tasks. Delegated.', ['Tasks.ReadWrite']),
    entry('graphperm.tasksReadAll', 'Tasks.Read.All', 'supported', 'Read all tasks across the tenant. Requires admin consent.', ['Tasks.Read.All']),
    entry('graphperm.tasksReadWriteAll', 'Tasks.ReadWrite.All', 'supported', 'Read and write all tasks. Requires admin consent.', ['Tasks.ReadWrite.All']),
    entry('graphperm.channelMessageRead', 'ChannelMessage.Read.All', 'supported', 'Read channel messages. Requires admin consent.', ['ChannelMessage.Read.All']),
    entry('graphperm.channelMessageSend', 'ChannelMessage.Send', 'supported', 'Send channel messages. Requires admin consent.', ['ChannelMessage.Send']),
    entry('graphperm.chatRead', 'Chat.Read', 'supported', 'Read 1:1 and group chats. Requires admin consent.', ['Chat.Read']),
    entry('graphperm.chatReadWrite', 'Chat.ReadWrite', 'supported', 'Read and write chats. Requires admin consent.', ['Chat.ReadWrite']),
    entry('graphperm.teamReadBasic', 'Team.ReadBasic.All', 'supported', 'Read team names and channels. Requires admin consent.', ['Team.ReadBasic.All']),
    entry('graphperm.channelReadBasic', 'Channel.ReadBasic.All', 'supported', 'Read channel names and descriptions.', ['Channel.ReadBasic.All']),
    entry('graphperm.orgReadAll', 'Organization.Read.All', 'supported', 'Read org info and subscriptions. Requires admin consent.', ['Organization.Read.All']),
    entry('graphperm.directoryReadAll', 'Directory.Read.All', 'supported', 'Read directory objects. Requires admin consent.', ['Directory.Read.All']),
    entry('graphperm.directoryReadWriteAll', 'Directory.ReadWrite.All', 'supported', 'Read and write directory objects. Requires admin consent.', ['Directory.ReadWrite.All']),
    entry('graphperm.applicationReadAll', 'Application.Read.All', 'supported', 'Read service principals. Requires admin consent.', ['Application.Read.All']),
    entry('graphperm.securityEventsRead', 'SecurityEvents.Read.All', 'supported', 'Read security alerts. Requires admin consent.', ['SecurityEvents.Read.All']),
    entry('graphperm.bookingsReadAll', 'Bookings.Read.All', 'experimental', 'Read Bookings businesses. Requires admin consent.', ['Bookings.Read.All']),
    entry('graphperm.sharepointTenantFull', 'Sites.FullControl.All', 'supported', 'Full control of all SharePoint sites. Requires admin consent.', ['Sites.FullControl.All']),

    // SharePoint
    entry('sp permission.site', 'SharePoint Site Permissions', 'known', 'SharePoint REST is bounded by the current user context and deployed app permissions.'),
    entry('sp permission.tenantAppCatalog', 'Tenant App Catalog Deploy', 'supported', 'Requires SiteCreator or FullControl on the app catalog site.')
  ],
  m365Integrations: [
    entry('m365.teams.tabs', 'Teams Tabs', 'known', 'SPFx web parts can be added as Teams static/custom tabs via manifest.'),
    entry('m365.teams.personalApp', 'Teams Personal App', 'known', 'SPFx web parts can be exposed as Teams personal apps.'),
    entry('m365.teams.messagingExtension', 'Teams Messaging Extension', 'experimental', 'Create messaging extensions from SPFx components.'),
    entry('m365.viva.connections', 'Viva Connections', 'known', 'ACE for dashboard, global nav, and mobile cards.'),
    entry('m365.viva.topics', 'Viva Topics', 'experimental', 'Topic experiences integration with SharePoint pages.'),
    entry('m365.viva.learning', 'Viva Learning', 'experimental', 'Learning content integration.'),
    entry('m365.planner.plans', 'Planner Plans', 'known', 'Graph Planner APIs for plan and task management.'),
    entry('m365.planner.buckets', 'Planner Buckets', 'known', 'Organize tasks into buckets via Graph.'),
    entry('m365.onedrive.files', 'OneDrive Files', 'known', 'Graph OneDrive/drive APIs for file management.'),
    entry('m365.onedrive.shared', 'OneDrive Shared', 'known', 'Shared files via Graph.'),
    entry('m365.outlook.mail', 'Outlook Mail', 'known', 'Read/send mail via Graph.'),
    entry('m365.outlook.calendar', 'Outlook Calendar', 'known', 'Read/manage calendar events via Graph.'),
    entry('m365.outlook.tasks', 'Outlook Tasks', 'known', 'To-Do/task management via Graph.'),
    entry('m365.powerAutomate.flows', 'Power Automate Flows', 'known', 'Trigger flows from SPFx via HTTP actions.'),
    entry('m365.powerAutomate.approvals', 'Power Automate Approvals', 'known', 'Approval workflows from SPFx components.'),
    entry('m365.azure.functions', 'Azure Functions', 'known', 'Custom backend APIs via AadHttpClient.'),
    entry('m365.azure.logicApps', 'Azure Logic Apps', 'known', 'HTTP-triggered logic apps.'),
    entry('m365.copilot.plugins', 'Copilot Plugins', 'experimental', 'Extend Microsoft Copilot with custom actions.'),
    entry('m365.copilot.declarativeAgents', 'Declarative Agents', 'experimental', 'Custom Copilot agents with SharePoint knowledge.'),
    entry('m365.loop.components', 'Loop Components', 'experimental', 'Embed Loop components in SPFx pages.')
  ],
  pnpDependencies: [
    dep('@pnp/sp', 'PnPjs SharePoint', 'browser-prebundle-required', 'known', 'TypeScript client for SharePoint REST. Needs browser prebundle.'),
    dep('@pnp/graph', 'PnPjs Graph', 'browser-prebundle-required', 'known', 'TypeScript client for Graph. Replace with MSGraphClient.'),
    dep('@pnp/spfx-controls-react', 'PnP React Controls', 'browser-prebundle-required', 'known', 'PeoplePicker, FilePicker, ListView, etc.'),
    dep('@pnp/spfx-property-controls', 'PnP Property Controls', 'browser-prebundle-required', 'known', 'Property pane controls (webPartPicker, taxPicker, etc.).'),
    dep('@pnp/spfx-provisioning', 'PnP Provisioning', 'browser-prebundle-required', 'experimental', 'Provisioning engine for site templates.'),
    dep('@pnp/logging', 'PnP Logging', 'browser-prebundle-required', 'known', 'Logging framework. Lightweight, often bundled.'),
    dep('@pnp/spashing', 'PnP Spashing', 'browser-prebundle-required', 'experimental', 'SharePoint migration assessment.')
  ],
  fluentUI: [
    dep('@fluentui/react', 'Fluent UI React v8', 'browser-prebundle-required', 'known', 'Primary UI library for SPFx web parts.'),
    dep('@fluentui/react-components', 'Fluent UI React v9', 'browser-prebundle-required', 'experimental', 'Newer Fluent UI. Check SPFx version compatibility.'),
    dep('office-ui-fabric-react', 'Office UI Fabric React', 'browser-prebundle-required', 'known', 'Legacy name for Fluent UI v8.'),
    dep('@fluentui/react-theme-provider', 'Fluent Theme Provider', 'browser-prebundle-required', 'known', 'Theme injection for Fluent UI.'),
    dep('@fluentui/date-time-utilities', 'Date Time Utilities', 'browser-prebundle-required', 'known', 'Date utilities used by calendar/date picker components.'),
    dep('@fluentui/react-hooks', 'Fluent React Hooks', 'browser-prebundle-required', 'known', 'Common hooks for Fluent UI components.')
  ],
  deployment: [
    entry('deploy.tenantAppCatalog', 'Tenant App Catalog', 'known', 'Central app catalog for the entire tenant.'),
    entry('deploy.siteCollectionAppCatalog', 'Site Collection App Catalog', 'known', 'Scoped app catalog for a site collection.'),
    entry('deploy.skipFeatureDeployment', 'Tenant-wide Deployment', 'supported', 'Deploy to all sites without site admin approval.'),
    entry('deploy.includeClientSideAssets', 'Include Client Side Assets', 'supported', 'Bundle client-side assets in the SPPKG.'),
    entry('deploy.apiApproval', 'API Permission Approval', 'supported', 'Admin approves Graph permissions in SharePoint admin center.'),
    entry('deploy.isolatedDomain', 'Isolated Domain', 'known', 'Run web part code in an isolated domain for enhanced security.'),
    entry('deploy.versionUpgrades', 'Version Upgrades', 'supported', 'Upgrade SPPKG versions with compatible manifests.'),
    entry('deploy.lazyLoading', 'Lazy Loading', 'supported', 'Components load on demand from CDN/SharePoint.'),
    entry('deploy.cdnExternalComponents', 'CDN External Components', 'known', 'Host large dependencies on Azure CDN or public CDN.'),
    entry('deploy.appOnlyPermissions', 'App-Only Permissions', 'known', 'Application permissions for background/daemon services.'),
    entry('deploy.storeSubmission', 'AppSource Submission', 'experimental', 'Submit to Microsoft AppSource marketplace.'),
    entry('deploy.multiTenant', 'Multi-Tenant Deployment', 'experimental', 'Deploy across multiple tenant instances.')
  ],
  sourcePatterns: [
    pattern('pattern.employeeDirectory', 'Employee Directory', 'supported', ['React Web Part'], ['Microsoft Graph /users'], ['User.Read.All']),
    pattern('pattern.quickLinks', 'Quick Links', 'supported', ['React Web Part'], ['Static links', 'SharePoint list optional'], []),
    pattern('pattern.dataTable', 'Data Table', 'supported', ['React Web Part'], ['SharePoint list items'], []),
    pattern('pattern.documentExplorer', 'Document Explorer', 'known', ['React Web Part'], ['SharePoint library files', 'Graph drives'], ['Files.Read.All']),
    pattern('pattern.approvals', 'Approval Dashboard', 'known', ['React Web Part'], ['SharePoint list', 'Power Automate optional'], []),
    pattern('pattern.orgChart', 'Organization Chart', 'known', ['React Web Part'], ['Graph /users/manager', '/directReports'], ['User.Read.All']),
    pattern('pattern.calendar', 'Calendar View', 'known', ['React Web Part'], ['Graph calendar', 'SharePoint Events list'], ['Calendars.Read']),
    pattern('pattern.faq', 'FAQ Accordion', 'supported', ['React Web Part'], ['Static data', 'SharePoint list'], []),
    pattern('pattern.newsFeed', 'News Feed', 'known', ['React Web Part'], ['SharePoint Site Pages list'], []),
    pattern('pattern.taskBoard', 'Task Board', 'known', ['React Web Part'], ['Graph Planner tasks'], ['Tasks.Read.All']),
    pattern('pattern.mailbox', 'Mailbox Viewer', 'known', ['React Web Part'], ['Graph /me/mailFolders', '/messages'], ['Mail.Read']),
    pattern('pattern.fileManager', 'File Manager', 'known', ['React Web Part'], ['Graph drives', '/root/children'], ['Files.Read.All']),
    pattern('pattern.searchResults', 'Search Results', 'experimental', ['React Web Part'], ['SharePoint Search API', 'Graph search'], []),
    pattern('pattern.imageGallery', 'Image Gallery', 'supported', ['React Web Part'], ['SharePoint Picture Library', 'Static images'], []),
    pattern('pattern.announcements', 'Announcements', 'supported', ['React Web Part'], ['SharePoint Announcements list'], []),
    pattern('pattern.kpiDashboard', 'KPI Dashboard', 'known', ['React Web Part'], ['SharePoint lists', 'Graph data'], []),
    pattern('pattern.contactCard', 'Contact Card', 'supported', ['React Web Part'], ['Graph /users/{id}'], ['User.Read.All']),
    pattern('pattern.eventRegistration', 'Event Registration', 'known', ['React Web Part'], ['SharePoint Events list', 'Registrations list'], []),
    pattern('pattern.leaveRequest', 'Leave Request', 'known', ['React Web Part'], ['SharePoint list', 'Power Automate'], []),
    pattern('pattern.inventoryTracker', 'Inventory Tracker', 'known', ['React Web Part'], ['SharePoint list with barcodes'], []),
    pattern('pattern.recipeBook', 'Recipe Book', 'known', ['React Web Part'], ['SharePoint list with images'], []),
    pattern('pattern.portfolio', 'Portfolio', 'known', ['React Web Part'], ['SharePoint library', 'Metadata'], []),
    pattern('pattern.surveyForm', 'Survey Form', 'known', ['React Web Part'], ['SharePoint list', 'Branching logic'], []),
    pattern('pattern.ticketing', 'Ticketing System', 'known', ['React Web Part'], ['SharePoint list', 'Status workflow'], []),
    pattern('pattern.weatherWidget', 'Weather Widget', 'experimental', ['React Web Part'], ['External API', 'Geolocation'], [])
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
