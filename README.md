# CODBSharePoint

Browser-native SharePoint compiler, validator, and packaging SDK.

Generate, validate, compile, and package SPFx solutions entirely in the browser without Node.js, Yeoman, Gulp, or any local SPFx toolchain.

## Credits & Ownership

Built and maintained by the **City of Daytona Beach**.

- **Lead Developer:** Daniel Gurczynski
- **Technology:** City of Daytona Beach Information Technology

> This is an internal City of Daytona Beach project designed to make SharePoint
> modern site customization fast, safe, and approachable for staff — no coding
> toolchain or approval backlog required.

## Why This Project Exists (Purpose)

SharePoint modern sites are customized through SPFx (SharePoint Framework)
web parts, extensions, and ACEs. Historically that demanded:

- A full **Node.js / Yeoman / Gulp / TypeScript** local toolchain on each machine
- A **CI/CD or backend pipeline** to build and package solutions
- Deep **developer knowledge** to validate, secure, and deploy safely

That made even small requests slow, risky, and dependent on infrastructure.

**CODBSharePoint solves this by moving the entire pipeline into the browser:**
analyze, validate, compile, bundle, and package an `.sppkg` — or even author
and persist a whole project — with no backend and no local developer stack.
It gives non-developers (and AI agents) a safe, repeatable path from idea to
deployed SharePoint component in minutes.

## How to Use It

### 1. Quick (one-shot) build — great for a single component

```javascript
import { CODBSharePoint } from 'codbsharepoint';
const sdk = new CODBSharePoint();

const result = await sdk.build({
  name: 'EmployeeDirectory',
  framework: 'react',
  solution: { name: 'EmployeeDirectory', version: '1.0.0', description: 'Employee directory' },
  components: [{ name: 'EmployeeDirectory', displayName: 'Employee Directory', framework: 'react' }],
  graph: ['User.Read.All']
});

if (result.sppkg) {
  const url = URL.createObjectURL(new Blob([result.sppkg], { type: 'application/octet-stream' }));
  const a = document.createElement('a');
  a.href = url;
  a.download = 'EmployeeDirectory.sppkg';
  a.click();
}
```

### 2. Author a full project (designer) — for iterative, persistent authoring

The **Designer** is the authoring-first core: create a project, add components,
templates, lists, themes, and permissions, then save/load and publish — all
serverless.

```javascript
import { CODBSharePoint } from 'codbsharepoint';
const sdk = new CODBSharePoint();

const designer = sdk.designer({ storage: 'indexeddb' }); // or 'memory' / 'localstorage'

// Create and author the project
await designer.create('EmployeeDirectory', { description: 'HR web part', spfxVersion: '1.18' });
await designer.addWebPart({ name: 'EmployeeDirectory', framework: 'react' });
await designer.addACE({ name: 'QuickLinks' });
await designer.addList({ name: 'Employees' });
await designer.addGraphPermission('User.Read.All');

// Persist (survives reload via the storage adapter)
await designer.save();

// Build into an .sppkg + source
const buildResult = await designer.build();

// Publish static artifacts (offline standalone host) for any static file host
const publishResult = await designer.publish({ title: 'EmployeeDirectory' });
```

### 3. Publish a standalone (serverless) host

```javascript
const { generateStaticPublish } = sdk.publishAPI;
const result = generateStaticPublish(ir, buildResult, { title: 'Preview' });
// result.files -> ['index.html', 'bundle/bundle.js', 'publish.json', ...]
```

Drop `result.files` on GitHub Pages, S3, or Cloudflare R2 and open `index.html`
to preview your components with no SharePoint or Node backend.

### 4. Integrate with an AI agent

```javascript
const tools = sdk.tools({ name: 'MySolution' });
tools.createWebPart({ name: 'EmployeeDirectory', framework: 'react' });
await tools.addGraphPermission(ir, 'User.Read.All');
const result = await tools.compile(ir);
```

## Installation

```html
<!-- CDN -->
<script type="module">
  import { CODBSharePoint } from 'https://cdn.jsdelivr.net/gh/CityofDaytonaBeach/codbsharepoint@v1.0.0/dist/codbsharepoint.mjs';
</script>
```

```bash
# npm
npm install codbsharepoint
```

## Quick Start

```javascript
import { CODBSharePoint } from 'codbsharepoint';

const sdk = new CODBSharePoint();

// Build a SharePoint web part
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
  ],
  graph: ['User.Read.All']
});

// Download the .sppkg file
if (result.sppkg) {
  const blob = new Blob([result.sppkg], { type: 'application/octet-stream' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'EmployeeDirectory.sppkg';
  a.click();
}
```

## Features

- **SPFx Web Parts** - Compile React/Vanilla JS web parts
- **Extensions** - Application Customizers, Field Customizers, Command Sets
- **Adaptive Card Extensions** - ACE components for Viva Connections
- **SPPKG Generation** - Create valid .sppkg packages
- **Source Project Export** - Download complete SPFx source projects
- **Validation** - 17-point validation suite
- **Security Scanner** - Detect secrets, XSS, code injection
- **Permission Analysis** - Auto-detect required Graph permissions
- **Compatibility Check** - SPFx version compatibility matrix
- **Bundle Analysis** - Size analysis and optimization recommendations
- **SharePoint Simulator** - Preview solutions in-browser
- **Import/Export** - Round-trip existing SPFx projects
- **Tool API** - Programmatic interface for AI agents
- **PWA Support** - Works offline once cached
- **Designer** - Authoring-first project service (create/save/load/build/publish)
- **Serverless Storage** - Memory, localStorage, and IndexedDB adapters
- **Template Registry** - React, vanilla, extensions, and ACE scaffolds
- **Static Publish** - Standalone offline host + deployable artifacts
- **Offline Bundling** - Real JSX/TSX compile via in-browser esbuild (wasm)

## API Reference

### `CODBSharePoint.build(spec)`
Build a SharePoint solution from specification.

### `CODBSharePoint.validate(ir)`
Validate a project with 17-point validation suite.

### `CODBSharePoint.analyze(ir)`
Analyze framework, permissions, and complexity.

### `CODBSharePoint.import(data)`
Import existing SPFx projects or SPPKG files.

### `CODBSharePoint.export(ir, result)`
Export with deployment artifacts and guides.

### `CODBSharePoint.tools(config)`
Get the Tool API for AI agents.

### `CODBSharePoint.simulator(config)`
Create a SharePoint simulator for previewing.

## Tool API (for AI Agents)

```javascript
const tools = sdk.tools({ name: 'MySolution' });

// Create components
tools.createWebPart({
  name: 'EmployeeDirectory',
  framework: 'react',
  displayName: 'Employee Directory'
});

// Add permissions
tools.addGraphPermission(ir, 'User.Read.All');

// Add lists
tools.createList({ title: 'Employees', template: 100 });

// Build and validate
const result = await tools.compile(ir);
const validation = await tools.validate(ir);
```

## Architecture

```
AI / Application
      │
  JavaScript
  React
  Vue
  Lovable
      │
      ▼
  CODBSharePoint SDK
      │
  Analyze
  Validate
  Compile
  Bundle
  Package
      │
      ▼
  .sppkg + deployment.json
      │
      ▼
  SharePoint
```

## Browser Compatibility

- Chrome 90+
- Firefox 90+
- Safari 15+
- Edge 90+

## License

MIT - **City of Daytona Beach**

Lead Developer: Daniel Gurczynski
