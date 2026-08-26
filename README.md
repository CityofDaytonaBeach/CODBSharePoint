# CODBSharePoint

Browser-native SharePoint compiler, validator, and packaging SDK.

Generate, validate, compile, and package SPFx solutions entirely in the browser without Node.js, Yeoman, Gulp, or any local SPFx toolchain.

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

MIT - City of Daytona Beach
