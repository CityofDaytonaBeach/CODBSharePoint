// ============================================================================
// SPFx Manifest Generator
// Generates client-side manifests, package-solution.json, and feature XML
// ============================================================================

import type {
  CODBIR,
  ComponentDefinition,
  ExtensionDefinition,
  SolutionConfig,
  SPFxVersion
} from '../types/index.js';
import { SPFx_COMPATIBILITY } from '../types/index.js';
import { escapeXml } from '../utils/helpers.js';

// ---------------------------------------------------------------------------
// Client-Side Manifests
// ---------------------------------------------------------------------------

export function generateComponentManifest(component: ComponentDefinition, namespace: string): Record<string, unknown> {
  const manifestId = `{${component.id}}`;

  const manifest: Record<string, unknown> = {
    id: manifestId,
    alias: `${namespace}-${component.name}WebPart`,
    componentType: component.type === 'webpart' ? 'WebPart' : component.type,
    version: component.version,
    manifestVersion: 2,
    requiresCustomScript: false,
    hiddenFromToolbox: false,
    preconfiguredEntries: component.preconfiguredEntries.map(entry => ({
      groupId: component.group.id,
      group: { default: component.group.name },
      title: { default: entry.defaultTitle },
      description: { default: entry.description },
      officeFabricIconFontName: entry.officeFabricIconFontName,
      properties: entry.properties
    }))
  };

  if (component.type === 'webpart') {
    manifest.loader = {
      id: 'c10f80bc-9e81-4838-90e2-0e3e164995b4',
      alias: 'WebPartLoader',
      internalModuleNames: ['baseBundle'],
      disablePreload: true
    };

    manifest.webPartData = {
      id: manifestId,
      alias: manifest.alias,
      componentType: 'WebPart',
      manifestVersion: 2,
      version: component.version,
      title: { default: component.displayName },
      description: { default: component.description },
      officeFabricIconFontName: component.officeFabricIconFontName,
      group: { default: component.group.name },
      properties: {},
      supportedHosts: component.supportedHosts || ['SharePointWebPart']
    };
  }

  return manifest;
}

export function generateExtensionManifest(ext: ExtensionDefinition, namespace: string): Record<string, unknown> {
  const manifestId = `{${ext.clientSideComponentId}}`;

  const manifest: Record<string, unknown> = {
    id: manifestId,
    alias: `${namespace}-${ext.name}Extension`,
    componentType: 'Extension',
    extensionType: ext.type,
    manifestVersion: 2,
    requiresCustomScript: false,
    hiddenFromToolbox: false,
    title: { default: ext.displayName },
    description: { default: ext.description }
  };

  if (ext.type === 'ApplicationCustomizer') {
    manifest.loader = {
      id: '4e967090-c0e1-4698-b3bb-2452e4997768',
      alias: 'ApplicationCustomizerLoader',
      internalModuleNames: ['baseBundle'],
      disablePreload: true
    };
  } else if (ext.type === 'FieldCustomizer') {
    manifest.loader = {
      id: '627e3e61-7205-4dd8-9d58-5b355d22e368',
      alias: 'FieldCustomizerLoader',
      internalModuleNames: ['baseBundle'],
      disablePreload: true
    };
  } else if (ext.type === 'ListViewCommandSet') {
    manifest.loader = {
      id: 'c10f80bc-9e81-4838-90e2-0e3e164995b5',
      alias: 'CommandSetLoader',
      internalModuleNames: ['baseBundle'],
      disablePreload: true
    };
  }

  return manifest;
}

// ---------------------------------------------------------------------------
// package-solution.json
// ---------------------------------------------------------------------------

export function generatePackageSolution(ir: CODBIR): Record<string, unknown> {
  const solution = ir.solution;
  const spfxVersion = ir.metadata.spfxVersion || '1.22.0';
  const compatibility = SPFx_COMPATIBILITY[spfxVersion as keyof typeof SPFx_COMPATIBILITY] || SPFx_COMPATIBILITY['1.22.0'];

  const solutionDef: Record<string, unknown> = {
    $schema: 'https://developer.microsoft.com/json-schemas/spfx-build/package-solution.schema.json',
    solution: {
      name: solution.name,
      id: solution.id,
      version: solution.version,
      includeClientSideAssets: solution.includeClientSideAssets,
      skipFeatureDeployment: solution.skipFeatureDeployment,
      isDomainIsolated: solution.isDomainIsolated,
      developer: {
        name: solution.developer.name,
        websiteUrl: solution.developer.websiteUrl,
        privacyUrl: solution.developer.privacyUrl || '',
        termsOfUseUrl: solution.developer.termsOfUseUrl || ''
      },
      metadata: {
        shortDescription: { default: solution.description },
        longDescription: { default: solution.description },
        screenshotPaths: [],
        videoUrl: solution.metadata.videoUrl || '',
        categories: solution.metadata.categories || []
      },
      features: ir.solution.features?.map(f => ({
        title: { default: f.title },
        description: { default: f.description },
        id: f.id,
        version: f.version,
        componentIds: f.components.map(c => c.id)
      })) || []
    },
    paths: {
      zippedPackage: `${solution.name}.sppkg`
    }
  };

  // Add permissions to solution
  if (ir.graph.length > 0) {
    (solutionDef.solution as Record<string, unknown>).webApiPermissionRequests = ir.graph.map(p => ({
      resource: p.resource,
      scope: p.scope
    }));
  }

  return solutionDef;
}

// ---------------------------------------------------------------------------
// Feature XML
// ---------------------------------------------------------------------------

export function generateFeatureXml(ir: CODBIR): string {
  const solution = ir.solution;
  const features = ir.solution.features || [];

  let xml = `<?xml version="1.0" encoding="utf-8"?>
<Feature xmlns="http://schemas.microsoft.com/sharepoint/"
         Id="{${solution.id}}"
         Title="${escapeXml(solution.name)}"
         Description="${escapeXml(solution.description)}"
         Version="${solution.version}"
         Hidden="FALSE"
         Scope="Web">`;

  for (const feature of features) {
    xml += `
  <ElementManifests>
    <ElementManifest Location="${feature.title.replace(/\s/g, '')}\\Elements.xml" />
  </ElementManifests>`;
  }

  xml += `
</Feature>`;

  return xml;
}

// ---------------------------------------------------------------------------
// Elements.xml
// ---------------------------------------------------------------------------

export function generateElementsXml(ir: CODBIR): string {
  const components = ir.components;
  const extensions = ir.extensions;

  let xml = `<?xml version="1.0" encoding="utf-8"?>
<Elements xmlns="http://schemas.microsoft.com/sharepoint/">`;

  for (const component of components) {
    if (component.type === 'webpart') {
      xml += `
  <ClientSideComponent
    Id="{${component.id}}"
    ComponentManifest="${component.name}.manifest.json"
    Properties="" />`;
    }
  }

  for (const ext of extensions) {
    xml += `
  <ClientSideExtension
    RegistrationType="${ext.type === 'ApplicationCustomizer' ? 'List' : ext.type === 'FieldCustomizer' ? 'Field' : 'ListView'}"
    ClientSideComponentId="{${ext.clientSideComponentId}}"
    ClientSideComponentProperties="" />`;
  }

  xml += `
</Elements>`;

  return xml;
}

// ---------------------------------------------------------------------------
// SharePoint project structure files
// ---------------------------------------------------------------------------

export function generateConfigJson(ir: CODBIR): Record<string, unknown> {
  return {
    $schema: 'https://developer.microsoft.com/json-schemas/spfx-build/config.2.0.schema.json',
    version: '2.0',
    manifests: ir.components.map(c => ({
      id: c.id,
      alias: `${ir.solution.namespace}-${c.name}WebPart`,
      componentType: 'WebPart',
      entry: `lib/webparts/${c.name}/${c.name}WebPart.js`,
      version: c.version,
      manifestVersion: 2
    })),
    externals: {},
    localizedResourcePath: 'lib/webparts/${componentName}/loc',
    additionalManifests: []
  };
}

export function generateTsConfig(ir: CODBIR): Record<string, unknown> {
  const spfxVersion = ir.metadata.spfxVersion || '1.22.0';
  return {
    '$schema': 'https://developer.microsoft.com/json-schemas/spfx/tsconfig.schema.json',
    'compilerOptions': {
      'target': 'es5',
      'forceConsistentCasingInFileNames': true,
      'module': 'esnext',
      'moduleResolution': 'node',
      'jsx': 'react',
      'declaration': true,
      'sourceMap': true,
      'experimentalDecorators': true,
      'skipLibCheck': true,
      'outDir': 'lib',
      'inlineSources': false,
      'strictNullChecks': false,
      'noUnusedLocals': false,
      'typeRoots': [
        './node_modules/@types',
        './node_modules/@microsoft'
      ],
      'types': [
        'webpack-env'
      ],
      'lib': [
        'es5',
        'dom',
        'es2015.collection',
        'es2015.promise'
      ]
    },
    'include': [
      'src/**/*.ts',
      'src/**/*.tsx'
    ],
    'exclude': [
      'node_modules',
      'lib'
    ]
  };
}

export function generatePackageJson(ir: CODBIR): Record<string, unknown> {
  const spfxVersion = ir.metadata.spfxVersion || '1.22.0';
  const compatibility = SPFx_COMPATIBILITY[spfxVersion as keyof typeof SPFx_COMPATIBILITY] || SPFx_COMPATIBILITY['1.22.0'];
  const dependencies: Record<string, string> = {
    '@microsoft/sp-core-library': spfxVersion,
    '@microsoft/sp-lodash-subset': spfxVersion,
    '@microsoft/sp-property-pane': spfxVersion,
    '@microsoft/sp-http': spfxVersion
  };

  // Add framework-specific dependencies
  const hasReact = ir.components.some(c => c.framework === 'react');
  if (hasReact) {
    const reactVersion = compatibility?.react === '18' ? '^18.2.0' : '^17.0.1';
    dependencies['react'] = reactVersion;
    dependencies['react-dom'] = reactVersion;
    dependencies['@types/react'] = compatibility?.react === '18' ? '^18.2.0' : '^17.0.45';
    dependencies['@types/react-dom'] = compatibility?.react === '18' ? '^18.2.0' : '^17.0.17';
    dependencies['@microsoft/sp-component-base'] = spfxVersion;
  }

  return {
    name: ir.solution.name.toLowerCase().replace(/\s+/g, '-'),
    version: ir.solution.version,
    private: true,
    main: 'lib/index.js',
    scripts: {
      build: 'gulp build',
      bundle: 'gulp bundle --ship',
      'package-solution': 'gulp package-solution --ship',
      clean: 'gulp clean',
      test: 'gulp test'
    },
    dependencies,
    devDependencies: {
      '@microsoft/sp-build-web': '1.18.0',
      '@microsoft/sp-module-interfaces': '1.18.0',
      '@microsoft/sp-tslint-rules': '1.18.0',
      '@microsoft/sp-webpart-workbench': '1.18.0',
      'gulp': '^4.0.2',
      'typescript': '4.7.4'
    }
  };
}
