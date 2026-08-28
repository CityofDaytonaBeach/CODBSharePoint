// ============================================================================
// Static Publish (Serverless Host)
// Produces static artifacts (HTML host + consolidated bundle + publish.json)
// that can be deployed to any static host (GitHub Pages, R2, S3, Cloudflare)
// WITHOUT a SharePoint backend. The host renders a standalone preview of each
// component and packages the compiled bundle for reuse.
// ============================================================================

import type { CODBIR, VFSFile, BuildResult } from '../types/index.js';

export interface StaticPublishOptions {
  title?: string;
  appName?: string;
  includeSPPKG?: boolean;
  minify?: boolean;
}

export interface StaticPublishResult {
  success: boolean;
  files: VFSFile[];
  errors: string[];
}

export function generateStaticPublish(
  ir: CODBIR,
  buildResult: BuildResult,
  options: StaticPublishOptions = {}
): StaticPublishResult {
  const errors: string[] = [];
  if (!buildResult.success) {
    errors.push('Cannot publish: build was not successful');
  }

  const appName = options.appName || ir.solution.name;
  const title = options.title || `${ir.solution.name} — Preview`;

  const jsFiles = buildResult.files.filter(f => f.path.endsWith('.js') && !f.path.endsWith('.map'));
  const cssFiles = buildResult.files.filter(f => f.path.endsWith('.css'));

  const css = cssFiles
    .map(f => (typeof f.content === 'string' ? f.content : new TextDecoder().decode(f.content)))
    .join('\n');

  const components = ir.components
    .filter(c => c.type === 'webpart' || c.type === 'ace')
    .map(c => ({
      name: c.name,
      displayName: c.displayName || c.name,
      description: c.description || '',
      type: c.type
    }));

  const jsInline = jsFiles
    .map(f => (typeof f.content === 'string' ? f.content : new TextDecoder().decode(f.content)))
    .join('\n');

  const html = generateHostHtml({ title, appName, css, components, jsInline });

  const files: VFSFile[] = [
    { path: 'index.html', content: html, encoding: 'utf-8' }
  ];

  // Consolidated bundle for reuse on the host
  if (jsInline.trim().length > 0) {
    files.push({ path: `bundle/bundle.js`, content: jsInline, encoding: 'utf-8' });
  }

  // Publish manifest
  files.push({
    path: 'publish.json',
    content: JSON.stringify(
      {
        schema: 'codbsharepoint/publish/1.0',
        app: appName,
        description: ir.solution.description,
        version: ir.solution.version,
        spfxVersion: ir.metadata.spfxVersion,
        artifacts: files.map(f => ({ path: f.path, size: String(f.content).length })),
        components,
        permissions: ir.graph.map(g => ({ scope: g.scope, requiresAdminApproval: g.requiresAdminApproval })),
        targets: [
          { provider: 'github-pages', note: 'Drop these files into a Pages branch/folder' },
          { provider: 's3', note: 'Upload to a static bucket' },
          { provider: 'cloudflare-r2', note: 'Upload to an R2 bucket fronted by a worker' }
        ]
      },
      null,
      2
    ),
    encoding: 'utf-8'
  });

  if (options.includeSPPKG && buildResult.sppkg) {
    files.push({ path: `${ir.solution.name}.sppkg`, content: buildResult.sppkg, encoding: 'binary' });
  }

  return { success: errors.length === 0, files, errors };
}

interface HostOptions {
  title: string;
  appName: string;
  css: string;
  components: { name: string; displayName: string; description: string; type: string }[];
  jsInline: string;
}

function generateHostHtml(o: HostOptions): string {
  const cards = o.components
    .map(
      c => `<div class="preview-card" data-component="${c.name}">
  <div class="preview-head">
    <span class="badge">${c.type}</span>
    <h3>${c.displayName}</h3>
  </div>
  <p class="preview-desc">${c.description || 'No description provided.'}</p>
  <div class="preview-slot" id="slot-${c.name}"></div>
</div>`
    )
    .join('\n');

  const componentJson = JSON.stringify(o.components);

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>${o.title}</title>
  <style>
    :root { --primary: #0078d4; --lighter: #deecf9; --text: #323130; }
    * { box-sizing: border-box; }
    body { font-family: 'Segoe UI', -apple-system, sans-serif; margin: 0; background: #faf9f8; color: var(--text); }
    header { background: var(--primary); color: #fff; padding: 16px 24px; }
    header h1 { margin: 0; font-size: 20px; }
    main { max-width: 1000px; margin: 24px auto; padding: 0 16px; display: grid; grid-template-columns: repeat(auto-fill, minmax(300px, 1fr)); gap: 16px; }
    .preview-card { background: #fff; border: 1px solid #edebe9; border-radius: 6px; padding: 16px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }
    .preview-head { display: flex; align-items: center; gap: 8px; }
    .badge { background: var(--lighter); color: var(--primary); border-radius: 10px; padding: 2px 10px; font-size: 11px; text-transform: uppercase; }
    .preview-desc { color: #605e5c; font-size: 13px; }
    .preview-slot { margin-top: 12px; border: 1px dashed #c8c6c4; border-radius: 4px; padding: 12px; min-height: 60px; }
    footer { text-align: center; color: #8a8886; font-size: 12px; padding: 24px; }
    ${o.css}
  </style>
</head>
<body>
  <header><h1>${o.appName}</h1></header>
  <main>
${cards}
  </main>
  <footer>Generated by CODBSharePoint • static host preview</footer>
  <script>
    // Standalone host runtime: renders a lightweight preview shell for each component.
    // The compiled component bundle (bundle.js) is loaded for embed/reuse.
    window.CODB_PREVIEW = ${componentJson};
    function mountPreview() {
      (window.CODB_PREVIEW || []).forEach(function (c) {
        var slot = document.getElementById('slot-' + c.name);
        if (slot) slot.textContent = c.displayName + ' (authoring preview)';
      });
    }
    if (document.readyState !== 'loading') mountPreview(); else document.addEventListener('DOMContentLoaded', mountPreview);
  </script>
</body>
</html>`;
}
