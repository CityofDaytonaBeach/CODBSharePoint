// ============================================================================
// Export - Generate deployment artifacts
// ============================================================================

import type {
  CODBIR,
  DeploymentManifest,
  DeploymentPermission,
  DeploymentProvisioning,
  DeploymentInstruction,
  SPFxVersion,
  BuildResult
} from '../types/index.js';
import { GRAPH_PERMISSIONS } from '../types/index.js';
import { createVFS, type VFS } from '../core/vfs.js';

export class SPFxExporter {

  generateDeploymentManifest(
    ir: CODBIR,
    buildResult: Partial<BuildResult>
  ): DeploymentManifest {
    const permissions: DeploymentPermission[] = ir.graph.map(p => ({
      resource: p.resource,
      permission: p.scope,
      type: p.type,
      requiresAdminApproval: p.requiresAdminApproval,
      description: p.description || ''
    }));

    const provisioning: DeploymentProvisioning[] = [
      ...ir.lists.map(l => ({
        type: 'list',
        name: l.title,
        description: l.description
      })),
      ...ir.libraries.map(l => ({
        type: 'library',
        name: l.title,
        description: l.description
      })),
      ...ir.themes.map(t => ({
        type: 'theme',
        name: t.name
      })),
      ...ir.formatting.map(f => ({
        type: f.type,
        name: f.name,
        description: f.description
      })),
      ...ir.provisioning.map(p => ({
        type: p.type,
        name: p.name,
        description: p.description
      })),
      ...ir.pages.map(p => ({
        type: 'page',
        name: p.name
      }))
    ];

    const requiresAdmin = permissions.some(p => p.requiresAdminApproval);

    const instructions: DeploymentInstruction[] = [
      {
        step: 1,
        action: 'Upload to Tenant App Catalog',
        target: `https://tenant.sharepoint.com/sites/appcatalog/Shared%20Documents/Forms/AllItems.aspx`,
        notes: requiresAdmin ? 'Admin approval will be required for Graph permissions' : undefined
      },
      {
        step: 2,
        action: 'Deploy solution',
        target: 'App Catalog > Manage apps > Deploy',
        notes: 'Check "Make this solution available to all sites in the organization" if skipFeatureDeployment is true'
      }
    ];

    if (requiresAdmin) {
      instructions.push({
        step: 3,
        action: 'Approve permissions',
        target: 'SharePoint Admin Center > API Access',
        notes: `Approve the following Graph permissions: ${permissions.filter(p => p.requiresAdminApproval).map(p => p.permission).join(', ')}`
      });
    }

    if (provisioning.length > 0) {
      instructions.push({
        step: instructions.length + 1,
        action: 'Provision resources',
        target: 'Site Settings > Site Scripts',
        notes: `Create lists/libraries: ${provisioning.map(p => p.name).join(', ')}`
      });
    }

    return {
      status: buildResult.success ? 'ready' : 'errors',
      artifact: `${ir.solution.name}.sppkg`,
      artifactSize: buildResult.sppkg?.length,
      destination: 'Tenant App Catalog',
      requiresAdmin,
      permissions,
      provisioning,
      warnings: (buildResult.warnings || []).map(w => typeof w === 'string' ? w : w.message),
      instructions,
      metadata: {
        generator: 'codbsharepoint',
        version: __VERSION__ || '1.0.0',
        buildTime: new Date().toISOString(),
        spfxVersion: ir.metadata.spfxVersion
      }
    };
  }

  generateDeploymentGuide(ir: CODBIR, deployment: DeploymentManifest): string {
    let guide = `# Deployment Guide - ${ir.solution.name}\n\n`;
    guide += `**Version:** ${ir.solution.version}\n`;
    guide += `**Generated:** ${new Date().toLocaleDateString()}\n\n`;

    guide += `## Status\n\n`;
    guide += `- Status: ${deployment.status === 'ready' ? 'Ready for deployment' : 'Has issues'}\n`;
    guide += `- Artifact: ${deployment.artifact}\n`;
    guide += `- Destination: ${deployment.destination}\n\n`;

    if (deployment.permissions.length > 0) {
      guide += `## Required Permissions\n\n`;
      guide += `| Resource | Permission | Type | Admin Approval |\n`;
      guide += `|----------|------------|------|----------------|\n`;
      for (const perm of deployment.permissions) {
        guide += `| ${perm.resource} | ${perm.permission} | ${perm.type} | ${perm.requiresAdminApproval ? 'Yes' : 'No'} |\n`;
      }
      guide += `\n`;
    }

    if (deployment.provisioning.length > 0) {
      guide += `## Provisioning Required\n\n`;
      for (const prov of deployment.provisioning) {
        guide += `- **${prov.type.charAt(0).toUpperCase() + prov.type.slice(1)}:** ${prov.name}`;
        if (prov.description) guide += ` - ${prov.description}`;
        guide += `\n`;
      }
      guide += `\n`;
    }

    guide += `## Deployment Steps\n\n`;
    for (const instruction of deployment.instructions) {
      guide += `### Step ${instruction.step}: ${instruction.action}\n\n`;
      guide += `**Target:** ${instruction.target}\n\n`;
      if (instruction.notes) {
        guide += `**Notes:** ${instruction.notes}\n\n`;
      }
    }

    if (deployment.warnings.length > 0) {
      guide += `## Warnings\n\n`;
      for (const warning of deployment.warnings) {
        guide += `- ${warning}\n`;
      }
    }

    return guide;
  }

  generatePermissionsJson(ir: CODBIR): Record<string, unknown> {
    return {
      graph: ir.graph.map(p => ({
        resource: p.resource,
        scope: p.scope,
        type: p.type,
        requiresAdminApproval: p.requiresAdminApproval,
        description: p.description
      })),
      sharepoint: ir.permissions.map(p => ({
        resource: p.resource,
        scope: p.scope,
        level: p.level,
        description: p.description
      }))
    };
  }

  async generateDeploymentZip(ir: CODBIR, buildResult: BuildResult): Promise<Uint8Array> {
    const vfs = createVFS();

    // Add SPPKG
    if (buildResult.sppkg) {
      vfs.addFile(`${ir.solution.name}.sppkg`, buildResult.sppkg, 'binary');
    }

    // Add deployment manifest
    const deployment = this.generateDeploymentManifest(ir, buildResult);
    vfs.addFile('deployment.json', JSON.stringify(deployment, null, 2));

    // Add deployment guide
    const guide = this.generateDeploymentGuide(ir, deployment);
    vfs.addFile('README.md', guide);

    // Add permissions
    const permissions = this.generatePermissionsJson(ir);
    vfs.addFile('permissions.json', JSON.stringify(permissions, null, 2));

    // Add source code if available
    if (buildResult.files.length > 0) {
      for (const file of buildResult.files) {
        vfs.addFile(`source/${file.path}`, file.content);
      }
    }

    return vfs.toZip();
  }
}
