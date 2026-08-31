import { CODBSharePoint, findKnowledgeEntries, getKnowledgeCatalog, summarizeKnowledge } from '../src/index.js';

describe('CODBSharePoint - SharePoint knowledge catalog', () => {
  it('exposes broad SharePoint, Graph, schema, dependency, and deployment knowledge', () => {
    const catalog = getKnowledgeCatalog();
    const summary = summarizeKnowledge();

    expect(catalog.version).toBe('codbsharepoint.knowledge/1.0');
    expect(summary.sharePointRest).toBeGreaterThanOrEqual(7);
    expect(summary.graph).toBeGreaterThanOrEqual(8);
    expect(summary.schema).toBeGreaterThanOrEqual(10);
    expect(summary.sourcePatterns).toBeGreaterThanOrEqual(8);
    expect(catalog.sharePointRest.map(entry => entry.id)).toContain('sp.listItems');
    expect(catalog.graph.map(entry => entry.id)).toContain('graph.groupMembers');
    expect(catalog.schema.map(entry => entry.id)).toContain('field.taxonomy');
    expect(catalog.pnpDependencies.find(entry => entry.packageName === '@pnp/sp')?.handling).toBe('browser-prebundle-required');
    expect(catalog.fluentUI.find(entry => entry.packageName === '@fluentui/react')?.handling).toBe('browser-prebundle-required');
    expect(catalog.deployment.map(entry => entry.id)).toContain('deploy.apiApproval');
  });

  it('searches knowledge entries for prompt planning', () => {
    const graphResults = findKnowledgeEntries('group members');
    expect(graphResults.some(entry => entry.id === 'graph.groupMembers')).toBe(true);

    const schemaResults = findKnowledgeEntries('taxonomy');
    expect(schemaResults.some(entry => entry.id === 'field.taxonomy')).toBe(true);
  });

  it('feeds knowledge into SDK and AI prompt context', () => {
    const sdk = new CODBSharePoint();

    expect(sdk.knowledgeAPI.summary().graph).toBeGreaterThanOrEqual(8);
    expect(sdk.knowledgeAPI.find('quick links').some(entry => entry.id === 'pattern.quickLinks')).toBe(true);

    const context = JSON.parse(sdk.ai().exportPromptContext());
    expect(context.capabilities.sharePointKnowledge.graph).toBeGreaterThanOrEqual(8);
    expect(context.knowledge.sharePointRest.some((entry: { id: string }) => entry.id === 'sp.listItems')).toBe(true);
    expect(context.knowledge.deployment.some((entry: { id: string }) => entry.id === 'deploy.tenantAppCatalog')).toBe(true);
  });
});
