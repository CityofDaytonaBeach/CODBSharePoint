import {
  CODBSharePoint,
  MemoryStorage,
  TemplateRegistry,
  createStorage,
  generateStaticPublish,
  createIR,
  addWebPart,
  type StorageAdapter
} from '../src/index.js';

const sdk = new CODBSharePoint();

describe('CODBSharePoint - Serverless authoring (storage, templates, designer, publish)', () => {
  describe('storage', () => {
    it('MemoryStorage round-trips, removes and clears', async () => {
      const s = new MemoryStorage();
      expect((await s.keys()).length).toBe(0);
      await s.setItem('k1', { hello: 'world' });
      expect(await s.getItem('k1')).toEqual({ hello: 'world' });
      expect(await s.keys()).toEqual(['k1']);
      await s.removeItem('k1');
      expect(await s.getItem('k1')).toBeUndefined();
      await s.setItem('a', { x: 1 });
      await s.setItem('b', { y: 2 });
      await s.clear();
      expect(await s.keys()).toEqual([]);
    });

    it('createStorage factory accepts a kind and instance', () => {
      expect(createStorage('memory')).toBeDefined();
      const mem = new MemoryStorage();
      expect(createStorage(mem)).toBe(mem);
    });
  });

  describe('templates', () => {
    it('registers, resolves and lists templates', () => {
      const reg = new TemplateRegistry();
      const list = reg.list();
      expect(list.length).toBeGreaterThan(0);
      const types = list.map(t => t.kind);
      expect(types).toContain('webpart');
      expect(types).toContain('ace');

      const wpTemplate = reg.get('react-webpart')!;
      const files = wpTemplate.render({
        component: { name: 'Hello', framework: 'react' } as never,
        namespace: 'acme'
      });
      expect(files.size).toBeGreaterThan(0);
      const paths = Array.from(files.keys());
      expect(paths.some(p => p.endsWith('.tsx'))).toBe(true);
    });
  });

  describe('designer end-to-end', () => {
    it('create -> author -> save -> load -> build -> publish', async () => {
      const memory = new MemoryStorage();
      const designer = sdk.designer({ storage: memory });

      await designer.create('DesignerDemo', { description: 'demo', spfxVersion: '1.18' });
      const wp = await designer.addWebPart({ name: 'Hello', framework: 'react' });
      expect(wp.name).toBe('Hello');
      const ace = await designer.addACE({ name: 'QuickLinks' });
      expect(ace.name).toBe('QuickLinks');
      await designer.addList({ name: 'Docs' });
      await designer.addGraphPermission('Sites.Read.All');

      await designer.save();

      const reloaded = new CODBSharePoint().designer({ storage: memory });
      const loaded = await reloaded.load('DesignerDemo');
      expect(loaded).toBe(true);
      expect(reloaded.getIR().components.length).toBeGreaterThanOrEqual(2);
      expect(reloaded.getAssets()).toBeDefined();

      const buildResult = await reloaded.build();
      expect(buildResult).toBeDefined();

      const publishResult = await reloaded.publish({ title: 'Demo' });
      const paths = publishResult.files.map(f => f.path);
      expect(paths).toContain('index.html');
      expect(paths).toContain('publish.json');
      const manifest = JSON.parse(String(publishResult.files.find(f => f.path === 'publish.json')!.content));
      expect(manifest.app).toBe('DesignerDemo');
    });

    it('serialize produces a JSON manifest', async () => {
      const designer = sdk.designer({ storage: 'memory' });
      await designer.create('SerDemo');
      const json = designer.serialize();
      const parsed = JSON.parse(json) as { settings: { name: string } };
      expect(parsed.settings.name).toBe('SerDemo');
    });
  });

  describe('static publish', () => {
    it('generateStaticPublish emits a host, bundle container and manifest', () => {
      const ir = createIR({ name: 'PublishDemo' });
      addWebPart(ir, { name: 'Card', framework: 'react' });
      const buildResult = {
        success: true,
        files: [{ path: 'lib/Card.js', content: 'console.log(1);', encoding: 'utf-8' as const }]
      };
      const result = generateStaticPublish(ir, buildResult as never, { title: 'X' });
      const paths = result.files.map(f => f.path);
      expect(paths).toContain('index.html');
      expect(paths).toContain('bundle/bundle.js');
      expect(paths).toContain('publish.json');
      const html = String(result.files.find(f => f.path === 'index.html')!.content);
      expect(html).toContain('<!DOCTYPE html>');
    });
  });
});
