import { CODBSharePoint, uploadSPPKG } from '../src/index.js';

describe('CODBSharePoint - SPPKG Deployment', () => {
  it('validates required options before network calls', async () => {
    const result = await uploadSPPKG(
      new Uint8Array([0x50, 0x4b, 0x03, 0x04]),
      'test.sppkg',
      { siteUrl: '', httpClient: undefined as any }
    );

    expect(result.success).toBe(false);
    expect(result.errors).toContain('Missing siteUrl: the app catalog site URL is required.');
    expect(result.message).toContain('aborted');
  });

  it('rejects empty packages', async () => {
    const result = await uploadSPPKG(
      new Uint8Array(0),
      'test.sppkg',
      { siteUrl: 'https://example.sharepoint.com/sites/apps', httpClient: undefined as any }
    );

    expect(result.success).toBe(false);
    expect(result.errors).toContain('SPPKG is empty.');
  });

  it('adds .sppkg extension to filename if missing', async () => {
    const mockClient: any = {
      post: jest.fn(async () => ({ ok: true, status: 200, text: async () => '', json: async () => ({}) }))
    };

    const result = await uploadSPPKG(
      new Uint8Array([0x50, 0x4b, 0x03, 0x04]),
      'MyPackage',
      { siteUrl: 'https://example.sharepoint.com', httpClient: mockClient }
    );

    expect(result.fileName).toBe('MyPackage.sppkg');
    expect(mockClient.post).toHaveBeenCalled();
  });

  it('reports upload failure with status code', async () => {
    const mockClient: any = {
      post: jest.fn(async () => ({ ok: false, status: 403, text: async () => 'Forbidden', json: async () => ({}) }))
    };

    const result = await uploadSPPKG(
      new Uint8Array([0x50, 0x4b, 0x03, 0x04]),
      'test.sppkg',
      { siteUrl: 'https://example.sharepoint.com', httpClient: mockClient }
    );

    expect(result.success).toBe(false);
    expect(result.errors?.some(e => e.includes('403'))).toBe(true);
  });

  it('sdk exposes deployAPI', () => {
    const sdk = new CODBSharePoint();
    expect(typeof sdk.deployAPI.upload).toBe('function');
    expect(sdk.deployAPI.DEFAULT_LIBRARY.tenant).toBe('sites/apps');
  });

  it('executes upload and deploy flow with successful client', async () => {
    const mockClient: any = {
      post: jest.fn(async () => ({ ok: true, status: 200, text: async () => '', json: async () => ({}) }))
    };

    const result = await uploadSPPKG(
      new Uint8Array([0x50, 0x4b, 0x03, 0x04]),
      'test.sppkg',
      { siteUrl: 'https://example.sharepoint.com/sites/contoso', httpClient: mockClient }
    );

    expect(result.success).toBe(true);
    expect(result.deployed).toBe(true);
    // Verify both upload and deploy POST calls happened
    expect(mockClient.post).toHaveBeenCalledTimes(2);
  });
});
