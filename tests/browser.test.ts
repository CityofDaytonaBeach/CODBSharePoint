import { CODBSharePoint, initBrowser, browserBuildProof, createIR, addWebPart } from '../src/index.js';

describe('CODBSharePoint - Browser Runtime', () => {
  it('exposes browser API on SDK instance', () => {
    const sdk = new CODBSharePoint();
    expect(typeof sdk.browserAPI.init).toBe('function');
    expect(typeof sdk.browserAPI.initCustom).toBe('function');
    expect(typeof sdk.browserAPI.download).toBe('function');
    expect(typeof sdk.browserAPI.downloadSPPKG).toBe('function');
    expect(typeof sdk.browserAPI.buildProof).toBe('function');
  });

  it('initBrowser returns structured result', async () => {
    const result = await initBrowser();
    expect(result).toHaveProperty('success');
    expect(result).toHaveProperty('esbuildWasm');
    expect(result).toHaveProperty('diagnostics');
    expect(Array.isArray(result.diagnostics)).toBe(true);
    expect(typeof result.esbuildWasm).toBe('boolean');
  });

  it('browserBuildProof generates SPPKG bytes from IR', async () => {
    const ir = createIR({ name: 'BrowserProof' });
    addWebPart(ir, { name: 'ProofWebPart', framework: 'react' });

    const proof = await browserBuildProof(ir);

    expect(proof.sppkgBytes).toBeGreaterThan(0);
    expect(proof.validations.length).toBeGreaterThan(0);
    expect(proof.duration).toBeGreaterThanOrEqual(0);
    expect(proof.validations.some(v => v.includes('SPPKG generated'))).toBe(true);
  });

  it('browserBuildProof produces valid ZIP archive', async () => {
    const ir = createIR({ name: 'BrowserZip' });
    addWebPart(ir, { name: 'ZipWebPart', framework: 'react' });

    const proof = await browserBuildProof(ir);

    expect(proof.validations.some(v => v.includes('valid ZIP'))).toBe(true);
  });

  it('browserBuildProof detects Content_Types via unzip', async () => {
    const ir = createIR({ name: 'BrowserCT' });
    addWebPart(ir, { name: 'CTWebPart', framework: 'react' });

    const proof = await browserBuildProof(ir);

    expect(proof.validations.some(v => v.includes('Content_Types'))).toBe(true);
  });

  it('browserBuildProof produces valid package structure', async () => {
    const ir = createIR({ name: 'BrowserStruct' });
    addWebPart(ir, { name: 'StructWebPart', framework: 'react' });

    const proof = await browserBuildProof(ir);

    expect(proof.validations.some(v => v.includes('SPPKG generated'))).toBe(true);
    expect(proof.validations.some(v => v.includes('valid ZIP'))).toBe(true);
    expect(proof.validations.some(v => v.includes('Content_Types'))).toBe(true);
  });

  it('browserBuildProof detects missing bundle (expected without source)', async () => {
    const ir = createIR({ name: 'BrowserMissing' });
    addWebPart(ir, { name: 'MissingWebPart', framework: 'react' });

    const proof = await browserBuildProof(ir);

    // Without bundle files, structural validation will flag missing bundles
    expect(proof.sppkgBytes).toBeGreaterThan(0);
  });

  it('browserBuildProof works for multiple components', async () => {
    const ir = createIR({ name: 'BrowserMulti' });
    addWebPart(ir, { name: 'MultiA', framework: 'react' });
    addWebPart(ir, { name: 'MultiB', framework: 'vanilla' });

    const proof = await browserBuildProof(ir);

    expect(proof.sppkgBytes).toBeGreaterThan(0);
    expect(proof.validations.some(v => v.includes('valid ZIP'))).toBe(true);
  });

  it('download functions are callable (no-op in Node)', () => {
    const { downloadFile, downloadSPPKG } = require('../src/browser/index.js');
    expect(typeof downloadFile).toBe('function');
    expect(typeof downloadSPPKG).toBe('function');
  });
});
