// ============================================================================
// SPPKG Deployment Helper
// Uploads .sppkg packages to a SharePoint tenant app catalog using the
// SharePoint REST API. Requires an authenticated SPHttpClient context.
// ============================================================================

export type AppCatalogType = 'tenant' | 'siteCollection';

export interface UploadSPPKGOptions {
  /** SharePoint web absolute URL of the app catalog site (or site for site-collection). */
  siteUrl: string;
  /** Path to the Apps library (defaults to 'sites/apps' or 'Apps'). */
  libraryPath?: string;
  /** Overwrite existing package + deployment files with the same name. */
  overwrite?: boolean;
  /** Skip feature deployment (deploy only to app catalog, require per-site approval). */
  skipFeatureDeployment?: boolean;
  /** Pass an SPHttpClient-compatible object for the HTTP request. */
  httpClient?: SPHttpClientLike;
}

export interface SPHttpClientLike {
  post(url: string, options: { body: unknown; headers?: Record<string, string> }): Promise<{ ok: boolean; status: number; text(): Promise<string>; json(): Promise<unknown> }>;
  validateDigest?(url: string): Promise<string>;
}

export interface UploadSPPKGResult {
  success: boolean;
  fileName: string;
  libraryUrl: string;
  message: string;
  deployed?: boolean;
  errors?: string[];
}

// Fallback browser fetch-based client if no SPHttpClient passed.
export interface FetchClientOptions {
  headers: Record<string, string>;
}

const DEFAULT_LIBRARY = {
  tenant: 'sites/apps',
  siteCollection: 'Apps'
};

/**
 * Upload an SPPKG package to the tenant or site-collection app catalog.
 *
 * IMPORTANT: This performs real network calls to SharePoint. It requires an
 * authenticated context (either SPHttpClient from SPFx, or fetch with a valid
 * request digest header). It is NOT safe to call from an unauthenticated page.
 */
export async function uploadSPPKG(
  sppkg: Uint8Array,
  packageName: string,
  options: UploadSPPKGOptions
): Promise<UploadSPPKGResult> {
  const errors: string[] = [];
  const libraryPath = options.libraryPath || DEFAULT_LIBRARY[options.skipFeatureDeployment ? 'tenant' : 'tenant'];
  if (!options.siteUrl) errors.push('Missing siteUrl: the app catalog site URL is required.');

  if (sppkg.length === 0) errors.push('SPPKG is empty.');
  if (!packageName) errors.push('Missing packageName.');
  if (!packageName.endsWith('.sppkg')) packageName = `${packageName}.sppkg`;

  if (errors.length > 0) {
    return { success: false, fileName: packageName, libraryUrl: '', message: 'Upload aborted.', errors };
  }

  const base = options.siteUrl.replace(/\/+$/, '');
  const client = options.httpClient;

  try {
    // Build upload URL to the Apps library
    const uploadUrl = `${base}/${libraryPath}/Add(name='${packageName}',overwrite=${!!options.overwrite})`;

    // Read SPPKG as base64 string for the REST body
    let binary = '';
    for (let i = 0; i < sppkg.length; i++) binary += String.fromCharCode(sppkg[i]);
    const base64 = btoa(binary);

    // Upload the file bytes
    if (client) {
      const uploadResp = await client.post(uploadUrl, {
        body: { File: { __metadata: { type: 'SP.File' }, Name: packageName } },
        headers: { 'X-HTTP-Method': 'POST' }
      });

      if (!uploadResp.ok) {
        errors.push(`Upload failed (${uploadResp.status}): ${await safeText(uploadResp)}`);
        return { success: false, fileName: packageName, libraryUrl: uploadUrl, message: 'Upload failed.', errors };
      }
    } else {
      // Fallback: use global fetch with content-type octet-stream + digest
      const digest = await resolveDigest(base, libraryPath);
      const resp = await fetch(uploadUrl, {
        method: 'POST',
        headers: {
          'X-RequestDigest': digest,
          'Content-Type': 'application/octet-stream'
        },
        body: sppkg as BodyInit
      });
      if (!resp.ok) {
        errors.push(`Upload failed (${resp.status}): ${await safeText(resp)}`);
        return { success: false, fileName: packageName, libraryUrl: uploadUrl, message: 'Upload failed.', errors };
      }
    }

    // Deploy the package (enables tenant-wide unless skipFeatureDeployment)
    if (!options.skipFeatureDeployment) {
      const deployUrl = `${base}/${libraryPath}/Add(name='${packageName}',overwrite=true)/Deploy(skipFeatureDeployment=false)`;
      if (client) {
        const deployResp = await client.post(deployUrl, { body: {} });
        if (!deployResp.ok) errors.push(`Deploy failed (${deployResp.status}): ${await safeText(deployResp)}`);
      } else {
        const digest = await resolveDigest(base, libraryPath);
        const deployResp = await fetch(deployUrl, {
          method: 'POST',
          headers: { 'X-RequestDigest': digest, 'Content-Type': 'application/json' },
          body: '{}'
        });
        if (!deployResp.ok) errors.push(`Deploy failed (${deployResp.status}): ${await safeText(deployResp)}`);
      }
    }

    return {
      success: errors.length === 0,
      fileName: packageName,
      libraryUrl: uploadUrl,
      deployed: !!options.skipFeatureDeployment ? undefined : errors.length === 0,
      message: errors.length === 0
        ? `Uploaded to ${libraryPath}/${packageName}${options.skipFeatureDeployment ? ' (deployment pending admin approval)' : ' and deployed.'}`
        : 'Uploaded but deployment had issues.',
      errors: errors.length > 0 ? errors : undefined
    };
  } catch (error) {
    errors.push(`Upload exception: ${error instanceof Error ? error.message : String(error)}`);
    return { success: false, fileName: packageName, libraryUrl: '', message: 'Upload exception.', errors };
  }
}

async function resolveDigest(base: string, libraryPath: string): Promise<string> {
  const digestUrl = `${base}/${libraryPath}/contextinfo`;
  const resp = await fetch(digestUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: '{}'
  });
  if (!resp.ok) throw new Error(`Could not obtain request digest (${resp.status}).`);
  const data: any = await resp.json();
  return data?.d?.GetContextWebInformation?.FormDigestValue || '';
}

async function safeText(resp: { text(): Promise<string> }): Promise<string> {
  try { return await resp.text(); } catch { return ''; }
}

export { DEFAULT_LIBRARY };
