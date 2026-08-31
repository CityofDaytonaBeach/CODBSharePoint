// ============================================================================
// Sass compiler abstraction
// Uses Dart Sass compileString when available. This keeps SCSS compilation real
// and lets browser hosts provide/cache the compiler module explicitly.
// ============================================================================

export interface SassCompileResult {
  success: boolean;
  css?: string;
  error?: string;
}

type SassModule = typeof import('sass');

let cachedSass: SassModule | undefined;
let availabilityError: string | undefined;

async function resolveSass(): Promise<SassModule | undefined> {
  if (cachedSass) return cachedSass;
  if (availabilityError !== undefined) return undefined;

  try {
    cachedSass = (await import(/* webpackIgnore: true */ 'sass')) as SassModule;
    return cachedSass;
  } catch (error) {
    availabilityError = error instanceof Error ? error.message : String(error);
    return undefined;
  }
}

export async function compileSassString(path: string, content: string): Promise<SassCompileResult> {
  const sass = await resolveSass();
  if (!sass) {
    return {
      success: false,
      error: availabilityError ? `Sass compiler is not available: ${availabilityError}` : 'Sass compiler is not available'
    };
  }

  try {
    const result = sass.compileString(content, {
      url: new URL(`file:///${path.replace(/\\/g, '/')}`),
      style: 'expanded',
      silenceDeprecations: ['import'],
      importers: [{
        canonicalize(url: string) {
          if (url === '~@microsoft/sp-tslint-theme/vars.scss' || url === '@microsoft/sp-tslint-theme/vars.scss') {
            return new URL('codb-spfx-theme:vars.scss');
          }
          return null;
        },
        load(canonicalUrl: URL) {
          if (canonicalUrl.protocol === 'codb-spfx-theme:') {
            return {
              contents: '',
              syntax: 'scss'
            };
          }
          return null;
        }
      }]
    });

    return {
      success: true,
      css: result.css
    };
  } catch (error) {
    return {
      success: false,
      error: error instanceof Error ? error.message : String(error)
    };
  }
}
