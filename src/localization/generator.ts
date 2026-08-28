// ============================================================================
// Localization Generator
// Produces .resx resources and TypeScript/JS localized string modules from the
// IR's LocalizationConfig. Used to round out localization support that was
// previously only analyzed/recommended.
// ============================================================================

import type { CODBIR, LocalizationConfig, VFSFile } from '../types/index.js';

export function resolveStrings(
  config: LocalizationConfig,
  language: string
): Record<string, string> {
  return config.strings?.[language] || config.strings?.[config.defaultLanguage] || {};
}

// ---------------------------------------------------------------------------
// .resx generation
// ---------------------------------------------------------------------------

export function generateResx(language: string, strings: Record<string, string>): string {
  let xml = `<?xml version="1.0" encoding="utf-8"?>
<root>
  <data name="LocaleName" xml:space="preserve">
    <value>${language}</value>
  </data>
`;
  for (const [key, value] of Object.entries(strings)) {
    xml += `  <data name="${xmlEscape(key)}" xml:space="preserve">
    <value>${xmlEscape(value)}</value>
  </data>
`;
  }
  xml += `</root>`;
  return xml;
}

// ---------------------------------------------------------------------------
// Localized strings module (loc/*.js)
// ---------------------------------------------------------------------------

export function generateStringsModule(language: string, strings: Record<string, string>): string {
  const json = JSON.stringify({ language, strings }, null, 2);
  return `/* eslint-disable */
// LOCALIZED STRINGS  (${language})
export const strings = ${json};
export default strings;
`;
}

// ---------------------------------------------------------------------------
// Top-level generator
// ---------------------------------------------------------------------------

export function generateLocalizationFiles(ir: CODBIR): VFSFile[] {
  const config = ir.localization || { defaultLanguage: 'en-us', languages: [] };
  const files: VFSFile[] = [];
  const languages = new Set<string>([config.defaultLanguage, ...(config.languages || [])]);

  for (const lang of languages) {
    const strings = resolveStrings(config, lang);

    files.push({
      path: `sharepoint/localization/${lang}.resx`,
      content: generateResx(lang, strings),
      encoding: 'utf-8'
    });

    files.push({
      path: `lib/providers/loc/${lang}.js`,
      content: generateStringsModule(lang, strings),
      encoding: 'utf-8'
    });
  }

  return files;
}

function xmlEscape(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}
