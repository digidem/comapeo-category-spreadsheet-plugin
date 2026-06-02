/**
 * Metadata application functions for the import category functionality.
 * This file contains functions related to applying metadata to the spreadsheet.
 */

/**
 * Applies metadata to the Metadata sheet.
 * @param sheet - The metadata sheet
 * @param metadata - Metadata object
 * @param configData - Optional full config for primary-language detection
 */
function applyMetadata(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  metadata: any,
  configData?: any,
) {
  // Set headers
  sheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]);
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold');

  // Add metadata rows
  const metadataRows = Object.entries(metadata).map(([key, value]) => [
    key,
    value,
  ]);

  // Ensure primaryLanguage is always present in Metadata so that lint and
  // generate flows don't fall back to Categories!A1 (which is now "Name").
  const hasPrimaryLanguage = metadataRows.some(
    ([key]) => String(key).trim() === 'primaryLanguage',
  );
  if (!hasPrimaryLanguage) {
    const detected = detectPrimaryLanguageFromConfig(configData);
    metadataRows.push(['primaryLanguage', detected]);
  }

  if (metadataRows.length > 0) {
    sheet.getRange(2, 1, metadataRows.length, 2).setValues(metadataRows);
  }
}

/**
 * Detects the primary language from config data, handling both the old import
 * shape (NormalizedConfig with presets/messages) and the newer BuildRequest
 * shape (categories/translations).
 */
function detectPrimaryLanguageFromConfig(configData: any): string {
  if (!configData) return 'English';

  // Try the dedicated detectPrimaryLanguage (expects BuildRequest shape)
  if (typeof detectPrimaryLanguage === 'function') {
    // Normalize: old import uses `presets`/`messages`, detector expects
    // `categories`/`translations`.
    const normalized = {
      ...configData,
      categories: configData.categories || configData.presets || [],
      translations: configData.translations || configData.messages || {},
    };
    try {
      return detectPrimaryLanguage(normalized);
    } catch {
      // Fall through to English default
    }
  }

  return 'English';
}
