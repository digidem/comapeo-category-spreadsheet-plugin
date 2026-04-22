/// <reference path="./loggingHelpers.ts" />
/// <reference path="./types.ts" />
/// <reference path="./languageLookup.ts" />

/**
 * Cache key for languages data
 */
const LANGUAGES_CACHE_KEY = "all_languages_data";
const LANGUAGES_CACHE_KEY_ENHANCED = "all_languages_data_enhanced";
const LANGUAGES_CACHE_TTL = 21600; // 6 hours (maximum allowed by Google Apps Script CacheService)

/**
 * Module-level cache for language lookup to avoid rebuilding indexes
 * Cleared when clearLanguagesCache() is called
 */
let _cachedLookup: LanguageLookup | null = null;

/**
 * Gets or creates the cached language lookup object
 * This avoids rebuilding Map-based indexes on every call
 *
 * @returns Cached LanguageLookup instance
 */
function getLanguageLookup(): LanguageLookup {
  if (!_cachedLookup) {
    const enhanced = getAllLanguagesEnhanced();
    _cachedLookup = createLanguageLookup(enhanced);
  }
  return _cachedLookup;
}

/**
 * Gets the primary language value from cell A1 of the Categories sheet.
 *
 * Prefers the Metadata sheet key `primaryLanguage` (set during migrations) and
 * falls back to Categories!A1 for legacy sheets or when metadata is missing.
 *
 * @returns Primary language value as entered (display name or locale code).
 */
function getPrimaryLanguageName(): string {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // 1) Prefer Metadata sheet entry if present
  const metadataSheet = spreadsheet.getSheetByName("Metadata");
  if (metadataSheet) {
    const values = metadataSheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]).trim() === "primaryLanguage") {
        const lang = String(values[i][1] || "").trim();
        if (lang) {
          return lang;
        }
      }
    }
  }

  // 2) Fallback to Categories!A1 (legacy behavior)
  const categoriesSheet = spreadsheet.getSheetByName("Categories");
  return categoriesSheet?.getRange("A1").getValue() as string;
}

/**
 * Filters languages based on whether they match the primary language.
 *
 * Now correctly handles both English and native language names by comparing
 * language codes instead of name strings. Uses cached lookup for performance.
 *
 * @param allLanguages - Map of ISO language codes to display names.
 * @param includePrimary - Whether to include the primary language in results.
 * @returns Filtered language map.
 *
 * @example
 * // Cell A1 contains "Portuguese" or "Português"
 * const filtered = filterLanguagesByPrimary(allLanguages, false);
 * // Returns all languages except Portuguese (pt)
 */
function filterLanguagesByPrimary(
  allLanguages: LanguageMap,
  includePrimary: boolean,
): LanguageMap {
  const log = getScopedLogger("SpreadsheetData");
  const primaryLanguageName = getPrimaryLanguageName();
  const resolvedPrimaryLanguage = resolvePrimaryLanguageInput(primaryLanguageName);

  // If primary language is invalid or not found, log warning and return all languages
  if (!resolvedPrimaryLanguage) {
    log.warn(
      `Primary language "${primaryLanguageName}" not recognized in language lookup. ` +
      `Including all languages. Set "primaryLanguage" in Metadata sheet or put a valid language name or locale code in Categories!A1.`
    );
    return allLanguages;
  }

  const primaryCode = resolvedPrimaryLanguage.comparisonCode;
  // comparisonCode may be a full locale (e.g. "pt-br") when the user
  // entered a locale tag.  The allLanguages map mixes base codes ("pt")
  // and locale codes ("zh-CN").
  const primaryBase = primaryCode.split("-")[0].toLowerCase();
  const primaryHasSubtag = primaryCode.includes("-");

  // Filter by comparing language codes:
  // - When primary is a base code (e.g. "pt"), exclude all variants with
  //   the same base (pt, pt-BR, pt-PT all excluded).
  // - When primary is a locale code (e.g. "zh-cn"), only exclude exact
  //   matches and base-code-only entries, NOT sibling locale variants
  //   (so zh-CN excludes "zh" but keeps "zh-TW").
  return Object.entries(allLanguages)
    .filter(([code, _]) => {
      const codeLower = code.toLowerCase();
      const langBase = codeLower.split("-")[0];
      const langHasSubtag = codeLower.includes("-");
      const matchesBase = langBase === primaryBase;

      if (includePrimary) {
        if (!primaryHasSubtag) {
          // Primary is base code — include all with same base
          return matchesBase;
        }
        // Primary is locale — include exact match or base-code-only
        return codeLower === primaryCode.toLowerCase() || (matchesBase && !langHasSubtag);
      }
      if (!primaryHasSubtag) {
        // Primary is base code — exclude all with same base
        return !matchesBase;
      }
      // Primary is locale — exclude exact match and base-code-only,
      // keep sibling locale variants
      if (!matchesBase) return true;
      if (!langHasSubtag) return false; // base code excluded
      return codeLower !== primaryCode.toLowerCase(); // keep siblings
    })
    .reduce(
      (acc, [code, name]) => {
        acc[code as LanguageCode] = name;
        return acc;
      },
      {} as LanguageMap,
    );
}

/**
 * Fetches the enhanced language map with both English and native names.
 * Prefers cached data and falls back to remote fetch or local fallback.
 *
 * @returns Enhanced language map with dual-name support
 */
function getAllLanguagesEnhanced(): LanguageMapEnhanced {
  const log = getScopedLogger("SpreadsheetData");
  // Try to get from cache first
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get(LANGUAGES_CACHE_KEY_ENHANCED);

  if (cachedData) {
    log.debug("Using cached enhanced languages data");
    try {
      return JSON.parse(cachedData);
    } catch (parseError) {
      log.warn("Failed to parse cached enhanced languages, fetching fresh data");
      // Continue to fetch fresh data
    }
  }

  log.info("Fetching enhanced languages from remote source");

  try {
    const languagesUrl = "https://raw.githubusercontent.com/digidem/comapeo-mobile/refs/heads/develop/src/frontend/languages.json";
    const response = UrlFetchApp.fetch(languagesUrl);
    const languagesData = JSON.parse(response.getContentText());

    // Convert to enhanced format: {code: {englishName, nativeName}}
    const allLanguages: LanguageMapEnhanced = {};
    for (const [code, lang] of Object.entries(languagesData)) {
      const langData = lang as { englishName?: string; nativeName?: string };
      const englishName = langData.englishName || code; // Fallback to code if missing
      const nativeName = langData.nativeName || englishName; // Fallback to English if native missing

      allLanguages[code] = {
        englishName,
        nativeName,
      };
    }

    // Cache the result
    try {
      cache.put(LANGUAGES_CACHE_KEY_ENHANCED, JSON.stringify(allLanguages), LANGUAGES_CACHE_TTL);
      log.info("Enhanced languages data cached for 6 hours");
    } catch (cacheError) {
      log.warn("Failed to cache enhanced languages data", cacheError);
      // Continue even if caching fails
    }

    // Clear the module-level lookup cache to force rebuild with fresh data
    _cachedLookup = null;
    log.debug("Cleared lookup cache after fetching fresh enhanced language data");

    return allLanguages;
  } catch (error) {
    log.warn("Failed to fetch languages from remote source, using enhanced fallback", error);
    // Fallback to enhanced language data (defined in data/languagesFallback.ts)
    return LANGUAGES_FALLBACK_ENHANCED;
  }
}

/**
 * Fetches the language map, preferring cached data and falling back to
 * remote fetch or local fallback data when necessary.
 *
 * @deprecated Use getAllLanguagesEnhanced() for new code to support dual-name recognition
 * @returns Legacy language map with English names only
 *
 * ## Migration Guide
 *
 * This function is maintained for backward compatibility but should be replaced
 * with `getAllLanguagesEnhanced()` in new code to support dual-name recognition.
 *
 * ### Before (Legacy API):
 * ```typescript
 * const languages = getAllLanguages();
 * // Returns: { "pt": "Portuguese", "es": "Spanish" }
 * const name = languages["pt"]; // "Portuguese"
 * ```
 *
 * ### After (Enhanced API):
 * ```typescript
 * const languages = getAllLanguagesEnhanced();
 * // Returns: { "pt": { englishName: "Portuguese", nativeName: "Português" } }
 * const english = languages["pt"].englishName; // "Portuguese"
 * const native = languages["pt"].nativeName;   // "Português"
 * ```
 *
 * ### Migration Strategy:
 * 1. Identify all `getAllLanguages()` calls in your codebase
 * 2. Update to `getAllLanguagesEnhanced()` and access `.englishName` or `.nativeName`
 * 3. Update related code to support both name forms
 * 4. Test thoroughly with both English and native language names
 *
 * ### Timeline:
 * - **Current**: Both APIs supported, legacy API maintained indefinitely
 * - **Future**: No removal planned - backward compatibility is a priority
 * - **Recommendation**: Use enhanced API for new features, migrate when convenient
 */
function getAllLanguages(): LanguageMap {
  // Get enhanced data and convert to legacy format for backward compatibility
  const enhanced = getAllLanguagesEnhanced();
  return toLegacyLanguageMap(enhanced, false); // Use English names
}

/**
 * Gets the display name for a language code
 *
 * @param code - ISO language code (e.g., "pt", "es")
 * @param preferNative - If true, return native name; otherwise return English name
 * @returns Display name in the requested format, or the code itself if not found
 *
 * @example
 * getLanguageDisplayName("pt", false) // => "Portuguese"
 * getLanguageDisplayName("pt", true)  // => "Português"
 * getLanguageDisplayName("es", false) // => "Spanish"
 * getLanguageDisplayName("es", true)  // => "Español"
 */
function getLanguageDisplayName(code: LanguageCode, preferNative = false): string {
  const enhanced = getAllLanguagesEnhanced();
  const languageData = enhanced[code];

  if (!languageData) {
    // Fallback to code if not found
    return code;
  }

  return preferNative ? languageData.nativeName : languageData.englishName;
}

/**
 * Gets both English and native names for a language code
 *
 * @param code - ISO language code (e.g., "pt", "es")
 * @returns Object with both names, or undefined if code not found
 *
 * @example
 * getLanguageNames("pt") // => { english: "Portuguese", native: "Português" }
 * getLanguageNames("es") // => { english: "Spanish", native: "Español" }
 */
function getLanguageNames(code: LanguageCode): { english: string; native: string } | undefined {
  const enhanced = getAllLanguagesEnhanced();
  const languageData = enhanced[code];

  if (!languageData) {
    return undefined;
  }

  return {
    english: languageData.englishName,
    native: languageData.nativeName,
  };
}

/**
 * Returns available languages filtered by whether to include the primary entry.
 *
 * @param includePrimary - When true, returns only the primary language.
 * @returns Map of language codes to display names.
 */
function languages(includePrimary = false): LanguageMap {
  const allLanguages = getAllLanguages();
  return filterLanguagesByPrimary(allLanguages, includePrimary);
}

/**
 * Resolves the primary language code and name configured in the spreadsheet.
 *
 * Supports English names, native names, canonical language codes, and locale
 * tags in Metadata!primaryLanguage or Categories!A1.
 *
 * @returns Primary language code and original user-entered value.
 * @throws {Error} When the configured value is unsupported.
 *
 * @example
 * // Value contains "Portuguese"
 * getPrimaryLanguage() // => { code: "pt", name: "Portuguese" }
 *
 * @example
 * // Value contains "Português"
 * getPrimaryLanguage() // => { code: "pt", name: "Português" }
 *
 * @example
 * // Value contains "pt-BR"
 * getPrimaryLanguage() // => { code: "pt-br", name: "pt-BR" }
 */
function getPrimaryLanguage(): { code: LanguageCode; name: string } {
  const primaryLanguage = getPrimaryLanguageName();
  const resolvedPrimaryLanguage = resolvePrimaryLanguageInput(primaryLanguage);

  if (!resolvedPrimaryLanguage) {
    throw new Error(
      `Invalid primary language: use a recognized language name or locale code (for example "English", ` +
        `"Português", "en", or "pt-BR").`,
    );
  }

  return {
    code: resolvedPrimaryLanguage.code,
    name: primaryLanguage,
  };
}

/**
 * Retrieves the set of languages that can be targeted for translation,
 * combining the canonical list with any custom sheet headers.
 */
function getAvailableTargetLanguages(): LanguageMap {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const primaryLanguage = getPrimaryLanguageName();
  const allLanguages = getAllLanguages();

  // Resolve primary language to a comparison code for consistent matching.
  const resolvedPrimary = resolvePrimaryLanguageInput(primaryLanguage);
  const primaryCode = resolvedPrimary ? resolvedPrimary.comparisonCode : null;

  // Get all languages except the primary one
  const targetLanguages = filterLanguagesByPrimary(allLanguages, false);

  // Add custom languages from translation sheets
  const translationSheets = sheets(true);
  for (const sheetName of translationSheets) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet) {
      const lastColumn = sheet.getLastColumn();

      // Skip empty sheets (no columns)
      if (lastColumn === 0) {
        continue;
      }

      const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
      // Look for custom language headers (format: "Language Name - ISO")
      headers.forEach((header, index) => {
        if (index > 2 && header && typeof header === "string" && header.includes(" - ")) {
          const [name, iso] = header.split(" - ");
          if (!name || !iso) return;
          // Resolve the header ISO through the same resolver used for the
          // primary language.  comparisonCode now preserves full locale
          // codes (e.g. "pt-br") so distinct variants are not collapsed.
          const resolvedHeader = resolvePrimaryLanguageInput(iso.trim());
          const headerCode = resolvedHeader
            ? resolvedHeader.comparisonCode.toLowerCase()
            : iso.trim().toLowerCase();
          // Compare using base-code matching: if either side is a base
          // code (no subtag), match against the other's base part so that
          // "Portuguese" (pt) excludes "Portuguese - pt-BR" (pt-br).
          // When both have subtags, compare full codes so pt-BR ≠ pt-PT.
          const primaryBase = primaryCode?.split("-")[0].toLowerCase();
          const headerBase = headerCode.split("-")[0];
          const primaryHasSubtag = primaryCode
            ? primaryCode.includes("-")
            : false;
          const headerHasSubtag = headerCode.includes("-");
          let isPrimary = false;
          if (primaryCode) {
            if (!primaryHasSubtag || !headerHasSubtag) {
              // At least one is a base code — compare base parts
              isPrimary = headerBase === primaryBase;
            } else {
              // Both have subtags — compare full codes
              isPrimary = headerCode === primaryCode.toLowerCase();
            }
          }
          if (!isPrimary) {
            targetLanguages[iso.toLowerCase()] = name;
          }
        }
      });
    }
  }

  return targetLanguages;
}

/**
 * Lists supported language names for validation in the Categories sheet.
 *
 * Returns BOTH English and native language names, plus canonical language codes,
 * to support dual-name recognition and locale-code parity.
 *
 * @returns Array of all supported primary-language tokens
 *
 * @example
 * getSupportedLanguagesForA1Cell()
 * // => ["English", "en", "Spanish", "Español", "es", "Portuguese", "Português", "pt", ...]
 */
function getSupportedLanguagesForA1Cell(): string[] {
  const enhanced = getAllLanguagesEnhanced();
  const allNames = new Set<string>();

  // Collect English names, native names, and canonical codes
  for (const [code, data] of Object.entries(enhanced)) {
    allNames.add(data.englishName);
    if (data.englishName !== data.nativeName) {
      allNames.add(data.nativeName);
    }
    allNames.add(code);
  }

  return Array.from(allNames).sort();
}

/**
 * Checks whether a language name is valid for the A1 configuration cell.
 *
 * Supports English names, native names, and canonical language codes with
 * case-insensitive matching. Uses cached lookup for performance.
 *
 * @param languageName - Language token in English, native, or canonical code form
 * @returns True if the token is recognized
 *
 * @example
 * isValidLanguageForA1Cell("Portuguese") // => true
 * isValidLanguageForA1Cell("Português")  // => true
 * isValidLanguageForA1Cell("pt")         // => true
 * isValidLanguageForA1Cell("PORTUGUESE") // => true
 * isValidLanguageForA1Cell("Invalid")    // => false
 */
function isValidLanguageForA1Cell(languageName: string): boolean {
  return resolvePrimaryLanguageInput(languageName) !== null;
}

/**
 * Returns the ordered list of spreadsheet sheet names used by the exporter.
 *
 * @param translationsOnly - When true, include only translation sheets.
 */
function sheets(translationsOnly = false): string[] {
  const translationSheets = [
    "Category Translations",
    "Detail Label Translations",
    "Detail Helper Text Translations",
    "Detail Option Translations",
  ];

  if (translationsOnly) {
    return translationSheets;
  }

  return [...translationSheets, "Categories", "Details", "Icons"];
}
/**
 * Reads spreadsheet data for all relevant sheets into a structured object.
 *
 * @returns SheetData containing sheet names and their value matrices.
 */
function getSpreadsheetData(): SheetData {
  const sheetNames = sheets();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const data: Record<string, unknown> = {
    documentName: spreadsheet.getName(),
  };

  for (const sheetName of sheetNames) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet) {
      const values = sheet.getDataRange().getValues(); // Get all data in the sheet
      data[sheetName] = values;
    }
  }

  return data as SheetData;
}

/**
 * Clears ALL languages cache (both legacy and enhanced) and the lookup cache.
 * Also clears dependent caches in validation.ts.
 * Useful for debugging or forcing fresh data.
 */
function clearLanguagesCache(): void {
  const cache = CacheService.getScriptCache();
  cache.remove(LANGUAGES_CACHE_KEY);
  cache.remove(LANGUAGES_CACHE_KEY_ENHANCED);
  _cachedLookup = null; // Clear module-level lookup cache

  // Clear validation.ts language names cache if function exists
  if (typeof clearLanguageNamesCache !== "undefined") {
    clearLanguageNamesCache();
  }

  const log = getScopedLogger("SpreadsheetData");
  log.info("Languages cache cleared (legacy, enhanced, lookup index, and validation cache)");
}
