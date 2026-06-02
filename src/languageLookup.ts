/// <reference path="./types.ts" />

/**
 * Language Lookup System
 *
 * Provides bidirectional mapping between language codes and names (both English and native).
 * Supports O(1) lookups for both directions with case-insensitive matching.
 *
 * @example
 * const lookup = createLanguageLookup(languageData);
 * lookup.getCodeByName("Portuguese") // => "pt"
 * lookup.getCodeByName("Português")  // => "pt"
 * lookup.getCodeByName("PORTUGUESE") // => "pt"
 * lookup.getNamesByCode("pt")        // => { english: "Portuguese", native: "Português" }
 */

/**
 * Language lookup result type
 */
interface LanguageLookup {
  /**
   * Get language code from any name form (English or native, case-insensitive)
   * Null-safe: returns undefined for null/undefined inputs
   * @param name - Language name in any form (null-safe)
   * @returns Language code or undefined if not found or name is null/undefined
   */
  getCodeByName(name: string | null | undefined): LanguageCode | undefined;

  /**
   * Get the canonical language code for a code input, case-insensitively.
   * Null-safe: returns undefined for null/undefined inputs.
   * @param code - Language code or locale token (null-safe)
   * @returns Canonical language code if known, otherwise undefined
   */
  getCanonicalCode(code: LanguageCode | null | undefined): LanguageCode | undefined;

  /**
   * Get both English and native names for a language code
   * Null-safe: returns undefined for null/undefined inputs
   * @param code - ISO language code (null-safe)
   * @returns Object with english and native names, or undefined if not found or code is null/undefined
   */
  getNamesByCode(code: LanguageCode | null | undefined): { english: string; native: string } | undefined;

  /**
   * Get all name aliases for a language code
   * Null-safe: returns empty array for null/undefined inputs
   * @param code - ISO language code (null-safe)
   * @returns Array of all recognized names (English and native), or empty array if not found or code is null/undefined
   */
  getAllAliases(code: LanguageCode | null | undefined): string[];

  /**
   * Check if a language code exists
   * Null-safe: returns false for null/undefined inputs
   * @param code - ISO language code (null-safe)
   * @returns True if code exists in the lookup, false if not found or code is null/undefined
   */
  hasCode(code: LanguageCode | null | undefined): boolean;

  /**
   * Check if a name (any form) is recognized
   * Null-safe: returns false for null/undefined inputs
   * @param name - Language name in any form (null-safe)
   * @returns True if name is recognized, false if not found or name is null/undefined
   */
  hasName(name: string | null | undefined): boolean;

  /**
   * Get all supported language codes
   * @returns Array of all language codes
   */
  getAllCodes(): LanguageCode[];

  /**
   * Normalize a language name for comparison
   * Null-safe: returns empty string for null/undefined inputs
   * Useful for debugging and testing
   * @param name - Language name to normalize (null-safe)
   * @returns Normalized name (lowercase, trimmed, locale-aware), or empty string if name is null/undefined
   */
  normalize(name: string | null | undefined): string;
}

/**
 * Normalizes a language name for comparison
 * - Trims whitespace
 * - Converts to lowercase using English locale (avoids Turkish 'I' issue)
 * - Handles common variations
 *
 * @param name - Language name to normalize
 * @returns Normalized name for comparison
 *
 * @example
 * normalizeLanguageName("Portuguese") // => "portuguese"
 * normalizeLanguageName("PORTUGUÊS")  // => "português"
 * normalizeLanguageName("  Spanish  ") // => "spanish"
 */
function normalizeLanguageName(name: string): string {
  if (!name) {
    return "";
  }
  // Use 'en-US' locale to avoid Turkish 'I' issue where 'I' becomes 'ı' instead of 'i'
  return name.trim().toLocaleLowerCase("en-US");
}

/**
 * Creates a bidirectional language lookup system from enhanced language data
 *
 * Builds two efficient Map-based indexes:
 * 1. Code → Names (both English and native)
 * 2. Name (normalized) → Code (supports both English and native)
 *
 * @param languageData - Enhanced language map with both English and native names
 * @returns Language lookup interface with O(1) operations
 * @note Logs warnings (not errors) if language name collisions detected; uses first match
 *
 * @example
 * const data: LanguageMapEnhanced = {
 *   "pt": { englishName: "Portuguese", nativeName: "Português" },
 *   "es": { englishName: "Spanish", nativeName: "Español" }
 * };
 * const lookup = createLanguageLookup(data);
 * lookup.getCodeByName("português") // => "pt" (case-insensitive)
 */
function createLanguageLookup(languageData: LanguageMapEnhanced): LanguageLookup {
  // Index 1: Code → Names (both forms)
  const codeToNames = new Map<LanguageCode, { english: string; native: string }>();

  // Index 2: Normalized Name → Code (bidirectional)
  const nameToCode = new Map<string, LanguageCode>();

  // Index 3: Normalized Code → Canonical Code
  const normalizedCodeToCode = new Map<string, LanguageCode>();

  // Build indexes from language data with collision detection
  for (const [code, data] of Object.entries(languageData)) {
    const english = data.englishName;
    const native = data.nativeName;

    // Store names for this code
    codeToNames.set(code, { english, native });
    normalizedCodeToCode.set(normalizeLanguageName(code), code);

    // Create bidirectional mappings (normalized for case-insensitive matching)
    const normalizedEnglish = normalizeLanguageName(english);
    const normalizedNative = normalizeLanguageName(native);

    // Collision detection for English name - use first match and warn
    const existingCodeForEnglish = nameToCode.get(normalizedEnglish);
    if (existingCodeForEnglish && existingCodeForEnglish !== code) {
      const existingNames = codeToNames.get(existingCodeForEnglish);
      // Log warning but use first match (don't throw)
      if (typeof Logger !== "undefined") {
        Logger.log(
          `Warning: Language name collision detected: "${english}" (${code}) conflicts with "${existingNames?.english}" (${existingCodeForEnglish}). Both normalize to "${normalizedEnglish}". Using first match (${existingCodeForEnglish}).`,
        );
      }
      // Skip adding this duplicate mapping - first match wins
    } else {
      nameToCode.set(normalizedEnglish, code);
    }

    // Only add native if it's different from English (avoid duplicate keys)
    if (normalizedEnglish !== normalizedNative) {
      // Collision detection for native name - use first match and warn
      const existingCodeForNative = nameToCode.get(normalizedNative);
      if (existingCodeForNative && existingCodeForNative !== code) {
        const existingNames = codeToNames.get(existingCodeForNative);
        // Log warning but use first match (don't throw)
        if (typeof Logger !== "undefined") {
          Logger.log(
            `Warning: Language name collision detected: "${native}" (${code}) conflicts with "${existingNames?.native || existingNames?.english}" (${existingCodeForNative}). Both normalize to "${normalizedNative}". Using first match (${existingCodeForNative}).`,
          );
        }
        // Skip adding this duplicate mapping - first match wins
      } else {
        nameToCode.set(normalizedNative, code);
      }
    }
  }

  // Return lookup interface with null-safe operations
  return {
    /**
     * Get language code from any name form (English or native, case-insensitive)
     * Handles null/undefined gracefully by returning undefined
     *
     * @param name - Language name in any form (null-safe)
     * @returns Language code or undefined if not found or name is null/undefined
     */
    getCodeByName(name: string | null | undefined): LanguageCode | undefined {
      if (!name) {
        return undefined;
      }
      const normalized = normalizeLanguageName(name);
      return nameToCode.get(normalized);
    },

    /**
     * Get canonical code from any code casing
     * Handles null/undefined gracefully by returning undefined
     *
     * @param code - Language code or locale token (null-safe)
     * @returns Canonical code or undefined if not found or code is null/undefined
     */
    getCanonicalCode(code: LanguageCode | null | undefined): LanguageCode | undefined {
      if (!code) {
        return undefined;
      }
      const normalized = normalizeLanguageName(code);
      return normalizedCodeToCode.get(normalized);
    },

    /**
     * Get both English and native names for a language code
     * Handles null/undefined gracefully by returning undefined
     *
     * @param code - ISO language code (null-safe)
     * @returns Object with english and native names, or undefined if not found or code is null/undefined
     */
    getNamesByCode(code: LanguageCode | null | undefined): { english: string; native: string } | undefined {
      if (!code) {
        return undefined;
      }
      return codeToNames.get(code);
    },

    /**
     * Get all name aliases for a language code
     * Handles null/undefined gracefully by returning empty array
     *
     * @param code - ISO language code (null-safe)
     * @returns Array of all recognized names (English and native), or empty array if not found or code is null/undefined
     */
    getAllAliases(code: LanguageCode | null | undefined): string[] {
      if (!code) {
        return [];
      }

      const names = codeToNames.get(code);
      if (!names) {
        return [];
      }

      // Return unique names (English and native, if different)
      const aliases = [names.english];
      if (names.english !== names.native) {
        aliases.push(names.native);
      }
      return aliases;
    },

    /**
     * Check if a language code exists
     * Handles null/undefined gracefully by returning false
     *
     * @param code - ISO language code (null-safe)
     * @returns True if code exists in the lookup, false if not found or code is null/undefined
     */
    hasCode(code: LanguageCode | null | undefined): boolean {
      if (!code) {
        return false;
      }
      return codeToNames.has(code);
    },

    /**
     * Check if a name (any form) is recognized
     * Handles null/undefined gracefully by returning false
     *
     * @param name - Language name in any form (null-safe)
     * @returns True if name is recognized, false if not found or name is null/undefined
     */
    hasName(name: string | null | undefined): boolean {
      if (!name) {
        return false;
      }
      const normalized = normalizeLanguageName(name);
      return nameToCode.has(normalized);
    },

    /**
     * Get all supported language codes
     * @returns Array of all language codes
     */
    getAllCodes(): LanguageCode[] {
      return Array.from(codeToNames.keys());
    },

    /**
     * Normalize a language name for comparison
     * Handles null/undefined gracefully by returning empty string
     * Useful for debugging and testing
     *
     * @param name - Language name to normalize (null-safe)
     * @returns Normalized name (lowercase, trimmed, locale-aware), or empty string if name is null/undefined
     */
    normalize(name: string | null | undefined): string {
      if (!name) {
        return "";
      }
      return normalizeLanguageName(name);
    },
  };
}

/**
 * Creates a legacy-compatible LanguageMap from enhanced language data
 *
 * Extracts English names only for backward compatibility with code
 * expecting the legacy LanguageMap format.
 *
 * @param languageData - Enhanced language map
 * @param preferNative - If true, use native names instead of English (default: false)
 * @returns Legacy language map with single names
 *
 * @example
 * const enhanced: LanguageMapEnhanced = {
 *   "pt": { englishName: "Portuguese", nativeName: "Português" }
 * };
 * toLegacyLanguageMap(enhanced) // => { "pt": "Portuguese" }
 * toLegacyLanguageMap(enhanced, true) // => { "pt": "Português" }
 */
function toLegacyLanguageMap(
  languageData: LanguageMapEnhanced,
  preferNative = false,
): LanguageMap {
  const legacy: LanguageMap = {};

  for (const [code, data] of Object.entries(languageData)) {
    legacy[code] = preferNative ? data.nativeName : data.englishName;
  }

  return legacy;
}

/**
 * Converts legacy LanguageMap to enhanced format
 *
 * Assumes the provided name is the English name and uses it for both
 * English and native (for backward compatibility with existing data).
 *
 * @param legacyMap - Legacy language map with single names
 * @returns Enhanced language map with dual names
 *
 * @example
 * const legacy: LanguageMap = { "pt": "Portuguese" };
 * fromLegacyLanguageMap(legacy) // => { "pt": { englishName: "Portuguese", nativeName: "Portuguese" } }
 */
function fromLegacyLanguageMap(legacyMap: LanguageMap): LanguageMapEnhanced {
  const enhanced: LanguageMapEnhanced = {};

  for (const [code, name] of Object.entries(legacyMap)) {
    enhanced[code] = {
      englishName: name,
      nativeName: name, // Fallback to same name if native not available
    };
  }

  return enhanced;
}
