/// <reference path="./types.ts" />
/// <reference path="./languageLookup.ts" />
/// <reference path="./spreadsheetData.ts" />

/**
 * Validation utilities for user inputs and sheet data
 *
 * This module provides validation functions for:
 * - Language codes and names (supports both English and native names with caching)
 * - Field types and configurations
 * - Category data
 * - Configuration schemas
 *
 * ## Performance Notes
 * - Uses cached lookup from spreadsheetData.ts via getLanguageLookup() for O(1) lookups
 * - Maintains language names cache for error message generation
 * - All caches are cleared via clearLanguagesCache() in spreadsheetData.ts
 */

/**
 * Valid CoMapeo field types
 */
const VALID_FIELD_TYPES = [
  "text",
  "number",
  "selectOne",
  "selectMultiple",
] as const;
type ValidFieldType = (typeof VALID_FIELD_TYPES)[number];

/**
 * Valid geometry types for CoMapeo presets
 */
const VALID_GEOMETRY_TYPES = ["point", "line", "area"] as const;

/**
 * Calculates Levenshtein distance between two strings
 * Used for fuzzy matching and "Did you mean..." suggestions
 *
 * @param str1 - First string
 * @param str2 - Second string
 * @returns Edit distance (number of single-character edits needed)
 *
 * @example
 * levenshteinDistance("portugese", "portuguese") // => 1
 * levenshteinDistance("spanish", "español") // => 5
 */
function levenshteinDistance(str1: string, str2: string): number {
  const len1 = str1.length;
  const len2 = str2.length;

  // Create a 2D array for dynamic programming
  const matrix: number[][] = [];

  // Initialize first row and column
  for (let i = 0; i <= len1; i++) {
    matrix[i] = [i];
  }
  for (let j = 0; j <= len2; j++) {
    matrix[0][j] = j;
  }

  // Fill in the rest of the matrix
  for (let i = 1; i <= len1; i++) {
    for (let j = 1; j <= len2; j++) {
      const cost = str1[i - 1] === str2[j - 1] ? 0 : 1;
      matrix[i][j] = Math.min(
        matrix[i - 1][j] + 1, // deletion
        matrix[i][j - 1] + 1, // insertion
        matrix[i - 1][j - 1] + cost, // substitution
      );
    }
  }

  return matrix[len1][len2];
}

/**
 * Finds the closest matching strings using Levenshtein distance
 *
 * @param target - The string to find matches for
 * @param candidates - Array of candidate strings to match against
 * @param maxResults - Maximum number of results to return (default: 3 for UI readability)
 * @param maxDistance - Maximum edit distance to consider (default: 3 for reasonable typos)
 * @returns Array of closest matches sorted by distance
 *
 * @example
 * findClosestMatches("portugese", ["Portuguese", "Spanish", "French"])
 * // => ["Portuguese"]
 *
 * @example
 * // With proportional distance for longer strings
 * findClosestMatches("malayalaam", languages, 3, 4)
 * // => ["Malayalam"] (10 chars, allows 4 edits = 40% tolerance)
 */
function findClosestMatches(
  target: string,
  candidates: string[],
  maxResults: number = 3, // Limit suggestions for UI readability
  maxDistance: number = 3, // Allow reasonable typos (1-3 character errors)
): string[] {
  const targetLower = target.toLowerCase();

  // Calculate distances and filter by maxDistance
  const distances = candidates
    .map((candidate) => ({
      value: candidate,
      distance: levenshteinDistance(targetLower, candidate.toLowerCase()),
    }))
    .filter((item) => item.distance <= maxDistance)
    .sort((a, b) => a.distance - b.distance);

  // Return top matches
  return distances.slice(0, maxResults).map((item) => item.value);
}

/**
 * Validation result interface
 */
interface SheetValidationResult {
  valid: boolean;
  error?: string;
  warnings?: string[];
}

/**
 * Validates a language code against supported languages
 *
 * @param languageCode - The language code to validate (e.g., "en", "es")
 * @param supportedLanguages - Map of supported language codes to names
 * @returns Validation result
 *
 * @example
 * validateLanguageCode("en", getAllLanguages()) // { valid: true }
 * validateLanguageCode("xyz", getAllLanguages()) // { valid: false, error: "..." }
 */
function validateLanguageCode(
  languageCode: string,
  supportedLanguages: LanguageMap,
): SheetValidationResult {
  if (!languageCode || languageCode.trim() === "") {
    return {
      valid: false,
      error: "Language code cannot be empty",
    };
  }

  const normalizedCode = languageCode.trim().toLowerCase();

  if (!supportedLanguages[normalizedCode]) {
    return {
      valid: false,
      error: `Unsupported language code: "${languageCode}". Please use a valid ISO 639-1 language code.`,
    };
  }

  return { valid: true };
}

/**
 * Cache for all language names (English + native) to avoid expensive rebuilding
 * Populated lazily and cleared when language data is refreshed
 */
let _allLanguageNamesCache: string[] | null = null;

/**
 * Gets all language names (English + native) with caching
 * Used for fuzzy matching suggestions in error messages
 *
 * @returns Array of all language names (both English and native)
 */
function getAllLanguageNamesCached(): string[] {
  if (!_allLanguageNamesCache) {
    const lookup = getLanguageLookup();
    const allCodes = lookup.getAllCodes();
    const allNames: string[] = [];

    for (const langCode of allCodes) {
      const names = lookup.getNamesByCode(langCode);
      if (names) {
        allNames.push(names.english);
        if (names.english !== names.native) {
          allNames.push(names.native);
        }
      }
    }

    _allLanguageNamesCache = allNames;
  }

  return _allLanguageNamesCache;
}

/**
 * Clears the language names cache
 * Should be called when language data is refreshed
 */
function clearLanguageNamesCache(): void {
  _allLanguageNamesCache = null;
}

/**
 * Validates a language name against supported languages
 *
 * Now supports BOTH English and native language names (e.g., "Portuguese" OR "Português").
 * Uses case-insensitive matching for better user experience.
 * Uses cached lookup for O(1) performance.
 *
 * @param languageName - The language name to validate in any form
 * @returns Validation result with the matched language code if valid
 *
 * @example
 * validateLanguageName("English") // { valid: true, code: "en" }
 * validateLanguageName("Português") // { valid: true, code: "pt" }
 * validateLanguageName("Portuguese") // { valid: true, code: "pt" }
 * validateLanguageName("PORTUGUESE") // { valid: true, code: "pt" }
 */
function validateLanguageName(
  languageName: string,
): SheetValidationResult & { code?: string } {
  if (!languageName || languageName.trim() === "") {
    return {
      valid: false,
      error: "Language name cannot be empty",
    };
  }

  const normalizedName = languageName.trim();

  // Use cached lookup for O(1) performance
  const lookup = getLanguageLookup();

  // Try to find language code using enhanced lookup (supports both English and native names)
  const code = lookup.getCodeByName(normalizedName);

  if (!code) {
    // Build helpful error message with fuzzy matching suggestions
    const allLanguageNames = getAllLanguageNamesCached();

    // Calculate proportional max distance based on input length
    // Allows up to 30% edit distance, minimum 2, maximum 5
    const maxDistance = Math.min(
      5,
      Math.max(2, Math.floor(normalizedName.length * 0.3)),
    );

    // Find close matches using fuzzy matching with proportional distance
    const suggestions = findClosestMatches(
      normalizedName,
      allLanguageNames,
      3,
      maxDistance,
    );

    let errorMessage = `Unsupported language: "${languageName}".`;

    if (suggestions.length > 0) {
      errorMessage += ` Did you mean: ${suggestions.map((s) => `"${s}"`).join(", ")}?`;
    } else {
      // Fallback to examples if no close matches
      const allCodes = lookup.getAllCodes();
      const exampleLanguages = allCodes
        .slice(0, 3)
        .map((c) => {
          const names = lookup.getNamesByCode(c);
          return names ? `"${names.english}"` : `"${c}"`;
        })
        .join(", ");
      errorMessage += ` Examples: ${exampleLanguages}`;
    }

    return {
      valid: false,
      error: errorMessage,
    };
  }

  return {
    valid: true,
    code,
  };
}

/**
 * Validates the primary language value (Metadata!primaryLanguage preferred,
 * Categories!A1 fallback). Supports English and native names.
 */
function validatePrimaryLanguage(primaryLanguage: string): SheetValidationResult {
  const result = validateLanguageName(primaryLanguage);

  if (!result.valid) {
    return {
      valid: false,
      error: `Invalid primary language (Metadata!primaryLanguage or Categories!A1): ${result.error}`,
    };
  }

  return { valid: true };
}

/**
 * Validates a field type string from the spreadsheet
 *
 * @param typeString - The type string from the spreadsheet
 * @returns Validation result
 *
 * @example
 * validateFieldType("Text") // { valid: true }
 * validateFieldType("Invalid") // { valid: false, error: "..." }
 */
function validateFieldType(typeString: string): SheetValidationResult {
  if (!typeString || typeString.trim() === "") {
    return {
      valid: false,
      error: "Field type cannot be empty",
    };
  }

  const firstChar = typeString.charAt(0).toLowerCase();
  const validFirstChars = ["t", "n", "s", "m"]; // text, number, select one, multiple choice

  if (!validFirstChars.includes(firstChar)) {
    return {
      valid: false,
      error: `Invalid field type: "${typeString}". Must start with T(ext), N(umber), S(elect one), or M(ultiple choice)`,
    };
  }

  return { valid: true };
}

/**
 * Checks for duplicate option values after slugification and returns warnings
 *
 * This catches two cases of duplicates:
 * 1. Same label repeated: "Red, Red" → both become value "red"
 * 2. Different labels with same slug: "Test 1, Test  1" → both become "test-1"
 *
 * Note: Returns valid=true with warnings since duplicates are auto-removed during linting/generation
 *
 * @param optionsString - The comma-separated options string
 * @param fieldKey - Optional field key for deterministic fallback values
 * @returns Validation result with warnings about any duplicates found
 */
function validateUniqueOptionValues(
  optionsString: string,
  fieldKey?: string,
): SheetValidationResult {
  const options = optionsString
    .split(",")
    .map((opt) => opt.trim())
    .filter((opt) => opt !== "");

  if (options.length === 0) {
    return { valid: true };
  }

  // Create values using same logic as getFieldOptions/parseOptions
  const values: string[] = [];
  const duplicates: Array<{ label: string; value: string; position: number }> =
    [];

  options.forEach((label, index) => {
    const value = typeof createOptionValue === "function"
      ? createOptionValue(label, fieldKey, index)
      : slugify(label) || `${fieldKey || "option"}-${index + 1}`;
    if (values.includes(value)) {
      duplicates.push({ label, value, position: index + 1 });
    }
    values.push(value);
  });

  if (duplicates.length > 0) {
    const details = duplicates
      .map(
        (d) =>
          `"${d.label}" (position ${d.position}) produces duplicate value "${d.value}"`,
      )
      .join("; ");
    // Return valid with warning - duplicates are auto-removed during linting/generation
    return {
      valid: true,
      warnings: [`Duplicate option values will be auto-removed: ${details}`],
    };
  }

  return { valid: true };
}

/**
 * Validates field options for select fields
 *
 * @param fieldType - The field type
 * @param optionsString - The options string from the spreadsheet
 * @param fieldKey - Optional field key for deterministic fallback values
 * @returns Validation result
 */
function validateFieldOptions(
  fieldType: string,
  optionsString: string,
  fieldKey?: string,
): SheetValidationResult {
  const type = getFieldType(fieldType);

  // Text and number fields don't need options
  if (type === "text" || type === "number") {
    return { valid: true };
  }

  // Select fields require options
  if (!optionsString || optionsString.trim() === "") {
    return {
      valid: false,
      error: `Field type "${fieldType}" requires options, but none provided`,
    };
  }

  const options = optionsString
    .split(",")
    .map((opt) => opt.trim())
    .filter((opt) => opt !== "");

  if (options.length === 0) {
    return {
      valid: false,
      error: `Field type "${fieldType}" requires at least one option`,
    };
  }

  // Check for duplicate option values (after slugification) - returns warnings not errors
  const uniqueResult = validateUniqueOptionValues(optionsString, fieldKey);

  return {
    valid: true,
    warnings: uniqueResult.warnings,
  };
}

/**
 * Validates a complete field definition
 *
 * @param fieldData - Array representing a field row from the Details sheet
 * @param rowIndex - The row index for error reporting (1-based, for user display)
 * @returns Validation result with any warnings
 */
function validateFieldDefinition(
  fieldData: FieldRow,
  rowIndex: number,
): SheetValidationResult {
  const warnings: string[] = [];

  // Validate field name (column A)
  if (!fieldData[0] || String(fieldData[0]).trim() === "") {
    return {
      valid: false,
      error: `Row ${rowIndex}: Field name (column A) cannot be empty`,
    };
  }

  // Validate field type (column C)
  const typeValidation = validateFieldType(fieldData[2]);
  if (!typeValidation.valid) {
    return {
      valid: false,
      error: `Row ${rowIndex}: ${typeValidation.error}`,
    };
  }

  // Validate options for select fields (column D)
  const fieldIndex = Math.max(rowIndex - 2, 0);
  const fieldKey = typeof createFieldTagKey === "function"
    ? createFieldTagKey(String(fieldData[0] || ""), fieldIndex)
    : slugify(String(fieldData[0] || "")) || "field";
  const optionsValidation = validateFieldOptions(fieldData[2], fieldData[3], fieldKey);
  if (!optionsValidation.valid) {
    return {
      valid: false,
      error: `Row ${rowIndex}: ${optionsValidation.error}`,
    };
  }

  // Warning if helper text is empty
  if (!fieldData[1] || String(fieldData[1]).trim() === "") {
    warnings.push(
      `Row ${rowIndex}: Helper text (column B) is empty - consider adding guidance for users`,
    );
  }

  return {
    valid: true,
    warnings: warnings.length > 0 ? warnings : undefined,
  };
}

/**
 * Validates a category name
 *
 * @param categoryName - The category name to validate
 * @param rowIndex - The row index for error reporting
 * @returns Validation result
 */
function validateCategoryName(
  categoryName: string,
  rowIndex: number,
): SheetValidationResult {
  if (!categoryName || String(categoryName).trim() === "") {
    return {
      valid: false,
      error: `Row ${rowIndex}: Category name cannot be empty`,
    };
  }

  return { valid: true };
}

/**
 * Validates a complete category definition
 *
 * @param categoryData - Array representing a category row from the Categories sheet
 * @param rowIndex - The row index for error reporting (1-based, for user display)
 * @returns Validation result with any warnings
 */
function validateCategoryDefinition(
  categoryData: CategoryRow,
  rowIndex: number,
): SheetValidationResult {
  const warnings: string[] = [];

  // Validate category name (column A)
  const nameValidation = validateCategoryName(categoryData[0], rowIndex);
  if (!nameValidation.valid) {
    return nameValidation;
  }

  // Warning if icon is not specified
  if (!categoryData[1] || String(categoryData[1]).trim() === "") {
    warnings.push(
      `Row ${rowIndex}: Icon (column B) is empty - a default icon will be used`,
    );
  }

  // Warning if no fields specified
  if (!categoryData[2] || String(categoryData[2]).trim() === "") {
    warnings.push(
      `Row ${rowIndex}: No fields specified (column C) - category will have no custom fields`,
    );
  }

  return {
    valid: true,
    warnings: warnings.length > 0 ? warnings : undefined,
  };
}

/**
 * Validates the entire sheet data before processing
 *
 * @param data - The SheetData object containing all sheet data
 * @returns Validation result with all errors and warnings
 */
function validateSheetData(data: SheetData): SheetValidationResult {
  const errors: string[] = [];
  const warnings: string[] = [];

  // Validate Categories sheet
  if (!data.Categories || data.Categories.length < 2) {
    errors.push("Categories sheet is empty or has no data rows");
  } else {
    // Skip header row, validate data rows
    data.Categories.slice(1).forEach((category, index: number) => {
      const result = validateCategoryDefinition(
        category as CategoryRow,
        index + 2,
      ); // +2 for header and 0-index
      if (!result.valid && result.error) {
        errors.push(result.error);
      }
      if (result.warnings) {
        warnings.push(...result.warnings);
      }
    });
  }

  // Validate Details sheet
  if (!data.Details || data.Details.length < 2) {
    errors.push("Details sheet is empty or has no data rows");
  } else {
    // Skip header row, validate data rows
    data.Details.slice(1).forEach((field, index: number) => {
      const result = validateFieldDefinition(field as FieldRow, index + 2); // +2 for header and 0-index
      if (!result.valid && result.error) {
        errors.push(result.error);
      }
      if (result.warnings) {
        warnings.push(...result.warnings);
      }
    });
  }

  if (errors.length > 0) {
    return {
      valid: false,
      error: `Validation failed with ${errors.length} error(s):\n${errors.join("\n")}`,
      warnings: warnings.length > 0 ? warnings : undefined,
    };
  }

  return {
    valid: true,
    warnings: warnings.length > 0 ? warnings : undefined,
  };
}

/**
 * Validates configuration data schema
 *
 * @param config - The CoMapeoConfig object to validate
 * @returns Validation result indicating schema compliance
 */
function validateConfigSchema(config: CoMapeoConfig): SheetValidationResult {
  const errors: string[] = [];

  // Validate required top-level properties
  if (!config.metadata) {
    errors.push("Configuration missing 'metadata' property");
  }

  if (!config.fields || !Array.isArray(config.fields)) {
    errors.push("Configuration missing 'fields' array");
  }

  if (!config.presets || !Array.isArray(config.presets)) {
    errors.push("Configuration missing 'presets' array");
  }

  if (!config.icons || typeof config.icons !== "object") {
    errors.push("Configuration missing 'icons' object");
  }

  if (!config.translations || typeof config.translations !== "object") {
    errors.push("Configuration missing 'translations' object");
  }

  // Validate metadata structure
  if (config.metadata) {
    if (!config.metadata.dataset_id) {
      errors.push("Metadata missing 'dataset_id'");
    }
    if (!config.metadata.name) {
      errors.push("Metadata missing 'name'");
    }
    if (!config.metadata.version) {
      errors.push("Metadata missing 'version'");
    }
  }

  if (errors.length > 0) {
    return {
      valid: false,
      error: `Schema validation failed:\n${errors.join("\n")}`,
    };
  }

  return { valid: true };
}
