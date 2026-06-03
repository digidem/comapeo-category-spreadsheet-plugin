/// <reference path="./types.ts" />
/// <reference path="./loggingHelpers.ts" />

function getTranslationServiceLogger() {
  return typeof getScopedLogger === "function"
    ? getScopedLogger("TranslationService")
    : console;
}

const MAIN_LANGUAGE_COLUMN = "A";
const CATEGORIES_SHEET = "Categories";
const DETAILS_SHEET = "Details";
const CATEGORY_TRANSLATIONS_SHEET = "Category Translations";
const DETAIL_LABEL_TRANSLATIONS_SHEET = "Detail Label Translations";
const DETAIL_HELPER_TEXT_TRANSLATIONS_SHEET = "Detail Helper Text Translations";
const DETAIL_OPTION_TRANSLATIONS_SHEET = "Detail Option Translations";
const DETAILS_HELPER_TEXT_COLUMN = "B";
const DETAILS_OPTIONS_COLUMN = "D";

/** Cross-sheet translation cache. Persists across multiple sheet translations in a single run. */
let crossSheetTranslationCache = new Map<string, string>();

/** Clear the cross-sheet translation cache. Call at the start of a full translation run. */
function clearTranslationCache(): void {
  crossSheetTranslationCache.clear();
}

function normalizeLanguageSelection(
  selection: TranslationLanguage[] | LanguageSelectionPayload | null | undefined,
): LanguageSelectionPayload {
  if (Array.isArray(selection)) {
    const autoLanguages = selection
      .map((lang) => (typeof lang === "string" ? lang.trim() : ""))
      .filter((lang): lang is TranslationLanguage => Boolean(lang));
    return { autoTranslateLanguages: autoLanguages, customLanguages: [] };
  }

  if (!selection) {
    return { autoTranslateLanguages: [], customLanguages: [] };
  }

  const autoTranslateLanguages = Array.isArray(selection.autoTranslateLanguages)
    ? selection.autoTranslateLanguages
        .map((lang) => (typeof lang === "string" ? lang.trim() : ""))
        .filter((lang): lang is TranslationLanguage => Boolean(lang))
    : [];

  const customLanguages = Array.isArray(selection.customLanguages)
    ? selection.customLanguages.map((lang) => ({
        name: typeof lang?.name === "string" ? lang.name : "",
        iso: typeof lang?.iso === "string" ? lang.iso : "",
      }))
    : [];

  return { autoTranslateLanguages, customLanguages };
}

function manageLanguagesAndTranslate(
  selection: TranslationLanguage[] | LanguageSelectionPayload | null | undefined,
): void {
  const log = getTranslationServiceLogger();
  const { autoTranslateLanguages, customLanguages } = normalizeLanguageSelection(selection);

  log.info("[LANGUAGE MANAGEMENT] Received selection", {
    autoTranslateLanguages,
    customLanguages,
  });

  if (customLanguages.length > 0) {
    log.info("[LANGUAGE MANAGEMENT] Adding custom language columns", customLanguages);
    addCustomLanguagesToTranslationSheets(customLanguages);
  } else {
    log.info("[LANGUAGE MANAGEMENT] No custom languages requested");
  }

  if (autoTranslateLanguages.length > 0) {
    log.info("[LANGUAGE MANAGEMENT] Running bidirectional translation", autoTranslateLanguages);
    autoTranslateSheetsBidirectional(autoTranslateLanguages);
  } else {
    log.info("[LANGUAGE MANAGEMENT] No auto-translate languages selected; skipping translation");
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (spreadsheet) {
    const appliedUpdates = customLanguages.length > 0 || autoTranslateLanguages.length > 0;
    const toastTitle = "Language Manager";
    const toastMessage = appliedUpdates
      ? "Language updates applied successfully."
      : "No language changes detected.";
    spreadsheet.toast(toastMessage, toastTitle, 5);
  }
}

function handleLanguageSelection(
  selection: TranslationLanguage[] | LanguageSelectionPayload | null | undefined,
): void {
  try {
    manageLanguagesAndTranslate(selection);
  } catch (error) {
    const ui = SpreadsheetApp.getUi();
    const message = error instanceof Error ? error.message : String(error);
    ui.alert(
      translateMenuTexts[locale].error,
      `${translateMenuTexts[locale].errorText}${message}`,
      ui.ButtonSet.OK,
    );
    throw error;
  }
}

/**
 * Translates a sheet from source language to a single target language
 *
 * Reads text from column A (source language) and translates to target language column.
 * Only translates cells that are empty in the target column. Uses Google Translate API.
 *
 * @param sheetName - Name of the sheet to translate
 * @param targetLanguage - ISO language code for target language (e.g., "es", "fr")
 * @param sourceLanguage - Optional source language code (defaults to primary language from cell A1)
 *
 * @example
 * translateSheet("Category Translations", "es");
 * // Translates from primary language to Spanish
 */
function translateSheet(
  sheetName: string,
  targetLanguage: TranslationLanguage,
  sourceLanguage?: TranslationLanguage,
): void {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found`);
  }
  const lastRow = sheet.getLastRow();
  const languagesList: LanguageMap = languages();
  const primaryLanguage = getPrimaryLanguage();
  const mainLanguage = sourceLanguage || primaryLanguage.code;

  getTranslationServiceLogger().info("Source language", mainLanguage);
  getTranslationServiceLogger().info("Target language", targetLanguage);
  getTranslationServiceLogger().info("Primary language from A1", primaryLanguage);

  for (const [code, name] of Object.entries(languagesList)) {
    getTranslationServiceLogger().info(`Language code: ${code}, name: ${name}`);
  }

  const targetColumn = Object.keys(languagesList).indexOf(targetLanguage) + 2;

  for (let i = 2; i <= lastRow; i++) {
    const originalText = sheet.getRange(i, 1).getValue() as string;
    if (originalText) {
      const targetCell = sheet.getRange(i, targetColumn);
      if (!targetCell.getValue()) {
        try {
          const translation = LanguageApp.translate(
            originalText,
            mainLanguage,
            targetLanguage,
          );
          targetCell.setValue(translation);
        } catch (error) {
          getTranslationServiceLogger().warn(`Translation failed for "${originalText}" from ${mainLanguage} to ${targetLanguage}: ${error.message}`);
          // Skip this translation but don't fail the entire process
        }
      }
    }
  }
}

/**
 * Translates a sheet from source language to multiple target languages bidirectionally
 *
 * Supports custom language columns (format: "Language Name - ISO") in addition to
 * standard languages. Validates all target languages before translation. Skips cells
 * that already have translations.
 *
 * @param sheetName - Name of the sheet to translate
 * @param targetLanguages - Array of ISO language codes for target languages
 * @param sourceLanguage - Optional source language code (defaults to primary language)
 * @throws Error if source and target languages are the same, or if target languages are invalid
 *
 * @example
 * translateSheetBidirectional("Category Translations", ["es", "fr", "pt"]);
 * // Translates from primary language to Spanish, French, and Portuguese
 */
function translateSheetBidirectional(
  sheetName: string,
  targetLanguages: TranslationLanguage[],
  sourceLanguage?: TranslationLanguage,
): void {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found`);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    getTranslationServiceLogger().warn(`Sheet "${sheetName}" has no data to translate`);
    return;
  }

  const primaryLanguage = getPrimaryLanguage();
  const actualSourceLanguage = sourceLanguage || primaryLanguage.code;
  const availableTargetLanguages: LanguageMap = getAvailableTargetLanguages();


  // Validate that source language is different from target languages
  const conflictingLanguages = targetLanguages.filter(lang => lang === actualSourceLanguage);
  if (conflictingLanguages.length > 0) {
    throw new Error(`Cannot translate from ${actualSourceLanguage} to itself`);
  }

  // Validate that target languages are available
  const invalidLanguages = targetLanguages.filter(lang => !availableTargetLanguages[lang]);
  if (invalidLanguages.length > 0) {
    throw new Error(`Invalid target languages: ${invalidLanguages.join(', ')}`);
  }

  getTranslationServiceLogger().info("Bidirectional translation:");
  getTranslationServiceLogger().info("Source language:", actualSourceLanguage);
  getTranslationServiceLogger().info("Target languages:", targetLanguages);
  getTranslationServiceLogger().info("Available target languages:", availableTargetLanguages);

  // Get headers to find column positions
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

  let translationErrors: string[] = [];
  let successfulTranslations = 0;

  const translationCache = crossSheetTranslationCache;
  const allLanguages: LanguageMap = getAllLanguages();
  const resolveHeaderCode = createTranslationHeaderResolver(allLanguages);
  const headerCodeToIndex = new Map<string, number>();
  const targetColumnIndexes = new Map<TranslationLanguage, number>();

  headers.forEach((header, index) => {
    const code = resolveHeaderCode(String(header || ""));
    if (!code) return;
    const normalizedCode = code.toLowerCase();
    if (headerCodeToIndex.has(normalizedCode)) {
      getTranslationServiceLogger().warn(
        `Duplicate translation header for ${code} in ${sheetName}; using first match`,
      );
      return;
    }
    headerCodeToIndex.set(normalizedCode, index);
  });

  for (const targetLang of targetLanguages) {
    const normalizedTarget = targetLang.toLowerCase();
    const targetColumn = headerCodeToIndex.get(normalizedTarget);

    if (typeof targetColumn !== "number") {
      translationErrors.push(`Target column not found for language: ${targetLang}`);
      continue;
    }

    targetColumnIndexes.set(targetLang, targetColumn);
  }

  const rowCount = lastRow - 1;
  if (rowCount <= 0 || targetColumnIndexes.size === 0) {
    getTranslationServiceLogger().info(`No translation updates required for ${sheetName}`);
  } else {
    const sheetValues = sheet.getRange(2, 1, rowCount, lastColumn).getValues();
    const updatedColumns = new Map<number, string[]>();
    const updatedColumnIndexes = new Set<number>();

    targetColumnIndexes.forEach((columnIndex) => {
      const columnValues = sheetValues.map((row) => row[columnIndex]);
      updatedColumns.set(columnIndex, columnValues.slice());
    });

    // Phase 1: Collect all unique (sourceText, targetLang) pairs needing translation
    const pendingTranslations = new Map<string, { sourceText: string; targetLang: string }>();

    for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
      const rawValue = sheetValues[rowIndex][0];
      const originalText = typeof rawValue === "string" ? rawValue.trim() : String(rawValue || "").trim();
      if (!originalText) {
        continue;
      }

      targetColumnIndexes.forEach((_columnIndex, targetLang) => {
        const columnValues = updatedColumns.get(_columnIndex);
        if (!columnValues) {
          return;
        }

        const existingValue = columnValues[rowIndex];
        if (existingValue && String(existingValue).trim() !== "") {
          return;
        }

        const cacheKey = `${actualSourceLanguage}::${targetLang}::${originalText}`;
        if (!translationCache.has(cacheKey) && !pendingTranslations.has(cacheKey)) {
          pendingTranslations.set(cacheKey, { sourceText: originalText, targetLang });
        }
      });
    }

    getTranslationServiceLogger().info(
      `[Phase 1] Collected ${pendingTranslations.size} unique translations needed for ${sheetName}`,
    );

    // Phase 2: Translate all unique texts
    let translateIndex = 0;
    for (const [cacheKey, { sourceText, targetLang }] of pendingTranslations) {
      translateIndex++;
      try {
        const translatedText = LanguageApp.translate(
          sourceText,
          actualSourceLanguage,
          targetLang,
        );
        const translation = translatedText ? translatedText.trim() : "";
        if (translation) {
          translationCache.set(cacheKey, translation);
        }
      } catch (error) {
        const errorMessage = `Translation error for ${targetLang} (pair ${translateIndex}/${pendingTranslations.size}, source row context): ${error.message}`;
        getTranslationServiceLogger().error(errorMessage);
        translationErrors.push(errorMessage);
      }
    }

    getTranslationServiceLogger().info(
      `[Phase 2] Translated ${pendingTranslations.size} unique pairs, ${translationErrors.length} errors`,
    );

    // Phase 3: Fill column values from cache
    for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
      const rawValue = sheetValues[rowIndex][0];
      const originalText = typeof rawValue === "string" ? rawValue.trim() : String(rawValue || "").trim();
      if (!originalText) {
        continue;
      }

      targetColumnIndexes.forEach((columnIndex, targetLang) => {
        const columnValues = updatedColumns.get(columnIndex);
        if (!columnValues) {
          return;
        }

        const existingValue = columnValues[rowIndex];
        if (existingValue && String(existingValue).trim() !== "") {
          return;
        }

        const cacheKey = `${actualSourceLanguage}::${targetLang}::${originalText}`;
        const translation = translationCache.get(cacheKey);

        if (translation) {
          columnValues[rowIndex] = translation;
          updatedColumnIndexes.add(columnIndex);
          successfulTranslations++;
        }
      });
    }

    updatedColumnIndexes.forEach((columnIndex) => {
      const columnValues = updatedColumns.get(columnIndex);
      if (!columnValues) {
        return;
      }

      const valuesToWrite = columnValues.map((value) => [value]);
      sheet.getRange(2, columnIndex + 1, rowCount, 1).setValues(valuesToWrite);
    });
  }

  getTranslationServiceLogger().info(`Translation summary for ${sheetName}:`, {
    successfulTranslations,
    errors: translationErrors.length,
  });

  if (translationErrors.length > 0) {
    getTranslationServiceLogger().warn("Translation errors:", translationErrors);
  }
}

/**
 * Creates or retrieves the Category Translations sheet
 *
 * Creates a new sheet with language headers (primary language + target languages).
 * Links column A to Categories sheet with formulas for automatic sync.
 * If target languages not specified, includes all available languages.
 *
 * @param targetLanguages - Optional array of ISO language codes to include as columns
 * @returns The Category Translations sheet
 *
 * @example
 * const sheet = createCategoryTranslationsSheet(["es", "fr"]);
 * // Creates sheet with columns: Primary Language, Spanish, French
 */
function createCategoryTranslationsSheet(targetLanguages?: TranslationLanguage[]): GoogleAppsScript.Spreadsheet.Sheet {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(CATEGORY_TRANSLATIONS_SHEET);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(CATEGORY_TRANSLATIONS_SHEET);
    const primaryLanguage = getPrimaryLanguage();
    const allLanguages = getAllLanguages();

    // Create headers: primary language + selected target languages only
    const headers = [primaryLanguage.name];

    if (targetLanguages && targetLanguages.length > 0) {
      // Add only selected target languages
      for (const langCode of targetLanguages) {
        if (allLanguages[langCode]) {
          headers.push(allLanguages[langCode]);
        }
      }
    } else {
      // Fallback to old behavior if no target languages specified
      const otherLanguages = Object.entries(languages());
      headers.push(...otherLanguages.map(([_, name]) => name));
    }

    sheet
      .getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight("bold");

    const categoriesSheet = spreadsheet.getSheetByName(CATEGORIES_SHEET);
    if (categoriesSheet) {
      const lastRow = categoriesSheet.getLastRow();
      const formula = `=${CATEGORIES_SHEET}!${MAIN_LANGUAGE_COLUMN}2:${MAIN_LANGUAGE_COLUMN}${lastRow}`;
      sheet.getRange(2, 1, lastRow - 1, 1).setFormula(formula);
    }
  }
  return sheet;
}

/**
 * Creates a translation sheet for field labels, helper text, or options
 *
 * Creates sheet with language headers and links column A to source sheet/column.
 * Used for Detail Label Translations, Detail Helper Text Translations, and
 * Detail Option Translations sheets.
 *
 * @param sheetName - Name for the new translation sheet
 * @param targetLanguages - Array of ISO language codes for translation columns
 * @param sourceSheet - Source sheet name (e.g., "Details")
 * @param sourceColumn - Source column letter (e.g., "A", "B", "D")
 * @returns The created translation sheet
 * @throws Error if source sheet not found
 *
 * @example
 * const sheet = createTranslationSheet("Detail Label Translations", ["es", "fr"], "Details", "A");
 * // Creates sheet linked to Details column A with Spanish and French columns
 */
function createTranslationSheet(sheetName: string, targetLanguages: TranslationLanguage[], sourceSheet: string, sourceColumn: string): GoogleAppsScript.Spreadsheet.Sheet {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    const primaryLanguage = getPrimaryLanguage();
    const allLanguages = getAllLanguages();

    // Create headers: primary language + selected target languages only
    const headers = [primaryLanguage.name];

    // Add only selected target languages
    for (const langCode of targetLanguages) {
      if (allLanguages[langCode]) {
        headers.push(allLanguages[langCode]);
      }
    }

    sheet
      .getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight("bold");

    const sourceSheetObj = spreadsheet.getSheetByName(sourceSheet);
    if (!sourceSheetObj) {
      throw new Error(`Source sheet ${sourceSheet} not found`);
    }
    const lastRow = sourceSheetObj.getLastRow();
    const formula = `=${sourceSheet}!${sourceColumn}2:${sourceColumn}${lastRow}`;
    sheet.getRange(2, 1, lastRow - 1, 1).setFormula(formula);
  }

  return sheet;
}

/**
 * Ensures translation sheet has columns for all target languages
 *
 * Checks existing sheet headers and adds missing language columns.
 * Maintains bold formatting for headers. Used when adding new languages
 * to existing translation sheets.
 *
 * @param sheet - Translation sheet to update
 * @param targetLanguages - Array of ISO language codes that should exist
 *
 * @example
 * const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Category Translations");
 * ensureLanguageColumnsExist(sheet, ["es", "fr", "pt"]);
 * // Adds any missing language columns to the sheet
 */
function ensureLanguageColumnsExist(sheet: GoogleAppsScript.Spreadsheet.Sheet, targetLanguages: TranslationLanguage[]): void {
  const allLanguages = getAllLanguages();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const resolveHeaderCode = createTranslationHeaderResolver(allLanguages);
  const existingCodes = new Set<string>();
  const codeToName = new Map<string, string>();

  for (const [code, name] of Object.entries(allLanguages)) {
    codeToName.set(code.toLowerCase(), name);
  }

  headers.forEach((header) => {
    const code = resolveHeaderCode(String(header || ""));
    if (code) {
      existingCodes.add(code.toLowerCase());
    }
  });

  // Check which target languages are missing from the sheet
  const missingLanguages: string[] = [];
  for (const langCode of targetLanguages) {
    const normalizedCode = langCode.toLowerCase();
    if (existingCodes.has(normalizedCode)) {
      continue;
    }
    const languageName = codeToName.get(normalizedCode);
    if (languageName) {
      missingLanguages.push(languageName);
    }
  }

  // Add missing language columns
  if (missingLanguages.length > 0) {
    const startColumn = headers.length + 1;
    const newHeaders = [...headers, ...missingLanguages];

    // Update the header row
    sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);

    // Set header formatting
    sheet.getRange(1, startColumn, 1, missingLanguages.length).setFontWeight("bold");

    getTranslationServiceLogger().info(`Added ${missingLanguages.length} new language columns: ${missingLanguages.join(', ')}`);
  }
}

/**
 * Automatically translates all translation sheets to all available languages
 *
 * Creates translation sheets if they don't exist and translates Categories,
 * Detail Helper Text, and Detail Options to all languages defined in
 * the languages() function. Uses unidirectional translation.
 *
 * @deprecated Use autoTranslateSheetsBidirectional() instead for better language support
 *
 * @example
 * autoTranslateSheets();
 * // Translates all sheets to all available languages
 */
function autoTranslateSheets(): void {
  const allSheets = sheets();
  const translationSheets = sheets(true);
  const categoriesAndDetailsSheets = allSheets.filter(
    (sheet) => !translationSheets.includes(sheet),
  );

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const categoriesSheet = spreadsheet.getSheetByName(
    categoriesAndDetailsSheets[0],
  );
  const detailsSheet = spreadsheet.getSheetByName(
    categoriesAndDetailsSheets[1],
  );

  if (!categoriesSheet || !detailsSheet) {
    throw new Error("Categories or Details sheet not found");
  }

  const primaryLanguage = getPrimaryLanguage();
  let totalErrors = 0;
  let totalSuccessful = 0;

  for (const sheetName of translationSheets) {
    let sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      if (sheetName === CATEGORY_TRANSLATIONS_SHEET) {
        sheet = createCategoryTranslationsSheet();
      } else {
        sheet = spreadsheet.insertSheet(sheetName);
        const mainLanguage = Object.entries(languages(true))[0];
        const otherLanguages = Object.entries(languages());
        const headers = [
          mainLanguage[1],
          ...otherLanguages.map(([_, name]) => name),
        ];
        sheet
          .getRange(1, 1, 1, headers.length)
          .setValues([headers])
          .setFontWeight("bold");

        let sourceSheet: string;
        let sourceColumn: string;
        if (sheetName === DETAIL_HELPER_TEXT_TRANSLATIONS_SHEET) {
          sourceSheet = DETAILS_SHEET;
          sourceColumn = DETAILS_HELPER_TEXT_COLUMN;
        } else if (sheetName === DETAIL_OPTION_TRANSLATIONS_SHEET) {
          sourceSheet = DETAILS_SHEET;
          sourceColumn = DETAILS_OPTIONS_COLUMN;
        } else {
          sourceSheet = DETAILS_SHEET;
          sourceColumn = MAIN_LANGUAGE_COLUMN;
        }

        const sourceSheetObj = spreadsheet.getSheetByName(sourceSheet);
        if (!sourceSheetObj) {
          throw new Error(`Source sheet ${sourceSheet} not found`);
        }
        const lastRow = sourceSheetObj.getLastRow();
        const formula = `=${sourceSheet}!${sourceColumn}2:${sourceColumn}${lastRow}`;
        sheet.getRange(2, 1, lastRow - 1, 1).setFormula(formula);
      }
    }

    for (const lang of Object.keys(languages())) {
      try {
        translateSheet(sheetName, lang as TranslationLanguage);
        totalSuccessful++;
      } catch (error) {
        getTranslationServiceLogger().warn(`Translation failed for ${sheetName} to ${lang}: ${error.message}`);
        totalErrors++;
        // Continue with other languages instead of stopping completely
      }
    }
  }

  getTranslationServiceLogger().info(`Translation completed. Successful: ${totalSuccessful}, Errors: ${totalErrors}`);

  if (totalErrors > 0) {
    getTranslationServiceLogger().warn(`Some translations failed. The CoMapeo config will be generated with available translations.`);
  }
}

/**
 * Automatically translates all translation sheets to selected target languages
 *
 * Creates or updates all translation sheets (Category Translations, Detail Label
 * Translations, Detail Helper Text Translations, Detail Option Translations).
 * Validates target languages and skips translation if no languages specified.
 * Uses bidirectional translation supporting custom language columns.
 *
 * @param targetLanguages - Array of ISO language codes to translate to
 * @throws Error if target languages are invalid or conflict with source language
 *
 * @example
 * autoTranslateSheetsBidirectional(["es", "fr", "pt"]);
 * // Creates/updates all translation sheets with Spanish, French, and Portuguese
 *
 * @example
 * autoTranslateSheetsBidirectional([]);
 * // Skips translation entirely (useful for single-language configs)
 */
function autoTranslateSheetsBidirectional(targetLanguages: TranslationLanguage[]): void {
  // Validation - return early if no languages specified (skip translation)
  if (!targetLanguages || targetLanguages.length === 0) {
    getTranslationServiceLogger().info("[TRANSLATION] ✓ SKIPPING TRANSLATION - No target languages specified");
    getTranslationServiceLogger().info("[TRANSLATION] targetLanguages parameter:", targetLanguages);
    getTranslationServiceLogger().info("[TRANSLATION] Returning early without translation");
    return;
  }

  getTranslationServiceLogger().info("[TRANSLATION] Starting bidirectional translation");
  getTranslationServiceLogger().info("[TRANSLATION] Target languages count:", targetLanguages.length);
  getTranslationServiceLogger().info("[TRANSLATION] Target languages:", targetLanguages);

  clearTranslationCache();

  const allSheets = sheets();
  const translationSheets = sheets(true);
  const categoriesAndDetailsSheets = allSheets.filter(
    (sheet) => !translationSheets.includes(sheet),
  );

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const categoriesSheet = spreadsheet.getSheetByName(
    categoriesAndDetailsSheets[0],
  );
  const detailsSheet = spreadsheet.getSheetByName(
    categoriesAndDetailsSheets[1],
  );

  if (!categoriesSheet || !detailsSheet) {
    throw new Error("Categories or Details sheet not found");
  }

  const primaryLanguage = getPrimaryLanguage();
  const availableTargetLanguages = getAvailableTargetLanguages();
  const allLanguages = getAllLanguages();

  // Validate target languages exist in the comprehensive language list
  const invalidTargetLanguages = targetLanguages.filter(lang => !allLanguages[lang]);
  if (invalidTargetLanguages.length > 0) {
    throw new Error(`Invalid target languages: ${invalidTargetLanguages.join(', ')}. Please check the language codes.`);
  }

  // Validate target languages are available (not the primary language)
  const unavailableTargetLanguages = targetLanguages.filter(lang => !availableTargetLanguages[lang]);
  if (unavailableTargetLanguages.length > 0) {
    const unavailableLanguageNames = unavailableTargetLanguages.map(lang => allLanguages[lang]).join(', ');
    throw new Error(`Target languages not available: ${unavailableLanguageNames}. These languages might be the same as your source language (${primaryLanguage.name}).`);
  }

  // Validate source language is different from target languages
  const conflictingLanguages = targetLanguages.filter(lang => lang === primaryLanguage.code);
  if (conflictingLanguages.length > 0) {
    throw new Error(`Cannot translate from ${primaryLanguage.name} to itself. Please change the language in cell A1 of the Categories sheet or select different target languages.`);
  }

  getTranslationServiceLogger().info("Auto-translating from primary language:", primaryLanguage);
  getTranslationServiceLogger().info("Target languages:", targetLanguages);
  getTranslationServiceLogger().info("Available target languages:", availableTargetLanguages);

  let totalErrors = 0;
  let totalSuccessful = 0;

  for (const sheetName of translationSheets) {
    let sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      if (sheetName === CATEGORY_TRANSLATIONS_SHEET) {
        sheet = createCategoryTranslationsSheet(targetLanguages);
      } else {
        let sourceSheet: string;
        let sourceColumn: string;
        if (sheetName === DETAIL_HELPER_TEXT_TRANSLATIONS_SHEET) {
          sourceSheet = DETAILS_SHEET;
          sourceColumn = DETAILS_HELPER_TEXT_COLUMN;
        } else if (sheetName === DETAIL_OPTION_TRANSLATIONS_SHEET) {
          sourceSheet = DETAILS_SHEET;
          sourceColumn = DETAILS_OPTIONS_COLUMN;
        } else {
          sourceSheet = DETAILS_SHEET;
          sourceColumn = MAIN_LANGUAGE_COLUMN;
        }

        sheet = createTranslationSheet(sheetName, targetLanguages, sourceSheet, sourceColumn);
      }
    } else {
      // Sheet exists, ensure it has columns for the target languages
      ensureLanguageColumnsExist(sheet, targetLanguages);
    }

    try {
      // Use bidirectional translation for selected target languages
      translateSheetBidirectional(sheetName, targetLanguages, primaryLanguage.code as TranslationLanguage);
    } catch (error) {
      getTranslationServiceLogger().error(`Error translating sheet ${sheetName}:`, error);
      totalErrors++;
    }
  }

  getTranslationServiceLogger().info("[TRANSLATION] Translation summary:");
  getTranslationServiceLogger().info("[TRANSLATION] - Total errors:", totalErrors);
  getTranslationServiceLogger().info("[TRANSLATION] - Sheets processed:", translationSheets.length);
  getTranslationServiceLogger().info("[TRANSLATION] ✅ Translation completed");

  if (totalErrors > 0) {
    getTranslationServiceLogger().warn("[TRANSLATION] ⚠️  Some translations failed. Check the logs for details.");
  }
}

function normalizeCustomLanguages(newLanguages: CustomLanguageInput[] | null | undefined): CustomLanguageInput[] {
  if (!newLanguages || newLanguages.length === 0) {
    return [];
  }

  const log = getTranslationServiceLogger();
  const seenIso = new Set<string>();
  const isoPattern = /^[a-z]{2,8}(?:-[a-z0-9]{2,8})?$/;
  const normalized: CustomLanguageInput[] = [];

  for (const lang of newLanguages) {
    if (!lang) {
      continue;
    }

    const name = (lang.name || "").trim();
    const iso = (lang.iso || "").trim().toLowerCase();

    if (!name || !iso) {
      log.warn("Skipping custom language with missing name or ISO", lang);
      continue;
    }

    if (!isoPattern.test(iso)) {
      log.warn(`Skipping custom language with invalid ISO code: ${iso}`);
      continue;
    }

    if (seenIso.has(iso)) {
      log.info(`Skipping duplicate custom language ISO: ${iso}`);
      continue;
    }

    seenIso.add(iso);
    normalized.push({ name, iso });
  }

  return normalized;
}

function addCustomLanguagesToSheet(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  languagesToAdd: CustomLanguageInput[],
): void {
  if (!sheet) {
    return;
  }

  const log = getTranslationServiceLogger();
  const currentLastColumn = sheet.getLastColumn();
  const headerWidth = Math.max(currentLastColumn, 1);
  const headers = sheet.getRange(1, 1, 1, headerWidth).getValues()[0];
  const headerSet = new Set(headers);
  const updatedHeaders = headers.slice();

  for (const lang of languagesToAdd) {
    const headerText = `${lang.name} - ${lang.iso}`;
    if (headerSet.has(headerText)) {
      continue;
    }
    headerSet.add(headerText);
    updatedHeaders.push(headerText);
  }

  if (updatedHeaders.length === headers.length) {
    log.info(`No new custom language columns required for ${sheet.getName()}`);
    return;
  }

  const additionalColumns = updatedHeaders.length - currentLastColumn;
  if (additionalColumns > 0) {
    if (currentLastColumn === 0) {
      sheet.insertColumns(1, additionalColumns);
    } else {
      sheet.insertColumnsAfter(currentLastColumn, additionalColumns);
    }
  }

  sheet.getRange(1, 1, 1, updatedHeaders.length).setValues([updatedHeaders]);

  if (additionalColumns > 0) {
    const boldStart = Math.max(currentLastColumn, 1) + 1;
    sheet.getRange(1, boldStart, 1, additionalColumns).setFontWeight("bold");
  }

  log.info(`Added ${additionalColumns} custom language columns to ${sheet.getName()}`);
}

function addCustomLanguagesToTranslationSheets(newLanguages: CustomLanguageInput[]): void {
  const normalized = normalizeCustomLanguages(newLanguages);
  if (normalized.length === 0) {
    return;
  }

  const log = getTranslationServiceLogger();
  log.info(`Preparing to add ${normalized.length} custom languages`, normalized);

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const categorySheet = spreadsheet.getSheetByName(CATEGORY_TRANSLATIONS_SHEET) || createCategoryTranslationsSheet();
  addCustomLanguagesToSheet(categorySheet, normalized);

  const translationSheetConfigs = [
    {
      name: DETAIL_LABEL_TRANSLATIONS_SHEET,
      sourceSheet: DETAILS_SHEET,
      sourceColumn: MAIN_LANGUAGE_COLUMN,
    },
    {
      name: DETAIL_HELPER_TEXT_TRANSLATIONS_SHEET,
      sourceSheet: DETAILS_SHEET,
      sourceColumn: DETAILS_HELPER_TEXT_COLUMN,
    },
    {
      name: DETAIL_OPTION_TRANSLATIONS_SHEET,
      sourceSheet: DETAILS_SHEET,
      sourceColumn: DETAILS_OPTIONS_COLUMN,
    },
  ];

  for (const config of translationSheetConfigs) {
    let sheet = spreadsheet.getSheetByName(config.name);
    if (!sheet) {
      sheet = createTranslationSheet(config.name, [], config.sourceSheet, config.sourceColumn);
    }
    addCustomLanguagesToSheet(sheet, normalized);
  }
}

/**
 * Adds custom language columns across all translation sheets.
 *
 * Deprecated alias maintained for backward compatibility with older dialogs.
 */
function addNewLanguages(newLanguages: { name: string; iso: string }[]): void {
  addCustomLanguagesToTranslationSheets(newLanguages);
}
