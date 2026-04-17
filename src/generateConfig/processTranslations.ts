/// <reference path="../loggingHelpers.ts" />
/// <reference path="../types.ts" />


function normalizeComparisonKey(value: unknown): string {
  if (typeof value === "string") {
    return value.trim().toLowerCase();
  }
  if (value === null || value === undefined) return "";
  return String(value).trim().toLowerCase();
}

function buildPresetLookup(presets: CoMapeoPreset[]) {
  const byName = new Map<string, CoMapeoPreset>();
  const byIcon = new Map<string, CoMapeoPreset>();

  presets.forEach((preset) => {
    if (preset?.name) {
      const key = normalizeComparisonKey(preset.name);
      if (key) byName.set(key, preset);
    }
    if (preset?.icon) {
      const key = normalizeComparisonKey(preset.icon);
      if (key) byIcon.set(key, preset);
    }
  });

  return {
    find(candidate: unknown): CoMapeoPreset | undefined {
      const key = normalizeComparisonKey(candidate);
      if (!key) return undefined;
      return byName.get(key) || byIcon.get(key);
    },
  };
}

function resolvePresetForTranslationRow(
  translationValues: string[],
  presets: CoMapeoPreset[],
  translationIndex: number,
  columnToLanguageMap: Record<number, string>,
  primaryLanguageCode: string,
  lookup: ReturnType<typeof buildPresetLookup>,
): CoMapeoPreset | undefined {
  const fallbackPreset = presets[translationIndex];

  const primaryColumnIndexEntry = Object.entries(columnToLanguageMap).find(
    ([, lang]) => lang === primaryLanguageCode,
  );
  const primaryColumnIndex = primaryColumnIndexEntry
    ? parseInt(primaryColumnIndexEntry[0], 10)
    : 0;

  const primaryCell =
    translationValues[primaryColumnIndex] ??
    translationValues[0] ??
    "";

  const matchedPreset = lookup.find(primaryCell);
  if (matchedPreset) {
    return matchedPreset;
  }

  if (fallbackPreset) {
    const debugValue = primaryCell || "(empty primary cell)";
    getScopedLogger("ProcessTranslations").warn(
      `⚠️  Could not match category translation row ${translationIndex + 2} using value "${debugValue}". Falling back to row order.`,
    );
  }

  return fallbackPreset;
}

/**
 * Build column-to-language mapping for a specific translation sheet.
 * This function reads the header row of a sheet and maps column indexes to language codes.
 *
 * @param sheetName - Name of the translation sheet to process
 * @returns Object with targetLanguages array and columnToLanguageMap record
 */
function buildColumnMapForSheet(sheetName: string): {
  targetLanguages: string[];
  columnToLanguageMap: Record<number, string>;
} {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  if (!sheet) {
    getScopedLogger("ProcessTranslations").warn(`⏭️  Sheet "${sheetName}" not found - using only primary language`);
    const primaryLanguage = getPrimaryLanguage();
    return {
      targetLanguages: [primaryLanguage.code],
      columnToLanguageMap: { 0: primaryLanguage.code },
    };
  }

  const lastColumn = sheet.getLastColumn();

  // Guard: Skip if sheet is empty or has no columns
  if (lastColumn === 0) {
    getScopedLogger("ProcessTranslations").warn(`⏭️  Sheet "${sheetName}" is empty - using only primary language`);
    const primaryLanguage = getPrimaryLanguage();
    return {
      targetLanguages: [primaryLanguage.code],
      columnToLanguageMap: { 0: primaryLanguage.code },
    };
  }

  const headerRow = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const targetLanguages: string[] = [];
  const columnToLanguageMap: Record<number, string> = {};

  // Extract language codes from header columns with explicit column index mapping
  const allLanguages = getAllLanguages();
  const resolveHeaderCode = createTranslationHeaderResolver(allLanguages);
  const seenLanguages = new Set<string>();
  const headerB = String(headerRow[1] || "").trim().toLowerCase();
  const headerC = String(headerRow[2] || "").trim().toLowerCase();
  const hasMetaColumns = headerB.includes("iso") && headerC.includes("source");
  const languageStartIndex = hasMetaColumns ? 3 : 1;

  for (let i = languageStartIndex; i < headerRow.length; i++) {
    const header = headerRow[i]?.toString().trim();
    if (!header) {
      getScopedLogger("ProcessTranslations").warn(`⚠️  Empty header at column ${i + 1} in "${sheetName}" - skipping`);
      continue;
    }

    const langCode = resolveHeaderCode(header);
    if (!langCode) {
      getScopedLogger("ProcessTranslations").warn(`⚠️  Could not parse language from header "${header}" at column ${i + 1} in "${sheetName}"`);
      continue;
    }

    const normalizedCode = langCode.toLowerCase();
    if (seenLanguages.has(normalizedCode)) {
      getScopedLogger("ProcessTranslations").warn(`⚠️  Duplicate language header "${header}" (${langCode}) in "${sheetName}" - skipping duplicate`);
      continue;
    }

    seenLanguages.add(normalizedCode);
    targetLanguages.push(langCode);
    columnToLanguageMap[i] = langCode;
    getScopedLogger("ProcessTranslations").info(`[${sheetName}] Column ${i} (${header}) → ${langCode}`);
  }

  getScopedLogger("ProcessTranslations").info(`[${sheetName}] Found ${targetLanguages.length} languages:`, targetLanguages);
  getScopedLogger("ProcessTranslations").info(`[${sheetName}] Column mapping:`, columnToLanguageMap);

  return { targetLanguages, columnToLanguageMap };
}

/**
 * Processes all translation sheets and generates CoMapeo translations object
 *
 * Reads translation data from Category Translations, Detail Label Translations,
 * Detail Helper Text Translations, and Detail Option Translations sheets.
 * Maps each language to its corresponding translations for categories, fields,
 * and options.
 *
 * @param data - Spreadsheet data object
 * @param fields - Array of processed CoMapeo fields
 * @param presets - Array of processed CoMapeo presets
 * @returns CoMapeoTranslations object with language codes as keys
 *
 * @example
 * const translations = processTranslations(data, fields, presets);
 * // Returns: { "en": {...}, "es": {...}, "fr": {...} }
 */
function processTranslations(data, fields, presets) {
  const log = getScopedLogger("ProcessTranslations");
  log.info("Starting processTranslations...");
  const primaryLanguage = getPrimaryLanguage();

  // Build initial column map from Category Translations to determine available languages
  const initialMapping = buildColumnMapForSheet("Category Translations");

  // Early return if Category Translations is empty
  if (initialMapping.targetLanguages.length === 1 &&
      initialMapping.targetLanguages[0] === getPrimaryLanguage().code) {
    const categorySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Category Translations");
    if (!categorySheet || categorySheet.getLastColumn() === 0) {
      getScopedLogger("ProcessTranslations").warn("⏭️  Category Translations sheet is empty - using only primary language");
      const messages: CoMapeoTranslations = Object.fromEntries(
        initialMapping.targetLanguages.map((lang) => [lang, {}]),
      );
      return messages;
    }
  }

  // Initialize messages object for all detected languages
  const messages: CoMapeoTranslations = Object.fromEntries(
    initialMapping.targetLanguages.map((lang) => [lang, {}]),
  );

  const translationSheets = sheets(true);
  getScopedLogger("ProcessTranslations").info("Processing translation sheets:", translationSheets);
  const presetLookup = buildPresetLookup(presets);

  for (const sheetName of translationSheets) {
    getScopedLogger("ProcessTranslations").info(`\nProcessing sheet: ${sheetName}`);

    // Guard check: Skip if translation sheet data doesn't exist
    if (!data[sheetName]) {
      getScopedLogger("ProcessTranslations").warn(`⏭️  Skipping sheet "${sheetName}" - sheet data not found (translation may have been skipped)`);
      continue;
    }

    // Build column map for THIS specific sheet (defense against manual edits)
    const { targetLanguages, columnToLanguageMap } = buildColumnMapForSheet(sheetName);

    const translations = data[sheetName].slice(1);
    getScopedLogger("ProcessTranslations").info(`Found ${translations.length} translations`);

    // Validation: Check that data columns match expected language count
    if (translations.length > 0) {
      const firstRowColumnCount = translations[0].length;

      // Missing columns is an ERROR - translations will be incomplete
      if (firstRowColumnCount < targetLanguages.length) {
        getScopedLogger("ProcessTranslations").error(`❌ MISSING COLUMNS in "${sheetName}":`, {
          expectedLanguages: targetLanguages.length,
          actualColumns: firstRowColumnCount,
          missingColumns: targetLanguages.length - firstRowColumnCount,
          targetLanguages: targetLanguages,
          firstRow: translations[0]
        });
        getScopedLogger("ProcessTranslations").error(`⚠️  Translation data incomplete - ${targetLanguages.length - firstRowColumnCount} language(s) missing!`);
        getScopedLogger("ProcessTranslations").error(`⚠️  Missing languages will have no translations for this sheet.`);
      }
      // Extra columns is just INFO - likely metadata columns, will be ignored
      else if (firstRowColumnCount > targetLanguages.length) {
        getScopedLogger("ProcessTranslations").info(`ℹ️  Extra columns detected in "${sheetName}":`, {
          expectedLanguages: targetLanguages.length,
          actualColumns: firstRowColumnCount,
          extraColumns: firstRowColumnCount - targetLanguages.length,
        });
        getScopedLogger("ProcessTranslations").info(`ℹ️  Extra columns will be ignored (likely metadata). Translation processing continues normally.`);
      }
    }

    for (
      let translationIndex = 0;
      translationIndex < translations.length;
      translationIndex++
    ) {
      const translationRow = translations[translationIndex];
      const translation = translationRow.map((t) =>
        t === null || t === undefined ? "" : t.toString().trim(),
      );
      getScopedLogger("ProcessTranslations").info(
        `\nProcessing translation ${translationIndex + 1}/${translations.length}:`,
        translation,
      );
      // Use explicit column mapping to handle gaps and ensure correct language assignment
      for (const [columnIndex, lang] of Object.entries(columnToLanguageMap)) {
        const colIdx = parseInt(columnIndex);
        const translationValue = translation[colIdx];

        // Defensive check: skip if translation value is missing
        if (translationValue === undefined || translationValue === null) {
          log.warn(`⚠️  Missing translation value at column ${colIdx} for language ${lang}`);
          continue;
        }

        const messageType = sheetName.startsWith("Category")
          ? "presets"
          : "fields";
        let item: CoMapeoPreset | CoMapeoField | undefined;
        if (messageType === "fields") {
          item = fields[translationIndex];
        } else {
          item = resolvePresetForTranslationRow(
            translation,
            presets,
            translationIndex,
            columnToLanguageMap,
            primaryLanguage.code,
            presetLookup,
          );
        }

        if (!item) {
          getScopedLogger("ProcessTranslations").warn(
            `⚠️  Skipping translation row ${translationIndex + 2} in "${sheetName}" - no matching ${
              messageType === "fields" ? "field" : "preset"
            } found.`,
          );
          continue;
        }
        const key =
          messageType === "presets"
            ? (item as CoMapeoPreset).icon
            : (item as CoMapeoField).tagKey;

        log.info(
          `Processing ${messageType} for language: ${lang} (column ${colIdx}), key: ${key}`,
        );

        switch (sheetName) {
          case "Category Translations":
            messages[lang][`${messageType}.${key}.name`] = {
              message: translationValue,
              description: `Name for preset '${key}'`,
            };
            log.info(`Added category translation for ${key}: "${translationValue}"`);
            break;
          case "Detail Label Translations":
            messages[lang][`${messageType}.${key}.label`] = {
              message: translationValue,
              description: `Label for field '${key}'`,
            };
            log.info(`Added label translation for ${key}: "${translationValue}"`);
            break;
          case "Detail Helper Text Translations":
            messages[lang][`${messageType}.${key}.helperText`] = {
              message: translationValue,
              description: `Helper text for field '${key}'`,
            };
            log.info(`Added helper text translation for ${key}: "${translationValue}"`);
            break;
          case "Detail Option Translations": {
            const field = item as CoMapeoField;
            const fieldType = getFieldType(field.type || "");
            log.info(`Processing options for field type: ${fieldType}`);

            if (
              fieldType !== "number" &&
              fieldType !== "text" &&
              translationValue &&
              translationValue.trim()
            ) {
              const options = translationValue
                .split(",")
                .map((opt) => opt.trim());
              log.info(`Found ${options.length} options to process`);

              for (const [optionIndex, option] of options.entries()) {
                if (field.options?.[optionIndex]) {
                  const optionKey = `${messageType}.${key}.options.${field.options[optionIndex].value}`;
                  const optionValue = {
                    message: {
                      label: option,
                      value: field.options[optionIndex].value,
                    },
                    description: `Option '${option}' for field '${field.label}'`,
                  };
                  messages[lang][optionKey] = optionValue;
                  log.info(`Added option translation: ${option} for ${key}`);
                }
              }
            }
            break;
          }
          default:
            log.info(`Unhandled sheet name: ${sheetName}`);
            break;
        }
      }
    }
  }

  getScopedLogger("ProcessTranslations").info("\nTranslation processing complete");
  getScopedLogger("ProcessTranslations").info(
    "Messages per language:",
    Object.keys(messages).map(
      (lang) => `${lang}: ${Object.keys(messages[lang]).length} messages`,
    ),
  );

  return messages;
}
