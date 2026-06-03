/// <reference path="shared.ts" />

function buildLintCategorySummaries(
  data: SheetData,
): Array<{ id: string; name: string }> {
  const rows = data.Categories?.slice(1) || [];
  const headers = data.Categories?.[0] || [];
  const headerMap = createLintHeaderMap(headers);
  const nameCol = getLintColumnIndex(headerMap, "name") ?? CATEGORY_COL.NAME;
  const idCol = getLintColumnIndex(headerMap, "category id", "id");

  return rows
    .map((row, index) => {
      const name = String(row[nameCol] || "").trim();
      if (!name) {
        return null;
      }

      const explicitId =
        idCol !== undefined ? String(row[idCol] || "").trim() : "";
      return {
        id: explicitId || slugify(name) || `category-${index + 1}`,
        name,
      };
    })
    .filter(
      (
        category,
      ): category is {
        id: string;
        name: string;
      } => Boolean(category),
    );
}

function buildLintFieldSummaries(
  data: SheetData,
): Array<{
  id: string;
  name: string;
  helperText?: string;
  options?: SelectOption[];
}> {
  const rows = data.Details?.slice(1) || [];
  const headers = data.Details?.[0] || [];
  const headerMap = createLintHeaderMap(headers);
  const nameCol = getLintColumnIndex(headerMap, "name", "label") ?? DETAILS_COL.NAME;
  const helperCol =
    getLintColumnIndex(headerMap, "helper text", "helper", "help") ??
    DETAILS_COL.HELPER_TEXT;
  const typeCol = getLintColumnIndex(headerMap, "type") ?? DETAILS_COL.TYPE;
  const optionsCol = getLintColumnIndex(headerMap, "options") ?? DETAILS_COL.OPTIONS;
  const idCol = getLintColumnIndex(headerMap, "id") ?? DETAILS_COL.ID;
  type LintFieldSummary = {
    id: string;
    name: string;
    helperText?: string;
    options?: SelectOption[];
  };

  return rows
    .map((row, index) => {
      const name = String(row[nameCol] || "").trim();
      if (!name) {
        return null;
      }

      const helperText = String(row[helperCol] || "");
      const typeRaw = String(row[typeCol] || "")
        .trim()
        .toLowerCase();
      const optionsStr = String(row[optionsCol] || "");
      const explicitId = String(row[idCol] || "").trim();

      let options: SelectOption[] | undefined;
      let resolvedType: "selectOne" | "selectMultiple" | "number" | "text";
      const typeKey = typeRaw.charAt(0);
      switch (typeKey) {
        case "m":
          resolvedType = "selectMultiple";
          options = parseOptions(optionsStr);
          break;
        case "n":
          resolvedType = "number";
          break;
        case "t":
          resolvedType = "text";
          break;
        case "s":
        case "":
        default:
          resolvedType = "selectOne";
          options = parseOptions(optionsStr);
          break;
      }

      const isSelectType =
        resolvedType === "selectOne" || resolvedType === "selectMultiple";
      if (isSelectType && (!options || options.length === 0)) {
        return null;
      }

      if (options && options.length > 0) {
        const seenValues = new Set<string>();
        options = options.filter((option) => {
          if (seenValues.has(option.value)) {
            return false;
          }
          seenValues.add(option.value);
          return true;
        });
      }

      const fieldSummary: LintFieldSummary = {
        id: explicitId || slugify(name) || `field-${index + 1}`,
        name,
      };

      if (helperText) {
        fieldSummary.helperText = helperText;
      }

      if (options && options.length > 0) {
        fieldSummary.options = options;
      }

      return fieldSummary;
    })
    .filter(
      (field): field is LintFieldSummary => field !== null,
    );
}

function buildLintTranslationsByLocale(
  data: SheetData,
  categories: Array<{ id: string; name: string }>,
  fields: Array<{
    id: string;
    name: string;
    helperText?: string;
    options?: SelectOption[];
  }>,
): {
  translationsByLocale: Record<string, any>;
  localeHeaderRefs: Record<
    string,
    Array<{ sheetName: string; column: number; header: string }>
  >;
} {
  const translationsByLocale: Record<string, any> = {};
  const localeHeaderRefs: Record<
    string,
    Array<{ sheetName: string; column: number; header: string }>
  > = {};

  const ensureLocaleEntry = (locale: string): Record<string, any> => {
    if (!translationsByLocale[locale]) {
      translationsByLocale[locale] = {};
    }
    return translationsByLocale[locale];
  };

  const ensureCategoryEntry = (
    locale: string,
    categoryId: string,
  ): Record<string, string> => {
    const localeEntry = ensureLocaleEntry(locale);
    localeEntry.category = localeEntry.category || {};
    localeEntry.category[categoryId] = localeEntry.category[categoryId] || {};
    return localeEntry.category[categoryId];
  };

  const ensureFieldEntry = (
    locale: string,
    fieldId: string,
  ): Record<string, string> => {
    const localeEntry = ensureLocaleEntry(locale);
    localeEntry.field = localeEntry.field || {};
    localeEntry.field[fieldId] = localeEntry.field[fieldId] || {};
    return localeEntry.field[fieldId];
  };

  const categoryIdByName = new Map<string, string>();
  categories.forEach((category) => {
    categoryIdByName.set(category.name, category.id);
  });

  const fieldIdByName = new Map<string, string>();
  fields.forEach((field) => {
    fieldIdByName.set(field.name, field.id);
  });

  const fieldsByHelperText = new Map<string, string[]>();
  fields.forEach((field) => {
    if (!field.helperText) {
      return;
    }
    const fieldIds = fieldsByHelperText.get(field.helperText) || [];
    fieldIds.push(field.id);
    fieldsByHelperText.set(field.helperText, fieldIds);
  });

  const fieldsByOptionsString = new Map<
    string,
    Array<{ id: string; optionCount: number }>
  >();
  fields.forEach((field) => {
    if (!field.options || field.options.length === 0) {
      return;
    }
    const optionsString = field.options
      .map((option) => {
        const value = option?.value || "";
        const label = option?.label || "";
        if (!label) {
          return "";
        }
        return value === slugify(label) ? label : `${value}:${label}`;
      })
      .filter(Boolean)
      .join(", ");
    if (!optionsString) {
      return;
    }
    const matches = fieldsByOptionsString.get(optionsString) || [];
    matches.push({ id: field.id, optionCount: field.options.length });
    fieldsByOptionsString.set(optionsString, matches);
  });

  const processSheet = (
    sheetName: string,
    handler: (row: any[], locale: string, value: string) => void,
  ) => {
    const sheetData = data[sheetName];
    if (!sheetData || sheetData.length <= 1) {
      return;
    }

    const headers = sheetData[0];
    const langColumns = extractLanguagesFromHeaders(headers);
    langColumns.forEach((entry) => {
      const refs = localeHeaderRefs[entry.lang] || [];
      refs.push({
        sheetName,
        column: entry.index + 1,
        header: String(headers[entry.index] || "").trim(),
      });
      localeHeaderRefs[entry.lang] = refs;
    });

    sheetData.slice(1).forEach((row) => {
      langColumns.forEach((entry) => {
        const value = String(row[entry.index] || "").trim();
        if (!value) {
          return;
        }
        handler(row, entry.lang, value);
      });
    });
  };

  processSheet("Category Translations", (row, locale, value) => {
    const sourceName = String(row[TRANSLATION_COL.SOURCE_TEXT] || "").trim();
    const categoryId = categoryIdByName.get(sourceName);
    if (!categoryId) {
      return;
    }
    ensureCategoryEntry(locale, categoryId).name = value;
  });

  processSheet("Detail Label Translations", (row, locale, value) => {
    const sourceName = String(row[TRANSLATION_COL.SOURCE_TEXT] || "").trim();
    const fieldId = fieldIdByName.get(sourceName);
    if (!fieldId) {
      return;
    }
    ensureFieldEntry(locale, fieldId).label = value;
  });

  processSheet("Detail Helper Text Translations", (row, locale, value) => {
    const sourceHelperText = String(
      row[TRANSLATION_COL.SOURCE_TEXT] || "",
    ).trim();
    const matchingFieldIds = fieldsByHelperText.get(sourceHelperText) || [];
    matchingFieldIds.forEach((fieldId) => {
      ensureFieldEntry(locale, fieldId).helperText = value;
    });
  });

  processSheet("Detail Option Translations", (row, locale, value) => {
    const sourceOptions = String(row[TRANSLATION_COL.SOURCE_TEXT] || "").trim();
    const matchingFields = fieldsByOptionsString.get(sourceOptions) || [];
    const translatedOptions = splitTranslatedOptions(value);
    matchingFields.forEach((field) => {
      for (
        let optionIndex = 0;
        optionIndex < translatedOptions.length &&
        optionIndex < field.optionCount;
        optionIndex++
      ) {
        ensureFieldEntry(locale, field.id)[`options.${optionIndex}`] =
          translatedOptions[optionIndex];
      }
    });
  });

  Object.keys(translationsByLocale).forEach((locale) => {
    const entry = translationsByLocale[locale];
    const hasCategories = Boolean(
      entry.category && Object.keys(entry.category).length > 0,
    );
    const hasFields = Boolean(entry.field && Object.keys(entry.field).length > 0);
    if (!hasCategories && !hasFields) {
      delete translationsByLocale[locale];
    }
  });

  return { translationsByLocale, localeHeaderRefs };
}

function collectStrictLintMetrics(): {
  categoryCount: number;
  fieldCount: number;
  iconCount: number;
  optionCount: number;
  translationEntryCount: number;
  translationsByLocale: Record<string, any>;
  localeHeaderRefs: Record<
    string,
    Array<{ sheetName: string; column: number; header: string }>
  >;
} {
  // Perf note (#29/#30): getSpreadsheetData() is invoked exactly once per lint
  // run — only here, and lintAllSheets() calls collectStrictLintMetrics() once.
  // A per-execution cache would therefore be dead code. The remaining overlap is
  // between this single full-workbook snapshot and the per-sheet getRange() reads
  // each lint check performs; eliminating that would require routing every check
  // through a shared in-memory snapshot, i.e. a lint-engine redesign that is
  // explicitly out of scope. The single read is acceptable for a manual lint.
  const data = getSpreadsheetData();
  const categories = buildLintCategorySummaries(data);
  const fields = buildLintFieldSummaries(data);
  const { translationsByLocale, localeHeaderRefs } = buildLintTranslationsByLocale(
    data,
    categories,
    fields,
  );

  let iconCount = 0;
  try {
    iconCount = buildIconsFromSheet(data).length;
  } catch (error) {
    getScopedLogger("LintStrictMetrics").warn(
      "Falling back to Icons sheet row count while calculating entity totals:",
      error,
    );
    const iconsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Icons");
    iconCount = iconsSheet ? Math.max(0, iconsSheet.getLastRow() - 1) : 0;
  }

  const optionCount = fields.reduce(
    (total, field) => total + (field.options?.length || 0),
    0,
  );
  const translationEntryCount =
    typeof countTranslationEntries === "function"
      ? countTranslationEntries(translationsByLocale)
      : 0;

  return {
    categoryCount: categories.length,
    fieldCount: fields.length,
    iconCount,
    optionCount,
    translationEntryCount,
    translationsByLocale,
    localeHeaderRefs,
  };
}

/**
 * Phase 6 Task 2: Checks per-locale translation payload sizes.
 * Uses the merged per-locale translation object so lint matches strict validation.
 */
function checkTranslationPayloadSizes(metrics: {
  translationsByLocale: Record<string, any>;
  localeHeaderRefs: Record<
    string,
    Array<{ sheetName: string; column: number; header: string }>
  >;
}): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const logger = getScopedLogger("LintTranslationPayloadSizes");

  const PAYLOAD_LIMIT_BYTES = 1048576; // 1 MB

  Object.keys(metrics.translationsByLocale).forEach((locale) => {
    const localePayload = metrics.translationsByLocale[locale];
    const localeBytes = getByteLength(JSON.stringify(localePayload || {}));
    if (localeBytes <= PAYLOAD_LIMIT_BYTES) {
      return;
    }

    const sizeKB = Math.round(localeBytes / 1024);
    const refs = metrics.localeHeaderRefs[locale] || [];
    refs.forEach((ref) => {
      const sheet = spreadsheet.getSheetByName(ref.sheetName);
      if (!sheet) {
        return;
      }
      appendLintNote(
        sheet.getRange(1, ref.column),
        `Combined translations for locale '${locale}' exceed 1MB limit (${sizeKB}KB) after merging all translation sheets. This will fail strict validation.`,
        "error",
      );
    });
    logger.warn(
      `Merged translation payload for locale ${locale} is ${sizeKB}KB across ${refs.length} sheet column(s)`,
    );
  });
}

/**
 * Phase 6 Task 3: Checks strict total entity count parity.
 * Includes categories, fields, icons, option entries, and translation entries.
 */
function checkTotalEntityCounts(metrics: {
  categoryCount: number;
  fieldCount: number;
  iconCount: number;
  optionCount: number;
  translationEntryCount: number;
}): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const logger = getScopedLogger("LintTotalEntityCounts");

  const ENTITY_LIMIT = 10000;
  const totalCount =
    metrics.categoryCount +
    metrics.fieldCount +
    metrics.iconCount +
    metrics.optionCount +
    metrics.translationEntryCount;

  if (totalCount > ENTITY_LIMIT) {
    const categoriesSheet = spreadsheet.getSheetByName("Categories");
    const cell = categoriesSheet
      ? categoriesSheet.getRange(1, 1)
      : spreadsheet.getActiveSheet().getRange(1, 1);
    // A1 doubles as the primary-language cell (see validatePrimaryLanguageInA1).
    // Use the preserve-background variant so this advisory does not overwrite any
    // user-set A1 background; it still signals the error via red font + note.
    appendLintNotePreserveBackground(
      cell,
      `Total entity count (${totalCount}) exceeds 10,000 limit. Categories: ${metrics.categoryCount}, Details: ${metrics.fieldCount}, Icons: ${metrics.iconCount}, Options: ${metrics.optionCount}, Translations: ${metrics.translationEntryCount}. This will fail strict validation.`,
      "error",
    );
    logger.warn(
      `Total entity count ${totalCount} exceeds limit (Categories: ${metrics.categoryCount}, Details: ${metrics.fieldCount}, Icons: ${metrics.iconCount}, Options: ${metrics.optionCount}, Translations: ${metrics.translationEntryCount})`,
    );
  }
}
