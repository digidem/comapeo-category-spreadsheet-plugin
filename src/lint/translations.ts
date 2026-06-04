/// <reference path="shared.ts" />

/**
 * Check for duplicate slugs in translation sheets.
 * Warns when two different translated values produce the same slug,
 * which would cause silent overwrites during config generation.
 */
function checkDuplicateTranslationSlugs(): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const translationSheets = sheets(true);

  for (const sheetName of translationSheets) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) continue;

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) continue;

    try {
      // Get all translated values and check for duplicate slugs
      const lastCol = sheet.getLastColumn();
      if (lastCol < 4) continue; // Need at least Name, ISO, Source, and one translation

      // Clear only duplicate-slug lint notes from prior runs, preserving
      // option-count mismatch backgrounds/fonts set by validateSheetConsistency().
      clearRangeLintNotesWithPrefix(
        sheet.getRange(2, 4, lastRow - 1, lastCol - 3),
        `${LINT_NOTE_PREFIX}Duplicate translation slug`,
      );

      // Check each translation column (starting from column 4)
      for (let col = 4; col <= lastCol; col++) {
        const rawValues = sheet
          .getRange(2, col, lastRow - 1, 1)
          .getValues();

        // Build slug frequency map — iterate raw values in-place so i+2 maps to
        // the correct sheet row (blanks are skipped with continue, not filtered).
        const slugCounts = new Map<string, number[]>();
        for (let i = 0; i < rawValues.length; i++) {
          const value = String(rawValues[i][0] || "").trim();
          if (!value) continue;

          const slug = slugify(value);
          if (!slug) continue;

          if (!slugCounts.has(slug)) {
            slugCounts.set(slug, [i + 2]); // +2 for header and 0-index
          } else {
            slugCounts.get(slug)?.push(i + 2);
          }
        }

        // Highlight cells with duplicate slugs
        for (const [slug, rows] of slugCounts.entries()) {
          if (rows.length > 1) {
            getScopedLogger("LintTranslationSlugs").info(
              `Duplicate slug "${slug}" in ${sheetName} column ${col} at rows: ${rows.join(", ")}`,
            );
            const otherRowsStr = rows.join(", ");
            for (const row of rows) {
              appendLintNote(
                sheet.getRange(row, col),
                `Duplicate translation slug "${slug}" in rows: ${otherRowsStr}`,
                "warning",
              );
            }
          }
        }
      }
    } catch (error) {
      getScopedLogger("LintTranslationSlugs").error(`Error checking duplicate slugs in ${sheetName}:`, error);
    }
  }
}

/**
 * Validate translation headers are valid language names, ISO codes, or "Name - ISO" format
 */
function validateTranslationHeaders(): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const translationSheets = sheets(true);
  const allLanguages = getAllLanguages();
  const resolveHeaderCode = createTranslationHeaderResolver(allLanguages);

  for (const sheetName of translationSheets) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) continue;

    try {
      const lastCol = sheet.getLastColumn();
      if (lastCol < 2) continue; // Need at least a source column and one language column

      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      // Clear lint artifacts on header columns 2..lastCol only.
      // Column A (1,1) is excluded because validateSheetConsistency may have
      // placed a row-count mismatch note there; clearing it here would lose
      // that diagnostic if validateSheetConsistency is later skipped.
      // (lastCol >= 2 is guaranteed by the `continue` above.)
      clearLintArtifacts(sheet.getRange(1, 2, 1, lastCol - 1));

      const headerB = String(headers[1] || "")
        .trim()
        .toLowerCase();
      const headerC = String(headers[2] || "")
        .trim()
        .toLowerCase();
      const hasMetaColumns =
        headerB.includes("iso") && headerC.includes("source");
      const languageStartIndex = hasMetaColumns ? 3 : 1;

      // Check language columns (skip meta columns if present)
      // Track resolved locale codes to detect duplicates across columns
      const localeToColumns = new Map<string, number[]>();

      for (let i = languageStartIndex; i < headers.length; i++) {
        const header = String(headers[i] || "").trim();
        if (!header) continue;

        const parsedCode = resolveHeaderCode(header);

        if (!parsedCode) {
          getScopedLogger("LintTranslationHeaders").info(
            `Invalid translation header "${header}" in ${sheetName} column ${i + 1} - should be a language name, ISO code, or "Name - ISO" format`,
          );
          setLintNote(
            sheet.getRange(1, i + 1),
            `Invalid translation header "${header}" — should be a language name, ISO code, or "Name - ISO" format`,
            "error",
          );
          continue;
        }

        // Track locale → column indices for duplicate detection
        if (!localeToColumns.has(parsedCode)) {
          localeToColumns.set(parsedCode, [i]);
        } else {
          localeToColumns.get(parsedCode)?.push(i);
        }
      }

      // Flag duplicate locale headers
      localeToColumns.forEach((columns, code) => {
        if (columns.length > 1) {
          const firstCol = columns[0] + 1; // 1-based for display
          for (let idx = 1; idx < columns.length; idx++) {
            const colIndex = columns[idx];
            const header = String(headers[colIndex] || "").trim();
            const cell = sheet.getRange(1, colIndex + 1);
            setLintNote(
              cell,
              `Duplicate locale: header "${header}" resolves to "${code}" which is already used by column ${firstCol}`,
              "error",
            );
          }
        }
      });
    } catch (error) {
      getScopedLogger("LintTranslationHeaders").error(`Error validating headers in ${sheetName}:`, error);
    }
  }
}

/**
 * Validates that translation sheets have consistent headers and row counts with their source sheets.
 */
function validateTranslationSheetConsistency(): void {
  getScopedLogger("LintTranslationConsistency").info("Validating translation sheet consistency...");

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // Validate Category Translations
    const categoriesSheet = spreadsheet.getSheetByName("Categories");
    const categoryTranslationsSheet = spreadsheet.getSheetByName(
      "Category Translations",
    );

    if (categoriesSheet && categoryTranslationsSheet) {
      validateSheetConsistency(
        categoriesSheet,
        categoryTranslationsSheet,
        "Category Translations",
        false,
      );
    }

    // Validate Detail translations
    const detailsSheet = spreadsheet.getSheetByName("Details");

    if (detailsSheet) {
      const detailLabelTranslations = spreadsheet.getSheetByName(
        "Detail Label Translations",
      );
      const detailHelperTranslations = spreadsheet.getSheetByName(
        "Detail Helper Text Translations",
      );
      const detailOptionTranslations = spreadsheet.getSheetByName(
        "Detail Option Translations",
      );

      if (detailLabelTranslations) {
        validateSheetConsistency(
          detailsSheet,
          detailLabelTranslations,
          "Detail Label Translations",
          false,
        );
      }

      if (detailHelperTranslations) {
        validateSheetConsistency(
          detailsSheet,
          detailHelperTranslations,
          "Detail Helper Text Translations",
          false,
        );
      }

      if (detailOptionTranslations) {
        validateSheetConsistency(
          detailsSheet,
          detailOptionTranslations,
          "Detail Option Translations",
          true, // Special handling for option count validation
        );
      }
    }

    getScopedLogger("LintTranslationConsistency").info("Translation sheet consistency validation complete");
  } catch (error) {
    getScopedLogger("LintTranslationConsistency").error("Error validating translation sheet consistency:", error);
  }
}

/**
 * Validates consistency between a source sheet and its translation sheet.
 *
 * @param sourceSheet - The source sheet (Categories or Details)
 * @param translationSheet - The translation sheet to validate
 * @param translationSheetName - Name for logging
 * @param validateOptionCounts - Whether to validate option counts (for Detail Option Translations)
 */
function validateSheetConsistency(
  sourceSheet: GoogleAppsScript.Spreadsheet.Sheet,
  translationSheet: GoogleAppsScript.Spreadsheet.Sheet,
  translationSheetName: string,
  validateOptionCounts: boolean,
): void {
  getScopedLogger("LintTranslationConsistency").debug(
    `Validating consistency for ${translationSheetName} against ${sourceSheet.getName()}`,
  );

  try {
    // OPTIMIZATION: Read row counts once
    const sourceRowCount = sourceSheet.getLastRow();
    const translationRowCount = translationSheet.getLastRow();

    if (translationRowCount > 1 && translationSheet.getLastColumn() > 0) {
      const dataRange = translationSheet.getRange(
        2,
        1,
        translationRowCount - 1,
        translationSheet.getLastColumn(),
      );
      clearRangeBackgroundIfMatches(
        dataRange,
        [LINT_ERROR_BG, LINT_WARNING_BG, LINT_CRITICAL_BG], // Include bright red for primary column mismatches
      );
      clearRangeFontColorIfMatches(
        dataRange,
        [LINT_CRITICAL_FONT], // White text paired with red backgrounds for primary column mismatches
      );
    }

    // Check row count consistency (excluding header)
    if (sourceRowCount !== translationRowCount) {
      getScopedLogger("LintTranslationConsistency").warn(
        `Row count mismatch in ${translationSheetName}: ` +
          `Source has ${sourceRowCount} rows, translation has ${translationRowCount} rows`,
      );

      // Clear any stale row-count mismatch note from a prior lint run to prevent
      // duplicate accumulation (the background/font clear above only covers dataRange,
      // rows 2+, so A1 is excluded and the note would persist indefinitely).
      clearRangeLintNotesWithPrefix(
        translationSheet.getRange(1, 1),
        `${LINT_NOTE_PREFIX}Row count mismatch:`,
      );

      // Highlight the discrepancy in the translation sheet
      // Use cached translationRowCount instead of calling getLastRow() again
      if (translationRowCount > 0) {
        appendLintNote(
          translationSheet.getRange(1, 1),
          `Row count mismatch: source has ${sourceRowCount} rows, translation has ${translationRowCount} rows. Re-sync translation sheets before generating config.`,
          "warning",
        );
      }
    }

    // CRITICAL: Validate primary language column values match between source and translation sheets
    // This prevents translation lookup failures during config generation (e.g., "Animal" vs "Animal Terrs")
    const minRowCheck = Math.min(sourceRowCount, translationRowCount) - 1; // Exclude header
    if (minRowCheck > 0) {
      // Determine source column based on translation sheet type
      let sourceColumn = 1; // Default to column A
      if (translationSheetName === "Detail Helper Text Translations") {
        sourceColumn = 2; // Column B for helper text
      } else if (translationSheetName === "Detail Option Translations") {
        sourceColumn = 4; // Column D for options
      }

      // Read source and translation primary columns
      const sourceData = sourceSheet
        .getRange(2, sourceColumn, minRowCheck, 1)
        .getValues();
      const translationData = translationSheet
        .getRange(2, 1, minRowCheck, 1)
        .getValues();

      // Collect mismatched cells
      const mismatchedCells: Array<{
        row: number;
        sourceValue: string;
        translationValue: string;
      }> = [];

      for (let i = 0; i < minRowCheck; i++) {
        const sourceValue = String(sourceData[i][0] || "").trim();
        const translationValue = String(translationData[i][0] || "").trim();

        // Normalize comparison (case-insensitive, whitespace-normalized)
        const normalizedSource = sourceValue.toLowerCase().replace(/\s+/g, " ");
        const normalizedTranslation = translationValue
          .toLowerCase()
          .replace(/\s+/g, " ");

        if (normalizedSource !== normalizedTranslation) {
          mismatchedCells.push({
            row: i + 2, // +2 for header row and 0-index
            sourceValue,
            translationValue,
          });

          getScopedLogger("LintTranslationConsistency").warn(
            `Primary column mismatch in ${translationSheetName} row ${i + 2}: ` +
              `Source="${sourceValue}", Translation="${translationValue}"`,
          );
        }
      }

      // Highlight mismatched cells in bright red to indicate critical error
      if (mismatchedCells.length > 0) {
        const rangeStrings = mismatchedCells.map(({ row }) =>
          translationSheet.getRange(row, 1).getA1Notation(),
        );
        const rangeList = translationSheet.getRangeList(rangeStrings);
        rangeList.setBackground(LINT_CRITICAL_BG); // Bright red for critical mismatch
        rangeList.setFontColor(LINT_CRITICAL_FONT); // White text for visibility
      }
    }

    // Validate option counts for Detail Option Translations
    if (validateOptionCounts && sourceRowCount > 1) {
      const detailsSheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Details");
      if (!detailsSheet) return;

      // OPTIMIZATION: Read all data once in a single batch operation
      const translationLastCol = translationSheet.getLastColumn();

      // Read options column from Details sheet (column 4)
      const detailsData = detailsSheet
        .getRange(2, 4, sourceRowCount - 1, 1)
        .getValues();

      // Read all translation data in one operation
      const translationData = translationSheet
        .getRange(2, 1, translationRowCount - 1, translationLastCol)
        .getValues();

      // OPTIMIZATION: Collect all cells that need highlighting instead of setting individually
      const cellsToHighlight: Array<{ row: number; col: number }> = [];

      // Validate each row in a single pass
      const minRows = Math.min(detailsData.length, translationData.length);
      for (let i = 0; i < minRows; i++) {
        const sourceOptions = String(detailsData[i][0] || "").trim();
        if (!sourceOptions) continue; // Skip if no options in source

        // Source options use comma-only splitting to match parseOptions() in the builder.
        const sourceOptionCount = sourceOptions
          .split(",")
          .map((opt) => opt.trim())
          .filter((opt) => opt !== "").length;

        // Check each translation column (starting from column 4, after Name, ISO, Source columns)
        for (let col = 3; col < translationData[i].length; col++) {
          const translatedOptions = String(
            translationData[i][col] || "",
          ).trim();
          if (!translatedOptions) continue;

          const translatedOptionCount = translatedOptions
            .split(/[;,，、]/)
            .map((opt) => opt.trim())
            .filter((opt) => opt !== "").length;

          if (sourceOptionCount !== translatedOptionCount) {
            getScopedLogger("LintTranslationConsistency").warn(
              `Option count mismatch in ${translationSheetName} at row ${i + 2}, column ${col + 1}: ` +
                `Expected ${sourceOptionCount} options, found ${translatedOptionCount}`,
            );

            // Collect cell for batch highlighting
            cellsToHighlight.push({ row: i + 2, col: col + 1 });
          }
        }
      }

      // OPTIMIZATION: Apply all highlights in a single batch operation using RangeList
      if (cellsToHighlight.length > 0) {
        const rangeStrings = cellsToHighlight.map(({ row, col }) =>
          translationSheet.getRange(row, col).getA1Notation(),
        );
        const rangeList = translationSheet.getRangeList(rangeStrings);
        rangeList.setBackground(LINT_ERROR_BG); // Light red for mismatch
      }
    }
  } catch (error) {
    getScopedLogger("LintTranslationConsistency").error(
      `Error validating ${translationSheetName} consistency:`,
      error,
    );
  }
}

/**
 * Warns when multiple translation rows in the same sheet have source values
 * that slugify to the same key. Later rows would silently overwrite earlier ones
 * during config generation.
 */
function checkTranslationSourceOverwrites(): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const translationSheets = sheets(true);

  for (const sheetName of translationSheets) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) continue;

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) continue;

    try {
      // Source column is always column A (index 0) for all translation sheets
      const sourceRange = sheet.getRange(2, 1, lastRow - 1, 1);
      clearSourceOverwriteLintArtifacts(sourceRange);
      const sourceValues = sourceRange.getValues();

      // Track slug → array of { row, originalValue }
      const slugMap = new Map<string, Array<{ row: number; value: string }>>();

      for (let i = 0; i < sourceValues.length; i++) {
        const value = String(sourceValues[i][0] || "").trim();
        if (!value) continue;

        const slug = slugify(value);
        if (!slug) continue;

        const entry = { row: i + 2, value };
        if (!slugMap.has(slug)) {
          slugMap.set(slug, [entry]);
        } else {
          slugMap.get(slug)?.push(entry);
        }
      }

      // Warn on duplicate slugs
      slugMap.forEach((entries, slug) => {
        if (entries.length > 1) {
          for (const entry of entries) {
            const otherRows = entries
              .filter((e) => e.row !== entry.row)
              .map((e) => e.row);
            const cell = sheet.getRange(entry.row, 1);
            appendLintNote(
              cell,
              `Source value "${entry.value}" produces the same key as row ${otherRows.join(", ")}. Later values may overwrite earlier ones.`,
              "warning",
            );
          }
        }
      });
    } catch (error) {
      getScopedLogger("LintTranslationSourceOverwrites").error(
        `Error checking source overwrites in ${sheetName}:`,
        error,
      );
    }
  }
}

function lintTranslationSheets(): void {
  // First validate translation headers
  getScopedLogger("LintTranslations").info("Validating translation headers...");
  validateTranslationHeaders();

  const translationSheets = sheets(true);
  translationSheets.forEach((sheetName) => {
    const sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      getScopedLogger("LintTranslations").error(`Sheet "${sheetName}" not found`);
      return;
    }

    try {
      getScopedLogger("LintTranslations").info("Linting translation sheet: " + sheetName);

      // First clean any whitespace-only cells
      cleanWhitespaceOnlyCells(
        sheet,
        1,
        1,
        sheet.getLastRow(),
        sheet.getLastColumn(),
      );

      // Get all data from the sheet
      const data = sheet.getDataRange().getValues();
      // Capitalize the first letter of each cell if it's a string and not empty
      const updatedData = data.map((row) =>
        row.map((cell) =>
          typeof cell === "string" && cell.trim() !== ""
            ? capitalizeFirstLetter(cell)
            : cell,
        ),
      );

      // Update the sheet with the capitalized data
      sheet.getDataRange().setValues(updatedData);
      getScopedLogger("LintTranslations").info("Finished linting translation sheet: " + sheetName);
    } catch (error) {
      getScopedLogger("LintTranslations").error(
        "Error linting translation sheet " + sheetName + ":",
        error,
      );
    }
  });

  // After basic linting, validate translation sheet consistency
  validateTranslationSheetConsistency();

  // Phase 4: Warn on source-value slug collisions that cause silent overwrites
  checkTranslationSourceOverwrites();

  // Phase 4 (continued): Warn on duplicate translated-value slugs per column
  checkDuplicateTranslationSlugs();
}

/**
 * Detects translation sheet mismatches without modifying anything.
 * Returns details about mismatches found.
 *
 * @returns Object with mismatch details or null if no mismatches found
 */
function detectTranslationMismatches(): {
  hasMismatches: boolean;
  details: Array<{
    sheetName: string;
    sourceSheet: string;
    sourceColumn: string;
    mismatches: Array<{
      row: number;
      sourceValue: string;
      translationValue: string;
    }>;
  }>;
} | null {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const allDetails: Array<{
    sheetName: string;
    sourceSheet: string;
    sourceColumn: string;
    mismatches: Array<{
      row: number;
      sourceValue: string;
      translationValue: string;
    }>;
  }> = [];

  // Check Category Translations
  const categoriesSheet = spreadsheet.getSheetByName("Categories");
  const categoryTranslationsSheet = spreadsheet.getSheetByName(
    "Category Translations",
  );

  if (categoriesSheet && categoryTranslationsSheet) {
    const catMismatches = checkSheetMismatches(
      categoriesSheet,
      categoryTranslationsSheet,
      1, // Column A
    );
    if (catMismatches.length > 0) {
      allDetails.push({
        sheetName: "Category Translations",
        sourceSheet: "Categories",
        sourceColumn: "A",
        mismatches: catMismatches,
      });
    }
  }

  // Check Detail translations
  const detailsSheet = spreadsheet.getSheetByName("Details");
  if (detailsSheet) {
    const detailLabelTranslations = spreadsheet.getSheetByName(
      "Detail Label Translations",
    );
    const detailHelperTranslations = spreadsheet.getSheetByName(
      "Detail Helper Text Translations",
    );
    const detailOptionTranslations = spreadsheet.getSheetByName(
      "Detail Option Translations",
    );

    if (detailLabelTranslations) {
      const labelMismatches = checkSheetMismatches(
        detailsSheet,
        detailLabelTranslations,
        1,
      );
      if (labelMismatches.length > 0) {
        allDetails.push({
          sheetName: "Detail Label Translations",
          sourceSheet: "Details",
          sourceColumn: "A",
          mismatches: labelMismatches,
        });
      }
    }

    if (detailHelperTranslations) {
      const helperMismatches = checkSheetMismatches(
        detailsSheet,
        detailHelperTranslations,
        2,
      );
      if (helperMismatches.length > 0) {
        allDetails.push({
          sheetName: "Detail Helper Text Translations",
          sourceSheet: "Details",
          sourceColumn: "B",
          mismatches: helperMismatches,
        });
      }
    }

    if (detailOptionTranslations) {
      const optionMismatches = checkSheetMismatches(
        detailsSheet,
        detailOptionTranslations,
        4,
      );
      if (optionMismatches.length > 0) {
        allDetails.push({
          sheetName: "Detail Option Translations",
          sourceSheet: "Details",
          sourceColumn: "D",
          mismatches: optionMismatches,
        });
      }
    }
  }

  if (allDetails.length === 0) {
    return null;
  }

  return {
    hasMismatches: true,
    details: allDetails,
  };
}

/**
 * Helper function to check mismatches between source and translation sheet
 */
function checkSheetMismatches(
  sourceSheet: GoogleAppsScript.Spreadsheet.Sheet,
  translationSheet: GoogleAppsScript.Spreadsheet.Sheet,
  sourceColumn: number,
): Array<{ row: number; sourceValue: string; translationValue: string }> {
  const sourceRowCount = sourceSheet.getLastRow();
  const translationRowCount = translationSheet.getLastRow();
  const minRowCheck = Math.min(sourceRowCount, translationRowCount) - 1;

  if (minRowCheck <= 0) {
    return [];
  }

  const sourceData = sourceSheet
    .getRange(2, sourceColumn, minRowCheck, 1)
    .getValues();
  const translationData = translationSheet
    .getRange(2, 1, minRowCheck, 1)
    .getValues();

  const mismatches: Array<{
    row: number;
    sourceValue: string;
    translationValue: string;
  }> = [];

  for (let i = 0; i < minRowCheck; i++) {
    const sourceValue = String(sourceData[i][0] || "").trim();
    const translationValue = String(translationData[i][0] || "").trim();

    const normalizedSource = sourceValue.toLowerCase().replace(/\s+/g, " ");
    const normalizedTranslation = translationValue
      .toLowerCase()
      .replace(/\s+/g, " ");

    if (normalizedSource !== normalizedTranslation) {
      mismatches.push({
        row: i + 2,
        sourceValue,
        translationValue,
      });
    }
  }

  return mismatches;
}

/**
 * Fixes translation sheet mismatches by re-syncing formulas from source sheets
 * and optionally re-running translation for configured languages.
 *
 * @param reTranslate - Whether to re-run translation after fixing formulas
 * @param mismatchData - Optional pre-detected mismatch data (avoids re-detection after formula sync)
 */
function fixTranslationMismatches(
  reTranslate: boolean = true,
  mismatchData?: {
    hasMismatches: boolean;
    details: Array<{
      sheetName: string;
      sourceSheet: string;
      sourceColumn: string;
      mismatches: Array<{
        row: number;
        sourceValue: string;
        translationValue: string;
      }>;
    }>;
  } | null,
): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Re-sync Category Translations formulas
  const categoriesSheet = spreadsheet.getSheetByName("Categories");
  const categoryTranslationsSheet = spreadsheet.getSheetByName(
    "Category Translations",
  );

  if (categoriesSheet && categoryTranslationsSheet) {
    const lastRow = categoriesSheet.getLastRow();
    if (lastRow > 1) {
      const formula = `=Categories!A2:A${lastRow}`;
      categoryTranslationsSheet
        .getRange(2, 1, lastRow - 1, 1)
        .setFormula(formula);
      getScopedLogger("LintFixTranslationSheets").info(
        `Re-synced Category Translations formulas (rows 2-${lastRow})`,
      );
    }
  }

  // Re-sync Detail translations formulas
  const detailsSheet = spreadsheet.getSheetByName("Details");
  if (detailsSheet) {
    const lastRow = detailsSheet.getLastRow();

    if (lastRow > 1) {
      // Detail Label Translations (column A)
      const detailLabelSheet = spreadsheet.getSheetByName(
        "Detail Label Translations",
      );
      if (detailLabelSheet) {
        const formula = `=Details!A2:A${lastRow}`;
        detailLabelSheet.getRange(2, 1, lastRow - 1, 1).setFormula(formula);
        getScopedLogger("LintFixTranslationSheets").info(
          `Re-synced Detail Label Translations formulas (rows 2-${lastRow})`,
        );
      }

      // Detail Helper Text Translations (column B)
      const detailHelperSheet = spreadsheet.getSheetByName(
        "Detail Helper Text Translations",
      );
      if (detailHelperSheet) {
        const formula = `=Details!B2:B${lastRow}`;
        detailHelperSheet.getRange(2, 1, lastRow - 1, 1).setFormula(formula);
        getScopedLogger("LintFixTranslationSheets").info(
          `Re-synced Detail Helper Text Translations formulas (rows 2-${lastRow})`,
        );
      }

      // Detail Option Translations (column D)
      const detailOptionSheet = spreadsheet.getSheetByName(
        "Detail Option Translations",
      );
      if (detailOptionSheet) {
        const formula = `=Details!D2:D${lastRow}`;
        detailOptionSheet.getRange(2, 1, lastRow - 1, 1).setFormula(formula);
        getScopedLogger("LintFixTranslationSheets").info(
          `Re-synced Detail Option Translations formulas (rows 2-${lastRow})`,
        );
      }
    }
  }

  // Re-translate if requested
  if (reTranslate) {
    getScopedLogger("LintFixTranslationSheets").info("Re-running translation for configured languages...");

    // Get configured target languages from Category Translations headers
    if (categoryTranslationsSheet) {
      const lastColumn = categoryTranslationsSheet.getLastColumn();
      if (lastColumn > 1) {
        const headers = categoryTranslationsSheet
          .getRange(1, 1, 1, lastColumn)
          .getValues()[0];
        const targetLanguages: TranslationLanguage[] = [];
        const allLanguages = getAllLanguages();
        const resolveHeaderCode = createTranslationHeaderResolver(allLanguages);
        const seenLanguages = new Set<string>();
        const headerB = String(headers[1] || "")
          .trim()
          .toLowerCase();
        const headerC = String(headers[2] || "")
          .trim()
          .toLowerCase();
        const hasMetaColumns =
          headerB.includes("iso") && headerC.includes("source");
        const languageStartIndex = hasMetaColumns ? 3 : 1;

        // Extract language codes from headers (skip first column which is primary language)
        for (let i = languageStartIndex; i < headers.length; i++) {
          const header = String(headers[i] || "").trim();
          if (!header) continue;

          const langCode = resolveHeaderCode(header);
          if (!langCode) continue;

          const normalizedCode = langCode.toLowerCase();
          if (seenLanguages.has(normalizedCode)) continue;
          seenLanguages.add(normalizedCode);
          targetLanguages.push(langCode as TranslationLanguage);
        }

        if (targetLanguages.length > 0) {
          getScopedLogger("LintFixTranslationSheets").info(
            `Re-translating to languages: ${targetLanguages.join(", ")}`,
          );

          // OPTIMIZATION: Only clear and re-translate rows that have mismatches
          // Use pre-detected mismatch data if provided, otherwise detect now
          const mismatchResult = mismatchData || detectTranslationMismatches();

          if (mismatchResult && mismatchResult.hasMismatches) {
            getScopedLogger("LintFixTranslationSheets").info(
              `Found ${mismatchResult.details.length} translation sheet(s) with mismatches`,
            );

            // Clear only the mismatched rows in each sheet
            mismatchResult.details.forEach((detail) => {
              const sheet = spreadsheet.getSheetByName(detail.sheetName);
              if (!sheet || sheet.getLastColumn() <= 1) return;

              const colCount = sheet.getLastColumn() - 1; // Exclude primary language column
              const mismatchedRows = detail.mismatches.map((m) => m.row);

              // Clear translation columns for only the mismatched rows
              mismatchedRows.forEach((row) => {
                sheet.getRange(row, 2, 1, colCount).clearContent();
              });

              getScopedLogger("LintFixTranslationSheets").info(
                `Cleared translations for ${mismatchedRows.length} mismatched row(s) in ${detail.sheetName}: ` +
                  `rows ${mismatchedRows.join(", ")}`,
              );
            });
          } else {
            getScopedLogger("LintFixTranslationSheets").info("No mismatches detected, skipping clearing step");
          }

          autoTranslateSheetsBidirectional(targetLanguages);
        } else {
          getScopedLogger("LintFixTranslationSheets").info(
            "No target languages found in headers, skipping re-translation",
          );
        }
      }
    }
  }

  // Clear any bright red highlighting and white font from the primary column after fixing
  // This removes the visual indicators since the issues have been resolved
  const translationSheetNames = [
    "Category Translations",
    "Detail Label Translations",
    "Detail Helper Text Translations",
    "Detail Option Translations",
  ];

  translationSheetNames.forEach((sheetName) => {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet && sheet.getLastRow() > 1) {
      const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1);
      clearRangeBackgroundIfMatches(range, [LINT_CRITICAL_BG]);
      clearRangeFontColorIfMatches(range, [LINT_CRITICAL_FONT]);
    }
  });

  getScopedLogger("LintFixTranslationSheets").info("Translation sheet fix complete");
}
