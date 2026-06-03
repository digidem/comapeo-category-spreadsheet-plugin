/// <reference path="shared.ts" />

// Generic linting function
function lintSheet(
  sheetName: string,
  columnValidations: ((value: string, row: number, col: number) => void)[],
  requiredColumns: number[] = [],
  preserveBackgroundColumns: number[] = [],
): void {
  console.time(`Linting ${sheetName}`);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    console.log(`${sheetName} sheet not found`);
    console.timeEnd(`Linting ${sheetName}`);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    console.log(`${sheetName} sheet is empty or contains only header`);
    console.timeEnd(`Linting ${sheetName}`);
    return;
  }

  try {
    if (lastRow > 1 && columnValidations.length > 0) {
      // Clear backgrounds, font colors, and stale lint notes for all columns
      // except those that should be preserved
      for (let col = 0; col < columnValidations.length; col++) {
        // Skip columns that should preserve their backgrounds (e.g., Categories icon column with user colors)
        const shouldClearBackground = !preserveBackgroundColumns.includes(col);
        const colRange = sheet.getRange(2, col + 1, lastRow - 1, 1);

        if (shouldClearBackground) {
          clearRangeBackgroundIfMatches(
            colRange,
            LINT_WARNING_BACKGROUND_COLORS,
          );
        }
        clearRangeFontColorIfMatches(colRange, LINT_WARNING_FONT_COLORS);
        // Clear stale [Lint] notes so re-linting produces a clean result
        clearRangeLintNotesWithPrefix(colRange, LINT_NOTE_PREFIX);
      }
    }

    // First clean any whitespace-only cells
    console.time(`Cleaning whitespace cells for ${sheetName}`);
    cleanWhitespaceOnlyCells(
      sheet,
      2,
      1,
      lastRow - 1,
      columnValidations.length,
    );
    console.timeEnd(`Cleaning whitespace cells for ${sheetName}`);

    // Check for duplicates in the first column (usually the name/identifier column)
    // Preserve background for columns in preserveBackgroundColumns (e.g., Categories
    // column A has user-managed category colors that the builder reads at export time).
    console.time(`Checking for duplicates in ${sheetName}`);
    checkForDuplicates(sheet, 1, 2, preserveBackgroundColumns.includes(0));
    console.timeEnd(`Checking for duplicates in ${sheetName}`);

    console.time(`Getting data for ${sheetName}`);
    // Get all data from the sheet, excluding the header row
    const dataRange = sheet.getRange(
      2,
      1,
      lastRow - 1,
      columnValidations.length,
    );
    const data = dataRange.getValues();
    console.timeEnd(`Getting data for ${sheetName}`);

    console.time(`Validating cells for ${sheetName}`);

    // Highlight required fields in batches before running column validations
    if (requiredColumns.length > 0) {
      const rangesToReset: string[] = [];
      requiredColumns.forEach((colIndex) => {
        // Skip columns that should preserve their backgrounds
        if (!preserveBackgroundColumns.includes(colIndex)) {
          const columnLetter = columnNumberToLetter(colIndex + 1);
          rangesToReset.push(`${columnLetter}2:${columnLetter}${lastRow}`);
        }
      });

      if (rangesToReset.length > 0) {
        sheet.getRangeList(rangesToReset).setBackground(null);
      }

      const requiredHighlights = new Map<number, number[]>();
      requiredColumns.forEach((colIndex) => {
        // Skip columns that should preserve their backgrounds
        if (!preserveBackgroundColumns.includes(colIndex)) {
          requiredHighlights.set(colIndex, []);
        }
      });

      data.forEach((row, rowIndex) => {
        requiredColumns.forEach((colIndex) => {
          // Skip columns that should preserve their backgrounds
          if (
            !preserveBackgroundColumns.includes(colIndex) &&
            isEmptyOrWhitespace(row[colIndex])
          ) {
            const rows = requiredHighlights.get(colIndex);
            if (rows) {
              rows.push(rowIndex + 2); // +2 accounts for header row
            }
          }
        });
      });

      requiredHighlights.forEach((rows, colIndex) => {
        if (!rows || rows.length === 0) {
          return;
        }
        const columnLetter = columnNumberToLetter(colIndex + 1);
        const rangeAddresses = rows.map(
          (rowNumber) => `${columnLetter}${rowNumber}`,
        );
        sheet.getRangeList(rangeAddresses).setBackground(LINT_WARNING_BG); // Light yellow for required fields
      });
    }

    // Iterate through each cell and apply the corresponding validation function
    data.forEach((row, rowIndex) => {
      row.forEach((cellValue, colIndex) => {
        if (columnValidations[colIndex]) {
          columnValidations[colIndex](
            String(cellValue || ""),
            rowIndex + 2,
            colIndex + 1,
          );
        }
      });
    });
    console.timeEnd(`Validating cells for ${sheetName}`);

    console.log(`${sheetName} sheet linting completed`);
  } catch (error) {
    console.error(`Error linting ${sheetName} sheet:`, error);
  } finally {
    console.timeEnd(`Linting ${sheetName}`);
  }
}

/**
 * Main linting function that validates all sheets in the spreadsheet.
 *
 * @param showAlerts - Whether to show UI alerts (default: true). Set to false when called from other functions.
 */
function lintAllSheets(showAlerts: boolean = true): void {
  try {
    // Clear cross-run caches so Drive file changes are picked up
    driveIconInfoCache.clear();

    console.log("Starting linting process...");

    console.log("Linting Categories sheet...");
    lintCategoriesSheet();

    console.log("Linting Details sheet...");
    lintDetailsSheet();

    console.log("Linting Icons sheet...");
    lintIconsSheet();

    // Phase 5: Cross-sheet icon checks
    checkCrossSheetIconCollisions();

    // Phase 6 Task 1: Inline SVG size warnings/errors
    checkInlineSvgSizes();

    console.log("Linting Translation sheets...");
    lintTranslationSheets();

    const strictLintMetrics = collectStrictLintMetrics();

    // Phase 6 Task 2: Per-locale translation payload size checks
    checkTranslationPayloadSizes(strictLintMetrics);

    console.log("Linting Metadata sheet...");
    lintMetadataSheet();

    // Phase 6 Task 3: Total entity count check
    checkTotalEntityCounts(strictLintMetrics);

    console.log("Finished linting all sheets.");

    // Add a summary of issues found, but only if showAlerts is true
    if (showAlerts) {
      const ui = SpreadsheetApp.getUi();
      ui.alert(
        "Linting Complete",
        "All sheets have been linted.\n\n" +
          "Review highlighted cells before generating config:\n" +
          "- Bright red with white text: Critical translation mismatch. Re-sync translations before export.\n" +
          "- Light red (#FFC7CE): Errors that can block or break config generation.\n" +
          "- Yellow (#FFF2CC): Warnings that need review.\n" +
          "- Light yellow (#FFFFCC): Advisory guidance.\n" +
          "- Red/orange text in icon columns: Icon source, access, or format issues.\n\n" +
          "Open the note on each flagged cell for the specific fix.",
        ui.ButtonSet.OK,
      );
    }
  } catch (error) {
    console.error("Error during linting process:", error);

    // Only show error alert if showAlerts is true
    if (showAlerts) {
      const ui = SpreadsheetApp.getUi();
      ui.alert(
        "Linting Error",
        "An error occurred during the linting process: " +
          error.message +
          "\n\nSome sheets may not have been fully processed.",
        ui.ButtonSet.OK,
      );
    }
  }
}
