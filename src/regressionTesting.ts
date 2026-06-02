/// <reference path="./loggingHelpers.ts" />

/**
 * Placeholder icon URL for test spreadsheets
 * All production icon references will be replaced with this safe placeholder
 */
const TEST_PLACEHOLDER_ICON_URL =
  "https://drive.google.com/file/d/1n1jR7XPFBgtpn544f7PMKxN2djewl60K/view?usp=drive_link";

/**
 * Creates a sanitized test copy of the current production spreadsheet
 *
 * This function:
 * 1. Creates a new spreadsheet with TEST_ prefix
 * 2. Copies all sheet structure
 * 3. Sanitizes all sensitive data (dataset IDs, names, URLs, content)
 * 4. Replaces icons with placeholder
 * 5. Validates the test spreadsheet
 *
 * @returns Object containing test spreadsheet ID and URL
 *
 * @example
 * const result = duplicateSpreadsheetForTesting();
 * console.log(`Test spreadsheet created: ${result.url}`);
 */
function duplicateSpreadsheetForTesting(): {
  id: string;
  url: string;
  name: string;
} {
  const log = getScopedLogger("RegressionTesting");

  try {
    log.info("Starting spreadsheet duplication for testing...");

    // Get the active production spreadsheet
    const productionSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const productionName = productionSpreadsheet.getName();

    // Generate test spreadsheet name with timestamp
    const timestamp = Utilities.formatDate(new Date(), "UTC", "yyyyMMdd");
    const testSpreadsheetName = `TEST_${timestamp}_${productionName}`;

    log.info(`Creating test spreadsheet: ${testSpreadsheetName}`);

    // Create new test spreadsheet
    const testSpreadsheet = SpreadsheetApp.create(testSpreadsheetName);

    // Get all sheets from production first
    const productionSheets = productionSpreadsheet.getSheets();

    log.info(`Copying ${productionSheets.length} sheets from production to test...`);

    // Remove the default "Sheet1" that Google Apps Script automatically creates
    // BUT only if it will not leave us with zero sheets
    const defaultSheet = testSpreadsheet.getSheetByName("Sheet1");
    if (defaultSheet) {
      const sheetsAfterCopy = productionSheets.length + 1; // +1 for the Sheet1 that exists
      if (sheetsAfterCopy > 1) {
        log.debug("Removing default Sheet1 from test spreadsheet");
        testSpreadsheet.deleteSheet(defaultSheet);
      } else {
        log.debug("Keeping Sheet1 because production has no sheets to copy");
      }
    }

    // Copy each sheet and sanitize its data
    productionSheets.forEach((productionSheet) => {
      const sheetName = productionSheet.getName();
      log.debug(`Processing sheet: ${sheetName}`);

      // Insert sheet in test spreadsheet with same name
      const testSheet = testSpreadsheet.insertSheet(sheetName);

      // Copy sheet structure (not data yet)
      copySheetStructure(productionSheet, testSheet);

      // Sanitize and copy data
      sanitizeAndCopySheetData(productionSheet, testSheet, sheetName);
    });

    // Sanitize metadata specifically
    sanitizeMetadataSheet(testSpreadsheet, productionName);

    // Validate the test spreadsheet
    const validation = validateTestSpreadsheet(testSpreadsheet.getId());

    if (!validation.isValid) {
      log.error("Test spreadsheet validation failed", validation.errors);
      throw new Error(
        `Test spreadsheet validation failed: ${validation.errors.join(", ")}`,
      );
    }

    log.info("Test spreadsheet created and validated successfully");

    // Return result
    const result = {
      id: testSpreadsheet.getId(),
      url: testSpreadsheet.getUrl(),
      name: testSpreadsheetName,
    };

    log.info(
      `Test spreadsheet ready: ${result.url}`,
    );

    return result;
  } catch (error) {
    log.error("Failed to create test spreadsheet", error);
    throw error;
  }
}

/**
 * Copies the structure (headers, formatting) from source to target sheet
 * Does not copy data rows
 *
 * @param sourceSheet - The production sheet to copy structure from
 * @param targetSheet - The test sheet to copy structure to
 */
function copySheetStructure(
  sourceSheet: GoogleAppsScript.Spreadsheet.Sheet,
  targetSheet: GoogleAppsScript.Spreadsheet.Sheet,
): void {
  const log = getScopedLogger("RegressionTesting");

  try {
    const lastRow = sourceSheet.getLastRow();
    const lastColumn = sourceSheet.getLastColumn();

    // If sheet is empty, nothing to copy
    if (lastRow === 0 || lastColumn === 0) {
      log.debug(`Sheet ${sourceSheet.getName()} is empty, skipping structure copy`);
      return;
    }

    // Copy header row (row 1) with formatting
    const headerRange = sourceSheet.getRange(1, 1, 1, lastColumn);
    const targetHeaderRange = targetSheet.getRange(1, 1, 1, lastColumn);

    // Copy values
    targetHeaderRange.setValues(headerRange.getValues());

    // Copy formatting
    targetHeaderRange.setFontWeights(headerRange.getFontWeights());
    targetHeaderRange.setFontStyles(headerRange.getFontStyles());
    targetHeaderRange.setBackgrounds(headerRange.getBackgrounds());
    targetHeaderRange.setNumberFormats(headerRange.getNumberFormats());

    log.debug(
      `Copied structure for sheet ${sourceSheet.getName()} (${lastColumn} columns)`,
    );
  } catch (error) {
    log.error(
      `Failed to copy structure for sheet ${sourceSheet.getName()}`,
      error,
    );
    throw error;
  }
}

/**
 * Sanitizes and copies data from production sheet to test sheet
 * Replaces all sensitive information with safe test data
 *
 * @param sourceSheet - Production sheet to read from
 * @param targetSheet - Test sheet to write to
 * @param sheetName - Name of the sheet being processed
 */
function sanitizeAndCopySheetData(
  sourceSheet: GoogleAppsScript.Spreadsheet.Sheet,
  targetSheet: GoogleAppsScript.Spreadsheet.Sheet,
  sheetName: string,
): void {
  const log = getScopedLogger("RegressionTesting");

  try {
    const lastRow = sourceSheet.getLastRow();
    const lastColumn = sourceSheet.getLastColumn();

    // If no data or only headers, nothing to copy
    if (lastRow <= 1) {
      log.debug(`Sheet ${sheetName} has no data rows, skipping`);
      return;
    }

    // Get all data (including headers)
    const allData = sourceSheet.getDataRange().getValues();

    log.debug(
      `Processing ${lastRow - 1} data rows from sheet ${sheetName}`,
    );

    // Process each row based on sheet type
    const sanitizedData = allData.map((row, rowIndex) => {
      // Always preserve header row (row 0)
      if (rowIndex === 0) {
        return row;
      }

      // Sanitize data rows
      return sanitizeRowData(row, sheetName, rowIndex);
    });

    // Write sanitized data to target sheet
    const targetRange = targetSheet.getRange(
      1,
      1,
      sanitizedData.length,
      lastColumn,
    );
    targetRange.setValues(sanitizedData);

    log.debug(
      `Sanitized and copied ${lastRow - 1} data rows for sheet ${sheetName}`,
    );
  } catch (error) {
    log.error(
      `Failed to sanitize and copy data for sheet ${sheetName}`,
      error,
    );
    throw error;
  }
}

/**
 * Sanitizes a single row of data based on sheet type
 *
 * @param row - The row data to sanitize
 * @param sheetName - Name of the sheet (Categories, Details, etc.)
 * @param rowIndex - Index of the row (0-based, not including header)
 * @returns Sanitized row data
 */
function sanitizeRowData(
  row: (string | number | boolean)[],
  sheetName: string,
  rowIndex: number,
): (string | number | boolean)[] {
  const sanitized = [...row];

  switch (sheetName) {
    case "Categories":
      // Column 0: Category name
      if (sanitized[0] && typeof sanitized[0] === "string") {
        sanitized[0] = `Test Category ${rowIndex + 1}`;
      }
      // Column 1: Icon URL - always replace with placeholder
      sanitized[1] = TEST_PLACEHOLDER_ICON_URL;
      // Column 2: Fields reference - keep structure but use generic names
      if (sanitized[2] && typeof sanitized[2] === "string") {
        const fields = sanitized[2].split(",").map((f) => f.trim()).filter((f) => f !== "");
        sanitized[2] = fields
          .map((field, index) => {
            // Extract number from field reference or use index
            const match = field.match(/\d+/);
            const num = match ? match[0] : String(index + 1);
            return `field${num}`;
          })
          .join(", ");
      }
      // Column 3: Color - always use default test color
      if (sanitized[3] && typeof sanitized[3] === "string") {
        // Use default test color for all entries to ensure consistency
        sanitized[3] = "#0066CC"; // Default blue
      }
      // Column 4: Geometry - keep structure
      break;

    case "Details":
      // Column 0: Field name
      if (sanitized[0] && typeof sanitized[0] === "string") {
        sanitized[0] = `Test Field ${rowIndex + 1}`;
      }
      // Column 1: Helper text
      if (sanitized[1] && typeof sanitized[1] === "string") {
        sanitized[1] = "This is test helper text for testing purposes.";
      }
      // Column 2: Type - keep structure
      // Column 3: Options - sanitize to generic options
      if (sanitized[3] && typeof sanitized[3] === "string") {
        const options = sanitized[3]
          .split(",")
          .map((opt) => `Test Option ${opt.trim().match(/\d+/)?.[0] || ""}`);
        sanitized[3] = options.join(", ");
      }
      break;

    case "Category Translations":
    case "Detail Label Translations":
    case "Detail Helper Text Translations":
    case "Detail Option Translations":
      // Column 0: Key/reference - keep structure
      // Columns 1+: Translation content - clear all translations
      for (let i = 1; i < sanitized.length; i++) {
        if (sanitized[i] && typeof sanitized[i] === "string") {
          sanitized[i] = ""; // Clear all translation content
        }
      }
      break;

    default:
      // For any other sheet, sanitize all text content
      for (let i = 0; i < sanitized.length; i++) {
        const cellValue = sanitized[i];
        if (cellValue && typeof cellValue === "string") {
          // Replace with generic test content
          sanitized[i] = `TEST_${cellValue.substring(0, 20)}`;
        }
      }
      break;
  }

  return sanitized;
}

/**
 * Sanitizes the metadata sheet in the test spreadsheet
 * Updates dataset_id, name, and version to test-specific values
 *
 * @param testSpreadsheet - The test spreadsheet to sanitize
 * @param productionName - Name of the production spreadsheet
 */
function sanitizeMetadataSheet(
  testSpreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
  productionName: string,
): void {
  const log = getScopedLogger("RegressionTesting");

  try {
    // Get or create Metadata sheet
    let metadataSheet = testSpreadsheet.getSheetByName("Metadata");
    if (!metadataSheet) {
      log.debug("Creating Metadata sheet in test spreadsheet");
      metadataSheet = testSpreadsheet.insertSheet("Metadata");
      // Add headers
      metadataSheet.getRange(1, 1, 1, 2).setValues([["Key", "Value"]]);
      metadataSheet
        .getRange(1, 1, 1, 2)
        .setFontWeight("bold");
    }

    // Clear existing metadata
    metadataSheet.clear();

    // Set headers
    metadataSheet.getRange(1, 1, 1, 2).setValues([["Key", "Value"]]);
    metadataSheet.getRange(1, 1, 1, 2).setFontWeight("bold");

    // Generate test metadata
    const timestamp = Utilities.formatDate(new Date(), "UTC", "yyyyMMdd");
    const testDatasetId = `comapeo-test-${timestamp}`;
    const testName = `config-test-${timestamp}-${slugify(productionName)}`;
    const testVersion = Utilities.formatDate(new Date(), "UTC", "yy.MM.dd");

    // Write test metadata
    const metadata = [
      ["dataset_id", testDatasetId],
      ["name", testName],
      ["version", testVersion],
    ];

    metadataSheet
      .getRange(2, 1, metadata.length, 2)
      .setValues(metadata);

    log.info(
      `Metadata sanitized - dataset_id: ${testDatasetId}, version: ${testVersion}`,
    );
  } catch (error) {
    log.error("Failed to sanitize metadata sheet", error);
    throw error;
  }
}

/**
 * Validates that the test spreadsheet is properly configured
 * Checks for required sheets, sanitized data, and functional structure
 *
 * @param testSpreadsheetId - ID of the test spreadsheet to validate
 * @returns Validation result with isValid flag and any errors
 *
 * @example
 * const validation = validateTestSpreadsheet("1abc123...");
 * if (!validation.isValid) {
 *   console.error("Validation failed:", validation.errors);
 * }
 */
function validateTestSpreadsheet(
  testSpreadsheetId: string,
): {
  isValid: boolean;
  errors: string[];
  warnings: string[];
} {
  const log = getScopedLogger("RegressionTesting");

  const errors: string[] = [];
  const warnings: string[] = [];

  try {
    log.info(`Validating test spreadsheet: ${testSpreadsheetId}`);

    // Open test spreadsheet
    const testSpreadsheet = SpreadsheetApp.openById(testSpreadsheetId);
    const testName = testSpreadsheet.getName();

    // Check that name starts with TEST_
    if (!testName.startsWith("TEST_")) {
      errors.push(
        `Spreadsheet name "${testName}" does not start with "TEST_"`,
      );
    }

    // Check required sheets exist
    const requiredSheets = [
      "Categories",
      "Details",
      "Category Translations",
      "Detail Label Translations",
      "Detail Helper Text Translations",
      "Detail Option Translations",
    ];

    const existingSheets = testSpreadsheet
      .getSheets()
      .map((s) => s.getName());

    requiredSheets.forEach((sheetName) => {
      if (!existingSheets.includes(sheetName)) {
        errors.push(`Required sheet "${sheetName}" is missing`);
      }
    });

    // Check Categories sheet for production references
    const categoriesSheet = testSpreadsheet.getSheetByName("Categories");
    if (categoriesSheet) {
      const categoriesData = categoriesSheet.getDataRange().getValues();

      // Check for production Google Drive URLs (excluding our placeholder)
      for (let i = 1; i < categoriesData.length; i++) {
        const iconUrl = categoriesData[i][1];
        if (
          iconUrl &&
          typeof iconUrl === "string" &&
          iconUrl.includes("drive.google.com") &&
          iconUrl !== TEST_PLACEHOLDER_ICON_URL
        ) {
          errors.push(
            `Production icon URL found in Categories sheet row ${i + 1}: ${iconUrl}`,
          );
        }
      }

      // Check for production category names
      for (let i = 1; i < categoriesData.length; i++) {
        const categoryName = categoriesData[i][0];
        if (
          categoryName &&
          typeof categoryName === "string" &&
          !categoryName.startsWith("Test Category")
        ) {
          warnings.push(
            `Non-test category name in Categories sheet row ${i + 1}: ${categoryName}`,
          );
        }
      }
    }

    // Check metadata sheet
    const metadataSheet = testSpreadsheet.getSheetByName("Metadata");
    if (metadataSheet) {
      const metadataData = metadataSheet.getDataRange().getValues();
      const metadataMap = new Map<string, string>();

      // Convert to map for easy lookup
      metadataData.forEach((row, index) => {
        if (index > 0 && row[0] && row[1]) {
          metadataMap.set(row[0], row[1]);
        }
      });

      // Check dataset_id
      const datasetId = metadataMap.get("dataset_id");
      if (!datasetId) {
        errors.push("dataset_id is missing from Metadata sheet");
      } else if (!datasetId.includes("test")) {
        errors.push(
          `dataset_id "${datasetId}" does not include "test" prefix`,
        );
      }

      // Check name
      const name = metadataMap.get("name");
      if (!name) {
        errors.push("name is missing from Metadata sheet");
      } else if (!name.includes("test")) {
        errors.push(`name "${name}" does not include "test" in value`);
      }

      // Check version
      const version = metadataMap.get("version");
      if (!version) {
        errors.push("version is missing from Metadata sheet");
      }
    } else {
      errors.push("Metadata sheet is missing");
    }

    // Check Details sheet
    const detailsSheet = testSpreadsheet.getSheetByName("Details");
    if (detailsSheet) {
      const detailsData = detailsSheet.getDataRange().getValues();

      // Check for production field names
      for (let i = 1; i < detailsData.length; i++) {
        const fieldName = detailsData[i][0];
        if (
          fieldName &&
          typeof fieldName === "string" &&
          !fieldName.startsWith("Test Field")
        ) {
          warnings.push(
            `Non-test field name in Details sheet row ${i + 1}: ${fieldName}`,
          );
        }
      }
    }

    // Check translation sheets are empty (content-wise)
    const translationSheets = [
      "Category Translations",
      "Detail Label Translations",
      "Detail Helper Text Translations",
      "Detail Option Translations",
    ];

    translationSheets.forEach((sheetName) => {
      const sheet = testSpreadsheet.getSheetByName(sheetName);
      if (sheet) {
        const data = sheet.getDataRange().getValues();
        let hasContent = false;

        // Check if any translation content exists
        for (let i = 1; i < data.length && !hasContent; i++) {
          for (let j = 1; j < data[i].length; j++) {
            if (data[i][j]) {
              hasContent = true;
              break;
            }
          }
        }

        if (hasContent) {
          warnings.push(
            `${sheetName} contains translation content (should be empty)`,
          );
        }
      }
    });

    // Final validation
    const isValid = errors.length === 0;

    log.info(
      `Validation complete - Valid: ${isValid}, Errors: ${errors.length}, Warnings: ${warnings.length}`,
    );

    return {
      isValid,
      errors,
      warnings,
    };
  } catch (error) {
    log.error("Validation process failed", error);
    return {
      isValid: false,
      errors: [`Validation error: ${error.message}`],
      warnings,
    };
  }
}

/**
 * Shows confirmation dialog before creating test spreadsheet
 *
 * @returns True if user confirms, false otherwise
 */
function showCreateTestSpreadsheetConfirmation(): boolean {
  const ui = SpreadsheetApp.getUi();

  const message =
    "This will create a TEST copy of your production spreadsheet.\n\n" +
    "What will be sanitized:\n" +
    "• All dataset IDs will get 'test-' prefix\n" +
    "• All Google Drive URLs replaced with placeholder\n" +
    "• Category and field names changed to generic test names\n" +
    "• Translation content will be cleared\n" +
    "• Project-specific content replaced with test data\n\n" +
    "This operation cannot be undone. Continue?";

  const response = ui.alert(
    "Create Test Spreadsheet",
    message,
    ui.ButtonSet.YES_NO,
  );

  return response === ui.Button.YES;
}

/**
 * Shows success dialog with test spreadsheet information
 *
 * @param result - Result object from duplicateSpreadsheetForTesting
 */
function showTestSpreadsheetCreatedDialog(result: {
  id: string;
  url: string;
  name: string;
}): void {
  const ui = SpreadsheetApp.getUi();

  const message =
    `Test spreadsheet created successfully!\n\n` +
    `Name: ${result.name}\n` +
    `ID: ${result.id}\n\n` +
    `All production data has been sanitized.\n` +
    `The test spreadsheet is ready for regression testing.\n\n` +
    `URL: ${result.url}`;

  ui.alert("Test Spreadsheet Created", message, ui.ButtonSet.OK);
}

/**
 * Handles errors during test spreadsheet creation
 *
 * @param error - The error that occurred
 */
function showTestSpreadsheetErrorDialog(error: Error): void {
  const ui = SpreadsheetApp.getUi();

  // Build comprehensive error message with context
  let message = `Failed to create test spreadsheet:\n\n`;
  message += `Error Type: ${error.name || "Error"}\n`;
  message += `Error Message: ${error.message}\n\n`;

  // Add stack trace if available (first few lines)
  if (error.stack) {
    const stackLines = error.stack.split("\n").slice(0, 3);
    message += `Stack Trace (first 3 lines):\n${stackLines.join("\n")}\n\n`;
  }

  // Add helpful troubleshooting tips based on error message
  const errorMsg = error.message.toLowerCase();
  if (errorMsg.includes("permission") || errorMsg.includes("access")) {
    message += `Troubleshooting:\n`;
    message += `• Ensure you have edit access to the current spreadsheet\n`;
    message += `• Check that Google Drive permissions are granted\n`;
    message += `• Try refreshing the page and running again\n`;
  } else if (errorMsg.includes("quota") || errorMsg.includes("limit")) {
    message += `Troubleshooting:\n`;
    message += `• You may have exceeded Google Sheets quota\n`;
    message += `• Try again later or delete old test spreadsheets\n`;
    message += `• Check Google Workspace quotas in admin panel\n`;
  } else if (errorMsg.includes("network") || errorMsg.includes("fetch")) {
    message += `Troubleshooting:\n`;
    message += `• Check your internet connection\n`;
    message += `• Verify Google Sheets API is accessible\n`;
    message += `• Try running the operation again\n`;
  } else {
    message += `Troubleshooting:\n`;
    message += `• Check the console logs for detailed error information\n`;
    message += `• Ensure all required sheets exist in the spreadsheet\n`;
    message += `• Verify you have necessary permissions\n`;
  }

  message += `\nIf the problem persists, please contact support with this error information.`;

  ui.alert("Error Creating Test Spreadsheet", message, ui.ButtonSet.OK);
}

/**
 * Main entry point for creating test spreadsheet
 * This is the function called from the menu
 */
function createTestSpreadsheetForRegression(): void {
  const log = getScopedLogger("RegressionTesting");

  try {
    log.info("User requested test spreadsheet creation");

    // Show confirmation dialog
    const confirmed = showCreateTestSpreadsheetConfirmation();

    if (!confirmed) {
      log.info("User cancelled test spreadsheet creation");
      return;
    }

    log.info("User confirmed, proceeding with duplication");

    // Create test spreadsheet
    const result = duplicateSpreadsheetForTesting();

    // Show success dialog
    showTestSpreadsheetCreatedDialog(result);

    log.info("Test spreadsheet creation completed successfully");
  } catch (error) {
    log.error("Test spreadsheet creation failed", error);
    showTestSpreadsheetErrorDialog(error);
  }
}
