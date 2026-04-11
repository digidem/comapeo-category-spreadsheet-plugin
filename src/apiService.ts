/**
 * CoMapeo Config API Service v2.0.0
 * JSON-only build endpoint, no ZIP workflow
 */
const logApiService = getScopedLogger("ApiService");

let AUTO_CREATED_APPLIES_COLUMN = false;
let AUTO_CREATED_CATEGORY_ID_COLUMN = false;

// =============================================================================
// Constants
// =============================================================================

// Column indices are declared globally in src/constants/columns.ts

/** Expected sheet headers for validation */
const EXPECTED_HEADERS = {
  CATEGORIES: ['Name', 'Icon', 'Fields', 'Applies'],
  DETAILS: ['Name', 'Helper Text', 'Type', 'Options', 'ID', 'Universal']
};

// =============================================================================
// Validation Functions
// =============================================================================

/**
 * Validates that sheet headers match expected format
 * Logs warnings if headers don't match but doesn't throw to allow flexibility
 *
 * @param sheetName - Name of the sheet being validated
 * @param headers - Actual headers from the sheet
 * @param expected - Expected headers
 * @returns true if valid, false if invalid (but doesn't throw)
 */
function validateSheetHeaders(
  sheetName: string,
  headers: string[],
  expected: string[]
): boolean {
  if (headers.length < expected.length) {
    logApiService.warn(
      `${sheetName} sheet has ${headers.length} columns but expected ${expected.length}. ` +
      `Expected: [${expected.join(', ')}], Found: [${headers.join(', ')}]. ` +
      `This may cause issues if columns are misaligned.`
    );
    return false;
  }

  for (let i = 0; i < expected.length; i++) {
    if (headers[i] !== expected[i]) {
      logApiService.warn(
        `${sheetName} sheet column ${i + 1} header mismatch: ` +
        `expected "${expected[i]}" but found "${headers[i]}". ` +
        `Data may not be read correctly.`
      );
      return false;
    }
  }

  return true;
}

// =============================================================================
// API Communication
// =============================================================================

/**
 * Sends a JSON build request to the API and returns the .comapeocat file
 *
 * @param buildRequest - The build request payload
 * @param maxRetries - Maximum number of total attempts (default: 3)
 * @returns URL to the saved .comapeocat file
 */
function sendBuildRequest(buildRequest: BuildRequest, maxRetries: number = RETRY_CONFIG.MAX_RETRIES): string {
  validateBuildRequest(buildRequest, { strict: true });
  // v2 JSON endpoint (legacy ZIP endpoint was /v1, prior code used /build)
  const apiUrl = `${API_BASE_URL}/v2`;
  let attemptNumber = 0;
  let lastError: Error | null = null;
  const startTime = Date.now();

  while (attemptNumber < maxRetries) {
    // Check total timeout before attempting
    const elapsedTime = Date.now() - startTime;
    if (elapsedTime >= RETRY_CONFIG.MAX_TOTAL_TIMEOUT_MS) {
      // Close any open modal dialogs before showing the error alert
      closeProcessingModalDialog();

      const ui = SpreadsheetApp.getUi();
      const errorMessage = `API request timed out after ${Math.round(elapsedTime / 1000)} seconds.\n\n` +
        `Last error: ${lastError?.message || "Maximum timeout exceeded"}\n\n` +
        `Please check your connection and try again.`;
      ui.alert("Request Timeout", errorMessage, ui.ButtonSet.OK);
      throw new Error(`API request exceeded maximum timeout of ${RETRY_CONFIG.MAX_TOTAL_TIMEOUT_MS}ms: ${lastError?.message || "timeout"}`);
    }

    attemptNumber++;

    try {
      if (attemptNumber > 1) {
        // Close any open modal dialogs before showing the retry alert
        closeProcessingModalDialog();

        const ui = SpreadsheetApp.getUi();
        ui.alert(
          "Retrying API Request",
          `Retry ${attemptNumber - 1} of ${maxRetries - 1}. Previous attempt failed: ${lastError?.message || "Unknown error"}`,
          ui.ButtonSet.OK
        );
      }

      logApiService.info(`Sending JSON build request to API (attempt ${attemptNumber} of ${maxRetries}):`, apiUrl);

      const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(buildRequest),
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch(apiUrl, options);
      const responseCode = response.getResponseCode();
      logApiService.info('Response code:', responseCode);

      if (responseCode === 200) {
        const responseBlob = response.getBlob();
        const contentType = response.getHeaders()['Content-Type'] || '';
        logApiService.info('Content-Type:', contentType);

        // Verify we got a binary file response
        if (contentType.includes('application/octet-stream') || contentType.includes('application/zip')) {
          const fileName = `${buildRequest.metadata.name}-${buildRequest.metadata.version}.comapeocat`;
          const blob = responseBlob.setName(fileName);
          return saveComapeocatToDrive(blob);
        }

        // Check if it's actually an error response in JSON
        try {
          const errorResponse: ApiErrorResponse = JSON.parse(response.getContentText());
          lastError = new Error(`API Error: ${errorResponse.message}` +
            (errorResponse.details?.errors ? ` - ${errorResponse.details.errors.join(', ')}` : ''));
        } catch {
          lastError = new Error("API returned unexpected response format");
        }
      } else {
        // Handle error response
        try {
          const errorResponse: ApiErrorResponse = JSON.parse(response.getContentText());
          lastError = new Error(`API Error (${responseCode}): ${errorResponse.message}` +
            (errorResponse.details?.errors ? ` - ${errorResponse.details.errors.join(', ')}` : ''));
        } catch {
          lastError = new Error(`API request failed with status ${responseCode}: ${response.getContentText()}`);
        }
      }

      logApiService.error(lastError?.message);

      // Only sleep if we're going to retry (not on last attempt)
      if (attemptNumber < maxRetries) {
        Utilities.sleep(RETRY_CONFIG.BASE_DELAY_MS * Math.pow(2, attemptNumber - 1));
      }
    } catch (error) {
      lastError = error;
      logApiService.error(`Error in API request (attempt ${attemptNumber} of ${maxRetries}):`, error);

      // Only sleep if we're going to retry (not on last attempt)
      if (attemptNumber < maxRetries) {
        Utilities.sleep(RETRY_CONFIG.BASE_DELAY_MS * Math.pow(2, attemptNumber - 1));
      }
    }
  }

  // Exhausted all retries
  // Close any open modal dialogs before showing the error alert
  closeProcessingModalDialog();

  const ui = SpreadsheetApp.getUi();
  const errorMessage = `Failed to generate the CoMapeo category file after ${maxRetries} attempts.\n\n` +
    `Last error: ${lastError?.message || "Unknown error"}\n\n` +
    `Please check your internet connection and try again.`;

  ui.alert("Error Generating CoMapeo Category", errorMessage, ui.ButtonSet.OK);

  throw new Error(`Failed to generate CoMapeo category after ${maxRetries} attempts: ${lastError?.message || "Unknown error"}`);
}

/**
 * Saves the .comapeocat blob to Google Drive
 *
 * @param blob - The .comapeocat file blob (name should already be set to name-version.comapeocat)
 * @returns URL to the saved file
 */
function saveComapeocatToDrive(blob: GoogleAppsScript.Base.Blob): string {
  logApiService.info('Saving .comapeocat file to Drive...');
  const configFolder = getConfigFolder();

  // Get or create builds folder
  let buildsFolderObj: GoogleAppsScript.Drive.Folder;
  const buildsFolders = configFolder.getFoldersByName('builds');
  if (buildsFolders.hasNext()) {
    buildsFolderObj = buildsFolders.next();
  } else {
    buildsFolderObj = configFolder.createFolder('builds');
  }

  // Use the blob's existing name (already set to name-version.comapeocat pattern)
  // This preserves both name and version to prevent file collisions
  const file = buildsFolderObj.createFile(blob);
  const fileUrl = file.getUrl();
  logApiService.info(`Download the .comapeocat file here: ${fileUrl}`);

  // Also create a ZIP archive of the .comapeocat file for easier sharing
  const zipName = `${blob.getName()}.zip`;
  const zipBlob = Utilities.zip([blob], zipName);
  const zipFile = buildsFolderObj.createFile(zipBlob);
  const zipUrl = zipFile.getUrl();
  logApiService.info(`Download the zipped .comapeocat file here: ${zipUrl}`);

  // Return ZIP URL by default
  return zipUrl;
}

// =============================================================================
// Payload Building
// =============================================================================

/**
 * Migrates old 4-column spreadsheet format to new 6-column format with Applies + ID columns
 * Old format: Name, Icon, Fields, Color
 * New format: Name, Icon, Fields, Applies, Category ID, Icon ID
 *
 * This function is idempotent - it checks if migration is needed before proceeding.
 * It validates the exact format before migrating to prevent data corruption.
 *
 * Safety features:
 * - Validates format before any modifications
 * - Warns user about unexpected formats and requires manual verification
 * - Logs all operations for debugging
 * - Preserves all existing data during column insertion
 *
 * MANUAL TESTING CHECKLIST (cannot be automated due to SpreadsheetApp dependency):
 * [ ] Test migration from old 4-column Categories format (Name, Icon, Fields, Color)
 * [ ] Test migration from old Details format without ID column
 * [ ] Test that already-migrated sheets are not re-migrated (idempotency)
 * [ ] Test data preservation during migration (verify no data loss)
 * [ ] Test edge case: empty sheets
 * [ ] Test edge case: sheets with unexpected number of columns
 * [ ] Test edge case: repeated calls to migration (should be safe)
 * [ ] Verify console logs show appropriate messages for each scenario
 */
function migrateSpreadsheetFormat(): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const migrationWarnings: string[] = [];
  const migrationNotices: string[] = [];

  AUTO_CREATED_APPLIES_COLUMN = false;
  AUTO_CREATED_CATEGORY_ID_COLUMN = false;

  // Get or create Metadata sheet first (needed for storing primary language)
  let metadataSheet = spreadsheet.getSheetByName('Metadata');
  if (!metadataSheet) {
    metadataSheet = spreadsheet.insertSheet('Metadata');
    metadataSheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]).setFontWeight('bold');
  }

  // Migrate Categories sheet
  const categoriesSheet = spreadsheet.getSheetByName('Categories');
  if (categoriesSheet) {
    const lastCol = categoriesSheet.getLastColumn();
    if (lastCol === 0) {
      logApiService.info('Categories sheet is empty, skipping migration');
    } else {
      let headers = categoriesSheet.getRange(1, 1, 1, lastCol).getValues()[0];
      let headersLower = headers.map(normalizeHeaderLabel);

      const refreshCategoryHeaders = () => {
        headers = categoriesSheet.getRange(1, 1, 1, categoriesSheet.getLastColumn()).getValues()[0];
        headersLower = headers.map(normalizeHeaderLabel);
      };

      const logCategoryWarning = (message: string) => {
        logApiService.warn(message);
        migrationWarnings.push(message);
      };

      const requireCategoryColumn = (
        index: number,
        predicate: (value: string) => boolean,
        description: string,
        canonical: string
      ): boolean => {
        if (headersLower.length <= index) {
          logCategoryWarning(`Categories sheet is missing the "${description}" column (position ${index + 1}).`);
          return false;
        }
        const headerValue = headersLower[index];
        if (!predicate(headerValue)) {
          logCategoryWarning(`Categories column ${index + 1} must be "${description}" but found "${headers[index] || ''}".`);
          return false;
        }
        if (headers[index] !== canonical) {
          setHeaderValue(categoriesSheet, index, canonical);
          refreshCategoryHeaders();
        }
        return true;
      };

      const headerA1Value = String(headers[0] || '');
      let isRecognizedLanguageHeader = false;
      if (headerA1Value && headerA1Value !== 'Name') {
        try {
          if (typeof validateLanguageName === 'function') {
            const validation = validateLanguageName(headerA1Value);
            isRecognizedLanguageHeader = Boolean(validation?.valid);
          }
        } catch (validationError) {
          logApiService.warn('Failed to validate Categories!A1 language header:', validationError);
        }
      }

      if (headerA1Value && headerA1Value !== 'Name' && isRecognizedLanguageHeader) {
        logApiService.info(`Preserving primary language '${headerA1Value}' from Categories!A1 before migration`);

        const metadataData = metadataSheet.getDataRange().getValues();
        let primaryLanguageExists = false;
        for (let i = 1; i < metadataData.length; i++) {
          if (metadataData[i][0] === 'primaryLanguage') {
            primaryLanguageExists = true;
            break;
          }
        }

        if (!primaryLanguageExists) {
          metadataSheet.appendRow(['primaryLanguage', headerA1Value]);
          logApiService.info(`Stored primaryLanguage='${headerA1Value}' in Metadata sheet`);
        }

        categoriesSheet.getRange(1, 1, 1, 4)
          .setValues([['Name', 'Icon', 'Fields', 'Applies']])
          .setFontWeight('bold');
        logApiService.info('Updated Categories headers to Name/Icon/Fields/Applies');
        refreshCategoryHeaders();
      }
      else if (headerA1Value && headerA1Value !== 'Name') {
        const response = ui.alert(
          'Convert Categories Header?',
          `Categories row 1 currently contains "${headerA1Value}" instead of the required "Name" header.` +
          '\n\nSelect "Yes" to automatically rename this header to "Name". ' +
          'You will then need to manually set the "primaryLanguage" entry in the Metadata sheet to the previous value.' +
          '\n\nSelect "No" to cancel and update the sheet manually.',
          ui.ButtonSet.YES_NO
        );

        if (response === ui.Button.YES) {
          categoriesSheet.getRange(1, 1).setValue('Name').setFontWeight('bold');
          refreshCategoryHeaders();

          ui.alert(
            'Primary Language Reminder',
            `The Categories header was updated to "Name". Please open the Metadata sheet and set the "primaryLanguage" row to "${headerA1Value}" before running the generator again.`,
            ui.ButtonSet.OK
          );
        } else {
          logCategoryWarning(`Categories sheet column 1 must be "Name" but found "${headerA1Value}".`);
        }
      }

      const nameOk = requireCategoryColumn(0, value => value === 'name', 'Name', 'Name');
      const iconOk = requireCategoryColumn(1, value => value.startsWith('icon'), 'Icon', 'Icon');
      const fieldsOk = requireCategoryColumn(
        2,
        value => value.startsWith('field') || value.startsWith('detail'),
        'Fields',
        'Fields'
      );

      const baseColumnsValid = nameOk && iconOk && fieldsOk;

      if (baseColumnsValid) {
        try {
          const appliesCreated = ensureAppliesColumn(categoriesSheet);
          refreshCategoryHeaders();
          if (appliesCreated) {
            AUTO_CREATED_APPLIES_COLUMN = true;
            migrationNotices.push('Applies column was missing. Added "Applies" header and preset row D2 to "track, observation". Review column D and update each category with the correct appliesTo values before running the generator again.');
          }
        } catch (error) {
          logCategoryWarning(`Failed to ensure Applies column on Categories sheet: ${(error as Error).message}`);
        }

        try {
          const idCreated = ensureCategoryIdColumn(categoriesSheet);
          refreshCategoryHeaders();
          if (idCreated) {
            AUTO_CREATED_CATEGORY_ID_COLUMN = true;
            migrationNotices.push('Category ID column was missing and has been added automatically. IDs were generated from category names; verify column E before rerunning.');
          }
        } catch (error) {
          logCategoryWarning(`Failed to ensure Category ID column on Categories sheet: ${(error as Error).message}`);
        }

        const categoriesLayoutError = validateCategoriesHeaderLayout(categoriesSheet);
        if (categoriesLayoutError) {
          logCategoryWarning(categoriesLayoutError);
        }
      } else {
        logCategoryWarning(`Categories sheet must start with columns Name, Icon, Fields. Found: ${headers.join(', ')}.`);
      }
    }
  }

  // Migrate Details sheet
  const detailsSheet = spreadsheet.getSheetByName('Details');
  if (detailsSheet) {
    const lastCol = detailsSheet.getLastColumn();
    if (lastCol === 0) {
      logApiService.info('Details sheet is empty, skipping migration');
    } else {
      let headers = detailsSheet.getRange(1, 1, 1, lastCol).getValues()[0];
      let headersLower = headers.map(normalizeHeaderLabel);

      const refreshDetailsHeaders = () => {
        headers = detailsSheet.getRange(1, 1, 1, detailsSheet.getLastColumn()).getValues()[0];
        headersLower = headers.map(normalizeHeaderLabel);
      };

      const renameDetailsHeader = (index: number, allowed: string[], canonical: string) => {
        if (headersLower.length <= index) return;
        const value = headersLower[index];
        if (allowed.includes(value) && headers[index] !== canonical) {
          setHeaderValue(detailsSheet, index, canonical);
          refreshDetailsHeaders();
        }
      };

      renameDetailsHeader(0, ['label'], 'Name');
      renameDetailsHeader(1, ['helper text', 'helpertext', 'help text', 'helper'], 'Helper Text');
      renameDetailsHeader(2, ['field type'], 'Type');
      renameDetailsHeader(3, ['option', 'options'], 'Options');
      renameDetailsHeader(4, ['field id', 'fieldid'], 'ID');
      renameDetailsHeader(5, ['universal'], 'Universal');

      // Check if already has ID column in correct position (idempotency check)
      const expectedNewHeaders = ['Name', 'Helper Text', 'Type', 'Options', 'ID', 'Universal'];
      if (headers.length >= 6 &&
          headers[0] === expectedNewHeaders[0] &&
          headers[4] === expectedNewHeaders[4] &&
          headers[5] === expectedNewHeaders[5]) {
        logApiService.info('Details sheet already in new format with ID column, skipping migration');
      }
      // Check if it's old format without ID column (4 or 5 columns)
      else if ((headers.length === 4 || headers.length === 5) &&
               headers[0] === 'Name' &&
               headers[1] === 'Helper Text' &&
               headers[2] === 'Type' &&
               headers[3] === 'Options') {
        logApiService.info('Migrating Details sheet to include ID column...');

        // Insert column E for ID (after OPTIONS in column D)
        detailsSheet.insertColumnAfter(4);

        // Preserve Universal column if it exists, otherwise add it
        const hasUniversal = headers.length >= 5 && headers[4];
        const newHeaders = ['Name', 'Helper Text', 'Type', 'Options', 'ID'];
        if (hasUniversal) {
          newHeaders.push(headers[4]);  // Preserve existing header name
        } else {
          newHeaders.push('Universal');
        }

        detailsSheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]).setFontWeight('bold');

        logApiService.info('Details sheet migrated successfully');
      }
      // Ambiguous format - log warning and skip to prevent data corruption
      else {
        const warning = `Details sheet must match [Name, Helper Text, Type, Options, ID, Universal]. Found ${headers.length} columns: ${headers.join(', ')}.`;
        logApiService.warn(warning);
        migrationWarnings.push(warning);
      }

      const detailsLayoutError = validateDetailsHeaderLayout(detailsSheet);
      if (detailsLayoutError) {
        logApiService.warn(detailsLayoutError);
        migrationWarnings.push(detailsLayoutError);
      }
    }
  }

  // Show warnings to user if any were encountered
  if (migrationWarnings.length > 0) {
    const warningMessage = 'Spreadsheet format errors detected during migration:\n\n' +
      migrationWarnings.map((w, i) => `${i + 1}. ${w}`).join('\n\n') +
      '\n\nPlease verify your sheet formats match the expected structure:\n' +
      '• Categories: Name, Icon, Fields, Applies, Category ID\n' +
      '• Details: Name, Helper Text, Type, Options, ID, Universal\n\n' +
      'Generation has been halted until the sheet layouts are corrected.';

    try {
      if (typeof closeProcessingModalDialog === 'function') {
        closeProcessingModalDialog();
      }
    } catch (error) {
      logApiService.warn('Unable to close processing dialog before migration alert:', error);
    }

    ui.alert(
      'Migration Errors',
      warningMessage,
      ui.ButtonSet.OK
    );

    throw new Error('Spreadsheet migration failed: invalid sheet format detected.');
  }

  if (migrationNotices.length > 0) {
    const noticeMessage = 'Your spreadsheet was missing some required columns. We added defaults so this run can finish, but you must review these columns before exporting again:\n\n' +
      migrationNotices.map((w, i) => `${i + 1}. ${w}`).join('\n\n');

    ui.alert(
      'Spreadsheet Updated',
      noticeMessage,
      ui.ButtonSet.OK
    );
  }
}

function normalizeHeaderLabel(value: unknown): string {
  return String(value || '').trim().toLowerCase();
}

function setHeaderValue(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  zeroBasedColumnIndex: number,
  value: string
): void {
  sheet
    .getRange(1, zeroBasedColumnIndex + 1, 1, 1)
    .setValue(value)
    .setFontWeight('bold');
}

function ensurePlainTextColumn(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  zeroBasedColumnIndex: number
): void {
  const column = zeroBasedColumnIndex + 1;
  const totalRows = Math.max(sheet.getMaxRows(), 1);
  const range = sheet.getRange(1, column, totalRows, 1);
  try {
    range.clearDataValidations();
  } catch (error) {
    logApiService.warn('Failed to clear data validations for column', column, error);
  }
  try {
    range.setNumberFormat('@STRING@');
  } catch (error) {
    logApiService.warn('Failed to set plain text format for column', column, error);
  }
}

function ensureAppliesColumn(categoriesSheet: GoogleAppsScript.Spreadsheet.Sheet): boolean {
  const lastCol = categoriesSheet.getLastColumn();
  if (lastCol < 3) {
    throw new Error('Categories sheet must include Name, Icon, and Fields before adding Applies.');
  }

  const totalRows = Math.max(categoriesSheet.getLastRow(), 1);
  const headers = categoriesSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const normalized = headers.map(normalizeHeaderLabel);

  const appliesIndex = normalized.findIndex(header => header.startsWith('appl') || header === 'tracks');
  if (appliesIndex === 3) {
    setHeaderValue(categoriesSheet, 3, 'Applies');
    ensurePlainTextColumn(categoriesSheet, 3);
    return false;
  }

  let existingValues: any[][] | null = null;
  if (appliesIndex !== -1) {
    const appliesColumnPosition = appliesIndex + 1; // 1-based index before insertion
    if (appliesColumnPosition <= 3) {
      throw new Error('Applies column is positioned before Fields. Please reorder columns to Name, Icon, Fields, Applies.');
    }
    existingValues = categoriesSheet.getRange(1, appliesColumnPosition, totalRows, 1).getValues();
  }

  categoriesSheet.insertColumnAfter(3); // New column becomes column 4 (0-based index 3)
  if (existingValues) {
    categoriesSheet.getRange(1, 4, totalRows, 1).setValues(existingValues);
    const deleteIndex = appliesIndex + 2; // Account for inserted column
    categoriesSheet.deleteColumn(deleteIndex);
  }
  setHeaderValue(categoriesSheet, 3, 'Applies');
  ensurePlainTextColumn(categoriesSheet, 3);
  if (!existingValues) {
    try {
      const firstAppliesCell = categoriesSheet.getRange(2, 4);
      if (!String(firstAppliesCell.getValue() || '').trim()) {
        firstAppliesCell.setValue('track, observation');
      }
    } catch (error) {
      logApiService.warn('Unable to seed Applies column default value:', error);
    }
  }
  return true;
}

function validateCategoriesHeaderLayout(
  categoriesSheet: GoogleAppsScript.Spreadsheet.Sheet
): string | null {
  const canonicalHeaders = ['Name', 'Icon', 'Fields', 'Applies', 'Category ID'];
  const normalizedExpected = canonicalHeaders.map(h => h.toLowerCase());
  const headers = categoriesSheet
    .getRange(1, 1, 1, categoriesSheet.getLastColumn())
    .getValues()[0];

  if (headers.length < canonicalHeaders.length) {
    return 'Categories sheet must include columns Name, Icon, Fields, Applies, Category ID in that order.';
  }

  const normalized = headers.map(normalizeHeaderLabel);
  for (let i = 0; i < canonicalHeaders.length; i++) {
    if (normalized[i] !== normalizedExpected[i]) {
      return `Categories sheet column ${i + 1} must be "${canonicalHeaders[i]}" but found "${headers[i] || ''}".`;
    }
  }

  return null;
}

function validateDetailsHeaderLayout(detailsSheet: GoogleAppsScript.Spreadsheet.Sheet): string | null {
  const canonicalHeaders = ['Name', 'Helper Text', 'Type', 'Options', 'ID', 'Universal'];
  const normalizedExpected = canonicalHeaders.map(h => h.toLowerCase());
  const headers = detailsSheet
    .getRange(1, 1, 1, detailsSheet.getLastColumn())
    .getValues()[0];

  if (headers.length < canonicalHeaders.length) {
    return 'Details sheet must include columns Name, Helper Text, Type, Options, ID, Universal in that order.';
  }

  const normalized = headers.map(normalizeHeaderLabel);
  for (let i = 0; i < canonicalHeaders.length; i++) {
    if (normalized[i] !== normalizedExpected[i]) {
      return `Details sheet column ${i + 1} must be "${canonicalHeaders[i]}" but found "${headers[i] || ''}".`;
    }
  }

  return null;
}

function ensureCategoryIdColumn(categoriesSheet: GoogleAppsScript.Spreadsheet.Sheet): boolean {
  const lastCol = categoriesSheet.getLastColumn();
  if (lastCol === 0) return;

  const headers = categoriesSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const normalized = headers.map(normalizeHeaderLabel);
  let idColIndex = normalized.findIndex((header) => header === 'category id' || header === 'id');

  let columnInserted = false;
  if (idColIndex === -1) {
    const appliesIndex = normalized.findIndex((header) => header.startsWith('appl'));
    const insertAfterPosition = appliesIndex >= 0 ? appliesIndex + 1 : headers.length;
    categoriesSheet.insertColumnAfter(insertAfterPosition);
    idColIndex = insertAfterPosition;
    columnInserted = true;
  }

  // Ensure header label is consistently "Category ID"
  categoriesSheet
    .getRange(1, idColIndex + 1, 1, 1)
    .setValue('Category ID')
    .setFontWeight('bold');

  populateMissingCategoryIds(categoriesSheet, idColIndex);
  ensurePlainTextColumn(categoriesSheet, idColIndex);
  return columnInserted;
}

function populateMissingCategoryIds(
  categoriesSheet: GoogleAppsScript.Spreadsheet.Sheet,
  idColIndex: number,
): void {
  const lastRow = categoriesSheet.getLastRow();
  if (lastRow <= 1) return;

  const nameRange = categoriesSheet.getRange(2, CATEGORY_COL.NAME + 1, lastRow - 1, 1);
  const idRange = categoriesSheet.getRange(2, idColIndex + 1, lastRow - 1, 1);
  const names = nameRange.getValues();
  const ids = idRange.getValues();

  let needsWrite = false;
  const updated = ids.map((row, index) => {
    let currentId = String(row[0] || '').trim();
    if (!currentId) {
      const name = String(names[index]?.[0] || '').trim();
      if (!name) {
        return [''];
      }
      currentId = slugify(name) || `category-${index + 1}`;
      needsWrite = true;
    }
    return [currentId];
  });

  if (needsWrite) {
    idRange.setValues(updated);
  }
}
