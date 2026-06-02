// Note: getVersionInfo, VERSION, COMMIT are defined in src/version.ts
// Apps Script will compile all .ts files together, making them globally available

let activeUserLocale = Session.getActiveUserLocale().split("_")[0];
const supportedLocales = ["en", "es"];
const defaultLocale = "en";
let activeLocale = supportedLocales.includes(activeUserLocale)
  ? activeUserLocale
  : defaultLocale;
// Compatibility alias: downstream files (dialog.ts, generateCoMapeoConfig.ts,
// translation.ts) reference `locale` as a global. GAS shares a single global
// scope across all .ts files, so both names must resolve to the same value.
let locale = activeLocale;


function onOpen() {
  // Log version info to Apps Script console
  // VERSION and COMMIT constants are defined in src/version.ts
  if (typeof VERSION !== 'undefined' && typeof COMMIT !== 'undefined') {
    getScopedLogger("Menu").info(`CoMapeo Config Spreadsheet Plugin v${VERSION} (${COMMIT})`);
    if (typeof getVersionInfo !== 'undefined') {
      getScopedLogger("Menu").info(`Full version: ${getVersionInfo()}`);
    }
  }
  const ui = SpreadsheetApp.getUi();
  const mainMenu = ui.createMenu(menuTexts[activeLocale].menu)
    .addItem(
      menuTexts[activeLocale].translateCoMapeoCategory,
      "translateCoMapeoCategory",
    )
    .addItem(menuTexts[activeLocale].generateIcons, "generateIcons")
    .addSeparator()
    .addItem(menuTexts[activeLocale].generateCoMapeoCategory, "generateCoMapeoCategory")
    .addItem(menuTexts[activeLocale].importCoMapeoCategory, "importCoMapeoCategory")
    .addSeparator()
    .addItem(menuTexts[activeLocale].lintAllSheets, "lintAllSheets")
    .addItem(menuTexts[activeLocale].cleanAllSheets, "cleanAllSheets")
    .addSeparator();

  const debugMenu = ui.createMenu(menuTexts[activeLocale].debugMenuTitle)
    .addItem("Create Test Spreadsheet for Regression", "createTestSpreadsheetForRegression")
    .addItem("Test Runner", "runAllTests")
    .addItem("Capture Baseline Performance Metrics", "captureAndDocumentBaselineMetrics")
    .addItem("Turn on legacy compatibility", "toggleLegacyCompatibility")
    .addItem(
      menuTexts[activeLocale].generateCoMapeoCategoryDebug,
      "generateCoMapeoCategoryDebug",
    );

  mainMenu
    .addSubMenu(debugMenu)
    .addSeparator()
    .addItem(menuTexts[activeLocale].openHelpPage, "openHelpPage")
    .addItem("About / Version", "showVersionInfo")
    .addToUi();

  // Add developer menu in development environment
  if (
    PropertiesService.getScriptProperties().getProperty("ENVIRONMENT") ===
    "development"
  ) {
    ui.createMenu("Developer")
      .addItem("Test Format Detection", "testFormatDetection")
      .addItem("Test Translation Extraction", "testTranslationExtraction")
      .addItem("Test Category Import", "testImportCategory")
      .addItem("Test Details and Icons", "testDetailsAndIcons")
      .addItem("Test Field Extraction", "testFieldExtraction")
      .addSeparator()
      .addItem("Run All Tests", "runAllTests")
      .addItem("Capture Baseline Metrics", "captureAndDocumentBaselineMetrics")
      .addItem("Generate Performance Report", "generatePerformanceReport")
      .addSeparator()
      .addItem("Clear Language Cache", "clearLanguagesCacheMenuItem")
      .addToUi();
  }
}

function translateCoMapeoCategory() {
  try {
    showSelectTranslationLanguagesDialog();
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      translateMenuTexts[activeLocale].error,
      translateMenuTexts[activeLocale].errorText + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
  }
}

function translateToSelectedLanguages(selectedLanguages: string[]) {
  const ui = SpreadsheetApp.getUi();
  try {
    autoTranslateSheetsBidirectional(selectedLanguages as TranslationLanguage[]);
    ui.alert(
      translateMenuTexts[activeLocale].completed,
      translateMenuTexts[activeLocale].completedText,
      ui.ButtonSet.OK,
    );
  } catch (error) {
    ui.alert(
      translateMenuTexts[activeLocale].error,
      translateMenuTexts[activeLocale].errorText + error.message,
      ui.ButtonSet.OK,
    );
  }
}

// This function is defined in generateCoMapeoConfig.ts and will be available globally

function generateIcons() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    iconMenuTexts[activeLocale].action,
    iconMenuTexts[activeLocale].actionText,
    ui.ButtonSet.YES_NO,
  );

  if (result === ui.Button.YES) {
    try {
      generateIconsConfig();
    } catch (error) {
      ui.alert(
        iconMenuTexts[activeLocale].error,
        iconMenuTexts[activeLocale].errorText + error.message,
        ui.ButtonSet.OK,
      );
    }
  }
}

function showVersionInfo() {
  const ui = SpreadsheetApp.getUi();
  // getVersionInfo is defined in src/version.ts
  const versionInfo = typeof getVersionInfo !== 'undefined'
    ? getVersionInfo()
    : (typeof VERSION !== 'undefined' ? VERSION : 'Unknown');
  ui.alert(
    "CoMapeo Config Spreadsheet Plugin",
    `Version: ${versionInfo}\n\nRepository: https://github.com/digidem/comapeo-config-spreadsheet-plugin`,
    ui.ButtonSet.OK
  );
}

function generateCoMapeoCategory() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    categoryMenuTexts[activeLocale].action,
    categoryMenuTexts[activeLocale].actionText,
    ui.ButtonSet.YES_NO,
  );

  if (result === ui.Button.YES) {
    try {
      generateCoMapeoConfig();
    } catch (error) {
      ui.alert(
        categoryMenuTexts[activeLocale].error,
        categoryMenuTexts[activeLocale].errorText + error.message,
        ui.ButtonSet.OK,
      );
    }
  }
}

function generateCoMapeoCategoryDebug() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    categoryDebugMenuTexts[activeLocale].action,
    categoryDebugMenuTexts[activeLocale].actionText,
    ui.ButtonSet.YES_NO,
  );

  if (result === ui.Button.YES) {
    try {
      generateCoMapeoConfigWithDriveWrites();
    } catch (error) {
      ui.alert(
        categoryDebugMenuTexts[activeLocale].error,
        categoryDebugMenuTexts[activeLocale].errorText + error.message,
        ui.ButtonSet.OK,
      );
    }
  }
}

function importCoMapeoCategory() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    importMenuTexts[activeLocale].action,
    importMenuTexts[activeLocale].actionText,
    ui.ButtonSet.YES_NO,
  );

  if (result === ui.Button.YES) {
    try {
      importCoMapeoCatFile();
    } catch (error) {
      ui.alert(
        importMenuTexts[activeLocale].error,
        importMenuTexts[activeLocale].errorText + error.message,
        ui.ButtonSet.OK,
      );
    }
  }
}

function lintCoMapeoCategory() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    lintMenuTexts[activeLocale].action,
    lintMenuTexts[activeLocale].actionText,
    ui.ButtonSet.YES_NO,
  );

  if (result === ui.Button.YES) {
    try {
      lintAllSheets();
      ui.alert(
        lintMenuTexts[activeLocale].completed,
        lintMenuTexts[activeLocale].completedText,
        ui.ButtonSet.OK,
      );
    } catch (error) {
      ui.alert(
        lintMenuTexts[activeLocale].error,
        lintMenuTexts[activeLocale].errorText + error.message,
        ui.ButtonSet.OK,
      );
    }
  }
}

function cleanAllSheets() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    cleanAllMenuTexts[activeLocale].action,
    cleanAllMenuTexts[activeLocale].actionText,
    ui.ButtonSet.YES_NO,
  );

  if (result === ui.Button.YES) {
    try {
      removeTranslationAndMetadataSheets();
      ui.alert(
        cleanAllMenuTexts[activeLocale].completed,
        cleanAllMenuTexts[activeLocale].completedText,
        ui.ButtonSet.OK,
      );
    } catch (error) {
      ui.alert(
        cleanAllMenuTexts[activeLocale].error,
        cleanAllMenuTexts[activeLocale].errorText + error.message,
        ui.ButtonSet.OK,
      );
    }
  }
}

function openHelpPage() {
  showHelpDialog();
}

/**
 * Menu item handler for clearing language cache
 * Useful for debugging or forcing fresh language data fetch
 */
function clearLanguagesCacheMenuItem() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    "Clear Language Cache",
    "This will clear the cached language data, forcing a fresh fetch from the remote source on next use. This is useful for debugging. Continue?",
    ui.ButtonSet.YES_NO,
  );

  if (result === ui.Button.YES) {
    try {
      clearLanguagesCache();
      ui.alert(
        "Cache Cleared",
        "Language cache has been successfully cleared. Next language operation will fetch fresh data.",
        ui.ButtonSet.OK,
      );
    } catch (error) {
      ui.alert(
        "Error",
        "An error occurred while clearing the cache: " + error.message,
        ui.ButtonSet.OK,
      );
    }
  }
}

/**
 * Toggle legacy compatibility flag in metadata sheet
 * Adds or updates legacyCompat key with TRUE/FALSE value
 */
function toggleLegacyCompatibility() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let metadataSheet = spreadsheet.getSheetByName('Metadata');

  if (!metadataSheet) {
    // Create Metadata sheet if it doesn't exist
    metadataSheet = spreadsheet.insertSheet('Metadata');
    metadataSheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]).setFontWeight('bold');
  }

  const data = metadataSheet.getDataRange().getValues();
  let currentValue = 'FALSE';
  let rowIndex = -1;

  // Find existing legacyCompat key
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === 'legacyCompat') {
      currentValue = String(data[i][1]).trim().toUpperCase();
      rowIndex = i;
      break;
    }
  }

  // Toggle value
  const newValue = currentValue === 'TRUE' ? 'FALSE' : 'TRUE';

  if (rowIndex === -1) {
    // Append new row
    metadataSheet.appendRow(['legacyCompat', newValue]);
  } else {
    // Update existing row
    metadataSheet.getRange(rowIndex + 1, 2).setValue(newValue);
  }

  // Show confirmation
  ui.alert(
    'Legacy Compatibility',
    `Legacy compatibility has been turned ${newValue === 'TRUE' ? 'ON' : 'OFF'}.`,
    ui.ButtonSet.OK
  );
}
