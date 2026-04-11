// Note: getVersionInfo, VERSION, COMMIT are defined in src/version.ts
// Apps Script will compile all .ts files together, making them globally available

let activeUserLocale = Session.getActiveUserLocale().split("_")[0];
const supportedLocales = ["en", "es"];
const defaultLocale = "en";
let locale = supportedLocales.includes(activeUserLocale)
  ? activeUserLocale
  : defaultLocale;

const logMenu = getScopedLogger("Menu");

function onOpen() {
  // Log version info to Apps Script console
  // VERSION and COMMIT constants are defined in src/version.ts
  if (typeof VERSION !== 'undefined' && typeof COMMIT !== 'undefined') {
    logMenu.info(`CoMapeo Config Spreadsheet Plugin v${VERSION} (${COMMIT})`);
    if (typeof getVersionInfo !== 'undefined') {
      logMenu.info(`Full version: ${getVersionInfo()}`);
    }
  }
  const ui = SpreadsheetApp.getUi();
  const mainMenu = ui.createMenu(menuTexts[locale].menu)
    .addItem(
      menuTexts[locale].translateCoMapeoCategory,
      "translateCoMapeoCategory",
    )
    .addItem(menuTexts[locale].generateIcons, "generateIcons")
    .addSeparator()
    .addItem(menuTexts[locale].generateCoMapeoCategory, "generateCoMapeoCategory")
    .addItem(menuTexts[locale].importCoMapeoCategory, "importCoMapeoCategory")
    .addSeparator()
    .addItem(menuTexts[locale].lintAllSheets, "lintAllSheets")
    .addItem(menuTexts[locale].cleanAllSheets, "cleanAllSheets")
    .addSeparator();

  const debugMenu = ui.createMenu(menuTexts[locale].debugMenuTitle)
    .addItem("Create Test Spreadsheet for Regression", "createTestSpreadsheetForRegression")
    .addItem("Test Runner", "runAllTests")
    .addItem("Capture Baseline Performance Metrics", "captureAndDocumentBaselineMetrics")
    .addItem("Turn on legacy compatibility", "toggleLegacyCompatibility")
    .addItem(
      menuTexts[locale].generateCoMapeoCategoryDebug,
      "generateCoMapeoCategoryDebug",
    );

  mainMenu
    .addSubMenu(debugMenu)
    .addSeparator()
    .addItem(menuTexts[locale].openHelpPage, "openHelpPage")
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
      translateMenuTexts[locale].error,
      translateMenuTexts[locale].errorText + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK,
    );
  }
}

function translateToSelectedLanguages(selectedLanguages: string[]) {
  const ui = SpreadsheetApp.getUi();
  try {
    autoTranslateSheetsBidirectional(selectedLanguages as TranslationLanguage[]);
    ui.alert(
      translateMenuTexts[locale].completed,
      translateMenuTexts[locale].completedText,
      ui.ButtonSet.OK,
    );
  } catch (error) {
    ui.alert(
      translateMenuTexts[locale].error,
      translateMenuTexts[locale].errorText + error.message,
      ui.ButtonSet.OK,
    );
  }
}

// This function is defined in generateCoMapeoConfig.ts and will be available globally

function generateIcons() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    iconMenuTexts[locale].action,
    iconMenuTexts[locale].actionText,
    ui.ButtonSet.YES_NO,
  );

  if (result === ui.Button.YES) {
    try {
      generateIconsConfig();
    } catch (error) {
      ui.alert(
        iconMenuTexts[locale].error,
        iconMenuTexts[locale].errorText + error.message,
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
    categoryMenuTexts[locale].action,
    categoryMenuTexts[locale].actionText,
    ui.ButtonSet.YES_NO,
  );

  if (result === ui.Button.YES) {
    try {
      generateCoMapeoConfig();
    } catch (error) {
      ui.alert(
        categoryMenuTexts[locale].error,
        categoryMenuTexts[locale].errorText + error.message,
        ui.ButtonSet.OK,
      );
    }
  }
}

function generateCoMapeoCategoryDebug() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    categoryDebugMenuTexts[locale].action,
    categoryDebugMenuTexts[locale].actionText,
    ui.ButtonSet.YES_NO,
  );

  if (result === ui.Button.YES) {
    try {
      generateCoMapeoConfigWithDriveWrites();
    } catch (error) {
      ui.alert(
        categoryDebugMenuTexts[locale].error,
        categoryDebugMenuTexts[locale].errorText + error.message,
        ui.ButtonSet.OK,
      );
    }
  }
}

function importCoMapeoCategory() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    importMenuTexts[locale].action,
    importMenuTexts[locale].actionText,
    ui.ButtonSet.YES_NO,
  );

  if (result === ui.Button.YES) {
    try {
      importCoMapeoCatFile();
    } catch (error) {
      ui.alert(
        importMenuTexts[locale].error,
        importMenuTexts[locale].errorText + error.message,
        ui.ButtonSet.OK,
      );
    }
  }
}

function lintCoMapeoCategory() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    lintMenuTexts[locale].action,
    lintMenuTexts[locale].actionText,
    ui.ButtonSet.YES_NO,
  );

  if (result === ui.Button.YES) {
    try {
      lintAllSheets();
      ui.alert(
        lintMenuTexts[locale].completed,
        lintMenuTexts[locale].completedText,
        ui.ButtonSet.OK,
      );
    } catch (error) {
      ui.alert(
        lintMenuTexts[locale].error,
        lintMenuTexts[locale].errorText + error.message,
        ui.ButtonSet.OK,
      );
    }
  }
}

function cleanAllSheets() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    cleanAllMenuTexts[locale].action,
    cleanAllMenuTexts[locale].actionText,
    ui.ButtonSet.YES_NO,
  );

  if (result === ui.Button.YES) {
    try {
      removeTranslationAndMetadataSheets();
      ui.alert(
        cleanAllMenuTexts[locale].completed,
        cleanAllMenuTexts[locale].completedText,
        ui.ButtonSet.OK,
      );
    } catch (error) {
      ui.alert(
        cleanAllMenuTexts[locale].error,
        cleanAllMenuTexts[locale].errorText + error.message,
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
