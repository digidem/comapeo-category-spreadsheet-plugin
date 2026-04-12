/**
 * Configuration application functions for the import category functionality.
 * This file contains functions related to applying the configuration to the spreadsheet.
 *
 * Uses the createOrClearSheet function from utils.ts
 * Calls applyMetadata, applyCategories, applyFields, and applyTranslations functions
 */


/**
 * Applies configuration data to the spreadsheet.
 * @param configData - Configuration data object
 * @param onProgress - Optional progress callback function
 */
function applyConfigurationToSpreadsheet(
  configData: any,
  onProgress?: (update: { percent: number; stage: string; detail?: string }) => void,
) {
  getScopedLogger("ApplyConfiguration").info("Applying configuration to spreadsheet...");

  if (onProgress) {
    onProgress({
      percent: 70,
      stage: "Updating spreadsheet",
      detail: "Preparing sheets...",
    });
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Clear validations only on managed sheets now that parsing succeeded
  try {
    if (typeof clearManagedSheetValidations === "function") {
      clearManagedSheetValidations(spreadsheet);
    }
  } catch (error) {
    getScopedLogger("ApplyConfiguration").error("Error clearing managed validations:", error);
  }

  // Create or clear necessary sheets
  const categoriesSheet = createOrClearSheet(spreadsheet, "Categories");
  const detailsSheet = createOrClearSheet(spreadsheet, "Details");
  const metadataSheet = createOrClearSheet(spreadsheet, "Metadata");

  // Apply metadata
  if (configData.metadata) {
    if (onProgress) {
      onProgress({
        percent: 75,
        stage: "Updating spreadsheet",
        detail: "Applying metadata...",
      });
    }
    applyMetadata(metadataSheet, configData.metadata);
  }
  getScopedLogger("ApplyConfiguration").debug("configData keys:", Object.keys(configData).join(", "));
  // Apply categories (presets)
  if (configData.presets && configData.presets.length > 0) {
    if (onProgress) {
      onProgress({
        percent: 80,
        stage: "Updating spreadsheet",
        detail: "Applying categories...",
      });
    }
    applyCategories(
      categoriesSheet,
      configData.presets,
      configData.fields,
      configData.icons,
    );
  }

  // Apply details (fields)
  if (configData.fields && configData.fields.length > 0) {
    if (onProgress) {
      onProgress({
        percent: 85,
        stage: "Updating spreadsheet",
        detail: "Applying field definitions...",
      });
    }
    applyFields(detailsSheet, configData.fields);
  }

  // Apply translations
  if (configData.messages && Object.keys(configData.messages).length > 0) {
    if (onProgress) {
      onProgress({
        percent: 90,
        stage: "Updating spreadsheet",
        detail: "Applying translations...",
      });
    }
    applyTranslations(
      spreadsheet,
      configData.messages,
      configData.presets,
      configData.fields,
    );
  }

  getScopedLogger("ApplyConfiguration").info("Configuration applied to spreadsheet successfully");

  // Add dropdowns after all data has been imported with a delay
  try {
    // Wait a moment to ensure all data is properly set
    Utilities.sleep(1000);

    // Only call if the function exists
    if (typeof addAllDropdowns === "function") {
      addAllDropdowns();
    }
  } catch (error) {
    getScopedLogger("ApplyConfiguration").info(
      "Note: Dropdowns not added. This is optional and won't affect the import.",
    );
  }
}
