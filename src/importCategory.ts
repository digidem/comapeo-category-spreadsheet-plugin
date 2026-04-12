// Main entry point for the import category functionality

/**
 * Extracts configuration data from files.
 * This is a compatibility function for tests and other modules.
 * It simply calls parseExtractedFiles internally.
 *
 * @param files - Array of extracted file blobs
 * @param tempFolder - Temporary folder containing the extracted files
 * @param options - Optional configuration options
 * @returns Configuration data object
 */
function extractConfigurationData(
  files: GoogleAppsScript.Base.Blob[] | undefined,
  tempFolder?: GoogleAppsScript.Drive.Folder,
  options?: any,
): any {
  getScopedLogger("ImportCategory").info(
    `Extracting configuration data from ${files ? files.length : "undefined"} files`,
  );
  if (options) {
    getScopedLogger("ImportCategory").info(`Using options: ${JSON.stringify(options)}`);
  }

  if (!files || !Array.isArray(files) || files.length === 0) {
    throw new Error("No extracted files available to parse");
  }

  return parseExtractedFiles(files, tempFolder);
}

/**
 * Shows the import category dialog.
 * This is the main entry point for the import category functionality.
 */
function showImportCategoryDialog() {
  // Get dialog texts from the importCategory/dialogTexts.ts file
  const title = importCategoryDialogTexts[locale].title;

  // Create HTML output using the function from importCategory/ui.ts
  const htmlOutput = HtmlService.createHtmlOutput(createImportCategoryHtml())
    .setWidth(800)
    .setHeight(600);

  // Show the dialog
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, title);
}

/**
 * Main function to process an imported category file.
 * This function is called from the UI when a file is selected.
 *
 * @param fileName - The name of the imported file
 * @param base64Data - The file content as base64 string
 * @returns Success message if import was successful
 */
function processImportedCategoryFile(
  fileName: string,
  base64Data: string,
): { success: boolean; message: string; details?: any } {
  try {
    getScopedLogger("ImportCategory").info(`Starting import of file: ${fileName}`);

    // Decode the base64 data
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data),
      "application/octet-stream",
      fileName,
    );
    getScopedLogger("ImportCategory").info(`Decoded file size: ${blob.getBytes().length} bytes`);

    // Extract the archive using the shared extractor
    const extractionResult = extractAndValidateFile(fileName, blob);
    if (!extractionResult.success || !extractionResult.files) {
      return {
        success: false,
        message: extractionResult.message || "Failed to extract file",
      };
    }

    // Parse the extracted files using the function from importCategory/parseFiles.ts
    const configData = parseExtractedFiles(
      extractionResult.files,
      extractionResult.tempFolder,
    );

    // Apply the configuration to the spreadsheet using the function from importCategory/applyConfiguration.ts
    applyConfigurationToSpreadsheet(configData);

    // Clean up temporary resources
    if (extractionResult.tempFolder) {
      try {
        extractionResult.tempFolder.setTrashed(true);
      } catch (error) {
        getScopedLogger("ImportCategory").warn("Error cleaning up temporary folder:", error);
      }
    }

    return {
      success: true,
      message: "Category file imported successfully",
      details: {
        presets: configData.presets.length,
        fields: configData.fields.length,
        icons: configData.icons.length,
        languages: Object.keys(configData.messages).length,
      },
      warnings: extractionResult.validationWarnings || [],
    };
  } catch (error) {
    getScopedLogger("ImportCategory").error("Error processing imported file:", error);
    return {
      success: false,
      message:
        "Error processing imported file: " +
        (error instanceof Error ? error.message : String(error)),
    };
  }
}

/**
 * Compatibility helper for progress-based Mapeo imports.
 * Reuses the standard parser to keep behavior in sync.
 */
function extractMapeoConfigurationData(
  files: GoogleAppsScript.Base.Blob[] | undefined,
  tempFolder: GoogleAppsScript.Drive.Folder,
  options?: any,
): any {
  return extractConfigurationData(files, tempFolder, options);
}

/**
 * Applies parsed configuration for classic Mapeo flows.
 */
function applyMapeoConfigurationToSpreadsheet(configData: any) {
  applyConfigurationToSpreadsheet(configData);
}
