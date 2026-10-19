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
  const locale = typeof activeLocale !== "undefined" ? activeLocale : "en";
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
): { success: boolean; message: string; details?: any; warnings?: string[] } {
  return processImportedCategoryFileCore(fileName, base64Data, "");
}

/**
 * Sidebar-driven entry point (#57): starts a job record, opens the progress
 * sidebar, then runs the import. The sidebar polls getJobStatus(jobId) while
 * this google.script.run execution works through the phases. Called from the
 * import dialog client instead of processImportedCategoryFile.
 */
function startImportCategoryJob(
  fileName: string,
  base64Data: string,
): { success: boolean; message: string; details?: any; warnings?: string[] } {
  const jobId = startJobStatus("import", {
    title: "Import Category File",
    detail: String(fileName || ""),
  });
  openJobSidebar(jobId);
  const result = processImportedCategoryFileCore(fileName, base64Data, jobId);
  if (result && result.success) {
    const details = result.details || {};
    const summary =
      `${result.message} (${details.presets || 0} categories, ` +
      `${details.fields || 0} fields, ${details.icons || 0} icons)`;
    completeJobStatus(jobId, summary);
  } else {
    failJobStatus(jobId, (result && result.message) || "Import failed");
  }
  return result;
}

/**
 * Import pipeline shared by the direct entry point and the sidebar job.
 * @param jobId - When non-empty, phase progress is written to the job store
 *   for the polling sidebar; empty for callers that bypass the sidebar.
 */
function processImportedCategoryFileCore(
  fileName: string,
  base64Data: string,
  jobId: string,
): { success: boolean; message: string; details?: any; warnings?: string[] } {
  try {
    getScopedLogger("ImportCategory").info(`Starting import of file: ${fileName}`);
    updateJobStatusProgress(jobId, {
      percent: 5,
      phase: "Decoding file",
      detail: String(fileName || ""),
      logMessage: `Import started: ${fileName}`,
    });

    // Decode the base64 data
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data),
      "application/octet-stream",
      fileName,
    );
    getScopedLogger("ImportCategory").info(`Decoded file size: ${blob.getBytes().length} bytes`);
    updateJobStatusProgress(jobId, {
      percent: 15,
      phase: "Extracting archive",
      detail: `${blob.getBytes().length} bytes`,
    });

    // Extract the archive using the shared extractor
    const extractionResult = extractAndValidateFile(fileName, blob);
    if (!extractionResult.success || !extractionResult.files) {
      return {
        success: false,
        message: extractionResult.message || "Failed to extract file",
      };
    }
    updateJobStatusProgress(jobId, {
      percent: 40,
      phase: "Parsing configuration",
      detail: `${extractionResult.files.length} file(s) extracted`,
      logMessage: `Extracted ${extractionResult.files.length} file(s) from archive`,
    });

    // Parse the extracted files using the function from importCategory/parseFiles.ts
    const configData = parseExtractedFiles(
      extractionResult.files,
      extractionResult.tempFolder,
    );
    updateJobStatusProgress(jobId, {
      percent: 60,
      phase: "Applying to spreadsheet",
      detail: `${configData.presets.length} categories, ${configData.fields.length} fields`,
      logMessage: `Parsed ${configData.presets.length} categories and ${configData.fields.length} fields`,
    });

    // Apply the configuration to the spreadsheet using the function from importCategory/applyConfiguration.ts
    applyConfigurationToSpreadsheet(configData);

    // Clean up temporary resources
    updateJobStatusProgress(jobId, {
      percent: 90,
      phase: "Cleaning up",
      detail: "Removing temporary extraction files",
    });
    if (extractionResult.tempFolder) {
      try {
        extractionResult.tempFolder.setTrashed(true);
      } catch (error) {
        getScopedLogger("ImportCategory").warn("Error cleaning up temporary folder:", error);
      }
    }

    // Collect all warnings: extraction + config validation
    const allWarnings = [
      ...(extractionResult.validationWarnings || []),
      ...(configData._validationWarnings || []),
    ];
    if (allWarnings.length > 0) {
      updateJobStatusProgress(jobId, {
        logMessage: `${allWarnings.length} validation warning(s) recorded`,
        logLevel: "warn",
      });
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
      warnings: allWarnings,
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

/**
 * Applies raw JSON imports after normalizing them into the spreadsheet format.
 *
 * Handles imports that aren't in the standard `.comapeocat` archive format.
 * First normalizes the raw JSON through `normalizeConfig()` to ensure the
 * expected schema, then delegates to `applyConfigurationToSpreadsheet()` to
 * populate the spreadsheet sheets.
 *
 * @param jsonData - Raw JSON configuration object (unvalidated). Expected to
 *   contain preset, field, and translation data in any CoMapeo-compatible schema.
 */
function applyMapeoJsonConfigToSpreadsheet(jsonData: any) {
  applyConfigurationToSpreadsheet(normalizeConfig(jsonData));
}
