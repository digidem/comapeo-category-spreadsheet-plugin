/**
 * Test function for importing category files
 * This function tests the extraction, parsing, and application of category files
 */

/**
 * Test the category import functionality
 * @param url - URL of the test file to download (optional)
 */
function testImportCategory(url?: string) {
  // const ui = SpreadsheetApp.getUi();

  try {
    // Show a progress dialog
    const htmlOutput = HtmlService.createHtmlOutput(
      '<div style="text-align: center; padding: 20px;">' +
        "<h3>Testing Category Import</h3>" +
        '<div id="progress-container" style="width: 100%; background-color: #f1f1f1; border-radius: 5px; margin: 20px 0;">' +
        '<div id="progress-bar" style="width: 0%; height: 30px; background-color: #4CAF50; border-radius: 5px; text-align: center; line-height: 30px; color: white;">0%</div>' +
        "</div>" +
        '<div id="status">Initializing test...</div>' +
        '<div id="log" style="text-align: left; margin-top: 20px; height: 200px; overflow-y: auto; border: 1px solid #ccc; padding: 10px; font-family: monospace; font-size: 12px;"></div>' +
        "</div>" +
        "<script>" +
        "function updateProgress(percent, status, logMessage, logType) {" +
        '  document.getElementById("progress-bar").style.width = percent + "%";' +
        '  document.getElementById("progress-bar").innerHTML = percent + "%";' +
        '  if (status) document.getElementById("status").innerHTML = status;' +
        "  if (logMessage) {" +
        '    const log = document.getElementById("log");' +
        '    const logEntry = document.createElement("div");' +
        '    if (logType) logEntry.style.color = logType === "error" ? "#F44336" : logType === "warning" ? "#FF9800" : logType === "success" ? "#4CAF50" : "#2196F3";' +
        "    logEntry.textContent = logMessage;" +
        "    log.appendChild(logEntry);" +
        "    log.scrollTop = log.scrollHeight;" +
        "  }" +
        "}" +
        "</script>",
    )
      .setWidth(600)
      .setHeight(400);

    // const dialog = ui.showModelessDialog(htmlOutput, "Testing Category Import");

    // Helper function to update progress
    const updateProgress = (
      percent: number,
      status: string,
      logMessage?: string,
      logType?: string,
    ) => {
      const script = `updateProgress(${percent}, "${status}", ${logMessage ? `"${logMessage.replace(/"/g, '\\"')}"` : "null"}, ${logType ? `"${logType}"` : "null"});`;
      try {
        htmlOutput.append(`<script>${script}</script>`);
      } catch (e) {
        console.error("Error updating progress:", e);
      }
    };

    // Log function
    const log = (
      message: string,
      type?: "info" | "success" | "warning" | "error",
    ) => {
      console.log(message);
      updateProgress(null, null, message, type);
    };

    // Start the test
    updateProgress(5, "Starting test...");
    log("Starting category import test", "info");

    // Use a default URL if none provided
    if (!url) {
      url = "https://luandro.com/dist/mapeo-default-min.mapeosettings";
      log(`Using default test file: ${url}`, "info");
    }

    // Download the test file
    updateProgress(10, "Downloading test file...");
    log(`Downloading file from: ${url}`, "info");

    let fileBlob: GoogleAppsScript.Base.Blob;
    try {
      const response = UrlFetchApp.fetch(url);
      fileBlob = response.getBlob();
      log(
        `Downloaded file: ${fileBlob.getName()} (${fileBlob.getBytes().length} bytes)`,
        "success",
      );
    } catch (error) {
      log(`Error downloading file: ${error}`, "error");
      throw new Error(`Failed to download file: ${error}`);
    }

    // Create a temporary folder
    updateProgress(20, "Creating temporary folder...");
    const tempFolder = DriveApp.createFolder(
      "Category_Import_Test_" + new Date().getTime(),
    );
    log(`Created temporary folder: ${tempFolder.getName()}`, "info");

    // Extract the file
    updateProgress(30, "Extracting file...");
    let extractionResult;

    try {
      extractionResult = extractAndValidateFile(fileBlob.getName(), fileBlob);

      if (!extractionResult.success) {
        log(`Extraction failed: ${extractionResult.message}`, "error");
        if (
          extractionResult.validationErrors &&
          extractionResult.validationErrors.length > 0
        ) {
          extractionResult.validationErrors.forEach((error) => {
            log(`Validation error: ${error}`, "error");
          });
        }
        throw new Error("Extraction failed");
      }

      log(
        `Extraction successful: ${extractionResult.files.length} files extracted`,
        "success",
      );

      // Log warnings if any
      if (
        extractionResult.validationWarnings &&
        extractionResult.validationWarnings.length > 0
      ) {
        extractionResult.validationWarnings.forEach((warning) => {
          log(`Validation warning: ${warning}`, "warning");
        });
      }
    } catch (error) {
      log(`Error during extraction: ${error}`, "error");
      throw new Error(`Failed to extract file: ${error}`);
    }

    // Extract configuration data
    updateProgress(50, "Extracting configuration data...");
    let configData;

    try {
      configData = extractConfigurationData(
        extractionResult.files,
        extractionResult.tempFolder,
        {
          onProgress: (stage: string, percent: number) => {
            const overallPercent = 50 + Math.round(percent * 0.2); // Map 0-100 to 50-70
            updateProgress(
              overallPercent,
              `Extracting: ${stage} (${percent}%)...`,
            );
          },
        },
      );

      log("Configuration data extracted:", "success");
      log(
        `- Metadata: ${configData.metadata ? "Present" : "Not present"}`,
        "info",
      );
      log(`- Presets: ${configData.presets.length}`, "info");
      log(`- Fields: ${configData.fields.length}`, "info");
      log(`- Icons: ${configData.icons.length}`, "info");
      log(`- Languages: ${Object.keys(configData.messages).length}`, "info");
    } catch (error) {
      log(`Error extracting configuration data: ${error}`, "error");
      throw new Error(`Failed to extract configuration data: ${error}`);
    }

    // Create a backup of the current sheets
    updateProgress(70, "Creating backup of current sheets...");
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const backupSheets: Record<string, unknown[][]> = {};

    // Backup Categories sheet
    const categoriesSheet = spreadsheet.getSheetByName("Categories");
    if (categoriesSheet) {
      log("Backing up Categories sheet...", "info");
      const categoriesData = categoriesSheet.getDataRange().getValues();
      backupSheets["Categories"] = categoriesData;
    }

    // Backup Details sheet
    const detailsSheet = spreadsheet.getSheetByName("Details");
    if (detailsSheet) {
      log("Backing up Details sheet...", "info");
      const detailsData = detailsSheet.getDataRange().getValues();
      backupSheets["Details"] = detailsData;
    }

    // Apply configuration to spreadsheet
    updateProgress(80, "Applying configuration to spreadsheet...");
    try {
      applyConfigurationToSpreadsheet(configData);
      log("Configuration applied to spreadsheet successfully", "success");

      // Verify Categories sheet
      const newCategoriesSheet = spreadsheet.getSheetByName("Categories");
      if (newCategoriesSheet) {
        const categoriesCount = Math.max(
          0,
          newCategoriesSheet.getLastRow() - 1,
        ); // Subtract header row
        log(`Categories sheet has ${categoriesCount} categories`, "info");

        // Check if icons are present
        let iconsCount = 0;
        if (categoriesCount > 0) {
          const iconRange = newCategoriesSheet.getRange(
            2,
            2,
            categoriesCount,
            1,
          );
          const iconValues = iconRange.getValues();
          iconsCount = iconValues.filter(
            (row) => row[0] && row[0].toString().trim() !== "",
          ).length;
          log(
            `Categories sheet has ${iconsCount} icons`,
            iconsCount > 0 ? "success" : "warning",
          );
        }
      } else {
        log("Categories sheet not found after import", "error");
      }

      // Verify Details sheet
      const newDetailsSheet = spreadsheet.getSheetByName("Details");
      if (newDetailsSheet) {
        const detailsCount = Math.max(0, newDetailsSheet.getLastRow() - 1); // Subtract header row
        log(
          `Details sheet has ${detailsCount} fields`,
          detailsCount > 0 ? "success" : "warning",
        );
      } else {
        log("Details sheet not found after import", "error");
      }
    } catch (error) {
      log(`Error applying configuration to spreadsheet: ${error}`, "error");

      // Restore backup
      log("Restoring backup...", "info");

      // Restore Categories sheet
      if (backupSheets["Categories"] && categoriesSheet) {
        categoriesSheet.clear();
        categoriesSheet
          .getRange(
            1,
            1,
            backupSheets["Categories"].length,
            backupSheets["Categories"][0].length,
          )
          .setValues(backupSheets["Categories"]);
        log("Restored Categories sheet", "success");
      }

      // Restore Details sheet
      if (backupSheets["Details"] && detailsSheet) {
        detailsSheet.clear();
        detailsSheet
          .getRange(
            1,
            1,
            backupSheets["Details"].length,
            backupSheets["Details"][0].length,
          )
          .setValues(backupSheets["Details"]);
        log("Restored Details sheet", "success");
      }

      throw new Error(`Failed to apply configuration to spreadsheet: ${error}`);
    }

    // Clean up
    updateProgress(95, "Cleaning up...");
    try {
      extractionResult.tempFolder.setTrashed(true);
      log("Temporary folder deleted", "info");
    } catch (error) {
      log(`Error cleaning up temporary folder: ${error}`, "warning");
    }

    // Complete
    updateProgress(100, "Test completed successfully");
    log("Category import test completed successfully", "success");

    return {
      success: true,
      message: "Category import test completed successfully",
    };
  } catch (error) {
    console.error("Error in category import test:", error);
    return {
      success: false,
      message: `Error in category import test: ${error instanceof Error ? error.message : String(error)}`,
    };
  }
}
