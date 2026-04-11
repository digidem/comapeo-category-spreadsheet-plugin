/**
 * Generates a CoMapeo configuration file from the spreadsheet data.
 * v2.0.0 - Uses JSON-only API, no ZIP workflow
 */
const logGenerateConfig = getScopedLogger("GenerateConfig");

function generateCoMapeoConfig() {
  let processingDialogOpen = false;
  const closeProcessingDialogIfOpen = () => {
    if (!processingDialogOpen) return;
    processingDialogOpen = false;
    if (typeof closeProcessingModalDialog === "function") {
      try {
        closeProcessingModalDialog();
      } catch (closeError) {
        logGenerateConfig.warn("Unable to close processing dialog:", closeError);
      }
    }
  };

  try {
    // Step 1: Initialize
    showProcessingModalDialog(processingDialogTexts[0][locale]);
    processingDialogOpen = true;
    logGenerateConfig.info("Generating CoMapeo config v2.0.0...");

    // Step 1.5: Migrate old spreadsheet format if needed
    migrateSpreadsheetFormat();

    // Step 1.6: Check for translation sheet mismatches BEFORE linting
    logGenerateConfig.info("Checking translation sheet consistency...");
    const mismatchResult = detectTranslationMismatches();

    if (mismatchResult && mismatchResult.hasMismatches) {
      closeProcessingDialogIfOpen();

      // Build detailed error message
      let errorDetails = "Translation sheet mismatches detected:\n\n";
      mismatchResult.details.forEach((detail) => {
        errorDetails += `${detail.sheetName}:\n`;
        detail.mismatches.forEach((m) => {
          errorDetails += `  Row ${m.row}: "${m.sourceValue}" → "${m.translationValue}"\n`;
        });
        errorDetails += "\n";
      });

      // Show dialog with fix/abort options
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert(
        "Translation Sheet Mismatch Detected",
        errorDetails +
          "Primary language values in translation sheets don't match their source sheets.\n" +
          "This will cause translation lookup failures during config generation.\n\n" +
          "Do you want to automatically fix these issues?\n" +
          "• YES: Re-sync formulas and re-translate to configured languages\n" +
          "• NO: Highlight issues in red and abort generation",
        ui.ButtonSet.YES_NO,
      );

      if (response === ui.Button.YES) {
        // Fix and continue
        const fixingDialogText = {
          en: {
            title: "Fixing Translation Mismatches",
            message: [
              "Re-syncing formulas and re-translating...",
              "This may take a moment",
            ],
          },
          es: {
            title: "Corrigiendo Desajustes de Traducción",
            message: [
              "Re-sincronizando fórmulas y volviendo a traducir...",
              "Esto puede tardar un momento",
            ],
          },
        };
        showProcessingModalDialog(fixingDialogText[locale]);
        processingDialogOpen = true;
        logGenerateConfig.info("Fixing translation mismatches...");
        fixTranslationMismatches(true, mismatchResult); // Pass pre-detected mismatch data
        logGenerateConfig.info("Translation fixes applied, continuing with generation...");
      } else {
        // Abort - run linting to highlight issues in red
        logGenerateConfig.info(
          "User chose to abort. Running linting to highlight issues...",
        );
        lintAllSheets(false); // This will highlight mismatches in bright red
        throw new Error(
          "Config generation aborted due to translation sheet mismatches. " +
            "Mismatched cells have been highlighted in bright red. " +
            "Please review and fix manually, or run generation again to auto-fix.",
        );
      }
    }

    // Step 2: Lint (passing false to prevent UI alerts)
    logGenerateConfig.info("Linting CoMapeo config...");
    showProcessingModalDialog(processingDialogTexts[2][locale]);
    processingDialogOpen = true;
    lintAllSheets(false);

    // Step 3: Get data
    const data = getSpreadsheetData();
    showProcessingModalDialog(processingDialogTexts[3][locale]);
    processingDialogOpen = true;

    // Step 4: Create build payload (JSON)
    showProcessingModalDialog(processingDialogTexts[4][locale]);
    processingDialogOpen = true;
    logGenerateConfig.info("Creating build payload...");
    const buildRequest = createBuildPayload(data);

    // Log category selection for verification
    logGenerateConfig.info("Category selection order:", getCategorySelection());

    // Step 5: Send to API (JSON mode)
    showProcessingModalDialog(processingDialogTexts[5][locale]);
    processingDialogOpen = true;
    logGenerateConfig.info("Sending JSON request to API...");
    const configUrl = sendBuildRequest(buildRequest);

    // Step 6: Show success dialog
    showProcessingModalDialog(processingDialogTexts[7][locale]);
    processingDialogOpen = true;
    closeProcessingDialogIfOpen();
    showConfigurationGeneratedDialog(configUrl);
  } catch (error) {
    logGenerateConfig.error("Error generating CoMapeo config:", error);
    closeProcessingDialogIfOpen();

    const ui = SpreadsheetApp.getUi();
    ui.alert(
      "Error Generating CoMapeo Category",
      "An error occurred while generating the CoMapeo category: " +
        error.message +
        "\n\nPlease try again. If the problem persists, contact support.",
      ui.ButtonSet.OK,
    );
  }
}

/**
 * Debug helper retained for developer menu.
 * Legacy behavior wrote artifacts to Drive; v2 flow is JSON-only, so this
 * delegates to the standard generator while keeping the menu entry alive.
 */
function generateCoMapeoConfigWithDriveWrites(): void {
  logGenerateConfig.info(
    "Debug build (Drive writes disabled in v2) - running standard generator.",
  );
  generateCoMapeoConfig();
}
