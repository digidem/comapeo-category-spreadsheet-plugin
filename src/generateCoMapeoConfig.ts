/**
 * Generates a CoMapeo configuration file from the spreadsheet data.
 * v2.0.0 - Uses JSON-only API, no ZIP workflow
 *
 * Progress reporting (#57): the long-running generation records its stages
 * in the job status store (src/jobStatus.ts) and a sidebar polls that store,
 * replacing the old close-and-reopen processing modal. The modal helpers in
 * src/dialog.ts remain for operations not yet migrated.
 */

function generateCoMapeoConfig() {
  const jobId = startJobStatus("generate", {
    title: "Generate CoMapeo Category",
    detail: "Reading spreadsheet and preparing generation",
  });
  openJobSidebar(jobId);

  /**
   * Reuses the localized processing dialog texts as stage labels so the
   * sidebar shows the same wording the modal used to.
   */
  const reportStage = (stageIndex: number, percent: number, extraDetail?: string) => {
    const stageTexts = processingDialogTexts[stageIndex]
      ? processingDialogTexts[stageIndex][locale]
      : null;
    const phase = stageTexts && stageTexts.title ? stageTexts.title : "Working";
    const firstMessage =
      stageTexts && Array.isArray(stageTexts.message) && stageTexts.message.length > 0
        ? stageTexts.message[0]
        : "";
    updateJobStatusProgress(jobId, {
      percent,
      phase,
      detail: extraDetail || firstMessage,
      logMessage: firstMessage || phase,
    });
  };

  try {
    // Step 1: Initialize
    reportStage(0, 5);
    getScopedLogger("GenerateConfig").info("Generating CoMapeo config v2.0.0...");

    // Step 1.5: Migrate old spreadsheet format if needed
    migrateSpreadsheetFormat();
    updateJobStatusProgress(jobId, { percent: 10, phase: "Format check" });

    // Step 1.6: Check for translation sheet mismatches BEFORE linting
    getScopedLogger("GenerateConfig").info("Checking translation sheet consistency...");
    const mismatchResult = detectTranslationMismatches();

    if (mismatchResult && mismatchResult.hasMismatches) {
      updateJobStatusProgress(jobId, {
        percent: 15,
        phase: "Waiting for your decision",
        detail: "Translation sheet mismatches detected — a dialog is asking how to proceed",
        logMessage: `Translation mismatches detected in ${mismatchResult.details.length} sheet(s)`,
        logLevel: "warn",
      });

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
        updateJobStatusProgress(jobId, {
          percent: 20,
          phase: "Fixing Translation Mismatches",
          detail: "Re-syncing formulas and re-translating",
          logMessage: "Fixing translation mismatches",
        });
        getScopedLogger("GenerateConfig").info("Fixing translation mismatches...");
        fixTranslationMismatches(true, mismatchResult); // Pass pre-detected mismatch data
        getScopedLogger("GenerateConfig").info("Translation fixes applied, continuing with generation...");
      } else {
        // Abort - run linting to highlight issues in red
        getScopedLogger("GenerateConfig").info(
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
    getScopedLogger("GenerateConfig").info("Linting CoMapeo config...");
    reportStage(2, 40);
    lintAllSheets(false);

    // Step 3: Get data
    const data = getSpreadsheetData();
    reportStage(3, 50);

    // Step 4: Create build payload (JSON)
    reportStage(4, 60);
    getScopedLogger("GenerateConfig").info("Creating build payload...");
    const buildRequest = createBuildPayload(data);

    // Log category selection for verification
    getScopedLogger("GenerateConfig").info("Category selection order:", getCategorySelection());

    // Step 5: Send to API (JSON mode)
    reportStage(5, 70);
    getScopedLogger("GenerateConfig").info("Sending JSON request to API...");
    const configUrl = sendBuildRequest(buildRequest);
    updateJobStatusProgress(jobId, {
      percent: 90,
      phase: "Packaging response",
      detail: "Build API returned successfully",
      logMessage: "Received .comapeocat from build API",
    });

    // Step 6: Show success dialog
    completeJobStatus(jobId, "Configuration generated successfully");
    showConfigurationGeneratedDialog(configUrl);
  } catch (error) {
    getScopedLogger("GenerateConfig").error("Error generating CoMapeo config:", error);
    failJobStatus(
      jobId,
      error instanceof Error ? error.message : String(error),
    );

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
  getScopedLogger("GenerateConfig").info(
    "Debug build (Drive writes disabled in v2) - running standard generator.",
  );
  generateCoMapeoConfig();
}
