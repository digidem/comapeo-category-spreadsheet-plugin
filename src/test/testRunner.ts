/// <reference path="../loggingHelpers.ts" />
/// <reference path="../../node_modules/@types/google-apps-script/index.d.ts" />

/**
 * Master Test Runner
 *
 * Executes all test suites in sequence and provides a comprehensive report.
 * This is the single entry point for running all tests at once.
 *
 * Usage:
 * - Run all tests: runAllTests()
 * - Run quick smoke tests: runAllTestsQuick()
 * - Run specific test category: runTestSuite('language')
 */

// Global flag to indicate tests are running in batch mode
const BATCH_TEST_MODE = true;

/**
 * Generate HTML for the test runner progress dialog
 * Used with the close-and-reopen pattern for progress updates
 */
function generateTestRunnerDialogHtml(
  percent: number,
  status: string,
  logMessages: string[],
  totalSuites: number,
  currentSuite: number
): string {
  const progressBarHtml = `
    <div style="text-align: center; padding: 20px; font-family: Arial, sans-serif;">
      <h2>Running Test Suite</h2>
      <div style="font-size: 14px; color: #666; margin-bottom: 10px;">Test ${currentSuite}/${totalSuites}</div>
      <div id="progress-container" style="width: 100%; background-color: #f1f1f1; border-radius: 5px; margin: 20px 0;">
        <div id="progress-bar" style="width: ${percent}%; height: 40px; background-color: #4CAF50; border-radius: 5px; text-align: center; line-height: 40px; color: white; font-weight: bold;">${percent}%</div>
      </div>
      <div id="status" style="font-size: 16px; margin: 20px 0; color: #333;">${status}</div>
      <div id="log" style="text-align: left; margin-top: 20px; height: 300px; overflow-y: auto; border: 1px solid #ccc; padding: 10px; font-family: monospace; font-size: 12px; background-color: #fafafa;">
        ${logMessages.map(msg => `<div style="margin-bottom: 4px;">${msg}</div>`).join('')}
      </div>
    </div>
  `;

  return progressBarHtml;
}

/**
 * Update test runner progress dialog - close and reopen pattern
 * This is the Apps Script way to update dialogs (modal dialogs can't be modified after showing)
 */
function updateTestRunnerProgress(
  percent: number,
  status: string,
  logMessage: string,
  totalSuites: number,
  currentSuite: number
): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) return;

  // Check if we're updating an existing dialog
  const scriptProperties = PropertiesService.getScriptProperties();
  const isUpdate = scriptProperties.getProperty('testRunnerDialogActive') === 'true';

  // Close any existing dialog first to prevent stacking
  if (isUpdate) {
    try {
      google.script.host.close();
      Utilities.sleep(200); // Give time for dialog to close
    } catch (e) {
      // Dialog might already be closed, ignore error
    }
  }

  // Collect recent log messages (keep last 50 to prevent dialog from getting too large)
  const recentLogs: string[] = [];
  const timestamp = new Date().toLocaleTimeString();
  recentLogs.push(`<div style="color: #2196F3;">${timestamp} - ${logMessage}</div>`);

  const html = generateTestRunnerDialogHtml(percent, status, recentLogs, totalSuites, currentSuite) +
    '<script>google.script.run._testRunnerCallback();</script>';

  // Close and reopen dialog with updated content
  // This is the Apps Script pattern for "updating" dialogs
  let ui: GoogleAppsScript.Spreadsheet.SpreadsheetUi;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    // Only catch the specific "from this context" error
    // Other errors should be allowed to propagate
    const errorMessage = String(e);
    if (errorMessage.includes('Cannot call SpreadsheetApp.getUi() from this context')) {
      console.log(`Progress: ${percent}% - ${status} - ${logMessage}`);
      return;
    }
    throw e; // Re-throw if it's a different error
  }

  ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(700).setHeight(500), "Test Runner Progress");
  scriptProperties.setProperty('testRunnerDialogActive', 'true');
}

/**
 * Callback function to mark that dialog has been closed
 * This is called from the HTML to indicate the dialog is no longer active
 */
function _testRunnerCallback(): void {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty('testRunnerDialogActive');
}

/**
 * Close any active test runner dialog
 */
function closeTestRunnerDialog(): void {
  try {
    google.script.host.close();
  } catch (e) {
    // Dialog might not be open, ignore error
  }
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty('testRunnerDialogActive');
}

function runAllTests(): void {
  const log = getScopedLogger("TestRunner");
  const totalSuites = 14;

  if (typeof Logger === "undefined") {
    throw new Error("Logger not available - tests must run in Apps Script environment");
  }

  // Clean up any existing dialog state
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty('testRunnerDialogActive');

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Initialize the progress dialog
  if (spreadsheet) {
    updateTestRunnerProgress(
      0,
      "Starting test suite...",
      "Initializing test runner...",
      totalSuites,
      0,
    );
  }

  log.info("========================================");
  log.info("Starting Comprehensive Test Suite");
  log.info("========================================");

  const startTime = Date.now();
  const results: {
    suiteName: string;
    passed: number;
    failed: number;
    duration: number;
    status: "PASSED" | "FAILED";
  }[] = [];

  // Helper function to update progress using close-and-reopen pattern
  const updateProgress = (current: number, total: number, suiteName: string) => {
    const percentage = Math.round((current / total) * 100);
    const message = `Running test ${current}/${total}: ${suiteName}`;
    log.info(`Progress: ${percentage}% - ${message}`);

    // Update dialog using close-and-reopen pattern
    if (spreadsheet) {
      updateTestRunnerProgress(percentage, message, message, total, current);
    }

    // Also show toast
    if (spreadsheet) {
      spreadsheet.toast(message, "Test Progress", 2);
    }

    // Add small delay to prevent rapid dialog updates
    Utilities.sleep(300);
  };

  let currentSuite = 0;

  // Test Suite 1: Language Recognition Tests
  currentSuite++;
  updateProgress(currentSuite, totalSuites, "Language Recognition");
  try {
    log.info("\n--- Test Suite 1: Language Recognition ---");
    const suiteStart = Date.now();
    testLanguageLookup();
    const duration = Date.now() - suiteStart;
    results.push({
      suiteName: "Language Recognition",
      passed: 1,
      failed: 0,
      duration,
      status: "PASSED",
    });
    log.info(`✓ Language Recognition tests passed (${duration}ms)`);
    if (spreadsheet) {
      spreadsheet.toast("✓ Language Recognition passed", "Test Results", 2);
    }
  } catch (error) {
    const duration = 0; // Duration already captured before test execution
    results.push({
      suiteName: "Language Recognition",
      passed: 0,
      failed: 1,
      duration,
      status: "FAILED",
    });
    log.error(`✗ Language Recognition tests failed: ${error.message}`);
    if (spreadsheet) {
      spreadsheet.toast("✗ Language Recognition FAILED", "Test Results", 3);
    }
  }

  // Test Suite 2: Language Recognition Integration Tests
  currentSuite++;
  updateProgress(currentSuite, totalSuites, "Language Recognition Integration");
  try {
    log.info("\n--- Test Suite 2: Language Recognition Integration ---");
    const suiteStart = Date.now();
    testLanguageRecognitionIntegration();
    const duration = Date.now() - suiteStart;
    results.push({
      suiteName: "Language Recognition Integration",
      passed: 1,
      failed: 0,
      duration,
      status: "PASSED",
    });
    log.info(`✓ Language Recognition Integration tests passed (${duration}ms)`);
    if (spreadsheet) {
      spreadsheet.toast("✓ Language Integration passed", "Test Results", 2);
    }
  } catch (error) {
    const duration = 0; // Duration already captured before test execution
    results.push({
      suiteName: "Language Recognition Integration",
      passed: 0,
      failed: 1,
      duration,
      status: "FAILED",
    });
    log.error(`✗ Language Recognition Integration tests failed: ${error.message}`);
    if (spreadsheet) {
      spreadsheet.toast("✗ Language Integration FAILED", "Test Results", 3);
    }
  }

  // Test Suite 3: Utils Slugify Tests
  currentSuite++;
  updateProgress(currentSuite, totalSuites, "Utils Slugify");
  try {
    log.info("\n--- Test Suite 3: Utils Slugify Functions ---");
    const suiteStart = Date.now();
    testUtilsSlugify();
    const duration = Date.now() - suiteStart;
    results.push({
      suiteName: "Utils Slugify Functions",
      passed: 1,
      failed: 0,
      duration,
      status: "PASSED",
    });
    log.info(`✓ Utils Slugify tests passed (${duration}ms)`);
    if (spreadsheet) {
      spreadsheet.toast("✓ Utils Slugify passed", "Test Results", 2);
    }
  } catch (error) {
    const duration = 0; // Duration already captured before test execution
    results.push({
      suiteName: "Utils Slugify Functions",
      passed: 0,
      failed: 1,
      duration,
      status: "FAILED",
    });
    log.error(`✗ Utils Slugify tests failed: ${error.message}`);
    if (spreadsheet) {
      spreadsheet.toast("✗ Utils Slugify FAILED", "Test Results", 3);
    }
  }

  // Test Suite 4: Format Detection Tests
  currentSuite++;
  updateProgress(currentSuite, totalSuites, "Format Detection");
  try {
    log.info("\n--- Test Suite 4: Format Detection ---");
    const suiteStart = Date.now();
    testFormatDetection();
    const duration = Date.now() - suiteStart;
    results.push({
      suiteName: "Format Detection",
      passed: 1,
      failed: 0,
      duration,
      status: "PASSED",
    });
    log.info(`✓ Format Detection tests passed (${duration}ms)`);
    if (spreadsheet) {
      spreadsheet.toast("✓ Format Detection passed", "Test Results", 2);
    }
  } catch (error) {
    const duration = 0; // Duration already captured before test execution
    results.push({
      suiteName: "Format Detection",
      passed: 0,
      failed: 1,
      duration,
      status: "FAILED",
    });
    log.error(`✗ Format Detection tests failed: ${error.message}`);
    if (spreadsheet) {
      spreadsheet.toast("✗ Format Detection FAILED", "Test Results", 3);
    }
  }

  // Test Suite 5: Field Extraction Tests
  currentSuite++;
  updateProgress(currentSuite, totalSuites, "Field Extraction");
  try {
    log.info("\n--- Test Suite 5: Field Extraction ---");
    const suiteStart = Date.now();
    testFieldExtraction();
    const duration = Date.now() - suiteStart;
    results.push({
      suiteName: "Field Extraction",
      passed: 1,
      failed: 0,
      duration,
      status: "PASSED",
    });
    log.info(`✓ Field Extraction tests passed (${duration}ms)`);
    if (spreadsheet) {
      spreadsheet.toast("✓ Field Extraction passed", "Test Results", 2);
    }
  } catch (error) {
    const duration = 0; // Duration already captured before test execution
    results.push({
      suiteName: "Field Extraction",
      passed: 0,
      failed: 1,
      duration,
      status: "FAILED",
    });
    log.error(`✗ Field Extraction tests failed: ${error.message}`);
    if (spreadsheet) {
      spreadsheet.toast("✗ Field Extraction FAILED", "Test Results", 3);
    }
  }

  // Test Suite 6: Extract and Validate Tests
  currentSuite++;
  updateProgress(currentSuite, totalSuites, "Extract and Validate");
  try {
    log.info("\n--- Test Suite 6: Extract and Validate ---");
    const suiteStart = Date.now();
    testExtractAndValidate();
    const duration = Date.now() - suiteStart;
    results.push({
      suiteName: "Extract and Validate",
      passed: 1,
      failed: 0,
      duration,
      status: "PASSED",
    });
    log.info(`✓ Extract and Validate tests passed (${duration}ms)`);
    if (spreadsheet) {
      spreadsheet.toast("✓ Extract and Validate passed", "Test Results", 2);
    }
  } catch (error) {
    const duration = 0; // Duration already captured before test execution
    results.push({
      suiteName: "Extract and Validate",
      passed: 0,
      failed: 1,
      duration,
      status: "FAILED",
    });
    log.error(`✗ Extract and Validate tests failed: ${error.message}`);
    if (spreadsheet) {
      spreadsheet.toast("✗ Extract and Validate FAILED", "Test Results", 3);
    }
  }

  // Test Suite 7: Details and Icons Tests
  currentSuite++;
  updateProgress(currentSuite, totalSuites, "Details and Icons");
  try {
    log.info("\n--- Test Suite 7: Details and Icons ---");
    const suiteStart = Date.now();
    testDetailsAndIcons();
    const duration = Date.now() - suiteStart;
    results.push({
      suiteName: "Details and Icons",
      passed: 1,
      failed: 0,
      duration,
      status: "PASSED",
    });
    log.info(`✓ Details and Icons tests passed (${duration}ms)`);
    if (spreadsheet) {
      spreadsheet.toast("✓ Details and Icons passed", "Test Results", 2);
    }
  } catch (error) {
    const duration = 0; // Duration already captured before test execution
    results.push({
      suiteName: "Details and Icons",
      passed: 0,
      failed: 1,
      duration,
      status: "FAILED",
    });
    log.error(`✗ Details and Icons tests failed: ${error.message}`);
    if (spreadsheet) {
      spreadsheet.toast("✗ Details and Icons FAILED", "Test Results", 3);
    }
  }

  // Test Suite 8: Translation Extraction Tests
  currentSuite++;
  updateProgress(currentSuite, totalSuites, "Translation Extraction");
  try {
    log.info("\n--- Test Suite 8: Translation Extraction ---");
    const suiteStart = Date.now();
    testTranslationExtraction();
    const duration = Date.now() - suiteStart;
    results.push({
      suiteName: "Translation Extraction",
      passed: 1,
      failed: 0,
      duration,
      status: "PASSED",
    });
    log.info(`✓ Translation Extraction tests passed (${duration}ms)`);
    if (spreadsheet) {
      spreadsheet.toast("✓ Translation Extraction passed", "Test Results", 2);
    }
  } catch (error) {
    const duration = 0; // Duration already captured before test execution
    results.push({
      suiteName: "Translation Extraction",
      passed: 0,
      failed: 1,
      duration,
      status: "FAILED",
    });
    log.error(`✗ Translation Extraction tests failed: ${error.message}`);
    if (spreadsheet) {
      spreadsheet.toast("✗ Translation Extraction FAILED", "Test Results", 3);
    }
  }

  // Test Suite 9: Import Category Tests
  currentSuite++;
  updateProgress(currentSuite, totalSuites, "Import Category");
  try {
    log.info("\n--- Test Suite 9: Import Category ---");
    const suiteStart = Date.now();
    testImportCategory();
    const duration = Date.now() - suiteStart;
    results.push({
      suiteName: "Import Category",
      passed: 1,
      failed: 0,
      duration,
      status: "PASSED",
    });
    log.info(`✓ Import Category tests passed (${duration}ms)`);
    if (spreadsheet) {
      spreadsheet.toast("✓ Import Category passed", "Test Results", 2);
    }
  } catch (error) {
    const duration = 0; // Duration already captured before test execution
    results.push({
      suiteName: "Import Category",
      passed: 0,
      failed: 1,
      duration,
      status: "FAILED",
    });
    log.error(`✗ Import Category tests failed: ${error.message}`);
    if (spreadsheet) {
      spreadsheet.toast("✗ Import Category FAILED", "Test Results", 3);
    }
  }

  // Test Suite 10: Zip to API Tests
  currentSuite++;
  updateProgress(currentSuite, totalSuites, "Zip to API");
  try {
    log.info("\n--- Test Suite 10: Zip to API ---");
    const suiteStart = Date.now();
    testZipToApi();
    const duration = Date.now() - suiteStart;
    results.push({
      suiteName: "Zip to API",
      passed: 1,
      failed: 0,
      duration,
      status: "PASSED",
    });
    log.info(`✓ Zip to API tests passed (${duration}ms)`);
    if (spreadsheet) {
      spreadsheet.toast("✓ Zip to API passed", "Test Results", 2);
    }
  } catch (error) {
    const duration = 0; // Duration already captured before test execution
    results.push({
      suiteName: "Zip to API",
      passed: 0,
      failed: 1,
      duration,
      status: "FAILED",
    });
    log.error(`✗ Zip to API tests failed: ${error.message}`);
    if (spreadsheet) {
      spreadsheet.toast("✗ Zip to API FAILED", "Test Results", 3);
    }
  }

  // Test Suite 11: End-to-End Tests
  currentSuite++;
  updateProgress(currentSuite, totalSuites, "End-to-End");
  try {
    log.info("\n--- Test Suite 11: End-to-End ---");
    const suiteStart = Date.now();
    testEndToEnd();
    const duration = Date.now() - suiteStart;
    results.push({
      suiteName: "End-to-End",
      passed: 1,
      failed: 0,
      duration,
      status: "PASSED",
    });
    log.info(`✓ End-to-End tests passed (${duration}ms)`);
    if (spreadsheet) {
      spreadsheet.toast("✓ End-to-End passed", "Test Results", 2);
    }
  } catch (error) {
    const duration = 0; // Duration already captured before test execution
    results.push({
      suiteName: "End-to-End",
      passed: 0,
      failed: 1,
      duration,
      status: "FAILED",
    });
    log.error(`✗ End-to-End tests failed: ${error.message}`);
    if (spreadsheet) {
      spreadsheet.toast("✗ End-to-End FAILED", "Test Results", 3);
    }
  }

  // Test Suite 12: Skip Translation Tests
  currentSuite++;
  updateProgress(currentSuite, totalSuites, "Skip Translation");
  try {
    log.info("\n--- Test Suite 12: Skip Translation ---");
    const suiteStart = Date.now();
    testSkipTranslation();
    const duration = Date.now() - suiteStart;
    results.push({
      suiteName: "Skip Translation",
      passed: 1,
      failed: 0,
      duration,
      status: "PASSED",
    });
    log.info(`✓ Skip Translation tests passed (${duration}ms)`);
    if (spreadsheet) {
      spreadsheet.toast("✓ Skip Translation passed", "Test Results", 2);
    }
  } catch (error) {
    const duration = 0; // Duration already captured before test execution
    results.push({
      suiteName: "Skip Translation",
      passed: 0,
      failed: 1,
      duration,
      status: "FAILED",
    });
    log.error(`✗ Skip Translation tests failed: ${error.message}`);
    if (spreadsheet) {
      spreadsheet.toast("✗ Skip Translation FAILED", "Test Results", 3);
    }
  }

  // Test Suite 13: Debug Logger Tests
  currentSuite++;
  updateProgress(currentSuite, totalSuites, "Debug Logger");
  try {
    log.info("\n--- Test Suite 13: Debug Logger ---");
    const suiteStart = Date.now();
    testDebugLogger();
    const duration = Date.now() - suiteStart;
    results.push({
      suiteName: "Debug Logger",
      passed: 1,
      failed: 0,
      duration,
      status: "PASSED",
    });
    log.info(`✓ Debug Logger tests passed (${duration}ms)`);
    if (spreadsheet) {
      spreadsheet.toast("✓ Debug Logger passed", "Test Results", 2);
    }
  } catch (error) {
    const duration = 0; // Duration already captured before test execution
    results.push({
      suiteName: "Debug Logger",
      passed: 0,
      failed: 1,
      duration,
      status: "FAILED",
    });
    log.error(`✗ Debug Logger tests failed: ${error.message}`);
    if (spreadsheet) {
      spreadsheet.toast("✗ Debug Logger FAILED", "Test Results", 3);
    }
  }

  // Test Suite 14: Lint Parity Tests
  currentSuite++;
  updateProgress(currentSuite, totalSuites, "Lint Parity");
  try {
    log.info("\n--- Test Suite 14: Lint Parity ---");
    const suiteStart = Date.now();
    runLintParityTests();
    const duration = Date.now() - suiteStart;
    results.push({
      suiteName: "Lint Parity",
      passed: 1,
      failed: 0,
      duration,
      status: "PASSED",
    });
    log.info(`✓ Lint parity tests passed (${duration}ms)`);
    if (spreadsheet) {
      spreadsheet.toast("✓ Lint Parity passed", "Test Results", 2);
    }
  } catch (error) {
    const duration = 0;
    results.push({
      suiteName: "Lint Parity",
      passed: 0,
      failed: 1,
      duration,
      status: "FAILED",
    });
    log.error(`✗ Lint parity tests failed: ${error.message}`);
    if (spreadsheet) {
      spreadsheet.toast("✗ Lint Parity FAILED", "Test Results", 3);
    }
  }

  const totalDuration = Date.now() - startTime;

  // Generate Summary Report
  log.info("\n========================================");
  log.info("Test Suite Summary");
  log.info("========================================");

  const passedSuites = results.filter(r => r.status === "PASSED").length;
  const failedSuites = results.filter(r => r.status === "FAILED").length;

  log.info(`Total Test Suites: ${results.length}`);
  log.info(`Passed: ${passedSuites}`);
  log.info(`Failed: ${failedSuites}`);
  log.info(`Total Duration: ${totalDuration}ms`);

  if (failedSuites > 0) {
    log.info("\nFailed Test Suites:");
    results
      .filter(r => r.status === "FAILED")
      .forEach(r => {
        log.info(`  ✗ ${r.suiteName} (${r.duration}ms)`);
      });
  }

  log.info("========================================");

  // Close progress dialog and show final completion dialog with summary
  if (spreadsheet) {
    // Close the progress dialog first
    closeTestRunnerDialog();
    Utilities.sleep(300); // Give time for dialog to close

    // Show completion summary
    let ui: GoogleAppsScript.Spreadsheet.SpreadsheetUi;
    try {
      ui = SpreadsheetApp.getUi();
    } catch (e) {
      const errorMessage = String(e);
      if (errorMessage.includes('Cannot call SpreadsheetApp.getUi() from this context')) {
        log.info(`Test completed: ${passedSuites} passed, ${failedSuites} failed in ${(totalDuration / 1000).toFixed(2)}s`);
        return;
      }
      throw e; // Re-throw if it's a different error
    }

    const title = failedSuites > 0 ? "Test Suite Completed with Failures" : "All Tests Passed Successfully! 🎉";
    const summaryHtml = `
      <div style="text-align: center; padding: 20px; font-family: Arial, sans-serif;">
        <h2 style="color: ${failedSuites > 0 ? '#F44336' : '#4CAF50'}">${title}</h2>
        <div style="background-color: #f5f5f5; padding: 20px; border-radius: 8px; margin: 20px 0; text-align: left;">
          <div style="font-size: 14px; margin-bottom: 10px;"><strong>Total Suites:</strong> ${results.length}</div>
          <div style="font-size: 14px; margin-bottom: 10px; color: #4CAF50;"><strong>✅ Passed:</strong> ${passedSuites}</div>
          <div style="font-size: 14px; margin-bottom: 10px; color: ${failedSuites > 0 ? '#F44336' : '#4CAF50'};"><strong>❌ Failed:</strong> ${failedSuites}</div>
          <div style="font-size: 14px; margin-bottom: 10px;"><strong>Duration:</strong> ${(totalDuration / 1000).toFixed(2)}s</div>
        </div>
        ${failedSuites > 0 ? '<div style="color: #F44336; font-weight: bold;">Check logs for details on failed suites.</div>' : '<div style="color: #4CAF50; font-weight: bold;">All tests completed successfully!</div>'}
      </div>
    `;
    ui.showModalDialog(HtmlService.createHtmlOutput(summaryHtml).setWidth(500).setHeight(400), "Test Runner Complete");
  }

  // Throw error if any tests failed
  if (failedSuites > 0) {
    throw new Error(`${failedSuites} test suite(s) failed. Check logs for details.`);
  }
}

/**
 * Quick Test Runner - Runs only the most critical tests
 * Much faster than full test suite
 */
function runAllTestsQuick(): void {
  const log = getScopedLogger("TestRunner");

  if (typeof Logger === "undefined") {
    throw new Error("Logger not available - tests must run in Apps Script environment");
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (spreadsheet) {
    spreadsheet.toast("Starting quick tests...", "Test Runner", 3);
  }

  log.info("========================================");
  log.info("Starting Quick Test Suite");
  log.info("========================================");

  const startTime = Date.now();

  // Quick Test 1: Language Recognition
  if (spreadsheet) {
    spreadsheet.toast("Test 1/3: Language Recognition", "Progress", 3);
  }
  try {
    log.info("\n--- Quick Test 1: Language Recognition ---");
    testLanguageLookupQuick();
    log.info("✓ Language Recognition quick test passed");
    if (spreadsheet) {
      spreadsheet.toast("✓ Language Recognition passed", "Test Results", 2);
    }
  } catch (error) {
    log.error(`✗ Language Recognition quick test failed: ${error.message}`);
    throw error;
  }

  // Quick Test 2: Utils Slugify
  if (spreadsheet) {
    spreadsheet.toast("Test 2/3: Utils Slugify", "Progress", 3);
  }
  try {
    log.info("\n--- Quick Test 2: Utils Slugify ---");
    testUtilsSlugifyQuick();
    log.info("✓ Utils Slugify quick test passed");
    if (spreadsheet) {
      spreadsheet.toast("✓ Utils Slugify passed", "Test Results", 2);
    }
  } catch (error) {
    log.error(`✗ Utils Slugify quick test failed: ${error.message}`);
    throw error;
  }

  // Quick Test 3: Format Detection
  if (spreadsheet) {
    spreadsheet.toast("Test 3/3: Format Detection", "Progress", 3);
  }
  try {
    log.info("\n--- Quick Test 3: Format Detection ---");
    if (typeof testFormatDetection === "function") {
      testFormatDetection();
      log.info("✓ Format Detection quick test passed");
      if (spreadsheet) {
        spreadsheet.toast("✓ Format Detection passed", "Test Results", 2);
      }
    } else {
      log.info("⚠ Format Detection test not available, skipping");
    }
  } catch (error) {
    log.error(`✗ Format Detection quick test failed: ${error.message}`);
    throw error;
  }

  const totalDuration = Date.now() - startTime;

  log.info("\n========================================");
  log.info("Quick Test Suite Complete");
  log.info("========================================");
  log.info(`Duration: ${totalDuration}ms`);
  log.info("========================================");

  // Show summary in UI
  if (spreadsheet) {
    let ui: GoogleAppsScript.Spreadsheet.SpreadsheetUi;
    try {
      ui = SpreadsheetApp.getUi();
    } catch (e) {
      const errorMessage = String(e);
      if (errorMessage.includes('Cannot call SpreadsheetApp.getUi() from this context')) {
        log.info(`Quick tests completed in ${(totalDuration / 1000).toFixed(2)}s`);
        return;
      }
      throw e; // Re-throw if it's a different error
    }

    const message =
      `Quick Tests Complete!\n\n` +
      `All critical tests passed\n` +
      `Duration: ${(totalDuration / 1000).toFixed(2)}s`;

    ui.alert("Quick Tests Complete", message, ui.ButtonSet.OK);
  }
}

/**
 * Run a specific test suite by name
 *
 * @param suiteName - Name of the test suite to run
 */
function runTestSuite(suiteName: string): void {
  const log = getScopedLogger("TestRunner");

  const suites: Record<string, () => void> = {
    language: testLanguageLookup,
    "language-integration": testLanguageRecognitionIntegration,
    utils: testUtilsSlugify,
    "utils-quick": testUtilsSlugifyQuick,
    format: testFormatDetection,
    fields: testFieldExtraction,
    extract: testExtractAndValidate,
    details: testDetailsAndIcons,
    translation: testTranslationExtraction,
    import: testImportCategory,
    zip: testZipToApi,
    endtoend: testEndToEnd,
    skip: testSkipTranslation,
    logger: testDebugLogger,
    lint: runLintParityTests,
  };

  const suiteNameLower = suiteName.toLowerCase();

  if (!(suiteNameLower in suites)) {
    const availableSuites = Object.keys(suites).join(", ");
    throw new Error(
      `Test suite "${suiteName}" not found.\nAvailable suites: ${availableSuites}`,
    );
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (spreadsheet) {
    spreadsheet.toast(`Running: ${suiteName}`, "Test Suite", 3);
  }

  log.info(`Running test suite: ${suiteName}`);
  suites[suiteNameLower]();
  log.info(`✓ Test suite "${suiteName}" completed successfully`);

  // Show confirmation in UI
  if (spreadsheet) {
    let ui: GoogleAppsScript.Spreadsheet.SpreadsheetUi;
    try {
      ui = SpreadsheetApp.getUi();
    } catch (e) {
      const errorMessage = String(e);
      if (errorMessage.includes('Cannot call SpreadsheetApp.getUi() from this context')) {
        log.info(`Test suite "${suiteName}" completed successfully`);
        return;
      }
      throw e; // Re-throw if it's a different error
    }

    ui.alert(
      "Test Suite Complete",
      `Test suite "${suiteName}" completed successfully!`,
      ui.ButtonSet.OK,
    );
  }
}
