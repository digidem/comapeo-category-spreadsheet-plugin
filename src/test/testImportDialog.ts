/**
 * Unit tests for the Import Dialog functionality
 *
 * Tests the HTML builder functions and format detection logic
 * used in the category file import dialog.
 */

/**
 * Test results interface
 */
interface ImportDialogTestResult {
  name: string;
  passed: boolean;
  message: string;
}

/**
 * Runs all import dialog tests
 */
function runImportDialogTests(): ImportDialogTestResult[] {
  const results: ImportDialogTestResult[] = [];

  console.log("===== Import Dialog Tests =====");

  // Test HTML builder functions exist and return strings
  results.push(testBuildImportDialogStylesReturnsString());
  results.push(testBuildImportDialogBodyReturnsString());
  results.push(testBuildImportDialogScriptReturnsString());
  results.push(testBuildImportDialogHtmlReturnsValidHtml());

  // Test HTML contains required elements
  results.push(testHtmlContainsUploadArea());
  results.push(testHtmlContainsDriveInput());
  results.push(testHtmlContainsProgressBar());
  results.push(testHtmlContainsImportButton());
  results.push(testHtmlContainsFormatBadges());

  // Test script contains required functions
  results.push(testScriptContainsDetectFormat());
  results.push(testScriptContainsHandleImport());
  results.push(testScriptContainsProgressFunctions());

  // Summary
  const passed = results.filter((r) => r.passed).length;
  const failed = results.filter((r) => !r.passed).length;
  console.log(`\n===== Results: ${passed} passed, ${failed} failed =====`);

  return results;
}

/**
 * Test that buildImportDialogStyles returns a non-empty string
 */
function testBuildImportDialogStylesReturnsString(): ImportDialogTestResult {
  const name = "buildImportDialogStyles returns string";
  try {
    const result = buildImportDialogStyles();
    if (typeof result !== "string") {
      return { name, passed: false, message: "Expected string, got " + typeof result };
    }
    if (result.length < 100) {
      return { name, passed: false, message: "CSS seems too short: " + result.length + " chars" };
    }
    if (!result.includes("body {") || !result.includes(".upload-area")) {
      return { name, passed: false, message: "Missing expected CSS selectors" };
    }
    console.log("✓ " + name);
    return { name, passed: true, message: "OK" };
  } catch (e) {
    console.log("✗ " + name + ": " + e);
    return { name, passed: false, message: String(e) };
  }
}

/**
 * Test that buildImportDialogBody returns a non-empty string with HTML
 */
function testBuildImportDialogBodyReturnsString(): ImportDialogTestResult {
  const name = "buildImportDialogBody returns HTML string";
  try {
    const result = buildImportDialogBody();
    if (typeof result !== "string") {
      return { name, passed: false, message: "Expected string, got " + typeof result };
    }
    if (result.length < 100) {
      return { name, passed: false, message: "HTML seems too short: " + result.length + " chars" };
    }
    console.log("✓ " + name);
    return { name, passed: true, message: "OK" };
  } catch (e) {
    console.log("✗ " + name + ": " + e);
    return { name, passed: false, message: String(e) };
  }
}

/**
 * Test that buildImportDialogScript returns a non-empty string with JS
 */
function testBuildImportDialogScriptReturnsString(): ImportDialogTestResult {
  const name = "buildImportDialogScript returns JS string";
  try {
    const result = buildImportDialogScript();
    if (typeof result !== "string") {
      return { name, passed: false, message: "Expected string, got " + typeof result };
    }
    if (result.length < 100) {
      return { name, passed: false, message: "JS seems too short: " + result.length + " chars" };
    }
    console.log("✓ " + name);
    return { name, passed: true, message: "OK" };
  } catch (e) {
    console.log("✗ " + name + ": " + e);
    return { name, passed: false, message: String(e) };
  }
}

/**
 * Test that buildImportDialogHtml returns valid HTML document
 */
function testBuildImportDialogHtmlReturnsValidHtml(): ImportDialogTestResult {
  const name = "buildImportDialogHtml returns valid HTML document";
  try {
    const result = buildImportDialogHtml();
    if (typeof result !== "string") {
      return { name, passed: false, message: "Expected string, got " + typeof result };
    }
    if (!result.includes("<!DOCTYPE html>")) {
      return { name, passed: false, message: "Missing DOCTYPE" };
    }
    if (!result.includes("<html>") || !result.includes("</html>")) {
      return { name, passed: false, message: "Missing html tags" };
    }
    if (!result.includes("<head>") || !result.includes("</head>")) {
      return { name, passed: false, message: "Missing head tags" };
    }
    if (!result.includes("<body>") || !result.includes("</body>")) {
      return { name, passed: false, message: "Missing body tags" };
    }
    if (!result.includes("<style>") || !result.includes("</style>")) {
      return { name, passed: false, message: "Missing style tags" };
    }
    if (!result.includes("<script>") || !result.includes("</script>")) {
      return { name, passed: false, message: "Missing script tags" };
    }
    console.log("✓ " + name);
    return { name, passed: true, message: "OK" };
  } catch (e) {
    console.log("✗ " + name + ": " + e);
    return { name, passed: false, message: String(e) };
  }
}

/**
 * Test that HTML contains upload area element
 */
function testHtmlContainsUploadArea(): ImportDialogTestResult {
  const name = "HTML contains upload area";
  try {
    const body = buildImportDialogBody();
    if (!body.includes('id="uploadArea"')) {
      return { name, passed: false, message: "Missing uploadArea element" };
    }
    if (!body.includes('class="upload-area"')) {
      return { name, passed: false, message: "Missing upload-area class" };
    }
    if (!body.includes('id="fileInput"')) {
      return { name, passed: false, message: "Missing fileInput element" };
    }
    console.log("✓ " + name);
    return { name, passed: true, message: "OK" };
  } catch (e) {
    console.log("✗ " + name + ": " + e);
    return { name, passed: false, message: String(e) };
  }
}

/**
 * Test that HTML contains Drive input element
 */
function testHtmlContainsDriveInput(): ImportDialogTestResult {
  const name = "HTML contains Drive input";
  try {
    const body = buildImportDialogBody();
    if (!body.includes('id="driveInput"')) {
      return { name, passed: false, message: "Missing driveInput element" };
    }
    if (!body.includes('class="drive-input"')) {
      return { name, passed: false, message: "Missing drive-input class" };
    }
    console.log("✓ " + name);
    return { name, passed: true, message: "OK" };
  } catch (e) {
    console.log("✗ " + name + ": " + e);
    return { name, passed: false, message: String(e) };
  }
}

/**
 * Test that HTML contains progress bar elements
 */
function testHtmlContainsProgressBar(): ImportDialogTestResult {
  const name = "HTML contains progress bar";
  try {
    const body = buildImportDialogBody();
    if (!body.includes('id="progressContainer"')) {
      return { name, passed: false, message: "Missing progressContainer element" };
    }
    if (!body.includes('id="progressBar"')) {
      return { name, passed: false, message: "Missing progressBar element" };
    }
    if (!body.includes('id="progressStage"')) {
      return { name, passed: false, message: "Missing progressStage element" };
    }
    if (!body.includes('id="progressPercent"')) {
      return { name, passed: false, message: "Missing progressPercent element" };
    }
    console.log("✓ " + name);
    return { name, passed: true, message: "OK" };
  } catch (e) {
    console.log("✗ " + name + ": " + e);
    return { name, passed: false, message: String(e) };
  }
}

/**
 * Test that HTML contains import button
 */
function testHtmlContainsImportButton(): ImportDialogTestResult {
  const name = "HTML contains import button";
  try {
    const body = buildImportDialogBody();
    if (!body.includes('id="importBtn"')) {
      return { name, passed: false, message: "Missing importBtn element" };
    }
    if (!body.includes("handleImport()")) {
      return { name, passed: false, message: "Missing handleImport onclick" };
    }
    console.log("✓ " + name);
    return { name, passed: true, message: "OK" };
  } catch (e) {
    console.log("✗ " + name + ": " + e);
    return { name, passed: false, message: String(e) };
  }
}

/**
 * Test that CSS contains format badge styles
 */
function testHtmlContainsFormatBadges(): ImportDialogTestResult {
  const name = "CSS contains format badge styles";
  try {
    const styles = buildImportDialogStyles();
    const requiredClasses = [
      ".format-comapeocat",
      ".format-mapeosettings",
      ".format-zip",
      ".format-unknown",
    ];
    for (const className of requiredClasses) {
      if (!styles.includes(className)) {
        return { name, passed: false, message: "Missing " + className + " style" };
      }
    }
    console.log("✓ " + name);
    return { name, passed: true, message: "OK" };
  } catch (e) {
    console.log("✗ " + name + ": " + e);
    return { name, passed: false, message: String(e) };
  }
}

/**
 * Test that script contains detectFormat function
 */
function testScriptContainsDetectFormat(): ImportDialogTestResult {
  const name = "Script contains detectFormat function";
  try {
    const script = buildImportDialogScript();
    if (!script.includes("function detectFormat(")) {
      return { name, passed: false, message: "Missing detectFormat function" };
    }
    if (!script.includes(".comapeocat")) {
      return { name, passed: false, message: "Missing .comapeocat detection" };
    }
    if (!script.includes(".mapeosettings")) {
      return { name, passed: false, message: "Missing .mapeosettings detection" };
    }
    if (!script.includes(".zip")) {
      return { name, passed: false, message: "Missing .zip detection" };
    }
    console.log("✓ " + name);
    return { name, passed: true, message: "OK" };
  } catch (e) {
    console.log("✗ " + name + ": " + e);
    return { name, passed: false, message: String(e) };
  }
}

/**
 * Test that script contains handleImport function
 */
function testScriptContainsHandleImport(): ImportDialogTestResult {
  const name = "Script contains handleImport function";
  try {
    const script = buildImportDialogScript();
    if (!script.includes("function handleImport(")) {
      return { name, passed: false, message: "Missing handleImport function" };
    }
    if (!script.includes("startImportCategoryJob")) {
      return { name, passed: false, message: "Missing local file import call" };
    }
    if (!script.includes("processImportFile")) {
      return { name, passed: false, message: "Missing Drive import call" };
    }
    console.log("✓ " + name);
    return { name, passed: true, message: "OK" };
  } catch (e) {
    console.log("✗ " + name + ": " + e);
    return { name, passed: false, message: String(e) };
  }
}

/**
 * Test that script contains progress update functions
 */
function testScriptContainsProgressFunctions(): ImportDialogTestResult {
  const name = "Script contains progress functions";
  try {
    const script = buildImportDialogScript();
    if (!script.includes("function showProgress(")) {
      return { name, passed: false, message: "Missing showProgress function" };
    }
    if (!script.includes("function updateProgress(")) {
      return { name, passed: false, message: "Missing updateProgress function" };
    }
    console.log("✓ " + name);
    return { name, passed: true, message: "OK" };
  } catch (e) {
    console.log("✗ " + name + ": " + e);
    return { name, passed: false, message: String(e) };
  }
}

/**
 * Menu entry point for running import dialog tests
 */
function runImportDialogTestsFromMenu(): void {
  const results = runImportDialogTests();
  const passed = results.filter((r) => r.passed).length;
  const failed = results.filter((r) => !r.passed).length;

  let message = `Import Dialog Tests\n\n`;
  message += `Passed: ${passed}\nFailed: ${failed}\n\n`;

  if (failed > 0) {
    message += "Failed tests:\n";
    for (const result of results) {
      if (!result.passed) {
        message += `- ${result.name}: ${result.message}\n`;
      }
    }
  }

  SpreadsheetApp.getUi().alert("Test Results", message, SpreadsheetApp.getUi().ButtonSet.OK);
}
