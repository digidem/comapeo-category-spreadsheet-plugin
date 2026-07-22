/// <reference path="../types.ts" />
/// <reference path="../regressionTesting.ts" />

/**
 * Unit Tests for the regression-test sanitizer (src/regressionTesting.ts).
 *
 * Guards the fix for the Categories-sheet sanitization bug: the sanitizer used
 * to hardcode index 3 as "Color" and overwrite it with "#0066CC". In the
 * current six-column layout (Name, Icon, Fields, Applies, Category ID, Icon ID)
 * index 3 is the Applies column, so the old code corrupted every category's
 * appliesTo value and broke generation with a track-coverage error.
 *
 * `sanitizeRowData` is a pure function (no SpreadsheetApp calls), so these
 * assertions run against its return value directly.
 *
 * Run by calling testRegressionTesting() from the Apps Script editor.
 */
function testRegressionTesting(): void {
  if (typeof Logger === "undefined") {
    throw new Error("Logger not available - tests must run in Apps Script environment");
  }

  const testResults: {
    passed: number;
    failed: number;
    tests: Array<{ name: string; passed: boolean; error?: string }>;
  } = { passed: 0, failed: 0, tests: [] };

  function runTest(name: string, testFn: () => void): void {
    try {
      testFn();
      testResults.tests.push({ name, passed: true });
      testResults.passed++;
    } catch (error) {
      testResults.tests.push({
        name,
        passed: false,
        error: error.message || String(error),
      });
      testResults.failed++;
    }
  }

  function assertEqual<T>(actual: T, expected: T, message?: string): void {
    if (JSON.stringify(actual) !== JSON.stringify(expected)) {
      throw new Error(
        `Assertion failed: ${message || ""}\nExpected: ${JSON.stringify(expected)}\nActual: ${JSON.stringify(actual)}`,
      );
    }
  }

  // Canonical current header (see src/constants/columns.ts CATEGORY_COL).
  const canonicalHeader = [
    "Name",
    "Icon",
    "Fields",
    "Applies",
    "Category ID",
    "Icon ID",
  ];

  // --- Test 1: Applies column is preserved, not overwritten ---
  runTest("Applies value preserved (track, observation)", () => {
    const row = ["River", "https://drive/x.svg", "Width, Depth", "track, observation", "cat-river", "icon-river"];
    const out = sanitizeRowData(row, "Categories", 0, canonicalHeader);
    assertEqual(out[3], "track, observation", "Applies must be preserved verbatim");
  });

  runTest("Applies value preserved (observation only)", () => {
    const row = ["Tree", "https://drive/y.svg", "Species", "observation", "cat-tree", "icon-tree"];
    const out = sanitizeRowData(row, "Categories", 4, canonicalHeader);
    assertEqual(out[3], "observation", "Single-token Applies must be preserved");
  });

  // --- Test 2: Category ID becomes a deterministic test id ---
  runTest("Category ID replaced with deterministic test id", () => {
    const row = ["River", "https://drive/x.svg", "Width", "track", "cat-river", "icon-river"];
    const out = sanitizeRowData(row, "Categories", 0, canonicalHeader);
    assertEqual(out[4], "test-category-1", "Category ID must be test-category-(rowIndex+1)");
  });

  // --- Test 3: Icon ID is cleared so placeholder falls back to category id ---
  runTest("Icon ID cleared", () => {
    const row = ["River", "https://drive/x.svg", "Width", "track", "cat-river", "icon-river"];
    const out = sanitizeRowData(row, "Categories", 0, canonicalHeader);
    assertEqual(out[5], "", "Icon ID must be cleared");
  });

  // --- Test 4: Name and Icon sanitized as before ---
  runTest("Name and Icon sanitized", () => {
    const row = ["River", "https://drive/secret.svg", "Width", "track", "cat-river", "icon-river"];
    const out = sanitizeRowData(row, "Categories", 2, canonicalHeader);
    assertEqual(out[0], "Test Category 3", "Name must be generic");
    assertEqual(out[1], TEST_PLACEHOLDER_ICON_URL, "Icon URL must be the placeholder");
  });

  // --- Test 5: row width unchanged (setValues requires a rectangular grid) ---
  runTest("Row width preserved", () => {
    const row = ["River", "https://drive/x.svg", "Width", "track", "cat-river", "icon-river"];
    const out = sanitizeRowData(row, "Categories", 0, canonicalHeader);
    assertEqual(out.length, row.length, "Sanitized row must keep the same width");
  });

  // --- Test 6: full canonical row end-to-end ---
  runTest("Full canonical row sanitized correctly", () => {
    const row = ["River", "https://drive/x.svg", "Width, Depth", "track, observation", "cat-river", "icon-river"];
    const out = sanitizeRowData(row, "Categories", 0, canonicalHeader);
    assertEqual(out, [
      "Test Category 1",
      TEST_PLACEHOLDER_ICON_URL,
      "field1, field2",
      "track, observation",
      "test-category-1",
      "",
    ], "Full sanitized row mismatch");
  });

  // --- Test 7: legacy four-column layout with a real Color column still works ---
  runTest("Legacy layout: explicit Color column sanitized, no Applies corruption", () => {
    const legacyHeader = ["Name", "Icon", "Fields", "Color"];
    const row = ["River", "https://drive/x.svg", "Width", "#abcdef"];
    const out = sanitizeRowData(row, "Categories", 0, legacyHeader);
    assertEqual(out[3], "#0066CC", "Explicit Color column must be set to the test color");
    assertEqual(out.length, 4, "Legacy row width preserved");
  });

  // --- Test 8: header-resolved columns beat fixed positions ---
  runTest("Reordered header resolves columns by name", () => {
    // Applies placed at index 5 instead of 3; sanitizer must find it by header.
    const reorderedHeader = ["Name", "Icon", "Fields", "Category ID", "Icon ID", "Applies"];
    const row = ["River", "https://drive/x.svg", "Width", "cat-river", "icon-river", "track"];
    const out = sanitizeRowData(row, "Categories", 0, reorderedHeader);
    assertEqual(out[5], "track", "Applies resolved by header must be preserved");
    assertEqual(out[3], "test-category-1", "Category ID resolved by header must be sanitized");
    assertEqual(out[4], "", "Icon ID resolved by header must be cleared");
  });

  // --- Test 9: numeric Category ID is scrubbed (not just strings) ---
  runTest("Numeric Category ID is scrubbed", () => {
    const row = ["River", "https://drive/x.svg", "Width", "track", 12345, "icon-river"];
    const out = sanitizeRowData(row, "Categories", 0, canonicalHeader);
    assertEqual(out[4], "test-category-1", "Numeric production ID must be replaced");
  });

  // --- Test 10: import-path layout (Name, Icon, Fields, ID, Color, Icon ID) ---
  runTest("Import layout (bare ID + Color, no Applies) sanitized", () => {
    const importHeader = ["Name", "Icon", "Fields", "ID", "Color", "Icon ID"];
    const row = ["River", "https://drive/x.svg", "Width", "cat-river", "#abcdef", "icon-river"];
    const out = sanitizeRowData(row, "Categories", 0, importHeader);
    assertEqual(out[3], "test-category-1", "Header 'ID' must resolve to the Category ID column");
    assertEqual(out[4], "#0066CC", "Color column must get the test color");
    assertEqual(out[5], "", "Icon ID must be cleared");
  });

  // --- Report ---
  Logger.log("\n=== Regression Sanitizer Test Results ===");
  testResults.tests.forEach((t) => {
    Logger.log(`${t.passed ? "✓" : "✗"} ${t.name}${t.error ? "\n    " + t.error : ""}`);
  });
  Logger.log(`\nPassed: ${testResults.passed}/${testResults.tests.length}`);

  if (testResults.failed > 0) {
    throw new Error(`${testResults.failed} regression sanitizer test(s) failed`);
  }
}
