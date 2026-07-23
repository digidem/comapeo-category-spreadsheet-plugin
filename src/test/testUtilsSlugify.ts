/// <reference path="../types.ts" />
/// <reference path="../utils.ts" />

/**
 * Unit Tests for Utils Module - Slugify Functions
 *
 * Tests slugify(), buildSlugWithFallback(), createFieldTagKey(),
 * createPresetSlug(), and other utility functions.
 *
 * Run these tests by calling testUtilsSlugify() from the Apps Script editor.
 */

/**
 * Test suite for utility functions
 */
function testUtilsSlugify(): void {
  // Environment check - ensure we're in Apps Script context
  if (typeof Logger === "undefined") {
    throw new Error("Logger not available - tests must run in Apps Script environment");
  }

  const testResults: {
    passed: number;
    failed: number;
    tests: Array<{ name: string; passed: boolean; error?: string }>;
  } = {
    passed: 0,
    failed: 0,
    tests: [],
  };

  /**
   * Helper function to run a test
   */
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

  /**
   * Helper function to assert equality
   */
  function assertEqual<T>(actual: T, expected: T, message?: string): void {
    if (actual !== expected) {
      throw new Error(
        `Assertion failed: ${message || ""}\nExpected: ${JSON.stringify(expected)}\nActual: ${JSON.stringify(actual)}`,
      );
    }
  }

  /**
   * Helper function to assert truthiness
   */
  function assertTrue(value: boolean, message?: string): void {
    if (!value) {
      throw new Error(`Assertion failed: ${message || "Expected true but got false"}`);
    }
  }

  /**
   * Helper function to assert falsiness
   */
  function assertFalse(value: boolean, message?: string): void {
    if (value) {
      throw new Error(`Assertion failed: ${message || "Expected false but got true"}`);
    }
  }

  // ═══════════════════════════════════════════════════════════════════════
  // GROUP: Basic Slugify Tests
  // ═══════════════════════════════════════════════════════════════════════

  // Test 1: Basic slugify functionality
  runTest("slugify: Basic lowercase conversion", () => {
    assertEqual(slugify("Hello World"), "hello-world", "Should convert to lowercase and replace spaces with hyphens");
    assertEqual(slugify("Test Case"), "test-case", "Should handle multiple words");
  });

  // Test 2: Special characters removal
  runTest("slugify: Remove special characters", () => {
    assertEqual(slugify("Hello@World#Test"), "hello-world-test", "Should remove special characters");
    assertEqual(slugify("Test!@#$%^&*()"), "test", "Should remove all special characters");
  });

  // Test 3: Whitespace handling
  runTest("slugify: Trim and normalize whitespace", () => {
    assertEqual(slugify("  Hello   World  "), "hello-world", "Should trim and normalize whitespace");
    assertEqual(slugify("Test\t\nCase"), "test-case", "Should handle tabs and newlines");
  });

  // Test 4: Consecutive separators
  runTest("slugify: Handle consecutive separators", () => {
    assertEqual(slugify("Hello--World"), "hello-world", "Should convert multiple hyphens to single hyphen");
    assertEqual(slugify("Hello  __  World"), "hello-world", "Should handle mixed separators");
  });

  // Test 5: Leading/trailing separators
  runTest("slugify: Remove leading and trailing separators", () => {
    assertEqual(slugify("-Hello-"), "hello", "Should remove leading and trailing hyphens");
    assertEqual(slugify("___Test___"), "test", "Should handle multiple leading/trailing separators");
  });

  // Test 6: Numbers and alphanumeric
  runTest("slugify: Handle numbers and alphanumeric", () => {
    assertEqual(slugify("Test123"), "test123", "Should preserve numbers");
    assertEqual(slugify("123 Test"), "123-test", "Should handle numbers at start");
  });

  // Test 7: Unicode characters
  runTest("slugify: Handle unicode characters", () => {
    assertEqual(slugify("Café"), "cafe", "Should remove accented characters");
    assertEqual(slugify("naïve"), "naive", "Should convert to ASCII");
  });

  // Test 8: Empty and null values
  runTest("slugify: Handle empty and null values", () => {
    assertEqual(slugify(""), "", "Should return empty string for empty input");
    assertEqual(slugify(null), "", "Should return empty string for null");
    assertEqual(slugify(undefined), "", "Should return empty string for undefined");
  });

  // Test 9: Non-string inputs
  runTest("slugify: Convert non-string inputs", () => {
    assertEqual(slugify(123), "123", "Should convert numbers to string");
    assertEqual(slugify(true), "true", "Should convert boolean to string");
  });

  // Test 10: Complex real-world examples
  runTest("slugify: Complex real-world examples", () => {
    assertEqual(slugify("Category: Water Quality Monitoring"), "category-water-quality-monitoring");
    assertEqual(slugify("Field (Required)"), "field-required");
    assertEqual(slugify("Test!@#$%^&*()_+={}[]|\\:\";<>?,./~`"), "test");
  });

  // ═══════════════════════════════════════════════════════════════════════
  // GROUP: Build Slug With Fallback Tests
  // ═══════════════════════════════════════════════════════════════════════

  // Test 11: buildSlugWithFallback with valid slug
  runTest("buildSlugWithFallback: Use source when valid", () => {
    assertEqual(buildSlugWithFallback("Test Field", "field"), "test-field", "Should use source slug when valid");
    assertEqual(buildSlugWithFallback("Category Name", "category"), "category-name", "Should handle category slugs");
  });

  // Test 12: buildSlugWithFallback with empty source
  runTest("buildSlugWithFallback: Fallback for empty source", () => {
    assertEqual(buildSlugWithFallback("", "field", 0), "field-1", "Should use fallback prefix with index");
    assertEqual(buildSlugWithFallback(null, "category", 5), "category-6", "Should handle null/empty with index");
  });

  // Test 13: buildSlugWithFallback with non-slugifiable source
  runTest("buildSlugWithFallback: Fallback for non-slugifiable source", () => {
    assertEqual(buildSlugWithFallback("!!!", "field", 2), "field-3", "Should fallback when source has no alphanumeric");
    assertEqual(buildSlugWithFallback("---", "test", 0), "test-1", "Should fallback for separator-only strings");
  });

  // Test 14: buildSlugWithFallback with empty fallback prefix
  runTest("buildSlugWithFallback: Handle empty fallback prefix", () => {
    assertEqual(buildSlugWithFallback("", "", 0), "item-1", "Should use 'item' as default prefix");
    assertEqual(buildSlugWithFallback("!!!", "   ", 3), "item-4", "Should trim whitespace from prefix");
  });

  // Test 15: buildSlugWithFallback with various indices
  runTest("buildSlugWithFallback: Different indices", () => {
    assertEqual(buildSlugWithFallback("", "field", 0), "field-1", "Index 0 becomes item 1");
    assertEqual(buildSlugWithFallback("", "field", 5), "field-6", "Index 5 becomes item 6");
    assertEqual(buildSlugWithFallback("", "field", 99), "field-100", "Should handle large indices");
  });

  // ═══════════════════════════════════════════════════════════════════════
  // GROUP: Create Field Tag Key Tests
  // ═══════════════════════════════════════════════════════════════════════

  // Test 16: createFieldTagKey with valid name
  runTest("createFieldTagKey: Generate key from field name", () => {
    assertEqual(createFieldTagKey("Water Quality", 0), "water-quality", "Should create slug from field name");
    assertEqual(createFieldTagKey("pH Level", 1), "ph-level", "Should handle field name with capitals");
  });

  // Test 17: createFieldTagKey with empty name
  runTest("createFieldTagKey: Fallback for empty field name", () => {
    assertEqual(createFieldTagKey("", 0), "field-1", "Should use fallback when name is empty");
    assertEqual(createFieldTagKey(null, 5), "field-6", "Should handle null name with index");
  });

  // Test 18: createFieldTagKey with non-slugifiable name
  runTest("createFieldTagKey: Fallback for non-slugifiable name", () => {
    assertEqual(createFieldTagKey("!!!", 2), "field-3", "Should use fallback when name has no alphanumeric");
    assertEqual(createFieldTagKey("---", 0), "field-1", "Should handle separator-only names");
  });

  // Test 19: createFieldTagKey with and without index
  runTest("createFieldTagKey: Index parameter handling", () => {
    assertEqual(createFieldTagKey("Test"), "test", "Should create slug when name is valid");
    assertEqual(createFieldTagKey("Test", undefined), "test", "Should create slug when name is valid (undefined index)");
    assertEqual(createFieldTagKey("Test", 10), "test", "Should create slug when name is valid (with index)");
  });

  // Test 20: createFieldTagKey real-world examples
  runTest("createFieldTagKey: Real-world field names", () => {
    assertEqual(createFieldTagKey("Temperature (°C)", 0), "temperature-c", "Should handle special characters");
    assertEqual(createFieldTagKey("GPS Location", 1), "gps-location", "Should handle multi-word names");
  });

  // ═══════════════════════════════════════════════════════════════════════
  // GROUP: Create Preset Slug Tests
  // ═══════════════════════════════════════════════════════════════════════

  // Test 21: createPresetSlug with valid name
  runTest("createPresetSlug: Generate slug from preset name", () => {
    assertEqual(createPresetSlug("Water Quality", 0), "water-quality", "Should create slug from preset name");
    assertEqual(createPresetSlug("Infrastructure", 5), "infrastructure", "Should handle infrastructure names");
  });

  // Test 22: createPresetSlug with empty name
  runTest("createPresetSlug: Fallback for empty preset name", () => {
    assertEqual(createPresetSlug("", 0), "category-1", "Should use 'category' fallback prefix");
    assertEqual(createPresetSlug(null, 3), "category-4", "Should handle null with index");
  });

  // Test 23: createPresetSlug with non-slugifiable name
  runTest("createPresetSlug: Fallback for non-slugifiable name", () => {
    assertEqual(createPresetSlug("!!!", 1), "category-2", "Should use category prefix as fallback");
    assertEqual(createPresetSlug("   ", 5), "category-6", "Should handle whitespace-only names");
  });

  // Test 24: createPresetSlug with and without index
  runTest("createPresetSlug: Index parameter handling", () => {
    assertEqual(createPresetSlug("Test"), "test", "Should create slug when name is valid");
    assertEqual(createPresetSlug("Test", 20), "test", "Should create slug when name is valid (with index)");
  });

  // Test 25: createPresetSlug real-world examples
  runTest("createPresetSlug: Real-world preset names", () => {
    assertEqual(createPresetSlug("Water Source", 0), "water-source", "Should handle category names");
    assertEqual(createPresetSlug("Monitoring Site", 1), "monitoring-site", "Should handle multi-word names");
  });

  // ═══════════════════════════════════════════════════════════════════════
  // GROUP: Integration Tests
  // ═══════════════════════════════════════════════════════════════════════

  // Test 26: slugify consistency across functions
  runTest("Integration: Consistent slug generation", () => {
    const fieldName = "Test Field";
    const fieldKey = createFieldTagKey(fieldName, 0);
    const presetName = "Test Category";
    const presetSlug = createPresetSlug(presetName, 0);

    assertEqual(fieldKey, "test-field", "Field key should be properly slugified");
    assertEqual(presetSlug, "test-category", "Preset slug should be properly slugified");
  });

  // Test 27: Fallback generation with different indices
  runTest("Integration: Fallback generation across indices", () => {
    // Simulate multiple empty/invalid field names
    const fallbacks = [
      buildSlugWithFallback("", "field", 0),
      buildSlugWithFallback(null, "field", 1),
      buildSlugWithFallback("   ", "field", 2),
    ];

    assertEqual(fallbacks[0], "field-1", "First fallback should be field-1");
    assertEqual(fallbacks[1], "field-2", "Second fallback should be field-2");
    assertEqual(fallbacks[2], "field-3", "Third fallback should be field-3");
  });

  // Test 28: Real-world scenario - field and preset generation
  runTest("Integration: Real-world field and preset generation", () => {
    const fields = [
      createFieldTagKey("pH Level", 0),
      createFieldTagKey("Turbidity", 1),
      createFieldTagKey("", 2), // Empty field
    ];

    const categories = [
      createPresetSlug("Water Quality", 0),
      createPresetSlug("Infrastructure", 1),
    ];

    assertEqual(fields[0], "ph-level", "Should generate field key for valid name");
    assertEqual(fields[1], "turbidity", "Should generate field key for valid name");
    assertEqual(fields[2], "field-3", "Should use fallback for empty name");

    assertEqual(categories[0], "water-quality", "Should generate preset slug for valid name");
    assertEqual(categories[1], "infrastructure", "Should generate preset slug for valid name");
  });

  // Test 29: Special characters in fallback prefix
  runTest("Integration: Slugify fallback prefix", () => {
    const result = buildSlugWithFallback("", "Field!@#", 0);
    // The fallback prefix should be slugified too
    assertEqual(result, "field-1", "Should slugify the fallback prefix");
  });

  // Test 30: Unicode handling across all functions
  runTest("Integration: Unicode character handling", () => {
    assertEqual(slugify("Café"), "cafe", "slugify should convert unicode");
    assertEqual(createFieldTagKey("café", 0), "cafe", "Field key should handle unicode");
    assertEqual(createPresetSlug("Categoría", 0), "categoria", "Preset slug should handle unicode");
  });

  // ═══════════════════════════════════════════════════════════════════════
  // GROUP: Edge Case Tests (Production Hardening)
  // ═══════════════════════════════════════════════════════════════════════

  // Test 31: Very long strings
  runTest("Edge case: Very long string handling", () => {
    const longString = "a".repeat(1000);
    assertEqual(slugify(longString).length, 1000, "Should handle very long strings");
    assertEqual(slugify(longString).indexOf("--"), -1, "Should not create double hyphens in long strings");
  });

  // Test 32: Only special characters
  runTest("Edge case: Only special characters", () => {
    assertEqual(slugify("!@#$%^&*()"), "", "Should return empty for only special chars");
    assertEqual(createFieldTagKey("!@#$%^&*()", 5), "field-6", "Should use fallback with index");
  });

  // Test 33: Mixed separators
  runTest("Edge case: Mixed separators", () => {
    assertEqual(slugify("Test_-_-\t\n   Case"), "test-case", "Should normalize all separator types");
    assertEqual(slugify("a---b___c.d,e f"), "a-b-c-d-e-f", "Should convert all separators to hyphens");
  });

  // Test 34: Numbers with special cases
  runTest("Edge case: Numbers and special characters", () => {
    assertEqual(slugify("123-456-789"), "123-456-789", "Should preserve number sequences");
    assertEqual(slugify("Test123!!!"), "test123", "Should remove special chars but keep numbers");
  });

  // Test 35: Single character inputs
  runTest("Edge case: Single character inputs", () => {
    assertEqual(slugify("a"), "a", "Should handle single character");
    assertEqual(slugify("!"), "", "Should return empty for single special char");
    assertEqual(buildSlugWithFallback("", "x", 0), "x-1", "Should handle single char prefix");
  });

  // Test 36: Whitespace-only strings
  runTest("Edge case: Whitespace-only strings", () => {
    assertEqual(slugify("   "), "", "Should return empty for whitespace only");
    assertEqual(slugify("\t\n\r"), "", "Should handle various whitespace characters");
  });

  // Test 37: Complex unicode characters (slugify is ASCII-only for entity IDs)
  runTest("Edge case: Complex unicode characters", () => {
    assertEqual(slugify("日本語"), "", "slugify is ASCII-only; non-Latin entity names fold to empty and fall back to prefix-N");
    assertEqual(slugify("Emoji's 🚀🎉"), "emojis", "Should remove emojis and handle apostrophes");
  });

  // Test 37b: canonicalizeOptionValue — select option values preserve EVERY
  // script. Regression: option values used to be slugify()'d (ASCII-only), so
  // all-Thai / all-Vietnamese / all-Cyrillic / all-Greek option lists collapsed
  // to one value and the linter deleted the "duplicates".
  runTest("canonicalizeOptionValue: preserves all scripts (no false duplicates)", () => {
    // Thai: distinct options stay distinct, combining marks intact.
    const thaiA = canonicalizeOptionValue("เห็นด้วยตา");
    const thaiB = canonicalizeOptionValue("ได้ยินเสียงร้อง");
    assertTrue(thaiA.length > 0, "Thai option value must not be empty");
    assertTrue(thaiB.length > 0, "Thai option value must not be empty");
    assertTrue(thaiA !== thaiB, "Distinct Thai options must not collide");
    assertEqual(canonicalizeOptionValue("เห็นด้วยตา"), "เห็นด้วยตา", "Thai combining marks must be preserved");
    assertEqual(canonicalizeOptionValue("ไม้ตอง (ทองหลาง)"), "ไม้ตอง-ทองหลาง", "Thai punctuation becomes a separator, script intact");

    // Vietnamese tones are semantic — must NOT collapse (slugify gave "ma" for all).
    const vn = ["má", "mà", "mả", "mã", "mạ"].map(canonicalizeOptionValue);
    assertEqual(new Set(vn).size, 5, "Vietnamese tone variants must stay distinct");

    // Cyrillic й ≠ и and Greek ά ≠ α (slugify collapsed both pairs).
    assertTrue(canonicalizeOptionValue("й") !== canonicalizeOptionValue("и"), "Cyrillic й and и must stay distinct");
    assertTrue(canonicalizeOptionValue("ά") !== canonicalizeOptionValue("α"), "Greek ά and α must stay distinct");

    // Diacritics preserved (unlike slugify's Café -> cafe).
    assertEqual(canonicalizeOptionValue("Café"), "café", "Diacritics are semantic and must be kept");

    // Emoji variation selectors must not leave an invisible colliding mark.
    assertEqual(canonicalizeOptionValue("❤️"), "", "Emoji-only option reduces to empty (falls back), not an invisible mark");
    assertEqual(canonicalizeOptionValue("☕️"), "", "Coffee emoji also reduces to empty (shares the same fallback path)");
    // Both canonicalize to "", but callers use `canonicalizeOptionValue(x) || x`
    // — verify that fallback actually keeps distinct emoji options distinct,
    // since two different options both mapping to "" would otherwise collide.
    const heartValue = canonicalizeOptionValue("❤️") || "❤️";
    const coffeeValue = canonicalizeOptionValue("☕️") || "☕️";
    assertTrue(heartValue !== coffeeValue, "Distinct emoji options must stay distinct via the raw-label fallback");

    // Empty / separator-only input.
    assertEqual(canonicalizeOptionValue(""), "", "Empty input -> empty value");
    assertEqual(canonicalizeOptionValue("   "), "", "Whitespace-only -> empty value");
  });

  // Test 38: Performance - many slugs
  runTest("Edge case: Generate many slugs efficiently", () => {
    const start = Date.now();
    for (let i = 0; i < 1000; i++) {
      slugify(`Test ${i}`);
      createFieldTagKey(`Field ${i}`, i);
      createPresetSlug(`Category ${i}`, i);
    }
    const duration = Date.now() - start;
    assertTrue(duration < 5000, `Should generate 1000 slugs in reasonable time (${duration}ms)`);
  });

  // Test 39: Index overflow protection
  runTest("Edge case: Large index values", () => {
    const result = buildSlugWithFallback("", "field", 999999);
    assertEqual(result, "field-1000000", "Should handle large indices correctly");
  });

  // Test 40: Regression - known issue cases
  runTest("Regression: Known slugify edge cases", () => {
    // Known edge cases that have caused issues
    assertEqual(slugify("Test_Case"), "test-case", "Should convert underscores to hyphens");
    assertEqual(slugify("Test--Case"), "test-case", "Should not create double hyphens");
    assertEqual(slugify("-Test-"), "test", "Should trim leading/trailing hyphens");
  });

  // Print test results
  Logger.log("\n=== Utils Slugify Test Results ===");
  Logger.log(`Total Tests: ${testResults.passed + testResults.failed}`);
  Logger.log(`Passed: ${testResults.passed}`);
  Logger.log(`Failed: ${testResults.failed}`);
  Logger.log("\nDetailed Results:");

  for (const test of testResults.tests) {
    const status = test.passed ? "✓ PASS" : "✗ FAIL";
    Logger.log(`${status}: ${test.name}`);
    if (!test.passed && test.error) {
      Logger.log(`  Error: ${test.error}`);
    }
  }

  // Show summary in UI
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (spreadsheet) {
    const message = `Tests: ${testResults.passed + testResults.failed}\nPassed: ${testResults.passed}\nFailed: ${testResults.failed}`;
    spreadsheet.toast(message, "Utils Slugify Tests", 10);
  }

  // Throw error if any tests failed
  if (testResults.failed > 0) {
    throw new Error(`${testResults.failed} test(s) failed. Check logs for details.`);
  }
}

/**
 * Quick smoke test for utils slugify functions
 * Tests the most critical functionality only
 */
function testUtilsSlugifyQuick(): void {
  // Environment check
  if (typeof Logger === "undefined") {
    throw new Error("Logger not available - tests must run in Apps Script environment");
  }

  // Critical tests
  const tests = [
    { name: "Basic slugify", actual: slugify("Hello World"), expected: "hello-world" },
    { name: "Empty string", actual: slugify(""), expected: "" },
    { name: "Field tag key", actual: createFieldTagKey("Test Field", 0), expected: "test-field" },
    { name: "Preset slug", actual: createPresetSlug("Test Category", 0), expected: "test-category" },
    { name: "Fallback generation", actual: buildSlugWithFallback("", "field", 0), expected: "field-1" },
    { name: "Special characters", actual: slugify("Test!@#"), expected: "test" },
    { name: "Unicode handling", actual: slugify("Café"), expected: "cafe" },
  ];

  let passed = 0;
  let failed = 0;

  for (const test of tests) {
    if (test.actual === test.expected) {
      Logger.log(`✓ ${test.name}`);
      passed++;
    } else {
      Logger.log(`✗ ${test.name}: expected ${test.expected}, got ${test.actual}`);
      failed++;
    }
  }

  Logger.log(`\nQuick Test Results: ${passed} passed, ${failed} failed`);

  if (failed > 0) {
    throw new Error(`${failed} quick test(s) failed`);
  }
}
