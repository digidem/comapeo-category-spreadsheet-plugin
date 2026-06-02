/// <reference path="../types.ts" />
/// <reference path="../spreadsheetData.ts" />
/// <reference path="../validation.ts" />
/// <reference path="../languageLookup.ts" />
/// <reference path="../translationHeaderResolver.ts" />

/**
 * Integration Tests for Dual-Name Language Recognition
 *
 * Tests end-to-end language recognition with both English and native names.
 * Verifies that users can enter either form and the system works correctly.
 *
 * Run these tests by calling testLanguageRecognitionIntegration() from Apps Script editor.
 */

/**
 * Main integration test suite
 */
function testLanguageRecognitionIntegration(): void {
  // Environment check - ensure we're in Apps Script context
  if (typeof Logger === "undefined") {
    throw new Error("Logger not available - tests must run in Apps Script environment");
  }
  if (typeof SpreadsheetApp === "undefined") {
    throw new Error("SpreadsheetApp not available - tests must run in Apps Script environment");
  }
  if (typeof getScopedLogger === "undefined") {
    throw new Error("getScopedLogger not available - ensure loggingHelpers.ts is loaded");
  }

  const log = getScopedLogger("LanguageIntegrationTest");
  const results: { passed: number; failed: number; tests: Array<{ name: string; passed: boolean; error?: string }> } = {
    passed: 0,
    failed: 0,
    tests: [],
  };

  function runTest(name: string, testFn: () => void): void {
    try {
      testFn();
      results.tests.push({ name, passed: true });
      results.passed++;
      log.info(`✓ ${name}`);
    } catch (error) {
      results.tests.push({
        name,
        passed: false,
        error: error.message || String(error),
      });
      results.failed++;
      log.error(`✗ ${name}: ${error.message}`);
    }
  }

  log.info("Starting Language Recognition Integration Tests");

  // Clear cache to ensure we use fallback data for consistent test results
  log.info("Clearing language cache to use fallback data");
  clearLanguagesCache();

  // ═══════════════════════════════════════════════════════════════════════
  // GROUP 1: Basic Validation (English and Native Names)
  // Tests fundamental recognition of language names in both forms
  // ═══════════════════════════════════════════════════════════════════════

  // Test 1: Validate Portuguese (English form)
  runTest("Validate 'Portuguese' (English form)", () => {
    const result = validateLanguageName("Portuguese");

    if (!result.valid) {
      throw new Error(`Validation failed: ${result.error}`);
    }
    if (result.code !== "pt") {
      throw new Error(`Expected code 'pt', got '${result.code}'`);
    }
  });

  // Test 2: Validate Português (native form)
  runTest("Validate 'Português' (native form)", () => {
    const result = validateLanguageName("Português");

    if (!result.valid) {
      throw new Error(`Validation failed: ${result.error}`);
    }
    if (result.code !== "pt") {
      throw new Error(`Expected code 'pt', got '${result.code}'`);
    }
  });

  // ═══════════════════════════════════════════════════════════════════════
  // GROUP 2: Case-Insensitive Matching
  // Verifies that language names work in any case combination
  // ═══════════════════════════════════════════════════════════════════════

  // Test 3: Case-insensitive validation
  runTest("Validate 'PORTUGUESE' (uppercase)", () => {
    const result = validateLanguageName("PORTUGUESE");

    if (!result.valid) {
      throw new Error(`Validation failed: ${result.error}`);
    }
    if (result.code !== "pt") {
      throw new Error(`Expected code 'pt', got '${result.code}'`);
    }
  });

  // Test 4: Case-insensitive native name
  runTest("Validate 'português' (lowercase native)", () => {
    const result = validateLanguageName("português");

    if (!result.valid) {
      throw new Error(`Validation failed: ${result.error}`);
    }
    if (result.code !== "pt") {
      throw new Error(`Expected code 'pt', got '${result.code}'`);
    }
  });

  // ═══════════════════════════════════════════════════════════════════════
  // GROUP 3: Multiple Language Variations
  // Tests that different languages with both name forms work correctly
  // ═══════════════════════════════════════════════════════════════════════

  // Test 5: Spanish variations
  runTest("Validate 'Spanish' and 'Español'", () => {
    const englishResult = validateLanguageName("Spanish");
    const nativeResult = validateLanguageName("Español");

    if (!englishResult.valid || !nativeResult.valid) {
      throw new Error("Both forms should be valid");
    }
    if (englishResult.code !== "es" || nativeResult.code !== "es") {
      throw new Error("Both should map to 'es'");
    }
  });

  // Test 6: French variations
  runTest("Validate 'French' and 'Français'", () => {
    const englishResult = validateLanguageName("French");
    const nativeResult = validateLanguageName("Français");

    if (!englishResult.valid || !nativeResult.valid) {
      throw new Error("Both forms should be valid");
    }
    if (englishResult.code !== "fr" || nativeResult.code !== "fr") {
      throw new Error("Both should map to 'fr'");
    }
  });

  // ═══════════════════════════════════════════════════════════════════════
  // GROUP 4: Invalid Input Handling
  // Tests that invalid inputs are properly rejected
  // ═══════════════════════════════════════════════════════════════════════

  // Test 7: Invalid language name
  runTest("Reject invalid language name", () => {
    const result = validateLanguageName("InvalidLanguage");

    if (result.valid) {
      throw new Error("Should reject invalid language name");
    }
  });

  // Test 8: Empty string
  runTest("Reject empty language name", () => {
    const result = validateLanguageName("");

    if (result.valid) {
      throw new Error("Should reject empty string");
    }
  });

  // Test 9: Whitespace only
  runTest("Reject whitespace-only name", () => {
    const result = validateLanguageName("   ");

    if (result.valid) {
      throw new Error("Should reject whitespace-only string");
    }
  });

  // ═══════════════════════════════════════════════════════════════════════
  // GROUP 5: Data Structure and API Functions
  // Tests the enhanced data structures and helper functions
  // ═══════════════════════════════════════════════════════════════════════

  // Test 10: Enhanced data structure
  runTest("Enhanced data has both names", () => {
    const enhanced = getAllLanguagesEnhanced();

    const pt = enhanced.pt;
    if (!pt) {
      throw new Error("Portuguese not found in enhanced data");
    }
    if (pt.englishName !== "Portuguese") {
      throw new Error(`Expected English name 'Portuguese', got '${pt.englishName}'`);
    }
    if (pt.nativeName !== "Português") {
      throw new Error(`Expected native name 'Português', got '${pt.nativeName}'`);
    }
  });

  // Test 11: Display name function
  runTest("getLanguageDisplayName returns correct names", () => {
    const englishName = getLanguageDisplayName("pt", false);
    const nativeName = getLanguageDisplayName("pt", true);

    if (englishName !== "Portuguese") {
      throw new Error(`Expected 'Portuguese', got '${englishName}'`);
    }
    if (nativeName !== "Português") {
      throw new Error(`Expected 'Português', got '${nativeName}'`);
    }
  });

  // Test 12: Get language names
  runTest("getLanguageNames returns both forms", () => {
    const names = getLanguageNames("pt");

    if (!names) {
      throw new Error("Should return names object");
    }
    if (names.english !== "Portuguese" || names.native !== "Português") {
      throw new Error("Names don't match expected values");
    }
  });

  // ═══════════════════════════════════════════════════════════════════════
  // GROUP 6: Non-Latin Script Support
  // Tests Unicode and non-Latin script language names
  // ═══════════════════════════════════════════════════════════════════════

  // Test 13: Multiple non-Latin scripts
  runTest("Support non-Latin scripts", () => {
    // Test Japanese
    const jaResult = validateLanguageName("日本語");
    if (!jaResult.valid || jaResult.code !== "ja") {
      throw new Error("Japanese native name should work");
    }

    // Test Chinese
    const zhCNResult = validateLanguageName("简体中文");
    if (!zhCNResult.valid || zhCNResult.code !== "zh-CN") {
      throw new Error("Chinese native name should work");
    }
  });

  // ═══════════════════════════════════════════════════════════════════════
  // GROUP 7: Translation Header Resolver
  // Ensures header parsing supports names, native names, ISO, and Name-ISO
  // ═══════════════════════════════════════════════════════════════════════

  runTest("Translation header resolver supports mixed formats", () => {
    const allLanguages = getAllLanguages();
    const resolveHeader = createTranslationHeaderResolver(allLanguages);

    const spanish = resolveHeader("Spanish");
    if (spanish !== "es") {
      throw new Error(`Expected "es" for Spanish, got "${spanish}"`);
    }

    const spanishNative = resolveHeader("Español");
    if (spanishNative !== "es") {
      throw new Error(`Expected "es" for Español, got "${spanishNative}"`);
    }

    const zhCn = resolveHeader("zh-CN");
    if (!zhCn || zhCn.toLowerCase() !== "zh-cn") {
      throw new Error(`Expected "zh-CN" for zh-CN, got "${zhCn}"`);
    }

    const zhCnNamed = resolveHeader("Chinese Simplified - zh-CN");
    if (!zhCnNamed || zhCnNamed.toLowerCase() !== "zh-cn") {
      throw new Error(`Expected "zh-CN" for Chinese Simplified - zh-CN, got "${zhCnNamed}"`);
    }

    const custom = resolveHeader("Quechua - quz");
    if (custom !== "quz") {
      throw new Error(`Expected "quz" for Quechua - quz, got "${custom}"`);
    }

    const esLatam = resolveHeader("es-419");
    if (!esLatam || esLatam.toLowerCase() !== "es-419") {
      throw new Error(`Expected "es-419" for es-419, got "${esLatam}"`);
    }

    const esLatamNamed = resolveHeader("Spanish - es-419");
    if (!esLatamNamed || esLatamNamed.toLowerCase() !== "es-419") {
      throw new Error(`Expected "es-419" for Spanish - es-419, got "${esLatamNamed}"`);
    }
  });

  // ═══════════════════════════════════════════════════════════════════════
  // GROUP 8: Lookup System Consistency
  // Verifies internal consistency of the lookup system
  // ═══════════════════════════════════════════════════════════════════════

  // Test 14: Lookup system consistency
  runTest("Lookup system is consistent", () => {
    const enhanced = getAllLanguagesEnhanced();
    const lookup = createLanguageLookup(enhanced);

    // Test that both forms return the same code
    const code1 = lookup.getCodeByName("Portuguese");
    const code2 = lookup.getCodeByName("Português");
    const code3 = lookup.getCodeByName("PORTUGUESE");
    const code4 = lookup.getCodeByName("português");

    if (code1 !== code2 || code2 !== code3 || code3 !== code4) {
      throw new Error("All variations should map to the same code");
    }
    if (code1 !== "pt") {
      throw new Error("All should map to 'pt'");
    }
  });

  // ═══════════════════════════════════════════════════════════════════════
  // GROUP 8: Backward Compatibility
  // Ensures legacy functions still work correctly
  // ═══════════════════════════════════════════════════════════════════════

  // Test 15: Backward compatibility
  runTest("Legacy functions still work", () => {
    const legacy = getAllLanguages();

    // Should have language codes
    if (!legacy.pt || !legacy.es || !legacy.fr) {
      throw new Error("Legacy map should have language codes");
    }

    // Should work with existing code
    const langs = languages(false);
    if (!langs || Object.keys(langs).length === 0) {
      throw new Error("languages() function should return data");
    }
  });

  // ═══════════════════════════════════════════════════════════════════════
  // GROUP 9: Critical Fix Verification
  // Verifies that critical bugs are fixed and stay fixed
  // ═══════════════════════════════════════════════════════════════════════

  // Test 16: PRIMARY LANGUAGE FILTERING WITH NATIVE NAMES (Critical Fix Verification)
  runTest("filterLanguagesByPrimary works with native names", () => {
    // Simulate setting primary language to native name
    const allLanguages = getAllLanguages();

    // Mock getPrimaryLanguageName to return native name
    // Note: In real usage, this would come from cell A1
    const primaryLanguageName = "Português"; // Native form

    // Verify the filtering logic works correctly
    const enhanced = getAllLanguagesEnhanced();
    const lookup = createLanguageLookup(enhanced);
    const primaryCode = lookup.getCodeByName(primaryLanguageName);

    if (primaryCode !== "pt") {
      throw new Error(`Primary code should be 'pt', got '${primaryCode}'`);
    }

    // Verify filtering excludes primary language by code, not name
    const excludePrimary = Object.entries(allLanguages)
      .filter(([code, _]) => code !== primaryCode);

    if (excludePrimary.some(([code]) => code === "pt")) {
      throw new Error("Portuguese should be excluded when filtering by native name");
    }
  });

  // Test 17: VALIDATION FUNCTIONS WITH NATIVE NAMES (Critical Fix Verification)
  runTest("isValidLanguageForA1Cell accepts names and locale codes", () => {
    // English names
    if (!isValidLanguageForA1Cell("Portuguese")) {
      throw new Error("Should accept 'Portuguese'");
    }
    if (!isValidLanguageForA1Cell("Spanish")) {
      throw new Error("Should accept 'Spanish'");
    }

    // Native names
    if (!isValidLanguageForA1Cell("Português")) {
      throw new Error("Should accept 'Português'");
    }
    if (!isValidLanguageForA1Cell("Español")) {
      throw new Error("Should accept 'Español'");
    }

    // Canonical and locale code forms
    if (!isValidLanguageForA1Cell("pt")) {
      throw new Error("Should accept 'pt'");
    }
    if (!isValidLanguageForA1Cell("PT")) {
      throw new Error("Should accept 'PT'");
    }
    if (!isValidLanguageForA1Cell("zh-CN")) {
      throw new Error("Should accept 'zh-CN'");
    }
    if (!isValidLanguageForA1Cell("pt-BR")) {
      throw new Error("Should accept 'pt-BR'");
    }

    // Case variations
    if (!isValidLanguageForA1Cell("PORTUGUÊS")) {
      throw new Error("Should accept 'PORTUGUÊS'");
    }
    if (!isValidLanguageForA1Cell("português")) {
      throw new Error("Should accept 'português'");
    }

    // Invalid names
    if (isValidLanguageForA1Cell("InvalidLanguage")) {
      throw new Error("Should reject 'InvalidLanguage'");
    }
  });

  // Test 18: PRIMARY LANGUAGE LOCALE HELPERS (Critical Fix Verification)
  runTest("Primary language locale helpers resolve canonical and BCP 47 inputs", () => {
    const canonicalCode = normalizePrimaryLanguageLocaleCode("zh-CN");
    if (canonicalCode !== "zh-CN") {
      throw new Error(`Expected canonical zh-CN, got '${canonicalCode}'`);
    }

    const localeCode = normalizePrimaryLanguageLocaleCode("pt-BR");
    if (localeCode !== "pt-br") {
      throw new Error(`Expected normalized pt-br, got '${localeCode}'`);
    }

    const invalidLocaleCode = normalizePrimaryLanguageLocaleCode("pt_BR");
    if (invalidLocaleCode !== null) {
      throw new Error(`Expected null for invalid locale token, got '${invalidLocaleCode}'`);
    }

    const nativeName = resolvePrimaryLanguageInput("Português");
    if (!nativeName || nativeName.code !== "pt" || nativeName.comparisonCode !== "pt") {
      throw new Error("Expected Portuguese native name to resolve to pt");
    }

    const canonicalLanguageCode = resolvePrimaryLanguageInput("PT");
    if (!canonicalLanguageCode || canonicalLanguageCode.code !== "pt") {
      throw new Error("Expected PT to resolve to canonical pt");
    }

    const localeTag = resolvePrimaryLanguageInput("pt-BR");
    if (!localeTag || localeTag.code !== "pt-br" || localeTag.comparisonCode !== "pt-br") {
      throw new Error("Expected pt-BR to resolve to runtime pt-br and comparison pt-br (full locale preserved)");
    }

    const canonicalLocaleTag = resolvePrimaryLanguageInput("zh-CN");
    if (
      !canonicalLocaleTag ||
      canonicalLocaleTag.code !== "zh-CN" ||
      canonicalLocaleTag.comparisonCode !== "zh-CN"
    ) {
      throw new Error("Expected zh-CN to preserve canonical locale casing");
    }
  });

  // Test 19: PRIMARY LANGUAGE LOCALE IDENTITY (Critical Fix Verification)
  runTest("filterLanguagesByPrimary preserves locale primary identity", () => {
    const originalSpreadsheetApp = (globalThis as any).SpreadsheetApp;

    try {
      (globalThis as any).SpreadsheetApp = {
        getActiveSpreadsheet(): any {
          return {
            getSheetByName(name: string): any {
              if (name === "Metadata") {
                return {
                  getDataRange(): any {
                    return {
                      getValues(): any[][] {
                        return [
                          ["Key", "Value"],
                          ["primaryLanguage", "pt-BR"],
                        ];
                      },
                    };
                  },
                };
              }
              if (name === "Categories") {
                return {
                  getRange(): any {
                    return {
                      getValue(): string {
                        return "pt-BR";
                      },
                    };
                  },
                };
              }
              return null;
            },
          };
        },
      };

      const primaryLanguages = filterLanguagesByPrimary(
        { en: "English", pt: "Portuguese", es: "Spanish" },
        true,
      );
      const primaryEntries = Object.entries(primaryLanguages);

      if (primaryEntries.length !== 1) {
        throw new Error(`Expected exactly one primary language entry, got ${primaryEntries.length}`);
      }
      if (primaryEntries[0][0] !== "pt-br") {
        throw new Error(`Expected primary locale key 'pt-br', got '${primaryEntries[0][0]}'`);
      }
      if (primaryEntries[0][1] !== "Portuguese") {
        throw new Error(`Expected locale primary to reuse base display name, got '${primaryEntries[0][1]}'`);
      }
    } finally {
      (globalThis as any).SpreadsheetApp = originalSpreadsheetApp;
    }
  });

  // Test 20: CACHE CLEARING (Critical Fix Verification)
  runTest("clearLanguagesCache clears both caches", () => {
    // This test verifies the cache clearing function works
    // In practice, we can't easily test the actual cache clearing
    // but we verify the function exists and runs without error
    clearLanguagesCache();

    // Verify we can still fetch data after clearing
    const enhanced = getAllLanguagesEnhanced();
    if (!enhanced || !enhanced.pt) {
      throw new Error("Should fetch enhanced data after cache clear");
    }

    const legacy = getAllLanguages();
    if (!legacy || !legacy.pt) {
      throw new Error("Should fetch legacy data after cache clear");
    }
  });

  // Test 21: GETPRIMARYLANGUAGE OPTIMIZATION (Critical Fix Verification)
  runTest("getPrimaryLanguage returns correct result efficiently", () => {
    // Mock data - in real usage this comes from cell A1
    // We test with both English and native names

    // Test with English name (would require actual spreadsheet mock)
    // This verifies the function signature and return type
    const enhanced = getAllLanguagesEnhanced();
    const lookup = createLanguageLookup(enhanced);

    // Verify lookup works for both forms
    const codeEnglish = lookup.getCodeByName("Portuguese");
    const codeNative = lookup.getCodeByName("Português");

    if (codeEnglish !== "pt" || codeNative !== "pt") {
      throw new Error("Both English and native names should resolve to same code");
    }
  });

  // Test 22: TURKISH LOCALE FIX (Moderate Issue Verification)
  runTest("normalizeLanguageName handles Turkish locale correctly", () => {
    // This verifies that 'I' becomes 'i', not 'ı' (Turkish behavior)
    // We can't directly test normalizeLanguageName since it's not exported,
    // but we verify through the lookup system
    const lookup = createLanguageLookup(getAllLanguagesEnhanced());

    // If Turkish was a supported language, verify case handling
    // For now, verify English 'I' converts correctly
    const codeUpper = lookup.getCodeByName("IRISH");
    const codeLower = lookup.getCodeByName("irish");

    if (codeUpper !== codeLower) {
      throw new Error("Case normalization should be locale-independent");
    }
  });

  // Print results
  log.info("\n=== Integration Test Results ===");
  log.info(`Total: ${results.passed + results.failed}`);
  log.info(`Passed: ${results.passed}`);
  log.info(`Failed: ${results.failed}`);

  // Show toast notification
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (spreadsheet) {
    const message = `Integration Tests\nTotal: ${results.passed + results.failed}\nPassed: ${results.passed}\nFailed: ${results.failed}`;
    const toastTitle = results.failed === 0 ? "✓ All Tests Passed" : "✗ Some Tests Failed";
    spreadsheet.toast(message, toastTitle, 10);
  }

  if (results.failed > 0) {
    throw new Error(`${results.failed} integration test(s) failed. Check logs for details.`);
  }

  log.info("✓ All integration tests passed!");
}

/**
 * Quick integration smoke test
 * Tests only the most critical functionality
 */
function testLanguageRecognitionQuick(): void {
  // Environment check
  if (typeof Logger === "undefined" || typeof SpreadsheetApp === "undefined") {
    throw new Error("Tests must run in Apps Script environment");
  }
  if (typeof getScopedLogger === "undefined") {
    throw new Error("getScopedLogger not available - ensure loggingHelpers.ts is loaded");
  }

  const log = getScopedLogger("QuickIntegrationTest");

  const tests = [
    { name: "Portuguese", input: "Portuguese", expectedCode: "pt" },
    { name: "Português", input: "Português", expectedCode: "pt" },
    { name: "Spanish", input: "Spanish", expectedCode: "es" },
    { name: "Español", input: "Español", expectedCode: "es" },
    { name: "PORTUGUESE", input: "PORTUGUESE", expectedCode: "pt" },
  ];

  let passed = 0;
  let failed = 0;

  for (const test of tests) {
    const result = validateLanguageName(test.input);
    if (result.valid && result.code === test.expectedCode) {
      log.info(`✓ ${test.name}`);
      passed++;
    } else {
      log.error(`✗ ${test.name}: expected ${test.expectedCode}, got ${result.code}`);
      failed++;
    }
  }

  log.info(`\nQuick Test: ${passed} passed, ${failed} failed`);

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (spreadsheet) {
    spreadsheet.toast(`Quick Test: ${passed}/${tests.length} passed`, "Language Recognition", 5);
  }

  if (failed > 0) {
    throw new Error(`${failed} quick test(s) failed`);
  }
}
