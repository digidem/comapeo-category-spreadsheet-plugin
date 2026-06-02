/**
 * Integration tests for import functionality using real configuration files.
 *
 * These tests download actual config files from GitHub releases and verify
 * the import pipeline works correctly with them.
 *
 * Test files:
 * - Latest CoMapeo v6: https://github.com/digidem/mapeo-default-config/releases/download/v6.0.0/build.zip
 * - Legacy CoMapeo v5: https://github.com/digidem/mapeo-default-config/releases/download/v5.0.0/build.zip
 * - Legacy Mapeo v3.6.1: https://github.com/digidem/mapeo-default-config/releases/download/v3.6.1/mapeo-default-settings-v3.6.1.mapeosettings
 */

/**
 * Test result interface
 */
interface ImportIntegrationTestResult {
  name: string;
  passed: boolean;
  message: string;
  duration?: number;
  details?: any;
}

/**
 * Test configuration file URLs
 */
const TEST_CONFIG_URLS = {
  coMapeoV6: 'https://github.com/digidem/mapeo-default-config/releases/download/v6.0.0/build.zip',
  coMapeoV5: 'https://github.com/digidem/mapeo-default-config/releases/download/v5.0.0/build.zip',
  mapeoV3: 'https://github.com/digidem/mapeo-default-config/releases/download/v3.6.1/mapeo-default-settings-v3.6.1.mapeosettings',
};

/**
 * Fetches a file from a URL and returns as blob
 */
function fetchTestFile(url: string): GoogleAppsScript.Base.Blob {
  const response = UrlFetchApp.fetch(url, {
    followRedirects: true,
    muteHttpExceptions: false,
  });

  const contentType = (response.getHeaders() as Record<string, string>)['Content-Type'] || 'application/octet-stream';
  const fileName = url.split('/').pop() || 'test-file';

  return response.getBlob().setName(fileName);
}

/**
 * Tests the import of a CoMapeo v6 configuration file
 */
function testImportCoMapeoV6(): ImportIntegrationTestResult {
  const name = 'Import CoMapeo v6.0.0 config (build.zip)';
  const startTime = Date.now();

  try {
    console.log('Fetching CoMapeo v6 config from GitHub...');
    const blob = fetchTestFile(TEST_CONFIG_URLS.coMapeoV6);
    console.log(`Downloaded file: ${blob.getName()}, size: ${blob.getBytes().length} bytes`);

    // Convert to base64 and process
    const base64Data = Utilities.base64Encode(blob.getBytes());
    const result = processImportedCategoryFile(blob.getName(), base64Data);

    const duration = Date.now() - startTime;

    if (!result.success) {
      console.log('✗ ' + name + ': ' + result.message);
      return { name, passed: false, message: result.message, duration };
    }

    console.log('✓ ' + name + ' (' + duration + 'ms)');
    return {
      name,
      passed: true,
      message: 'OK',
      duration,
      details: result.details,
    };
  } catch (e) {
    const duration = Date.now() - startTime;
    console.log('✗ ' + name + ': ' + e);
    return { name, passed: false, message: String(e), duration };
  }
}

/**
 * Tests the import of a CoMapeo v5 configuration file
 */
function testImportCoMapeoV5(): ImportIntegrationTestResult {
  const name = 'Import CoMapeo v5.0.0 config (build.zip)';
  const startTime = Date.now();

  try {
    console.log('Fetching CoMapeo v5 config from GitHub...');
    const blob = fetchTestFile(TEST_CONFIG_URLS.coMapeoV5);
    console.log(`Downloaded file: ${blob.getName()}, size: ${blob.getBytes().length} bytes`);

    // Convert to base64 and process
    const base64Data = Utilities.base64Encode(blob.getBytes());
    const result = processImportedCategoryFile(blob.getName(), base64Data);

    const duration = Date.now() - startTime;

    if (!result.success) {
      console.log('✗ ' + name + ': ' + result.message);
      return { name, passed: false, message: result.message, duration };
    }

    console.log('✓ ' + name + ' (' + duration + 'ms)');
    return {
      name,
      passed: true,
      message: 'OK',
      duration,
      details: result.details,
    };
  } catch (e) {
    const duration = Date.now() - startTime;
    console.log('✗ ' + name + ': ' + e);
    return { name, passed: false, message: String(e), duration };
  }
}

/**
 * Tests the import of a legacy Mapeo v3.6.1 .mapeosettings file
 */
function testImportMapeoV3(): ImportIntegrationTestResult {
  const name = 'Import Mapeo v3.6.1 (.mapeosettings)';
  const startTime = Date.now();

  try {
    console.log('Fetching Mapeo v3.6.1 config from GitHub...');
    const blob = fetchTestFile(TEST_CONFIG_URLS.mapeoV3);
    console.log(`Downloaded file: ${blob.getName()}, size: ${blob.getBytes().length} bytes`);

    // Convert to base64 and process
    const base64Data = Utilities.base64Encode(blob.getBytes());
    const result = processImportedCategoryFile(blob.getName(), base64Data);

    const duration = Date.now() - startTime;

    if (!result.success) {
      console.log('✗ ' + name + ': ' + result.message);
      return { name, passed: false, message: result.message, duration };
    }

    console.log('✓ ' + name + ' (' + duration + 'ms)');
    return {
      name,
      passed: true,
      message: 'OK',
      duration,
      details: result.details,
    };
  } catch (e) {
    const duration = Date.now() - startTime;
    console.log('✗ ' + name + ': ' + e);
    return { name, passed: false, message: String(e), duration };
  }
}

/**
 * Runs all import integration tests
 */
function runImportIntegrationTests(): ImportIntegrationTestResult[] {
  const results: ImportIntegrationTestResult[] = [];

  console.log('===== Import Integration Tests =====');
  console.log('Testing with real configuration files from GitHub releases');
  console.log('');

  // Test each config file
  results.push(testImportCoMapeoV6());
  results.push(testImportCoMapeoV5());
  results.push(testImportMapeoV3());

  // Summary
  const passed = results.filter(r => r.passed).length;
  const failed = results.filter(r => !r.passed).length;
  const totalDuration = results.reduce((sum, r) => sum + (r.duration || 0), 0);

  console.log('');
  console.log(`===== Results: ${passed} passed, ${failed} failed =====`);
  console.log(`Total time: ${(totalDuration / 1000).toFixed(2)} seconds`);

  return results;
}

/**
 * Menu entry point for running import integration tests
 */
function runImportIntegrationTestsFromMenu(): void {
  const results = runImportIntegrationTests();
  const passed = results.filter(r => r.passed).length;
  const failed = results.filter(r => !r.passed).length;

  let message = `Import Integration Tests\n\n`;
  message += `Passed: ${passed}\nFailed: ${failed}\n\n`;

  if (failed > 0) {
    message += 'Failed tests:\n';
    for (const result of results) {
      if (!result.passed) {
        message += `- ${result.name}:\n  ${result.message}\n`;
      }
    }
  } else {
    message += 'All tests passed!\n\n';
    for (const result of results) {
      message += `✓ ${result.name} (${result.duration}ms)\n`;
      if (result.details) {
        message += `  Categories: ${result.details.presets || 0}, Fields: ${result.details.fields || 0}\n`;
      }
    }
  }

  SpreadsheetApp.getUi().alert('Integration Test Results', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Quick smoke test with just CoMapeo v6
 */
function runQuickImportTest(): void {
  console.log('===== Quick Import Test (CoMapeo v6 only) =====');
  const result = testImportCoMapeoV6();

  const message = result.passed
    ? `✓ Import test passed!\n\nDuration: ${result.duration}ms\nCategories: ${result.details?.presets || 0}\nFields: ${result.details?.fields || 0}`
    : `✗ Import test failed!\n\n${result.message}`;

  SpreadsheetApp.getUi().alert('Quick Import Test', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Diagnostic function to test processImportedCategoryFile directly.
 * Run this from Apps Script editor to verify the import pipeline works.
 */
function diagnoseImportPipeline(): void {
  const ui = SpreadsheetApp.getUi();

  console.log('=== Import Pipeline Diagnostic ===');
  console.log('Step 1: Fetching test file from GitHub...');

  try {
    const url = TEST_CONFIG_URLS.coMapeoV6;
    console.log(`Fetching: ${url}`);

    const response = UrlFetchApp.fetch(url, {
      followRedirects: true,
      muteHttpExceptions: false,
    });

    const blob = response.getBlob().setName('build.zip');
    const fileSize = blob.getBytes().length;
    console.log(`Step 2: File fetched. Size: ${fileSize} bytes (${(fileSize / 1024).toFixed(1)} KB)`);

    // Convert to base64
    console.log('Step 3: Converting to base64...');
    const base64Data = Utilities.base64Encode(blob.getBytes());
    console.log(`Base64 length: ${base64Data.length} characters`);

    // Call the import function
    console.log('Step 4: Calling processImportedCategoryFile...');
    const startTime = Date.now();
    const result = processImportedCategoryFile('build.zip', base64Data);
    const duration = Date.now() - startTime;

    console.log(`Step 5: Import completed in ${duration}ms`);
    console.log(`Result: ${JSON.stringify(result, null, 2)}`);

    if (result.success) {
      ui.alert(
        'Import Diagnostic: SUCCESS',
        `The import pipeline works correctly!\n\n` +
        `Duration: ${duration}ms\n` +
        `Categories: ${result.details?.presets || 0}\n` +
        `Fields: ${result.details?.fields || 0}\n` +
        `Icons: ${result.details?.icons || 0}\n` +
        `Languages: ${result.details?.languages || 0}`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        'Import Diagnostic: FAILED',
        `The import pipeline failed:\n\n${result.message}`,
        ui.ButtonSet.OK
      );
    }
  } catch (error) {
    console.error('Diagnostic error:', error);
    ui.alert(
      'Import Diagnostic: ERROR',
      `An error occurred:\n\n${error instanceof Error ? error.message : String(error)}\n\nCheck the Apps Script logs for details.`,
      ui.ButtonSet.OK
    );
  }
}
