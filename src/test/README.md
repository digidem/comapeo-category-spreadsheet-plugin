# Test Suite Documentation

## Overview

This directory contains the test suite for the CoMapeo Category Set Spreadsheet Plugin. The tests verify functionality across icon generation, translation extraction, import/export workflows, and end-to-end integration.

## Quick Start

### Running Tests

Tests must be run from the **Apps Script Editor** (not locally):

1. Open your Google Spreadsheet
2. Go to `Extensions > Apps Script`
3. In the Apps Script editor, select the test function you want to run from the dropdown
4. Click the "Run" button (▶)
5. Check the **Execution log** (View > Logs or Ctrl/Cmd+Enter) for results

### Available Test Functions

| Function | Purpose | Runtime | Prerequisites |
|----------|---------|---------|---------------|
| `testEndToEnd()` | Complete end-to-end suite | ~5-10 min | Empty or test spreadsheet |
| `testImportCategory()` | Import .comapeocat/.mapeosettings | ~2-3 min | Test file URL |
| `testExtractAndValidate()` | Extract and validate config | ~2-3 min | Test file URL |
| `testFieldExtraction()` | Field extraction only | ~1-2 min | Test file URL |
| `testTranslationExtraction()` | Translation extraction only | ~1-2 min | Test file URL |
| `testFormatDetection()` | File format detection | ~30 sec | Test file URL |
| `testDetailsAndIcons()` | Details/icons processing | ~2-3 min | Test file URL |
| `testSkipTranslation()` | Skip translation workflow | ~1-2 min | Configured spreadsheet |
| `testDebugLogger()` | Debug logger functionality | ~10 sec | None |
| `runLintParityTests()` | Lint/builder parser parity checks | ~10 sec | None |

## Test Helpers (`testHelpers.ts`)

**NEW**: Test helper functions for setup, teardown, and state management.

### Key Functions

#### `testSetup()` / `testTeardown(backup)`
Automatically backup and restore spreadsheet state:

```typescript
function myTest() {
  const backup = testSetup();
  try {
    // ... run test ...
  } finally {
    testTeardown(backup);
  }
}
```

#### `withTestSetup(testFn, testName)`
Convenience wrapper that handles setup/teardown automatically:

```typescript
const result = withTestSetup(() => {
  // ... test code ...
  return { success: true, message: "Test passed" };
}, "My Feature Test");
```

#### `backupSpreadsheet()` / `restoreSpreadsheet(backup)`
Manual backup/restore for fine-grained control:

```typescript
const backup = backupSpreadsheet();
// ... modify spreadsheet ...
restoreSpreadsheet(backup);
```

#### `createTestSpreadsheet(name)` / `deleteTestSpreadsheet(id)`
Create isolated test spreadsheets:

```typescript
const testSheet = createTestSpreadsheet("MyTest");
// ... test on testSheet ...
deleteTestSpreadsheet(testSheet.getId());
```

#### `cleanupTestArtifacts(olderThanMinutes)`
Clean up temporary test files and folders:

```typescript
const result = cleanupTestArtifacts(60); // Clean artifacts older than 1 hour
console.log(`Cleaned ${result.itemsCleaned} items`);
```

### Assertion Functions

```typescript
assert(condition, "Error message");
assertEqual(actual, expected, "Optional message");
assertTruthy(value, "Should be truthy");
assertFalsy(value, "Should be falsy");
assertThrows(() => { throw new Error(); }, "Should throw");
```

### Usage Example

```typescript
function testImportWithCleanup() {
  return withTestSetup(() => {
    // Test runs with automatic backup/restore
    const result = importCategoryFile(testUrl);

    assert(result.success, "Import should succeed");
    assertEqual(result.categoriesCount, 5, "Should import 5 categories");

    return {
      success: true,
      message: "Import test passed",
    };
  }, "Import Category with Cleanup");
}
```

## Test Suite Structure

### 1. End-to-End Tests (`testEndToEnd.ts`)

**Purpose**: Comprehensive validation of the entire plugin workflow

**What It Tests**:
- ✅ Icons persistence after import
- ✅ Translation sheet population
- ✅ Dropdown functionality in Details column
- ✅ Complete round-trip (export → import → compare)

**How to Run**:
```javascript
testEndToEnd()  // Uses default test file
// OR
testEndToEnd("https://example.com/your-test-file.mapeosettings")
```

**Expected Results**:
```
=== Starting End-to-End Test Suite ===
--- Test 1: Icons Test ---
✅ Icons Test PASSED

--- Test 2: Translations Test ---
✅ Translations Test PASSED

--- Test 3: Dropdown Test ---
✅ Dropdown Test PASSED

--- Test 4: Round-Trip Test ---
✅ Round-Trip Test PASSED

=== All Tests PASSED ===
```

**What to Verify**:
1. **Icons**: Check Categories sheet column B for icon URLs
2. **Translations**: Verify translation sheets exist with populated data
3. **Dropdowns**: Open Details sheet, click cells to see dropdown lists
4. **Round-trip**: Compare exported config with original

### 2. Import Tests (`testImportCategory.ts`)

**Purpose**: Validate import functionality for .comapeocat and .mapeosettings files

**What It Tests**:
- ✅ File format detection (TAR, ZIP, JSON)
- ✅ Category data extraction
- ✅ Field data extraction
- ✅ Icon extraction from sprite
- ✅ Translation data extraction
- ✅ Spreadsheet population

**How to Run**:
```javascript
testImportCategory()  // Uses default test file
// OR
testImportCategory("https://example.com/custom-file.comapeocat")
```

**Expected Results**:
- Categories sheet populated with data
- Details sheet populated with fields
- Translation sheets created and populated
- Icons extracted to temp folder
- Console shows: "✅ Import test completed successfully"

### 3. Extraction and Validation Tests (`testExtractAndValidate.ts`)

**Purpose**: Test extraction pipeline without modifying spreadsheet

**What It Tests**:
- ✅ TAR/ZIP file extraction
- ✅ JSON parsing
- ✅ Icon sprite parsing
- ✅ Data validation
- ✅ Structure verification

**How to Run**:
```javascript
testExtractAndValidate()
// OR
testExtractAndValidate("https://example.com/test-file.mapeosettings")
```

**Expected Results**:
```
Starting extraction and validation test...
✅ File extracted successfully
✅ Icons parsed: 25 icons found
✅ Config data validated
✅ All validation checks passed
```

### 4. Field Extraction Tests (`testFieldExtraction.ts`)

**Purpose**: Validate field extraction from config files

**What It Tests**:
- ✅ Field types (text, number, selectOne, selectMultiple)
- ✅ Field options parsing
- ✅ Helper text extraction
- ✅ Universal field flags

**How to Run**:
```javascript
testFieldExtraction()
```

**Expected Results**:
- Console shows field count
- Field types correctly identified
- Options parsed for select fields
- No validation errors

### 5. Translation Extraction Tests (`testTranslationExtraction.ts`)

**Purpose**: Test translation data extraction and organization

**What It Tests**:
- ✅ Category name translations
- ✅ Field label translations
- ✅ Helper text translations
- ✅ Option translations
- ✅ Language code detection

**How to Run**:
```javascript
testTranslationExtraction()
```

**Expected Results**:
- Translation data extracted by language
- All translation types present
- No missing translations logged

### 6. Format Detection Tests (`testFormatDetection.ts`)

**Purpose**: Validate file format detection logic

**What It Tests**:
- ✅ TAR file detection
- ✅ ZIP file detection
- ✅ JSON file detection
- ✅ Content-based detection (magic numbers)

**How to Run**:
```javascript
testFormatDetection()
```

**Expected Results**:
```
Testing format detection...
✅ TAR format detected correctly
✅ ZIP format detected correctly
✅ JSON format detected correctly
```

### 7. Details and Icons Tests (`testDetailsAndIcons.ts`)

**Purpose**: Test Details sheet and icon processing

**What It Tests**:
- ✅ Details sheet population
- ✅ Icon URL generation
- ✅ Field type validation
- ✅ Option parsing

**How to Run**:
```javascript
testDetailsAndIcons()
```

**Expected Results**:
- Details sheet populated correctly
- Icon URLs valid
- Field types match expected

### 8. Skip Translation Tests (`testSkipTranslation.ts`)

**Purpose**: Test the "skip translation" workflow

**What It Tests**:
- ✅ Config generation without translation
- ✅ Primary language handling
- ✅ Translation sheets not modified

**How to Run**:
```javascript
testSkipTranslation()
```

**Expected Results**:
```
Testing skip translation workflow...
✅ Config generated without translation
✅ Primary language used only
✅ Translation sheets unchanged
```

### 9. Debug Logger Tests (`testDebugLogger.ts`)

**Purpose**: Validate debug logging functionality

**What It Tests**:
- ✅ Logger initialization
- ✅ Log level filtering
- ✅ Structured logging

**How to Run**:
```javascript
testDebugLogger()
```

**Expected Results**:
- Logs output to console
- All log levels work
- No errors

## Test Data Files

The `src/test/` directory includes test data files:

| File | Purpose | Size |
|------|---------|------|
| `mapeo-default-min.mapeosettings` | Default test configuration | 52 KB |
| `generated-config.comapeocat` | Generated test output | 698 KB |
| `config.json` | Config data structure | 61 KB |
| `fields.json` | Field definitions | 25 KB |
| `presets.json` | Preset data | 6.6 KB |
| `icons.json` | Icon definitions | 2.4 KB |
| `translations.json` | Translation data | 20 KB |
| `metadata.json` | Metadata | 203 bytes |
| `icons/` | Directory of test icons | 157 files |

## Troubleshooting

### Common Issues

#### 1. "Authorization Required" Error

**Problem**: Tests fail with authorization error

**Solution**:
1. Run the test function once
2. Click "Review Permissions" when prompted
3. Select your Google account
4. Click "Advanced" then "Go to [Project Name] (unsafe)"
5. Click "Allow"
6. Run the test again

#### 2. "Exceeded maximum execution time" Error

**Problem**: Test runs longer than 6 minutes (Apps Script limit)

**Solution**:
- Break tests into smaller chunks
- Run individual test functions instead of full suite
- Reduce test data size
- Add manual checkpoints with `Utilities.sleep()` to avoid timeout

#### 3. "Service invoked too many times" Error

**Problem**: Too many API calls in short time

**Solution**:
- Add delays between API calls: `Utilities.sleep(1000)`
- Reduce test iterations
- Use cached test data when possible

#### 4. Test Passes But Results Look Wrong

**Problem**: Test reports success but spreadsheet looks incorrect

**Solution**:
1. Check the console logs for warnings
2. Manually verify spreadsheet data
3. Compare with expected test data files
4. Run `lintAllSheets()` to validate data

#### 5. "Cannot read property of undefined" Error

**Problem**: Test tries to access missing data

**Solution**:
- Check test file URL is valid
- Verify test file is accessible (not behind authentication)
- Check Apps Script has Drive API access enabled

### Debugging Tips

1. **Enable Debug Logging**:
   ```javascript
   Logger.setLevel(LogLevel.DEBUG);
   ```

2. **Check Execution Log**:
   - Go to Apps Script Editor
   - View > Logs (or Ctrl/Cmd+Enter)
   - Look for error messages and stack traces

3. **Use Console.log**:
   ```javascript
   console.log("Variable value:", variable);
   ```

4. **Breakpoints** (not supported in Apps Script):
   - Use multiple `console.log()` statements
   - Add `return` statements to stop execution at specific points

5. **Test Incrementally**:
   - Comment out parts of test
   - Run smaller sections
   - Isolate the failing part

## Best Practices

### Before Running Tests

1. **Backup Your Spreadsheet**: Tests may modify data
2. **Use Test Spreadsheet**: Create a copy for testing
3. **Check Permissions**: Ensure Apps Script has necessary permissions
4. **Review Test File**: Verify test file URL is correct

### During Testing

1. **Monitor Console**: Watch execution log for errors
2. **Check Memory**: Large tests may use significant memory
3. **Verify Results**: Don't rely on test status alone - check data

### After Testing

1. **Review Logs**: Check for warnings even if tests pass
2. **Validate Data**: Run `lintAllSheets()` to verify data integrity
3. **Clean Up**: Delete temp folders created during testing
4. **Document Issues**: Note any unexpected behavior

## Test Coverage

### What's Tested

| Feature | Coverage | Test File |
|---------|----------|-----------|
| Import .comapeocat | ✅ 95% | testImportCategory.ts |
| Import .mapeosettings | ✅ 95% | testImportCategory.ts |
| Field extraction | ✅ 90% | testFieldExtraction.ts |
| Translation extraction | ✅ 85% | testTranslationExtraction.ts |
| Icon processing | ✅ 90% | testDetailsAndIcons.ts |
| Format detection | ✅ 100% | testFormatDetection.ts |
| Skip translation | ✅ 85% | testSkipTranslation.ts |
| End-to-end workflow | ✅ 80% | testEndToEnd.ts |

### What's NOT Tested

❌ **Unit tests**: No isolated unit tests (only integration/E2E)
❌ **API mocking**: Tests use real API calls
❌ **Edge cases**: Limited edge case coverage
❌ **Performance**: No performance benchmarking
❌ **Concurrent access**: No multi-user testing

## Contributing

### Adding New Tests

1. **Create Test File**:
   ```typescript
   // src/test/testNewFeature.ts
   function testNewFeature(): TestResult {
     // Test implementation
   }
   ```

2. **Follow Naming Convention**:
   - File: `test<FeatureName>.ts`
   - Function: `test<FeatureName>()`
   - Use camelCase

3. **Return Test Results**:
   ```typescript
   return {
     success: true,
     message: "Test passed",
     details: ["Detail 1", "Detail 2"],
   };
   ```

4. **Document in README**:
   - Add to test table
   - Describe what it tests
   - Provide expected results

### Test Structure

```typescript
function testFeature(): TestResult {
  try {
    console.log("Starting test...");

    // Arrange: Set up test data
    const testData = prepareTestData();

    // Act: Execute the feature
    const result = executeFeature(testData);

    // Assert: Verify results
    if (result !== expected) {
      return {
        success: false,
        message: "Test failed",
        errors: [`Expected ${expected}, got ${result}`],
      };
    }

    console.log("✅ Test passed");
    return {
      success: true,
      message: "Test passed",
      details: ["All checks completed"],
    };
  } catch (error) {
    console.error("❌ Test failed:", error);
    return {
      success: false,
      message: "Test failed with error",
      errors: [error.message],
    };
  }
}
```

## Additional Resources

- **[Coding Assistant Guide](../../context/process/assistant-guide.md)**: Project development guide
- **[import-cat.md](../../context/reference/import-cat.md)**: Import feature documentation
- **[architecture.md](../../context/reference/architecture.md)**: System architecture
- **[User docs](../../docs/index.md)**: End-user documentation

## Support

For issues with tests:
1. Check this README troubleshooting section
2. Review console logs for error messages
3. Check [GitHub Issues](https://github.com/digidem/comapeo-config-spreadsheet-plugin/issues)
4. Contact development team

---

**Last Updated**: 2025-10-12
**Maintained By**: Digital Democracy Development Team
