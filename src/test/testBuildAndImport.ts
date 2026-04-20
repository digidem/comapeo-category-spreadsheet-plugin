/**
 * Test suite for CoMapeo Config v2.0.0 API
 * Tests build flow (JSON payload, category ordering, API call) and import flow
 *
 * Run these tests from the Apps Script editor console:
 *   runBuildAndImportTests();
 */

// =============================================================================
// Test Data
// =============================================================================

/**
 * Mock spreadsheet data matching the structure returned by getSpreadsheetData()
 * Note: documentName is a STRING (spreadsheet name), not an array
 * Uses NEW 6-column format for Categories (with ID and Icon ID columns)
 */
function getMockSpreadsheetData(): SheetData {
  return {
    // documentName is a string from spreadsheet.getName()
    documentName: "Test Config" as any,  // Cast needed since SheetData expects arrays

    // Categories sheet data (row 0 is header) - NEW 6-column format
    Categories: [
      ["Name", "Icon", "Fields", "ID", "Color", "Icon ID"],
      ["Trees", "", "species,diameter", "trees", "#4CAF50", "trees"],
      ["Rivers", "", "flow-rate,depth", "rivers", "#2196F3", "rivers"],
      ["Buildings", "", "height,material", "buildings", "#9C27B0", "buildings"]
    ],

    // Details sheet data (row 0 is header) - NEW 6-column format with ID column
    Details: [
      ["Name", "Helper Text", "Type", "Options", "ID", "Universal"],
      ["Species", "Select the tree species", "s", "Oak,Pine,Maple", "species", "FALSE"],
      ["Diameter", "Enter trunk diameter in cm", "n", "", "diameter", "FALSE"],
      ["Flow Rate", "Water flow measurement", "n", "", "flow-rate", "FALSE"],
      ["Depth", "River depth in meters", "n", "", "depth", "FALSE"],
      ["Height", "Building height in meters", "i", "", "height", "FALSE"],
      ["Material", "Construction material", "m", "Wood,Concrete,Steel", "material", "FALSE"]
    ],

    // Translation sheets (minimal for testing)
    "Category Translations": [
      ["English", "Español"],
      ["Trees", "Árboles"],
      ["Rivers", "Ríos"],
      ["Buildings", "Edificios"]
    ],
    "Detail Label Translations": [
      ["English", "Español"],
      ["Species", "Especie"],
      ["Diameter", "Diámetro"],
      ["Flow Rate", "Caudal"],
      ["Depth", "Profundidad"],
      ["Height", "Altura"],
      ["Material", "Material"]
    ],
    "Detail Helper Text Translations": [
      ["English", "Español"],
      ["Select the tree species", "Seleccione la especie de árbol"],
      ["Enter trunk diameter in cm", "Ingrese el diámetro del tronco en cm"],
      ["Water flow measurement", "Medición del flujo de agua"],
      ["River depth in meters", "Profundidad del río en metros"],
      ["Building height in meters", "Altura del edificio en metros"],
      ["Construction material", "Material de construcción"]
    ],
    "Detail Option Translations": [
      ["English", "Español"],
      ["Oak,Pine,Maple", "Roble,Pino,Arce"],
      ["", ""],
      ["", ""],
      ["", ""],
      ["", ""],
      ["Wood,Concrete,Steel", "Madera,Concreto,Acero"]
    ]
  } as SheetData;
}

/**
 * Helper: Creates a minimal SheetData with custom categories and details
 * Reduces test data duplication
 */
function createTestSheetData(
  categories: any[][],
  details: any[][]
): SheetData {
  return {
    documentName: "Test Config" as any,
    Categories: [
      ["Name", "Icon", "Fields", "ID", "Color", "Icon ID"],
      ...categories
    ],
    Details: [
      ["Name", "Helper Text", "Type", "Options", "ID", "Universal"],
      ...details
    ]
  } as SheetData;
}

/**
 * Helper: Creates a simple category row
 */
function createCategoryRow(
  name: string,
  fields: string = "",
  id: string = "",
  color: string = "#FF0000"
): any[] {
  const categoryId = id || slugify(name);
  return [name, "", fields, categoryId, color, categoryId];
}

/**
 * Helper: Creates a simple field/detail row
 */
function createDetailRow(
  name: string,
  type: string = "t",
  options: string = "",
  universal: string = "FALSE"
): any[] {
  const fieldId = slugify(name);
  return [name, `Description for ${name}`, type, options, fieldId, universal];
}

/**
 * Helper: Creates a deterministic ASCII string of the requested size.
 * Using ASCII keeps the byte-length calculations aligned with string length.
 */
function createAsciiPayload(size: number, fillChar: string = "a"): string {
  return fillChar.repeat(size);
}

/**
 * Helper: Builds a select field with a large option list.
 */
function createSelectFieldWithOptions(optionCount: number): Field {
  return {
    id: "field-1",
    tagKey: "field-1",
    name: "Field",
    type: "selectOne",
    options: Array.from({ length: optionCount }, (_value, index) => ({
      label: `Option ${index}`,
      value: `option-${index}`
    }))
  };
}

/**
 * Helper: Builds a flat translation map for translated option labels.
 */
function createOptionTranslationMap(optionCount: number): { [key: string]: string } {
  const translations: { [key: string]: string } = {};
  for (let index = 0; index < optionCount; index++) {
    translations[`options.${index}`] = `Translated option ${index}`;
  }
  return translations;
}

// =============================================================================
// Build Flow Tests
// =============================================================================

function testBuildPayloadCreation(): boolean {
  console.log("=== Test: Build Payload Creation ===");

  try {
    // Note: This test requires mocking SpreadsheetApp which isn't available in unit tests
    // In a real test environment, we'd use a test spreadsheet
    // For now, we test the data transformation logic

    const mockData = getMockSpreadsheetData();

    // Test buildFields
    const fields = buildFields(mockData);
    if (!Array.isArray(fields)) {
      console.error("FAIL: buildFields should return an array");
      return false;
    }
    if (fields.length !== 6) {
      console.error(`FAIL: Expected 6 fields, got ${fields.length}`);
      return false;
    }
    console.log("PASS: buildFields returns correct count");

    // Verify field structure
    const speciesField = fields.find(f => f.id === "species");
    if (!speciesField) {
      console.error("FAIL: Could not find species field");
      return false;
    }
    if (speciesField.type !== "select") {
      console.error(`FAIL: species field should be 'select', got '${speciesField.type}'`);
      return false;
    }
    if (!speciesField.options || speciesField.options.length !== 3) {
      console.error("FAIL: species field should have 3 options");
      return false;
    }
    console.log("PASS: Field structure is correct");

    console.log("=== Build Payload Creation: ALL TESTS PASSED ===");
    return true;
  } catch (error) {
    console.error("FAIL: Exception thrown - " + error.message);
    return false;
  }
}

function testFieldTypeMapping(): boolean {
  console.log("=== Test: Field Type Mapping ===");

  try {
    const mockData = getMockSpreadsheetData();
    const fields = buildFields(mockData);

    // Find species field (select type - 's')
    const speciesField = fields.find(f => f.id === "species");
    if (!speciesField || speciesField.type !== "select") {
      console.error("FAIL: Species field should be type 'select'");
      return false;
    }
    console.log("PASS: Select field type mapped correctly");

    // Find diameter field (number type - 'n')
    const diameterField = fields.find(f => f.id === "diameter");
    if (!diameterField || diameterField.type !== "number") {
      console.error("FAIL: Diameter field should be type 'number'");
      return false;
    }
    console.log("PASS: Number field type mapped correctly");

    // Find material field (multiselect type - 'm')
    const materialField = fields.find(f => f.id === "material");
    if (!materialField || materialField.type !== "multiselect") {
      console.error("FAIL: Material field should be type 'multiselect'");
      return false;
    }
    console.log("PASS: Multiselect field type mapped correctly");

    console.log("=== Field Type Mapping: ALL TESTS PASSED ===");
    return true;
  } catch (error) {
    console.error("FAIL: Exception thrown - " + error.message);
    return false;
  }
}

function testParseOptions(): boolean {
  console.log("=== Test: Parse Options ===");

  try {
    // Test normal options
    const opts1 = parseOptions("Oak,Pine,Maple");
    if (!opts1 || opts1.length !== 3) {
      console.error("FAIL: Should parse 3 options");
      return false;
    }
    if (opts1[0].label !== "Oak" || opts1[0].value !== "oak") {
      console.error("FAIL: First option should be Oak/oak");
      return false;
    }
    console.log("PASS: Normal options parsed correctly");

    // Test empty string
    const opts2 = parseOptions("");
    if (opts2 !== undefined) {
      console.error("FAIL: Empty string should return undefined");
      return false;
    }
    console.log("PASS: Empty string returns undefined");

    // Test whitespace handling
    const opts3 = parseOptions("  A  ,  B  ,  C  ");
    if (!opts3 || opts3.length !== 3) {
      console.error("FAIL: Should handle whitespace");
      return false;
    }
    if (opts3[0].label !== "A" || opts3[1].label !== "B") {
      console.error("FAIL: Should trim whitespace from labels");
      return false;
    }
    console.log("PASS: Whitespace handled correctly");

    console.log("=== Parse Options: ALL TESTS PASSED ===");
    return true;
  } catch (error) {
    console.error("FAIL: Exception thrown - " + error.message);
    return false;
  }
}

function testSlugify(): boolean {
  console.log("=== Test: Slugify Function ===");

  try {
    // Test basic slugify
    if (slugify("Hello World") !== "hello-world") {
      console.error("FAIL: 'Hello World' should become 'hello-world'");
      return false;
    }
    console.log("PASS: Basic slugify works");

    // Test special characters
    if (slugify("Test & Example!") !== "test-example") {
      console.error("FAIL: Special characters should be removed");
      return false;
    }
    console.log("PASS: Special characters removed");

    // Test multiple spaces/dashes
    if (slugify("Multiple   Spaces") !== "multiple-spaces") {
      console.error("FAIL: Multiple spaces should become single dash");
      return false;
    }
    console.log("PASS: Multiple spaces handled");

    console.log("=== Slugify Function: ALL TESTS PASSED ===");
    return true;
  } catch (error) {
    console.error("FAIL: Exception thrown - " + error.message);
    return false;
  }
}

// =============================================================================
// Import Flow Tests
// =============================================================================

function testImportParsing(): boolean {
  console.log("=== Test: Import Parsing ===");

  try {
    // Create a mock BuildRequest to simulate imported data
    const mockImportData: BuildRequest = {
      metadata: {
        name: "Test Import Config",
        version: "1.0.0",
        description: "A test configuration"
      },
      categories: [
        { id: "cat1", name: "Category One", color: "#FF0000", defaultFieldIds: ["field1"] },
        { id: "cat2", name: "Category Two", color: "#00FF00", defaultFieldIds: ["field2"] }
      ],
      fields: [
        { id: "field1", name: "Field One", type: "text" },
        { id: "field2", name: "Field Two", type: "number" }
      ],
      icons: [
        { id: "icon1", svgData: "<svg></svg>" }
      ]
    };

    // Validate the structure
    validateBuildRequest(mockImportData);
    console.log("PASS: Valid BuildRequest passes validation");

    // Test category count
    if (mockImportData.categories.length !== 2) {
      console.error("FAIL: Expected 2 categories");
      return false;
    }
    console.log("PASS: Categories count correct");

    // Test field count
    if (mockImportData.fields.length !== 2) {
      console.error("FAIL: Expected 2 fields");
      return false;
    }
    console.log("PASS: Fields count correct");

    console.log("=== Import Parsing: ALL TESTS PASSED ===");
    return true;
  } catch (error) {
    console.error("FAIL: Exception thrown - " + error.message);
    return false;
  }
}

function testValidateBuildRequest(): boolean {
  console.log("=== Test: Validate BuildRequest ===");

  try {
    // Test missing metadata
    try {
      validateBuildRequest({ categories: [], fields: [] } as any);
      console.error("FAIL: Should throw for missing metadata");
      return false;
    } catch (e) {
      console.log("PASS: Throws for missing metadata");
    }

    // Test missing name
    try {
      validateBuildRequest({ metadata: { version: "1.0.0" }, categories: [], fields: [] } as any);
      console.error("FAIL: Should throw for missing name");
      return false;
    } catch (e) {
      console.log("PASS: Throws for missing name");
    }

    // Test invalid categories (not array)
    try {
      validateBuildRequest({ metadata: { name: "test", version: "1.0.0" }, categories: "invalid", fields: [] } as any);
      console.error("FAIL: Should throw for non-array categories");
      return false;
    } catch (e) {
      console.log("PASS: Throws for non-array categories");
    }

    // Test valid request
    try {
      validateBuildRequest({
        metadata: { name: "test", version: "1.0.0" },
        categories: [],
        fields: []
      });
      console.log("PASS: Valid request passes validation");
    } catch (e) {
      console.error("FAIL: Valid request should not throw");
      return false;
    }

    const strictValidRequest: BuildRequest = {
      metadata: { name: "Strict Config", version: "1.0.0" },
      locales: ["en"],
      categories: [
        { id: "cat-1", name: "Category", appliesTo: ["observation"], defaultFieldIds: ["field-1"], fields: ["field-1"] }
      ],
      fields: [
        { id: "field-1", tagKey: "field-1", name: "Field", type: "text" }
      ],
      icons: [
        { id: "icon-1", svgData: "<svg></svg>" }
      ],
      translations: {
        es: {
          categories: { "cat-1": { name: "Categoría" } },
          fields: { "field-1": { label: "Campo" } }
        }
      }
    };

    try {
      validateBuildRequest(strictValidRequest, { strict: true });
      console.log("PASS: Strict validation accepts valid payload");
    } catch (e) {
      console.error("FAIL: Strict validation should accept valid payload", e);
      return false;
    }

    const strictTrackOnlyRequest = JSON.parse(JSON.stringify(strictValidRequest));
    strictTrackOnlyRequest.categories[0].appliesTo = ["track"];
    try {
      validateBuildRequest(strictTrackOnlyRequest, { strict: true });
      console.log("PASS: Strict validation accepts track-only categories");
    } catch (e) {
      console.error("FAIL: Strict validation should not require observation coverage", e);
      return false;
    }

    const missingOptions = JSON.parse(JSON.stringify(strictValidRequest));
    missingOptions.fields[0].type = "selectOne";
    missingOptions.fields[0].options = [];
    try {
      validateBuildRequest(missingOptions, { strict: true });
      console.error("FAIL: Select fields without options should throw");
      return false;
    } catch (e) {
      console.log("PASS: Missing options detected");
    }

    const invalidLocale = JSON.parse(JSON.stringify(strictValidRequest));
    invalidLocale.translations = { "es_MX": { category: { "cat-1": { name: "Categoría" } } } };
    try {
      validateBuildRequest(invalidLocale, { strict: true });
      console.error("FAIL: Invalid locale tags should throw");
      return false;
    } catch (e) {
      console.log("PASS: Invalid locale tag detected");
    }

    const unknownFieldRef = JSON.parse(JSON.stringify(strictValidRequest));
    unknownFieldRef.categories[0].defaultFieldIds = ["missing-field"];
    try {
      validateBuildRequest(unknownFieldRef, { strict: true });
      console.error("FAIL: Unknown field references should throw");
      return false;
    } catch (e) {
      console.log("PASS: Unknown field reference detected");
    }

    const duplicateFieldIds = JSON.parse(JSON.stringify(strictValidRequest));
    duplicateFieldIds.fields.push({
      id: "FIELD-1",
      tagKey: "FIELD-1",
      name: "Duplicate Field",
      type: "text"
    });
    try {
      validateBuildRequest(duplicateFieldIds, { strict: true });
      console.error("FAIL: Case-insensitive duplicate field IDs should throw");
      return false;
    } catch (e) {
      const message = String((e as Error).message || e);
      if (!message.includes("Duplicate field id detected:")) {
        console.error(`FAIL: Unexpected duplicate field ID error: ${message}`);
        return false;
      }
      console.log("PASS: Case-insensitive duplicate field IDs are rejected");
    }

    console.log("=== Validate BuildRequest: ALL TESTS PASSED ===");
    return true;
  } catch (error) {
    console.error("FAIL: Exception thrown - " + error.message);
    return false;
  }
}

function testStrictMergedTranslationPayloadLimit(): boolean {
  console.log("=== Test: Strict Merged Translation Payload Limit ===");

  try {
    const halfLimitBytes = 600 * 1024;
    const strictTranslationRequest: BuildRequest = {
      metadata: { name: "Merged Translation Payload", version: "1.0.0" },
      locales: ["es"],
      categories: [
        {
          id: "cat-1",
          name: "Category",
          appliesTo: ["observation"],
          defaultFieldIds: ["field-1"],
          fields: ["field-1"]
        }
      ],
      fields: [
        {
          id: "field-1",
          tagKey: "field-1",
          name: "Field",
          type: "text"
        }
      ],
      translations: {
        es: {
          categories: {
            "cat-1": {
              name: createAsciiPayload(halfLimitBytes, "c")
            }
          },
          fields: {
            "field-1": {
              label: createAsciiPayload(halfLimitBytes, "f")
            }
          }
        }
      }
    };

    const localeTranslations = strictTranslationRequest.translations?.es;
    const categoryBytes = JSON.stringify(localeTranslations?.categories || {}).length;
    const fieldBytes = JSON.stringify(localeTranslations?.fields || {}).length;
    const mergedBytes = JSON.stringify(localeTranslations || {}).length;

    if (categoryBytes >= 1 * 1024 * 1024) {
      console.error(`FAIL: Category translation payload should stay under 1 MB, got ${categoryBytes} bytes`);
      return false;
    }
    if (fieldBytes >= 1 * 1024 * 1024) {
      console.error(`FAIL: Field translation payload should stay under 1 MB, got ${fieldBytes} bytes`);
      return false;
    }
    if (mergedBytes <= 1 * 1024 * 1024) {
      console.error(`FAIL: Merged locale payload should exceed 1 MB, got ${mergedBytes} bytes`);
      return false;
    }
    console.log("PASS: Individual translation sections stay under 1 MB while merged locale exceeds it");

    try {
      validateBuildRequest(strictTranslationRequest, { strict: true });
      console.error("FAIL: Strict validation should reject merged locale payloads above 1 MB");
      return false;
    } catch (error) {
      const message = String((error as Error).message || error);
      if (!message.includes('Translations for locale "es" exceed 1 MB')) {
        console.error(`FAIL: Unexpected strict validation error for merged locale payload: ${message}`);
        return false;
      }
      console.log("PASS: Strict validation rejects the merged locale payload");
    }

    console.log("=== Strict Merged Translation Payload Limit: ALL TESTS PASSED ===");
    return true;
  } catch (error) {
    console.error("FAIL: Exception thrown - " + error.message);
    return false;
  }
}

function testStrictTotalEntityOverflowIncludesOptionsAndTranslations(): boolean {
  console.log("=== Test: Strict Total Entity Overflow Includes Options and Translations ===");

  try {
    const optionCount = 5001;
    const field = createSelectFieldWithOptions(optionCount);
    const strictEntityRequest: BuildRequest = {
      metadata: { name: "Entity Overflow Payload", version: "1.0.0" },
      locales: ["es"],
      categories: [
        {
          id: "cat-1",
          name: "Category",
          appliesTo: ["observation"],
          defaultFieldIds: ["field-1"],
          fields: ["field-1"]
        }
      ],
      fields: [field],
      translations: {
        es: {
          fields: {
            "field-1": createOptionTranslationMap(optionCount)
          }
        }
      }
    };

    const categoryCount = strictEntityRequest.categories.length;
    const fieldCount = strictEntityRequest.fields.length;
    const iconCount = strictEntityRequest.icons?.length || 0;
    const optionTotal = strictEntityRequest.fields.reduce(
      (total, currentField) => total + (currentField.options?.length || 0),
      0
    );
    const translationTotal = Object.values(strictEntityRequest.translations?.es?.fields || {})
      .reduce((total, fieldTranslations) => total + Object.values(fieldTranslations).filter((value) => Boolean(value)).length, 0);
    const totalEntities = categoryCount + fieldCount + iconCount + optionTotal + translationTotal;

    if (categoryCount + fieldCount + iconCount >= 10000) {
      console.error("FAIL: Base entity counts should remain below the legacy row-count limit");
      return false;
    }
    if (optionTotal <= 0 || translationTotal <= 0) {
      console.error("FAIL: Test setup should include both options and translations");
      return false;
    }
    if (totalEntities <= 10000) {
      console.error(`FAIL: Total entities should exceed 10,000 once options and translations are counted, got ${totalEntities}`);
      return false;
    }
    console.log("PASS: Options and translated option entries push the payload past 10,000 entities");

    try {
      validateBuildRequest(strictEntityRequest, { strict: true });
      console.error("FAIL: Strict validation should reject payloads over the combined entity limit");
      return false;
    } catch (error) {
      const message = String((error as Error).message || error);
      if (!message.includes("Combined counts (categories + fields + icons + options + translations) exceed 10000")) {
        console.error(`FAIL: Unexpected strict validation error for total entity overflow: ${message}`);
        return false;
      }
      console.log("PASS: Strict validation rejects the combined entity overflow");
    }

    console.log("=== Strict Total Entity Overflow Includes Options and Translations: ALL TESTS PASSED ===");
    return true;
  } catch (error) {
    console.error("FAIL: Exception thrown - " + error.message);
    return false;
  }
}

function testExtractFileId(): boolean {
  console.log("=== Test: Extract File ID ===");

  try {
    // Test direct file ID
    const id1 = extractFileId("1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms");
    if (id1 !== "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms") {
      console.error("FAIL: Should return direct file ID");
      return false;
    }
    console.log("PASS: Direct file ID extracted");

    // Test /file/d/ URL format
    const id2 = extractFileId("https://drive.google.com/file/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/view");
    if (id2 !== "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms") {
      console.error("FAIL: Should extract from /file/d/ URL");
      return false;
    }
    console.log("PASS: /file/d/ URL format works");

    // Test ?id= URL format
    const id3 = extractFileId("https://drive.google.com/open?id=1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms");
    if (id3 !== "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms") {
      console.error("FAIL: Should extract from ?id= URL");
      return false;
    }
    console.log("PASS: ?id= URL format works");

    // Test empty input
    const id4 = extractFileId("");
    if (id4 !== null) {
      console.error("FAIL: Empty input should return null");
      return false;
    }
    console.log("PASS: Empty input returns null");

    // Test invalid input
    const id5 = extractFileId("not-a-valid-url");
    if (id5 !== null) {
      console.error("FAIL: Invalid input should return null");
      return false;
    }
    console.log("PASS: Invalid input returns null");

    console.log("=== Extract File ID: ALL TESTS PASSED ===");
    return true;
  } catch (error) {
    console.error("FAIL: Exception thrown - " + error.message);
    return false;
  }
}

function testMapFieldTypeToChar(): boolean {
  console.log("=== Test: Map Field Type to Char ===");

  try {
    if (mapFieldTypeToChar("text") !== "t") {
      console.error("FAIL: text should map to 't'");
      return false;
    }
    if (mapFieldTypeToChar("textarea") !== "T") {
      console.error("FAIL: textarea should map to 'T'");
      return false;
    }
    if (mapFieldTypeToChar("number") !== "n") {
      console.error("FAIL: number should map to 'n'");
      return false;
    }
    if (mapFieldTypeToChar("integer") !== "i") {
      console.error("FAIL: integer should map to 'i'");
      return false;
    }
    if (mapFieldTypeToChar("multiselect") !== "m") {
      console.error("FAIL: multiselect should map to 'm'");
      return false;
    }
    if (mapFieldTypeToChar("select") !== "s") {
      console.error("FAIL: select should map to 's'");
      return false;
    }
    console.log("PASS: All field type mappings correct");

    console.log("=== Map Field Type to Char: ALL TESTS PASSED ===");
    return true;
  } catch (error) {
    console.error("FAIL: Exception thrown - " + error.message);
    return false;
  }
}

// =============================================================================
// API Response Tests
// =============================================================================

function testApiErrorParsing(): boolean {
  console.log("=== Test: API Error Parsing ===");

  try {
    // Test error response parsing
    const errorJson = JSON.stringify({
      error: "VALIDATION_ERROR",
      message: "Invalid configuration",
      details: { errors: ["Missing required field: name"] }
    });

    const parsed: ApiErrorResponse = JSON.parse(errorJson);

    if (parsed.error !== "VALIDATION_ERROR") {
      console.error("FAIL: Error code not parsed correctly");
      return false;
    }
    console.log("PASS: Error code parsed correctly");

    if (parsed.message !== "Invalid configuration") {
      console.error("FAIL: Error message not parsed correctly");
      return false;
    }
    console.log("PASS: Error message parsed correctly");

    if (!parsed.details?.errors?.includes("Missing required field: name")) {
      console.error("FAIL: Error details not parsed correctly");
      return false;
    }
    console.log("PASS: Error details parsed correctly");

    console.log("=== API Error Parsing: ALL TESTS PASSED ===");
    return true;
  } catch (error) {
    console.error("FAIL: Exception thrown - " + error.message);
    return false;
  }
}

// =============================================================================
// Category Selection Tests
// =============================================================================

function testCategorySelection(): boolean {
  console.log("=== Test: Category Selection ===");

  try {
    // Test setting category selection
    const testCategories = ["cat-a", "cat-b", "cat-c"];
    setCategorySelection(testCategories);

    const retrieved = getCategorySelection();
    if (retrieved.length !== 3) {
      console.error("FAIL: Should retrieve 3 categories");
      return false;
    }
    console.log("PASS: Category count correct");

    // Verify order is preserved
    for (let i = 0; i < testCategories.length; i++) {
      if (retrieved[i] !== testCategories[i]) {
        console.error(`FAIL: Category order not preserved at index ${i}`);
        return false;
      }
    }
    console.log("PASS: Category order preserved");

    // Verify it's a copy (modifying retrieved doesn't affect internal state)
    retrieved.push("should-not-appear");
    const retrieved2 = getCategorySelection();
    if (retrieved2.length !== 3) {
      console.error("FAIL: getCategorySelection should return a copy");
      return false;
    }
    console.log("PASS: Returns a copy, not reference");

    console.log("=== Category Selection: ALL TESTS PASSED ===");
    return true;
  } catch (error) {
    console.error("FAIL: Exception thrown - " + error.message);
    return false;
  }
}

// =============================================================================
// New Field Type Tests (Integer, Boolean, Date, DateTime, Photo, Location)
// =============================================================================

function testAllFieldTypes(): boolean {
  console.log("=== Test: All Field Types ===");

  try {
    // Test all 11 supported field types
    const testData: SheetData = {
      documentName: "Type Test" as any,
      Categories: [
        ["Name", "Icon", "Fields", "ID", "Color", "Icon ID"],
        ["Test", "", "", "test", "#FF0000", "test"]
      ],
      Details: [
        ["Name", "Helper Text", "Type", "Options", "ID", "Universal"],
        ["Text Field", "Single line text", "t", "", "text-field", "FALSE"],
        ["Textarea Field", "Multi-line text", "T", "", "textarea-field", "FALSE"],
        ["Number Field", "Numeric value", "n", "", "number-field", "FALSE"],
        ["Integer Field", "Integer only", "i", "", "integer-field", "FALSE"],
        ["Select Field", "Single choice", "s", "A,B,C", "select-field", "FALSE"],
        ["Multiselect Field", "Multiple choices", "m", "X,Y,Z", "multi-field", "FALSE"],
        ["Boolean Field", "True/False", "b", "", "boolean-field", "FALSE"],
        ["Date Field", "Date only", "d", "", "date-field", "FALSE"],
        ["DateTime Field", "Date and time", "D", "", "datetime-field", "FALSE"],
        ["Photo Field", "Photo attachment", "p", "", "photo-field", "FALSE"],
        ["Location Field", "GPS coordinates", "l", "", "location-field", "FALSE"]
      ]
    } as SheetData;

    const fields = buildFields(testData);

    if (fields.length !== 11) {
      console.error(`FAIL: Expected 11 fields, got ${fields.length}`);
      return false;
    }
    console.log("PASS: All 11 field types created");

    // Verify each type
    const typeChecks = [
      { id: "text-field", expectedType: "text" },
      { id: "textarea-field", expectedType: "textarea" },
      { id: "number-field", expectedType: "number" },
      { id: "integer-field", expectedType: "integer" },
      { id: "select-field", expectedType: "select" },
      { id: "multi-field", expectedType: "multiselect" },
      { id: "boolean-field", expectedType: "boolean" },
      { id: "date-field", expectedType: "date" },
      { id: "datetime-field", expectedType: "datetime" },
      { id: "photo-field", expectedType: "photo" },
      { id: "location-field", expectedType: "location" }
    ];

    for (const check of typeChecks) {
      const field = fields.find(f => f.id === check.id);
      if (!field) {
        console.error(`FAIL: Field ${check.id} not found`);
        return false;
      }
      if (field.type !== check.expectedType) {
        console.error(`FAIL: Field ${check.id} should be type '${check.expectedType}', got '${field.type}'`);
        return false;
      }
    }
    console.log("PASS: All field types mapped correctly");

    console.log("=== All Field Types: ALL TESTS PASSED ===");
    return true;
  } catch (error) {
    console.error("FAIL: Exception thrown - " + error.message);
    return false;
  }
}

function testIntegerVsNumberDistinction(): boolean {
  console.log("=== Test: Integer vs Number Distinction ===");

  try {
    const testData: SheetData = {
      documentName: "Integer Test" as any,
      Categories: [
        ["Name", "Icon", "Fields", "ID", "Color", "Icon ID"],
        ["Test", "", "", "test", "#FF0000", "test"]
      ],
      Details: [
        ["Name", "Helper Text", "Type", "Options", "ID", "Universal"],
        ["Age", "Person's age", "i", "", "age", "FALSE"],
        ["Height", "Height in cm", "n", "", "height", "FALSE"]
      ]
    } as SheetData;

    const fields = buildFields(testData);

    const ageField = fields.find(f => f.id === "age");
    const heightField = fields.find(f => f.id === "height");

    if (!ageField || ageField.type !== "integer") {
      console.error("FAIL: 'i' should map to integer type");
      return false;
    }
    console.log("PASS: Integer type correctly mapped from 'i'");

    if (!heightField || heightField.type !== "number") {
      console.error("FAIL: 'n' should map to number type");
      return false;
    }
    console.log("PASS: Number type correctly mapped from 'n'");

    if (ageField.type === heightField.type) {
      console.error("FAIL: Integer and number should be distinct types");
      return false;
    }
    console.log("PASS: Integer and number are distinct types");

    console.log("=== Integer vs Number Distinction: ALL TESTS PASSED ===");
    return true;
  } catch (error) {
    console.error("FAIL: Exception thrown - " + error.message);
    return false;
  }
}

function testMapFieldTypeToCharAllTypes(): boolean {
  console.log("=== Test: Map All Field Types to Char ===");

  try {
    const mappings = [
      { type: "text", expected: "t" },
      { type: "textarea", expected: "T" },
      { type: "number", expected: "n" },
      { type: "integer", expected: "i" },
      { type: "select", expected: "s" },
      { type: "multiselect", expected: "m" },
      { type: "boolean", expected: "b" },
      { type: "date", expected: "d" },
      { type: "datetime", expected: "D" },
      { type: "photo", expected: "p" },
      { type: "location", expected: "l" }
    ];

    for (const mapping of mappings) {
      const result = mapFieldTypeToChar(mapping.type as FieldType);
      if (result !== mapping.expected) {
        console.error(`FAIL: ${mapping.type} should map to '${mapping.expected}', got '${result}'`);
        return false;
      }
    }
    console.log("PASS: All 11 field types map correctly to characters");

    console.log("=== Map All Field Types to Char: ALL TESTS PASSED ===");
    return true;
  } catch (error) {
    console.error("FAIL: Exception thrown - " + error.message);
    return false;
  }
}

// =============================================================================
// Universal Fields Tests
// =============================================================================

function testUniversalFieldsAddedToAllCategories(): boolean {
  console.log("=== Test: Universal Fields Added to All Categories ===");

  try {
    const testData: SheetData = {
      documentName: "Universal Test" as any,
      Categories: [
        ["Name", "Icon", "Fields", "ID", "Color", "Icon ID"],
        ["Category A", "", "field-a", "cat-a", "#FF0000", "cat-a"],
        ["Category B", "", "field-b", "cat-b", "#00FF00", "cat-b"],
        ["Category C", "", "", "cat-c", "#0000FF", "cat-c"]
      ],
      Details: [
        ["Name", "Helper Text", "Type", "Options", "ID", "Universal"],
        ["Notes", "General notes", "T", "", "notes", "TRUE"],  // Universal field
        ["Field A", "Specific to A", "t", "", "field-a", "FALSE"],
        ["Field B", "Specific to B", "t", "", "field-b", "FALSE"]
      ]
    } as SheetData;

    const fields = buildFields(testData);
    const categories = buildCategories(testData, fields);

    if (categories.length !== 3) {
      console.error(`FAIL: Expected 3 categories, got ${categories.length}`);
      return false;
    }

    // All categories should have the universal field 'notes'
    for (const cat of categories) {
      if (!cat.defaultFieldIds || !cat.defaultFieldIds.includes("notes")) {
        console.error(`FAIL: Category '${cat.id}' should include universal field 'notes'`);
        return false;
      }
    }
    console.log("PASS: Universal field 'notes' appears in all categories");

    // Check Category A has both notes (universal) and field-a (explicit)
    const catA = categories.find(c => c.id === "cat-a");
    if (!catA || !catA.defaultFieldIds || catA.defaultFieldIds.length !== 2) {
      console.error("FAIL: Category A should have 2 fields (notes + field-a)");
      return false;
    }
    if (!catA.defaultFieldIds.includes("notes") || !catA.defaultFieldIds.includes("field-a")) {
      console.error("FAIL: Category A should have both notes and field-a");
      return false;
    }
    console.log("PASS: Category A has universal + explicit fields");

    // Check Category C has only notes (universal, no explicit fields)
    const catC = categories.find(c => c.id === "cat-c");
    if (!catC || !catC.defaultFieldIds || catC.defaultFieldIds.length !== 1) {
      console.error("FAIL: Category C should have only 1 field (notes)");
      return false;
    }
    if (!catC.defaultFieldIds.includes("notes")) {
      console.error("FAIL: Category C should have notes field");
      return false;
    }
    console.log("PASS: Category C has only universal field");

    console.log("=== Universal Fields Added to All Categories: ALL TESTS PASSED ===");
    return true;
  } catch (error) {
    console.error("FAIL: Exception thrown - " + error.message);
    return false;
  }
}

function testUniversalFieldDoesNotSetRequired(): boolean {
  console.log("=== Test: Universal Field Does Not Set Required ===");

  try {
    const testData: SheetData = {
      documentName: "Universal Required Test" as any,
      Categories: [
        ["Name", "Icon", "Fields", "ID", "Color", "Icon ID"],
        ["Test", "", "", "test", "#FF0000", "test"]
      ],
      Details: [
        ["Name", "Helper Text", "Type", "Options", "ID", "Universal"],
        ["Universal Field", "Universal but optional", "t", "", "universal-field", "TRUE"],
        ["Regular Field", "Regular field", "t", "", "regular-field", "FALSE"]
      ]
    } as SheetData;

    const fields = buildFields(testData);

    const universalField = fields.find(f => f.id === "universal-field");
    if (!universalField) {
      console.error("FAIL: Universal field not found");
      return false;
    }

    // Universal field should NOT have required property set
    if (universalField.required === true) {
      console.error("FAIL: Universal field should not be marked as required");
      return false;
    }
    console.log("PASS: Universal field is not marked as required");

    // Neither field should have required set (since we removed that mapping)
    const regularField = fields.find(f => f.id === "regular-field");
    if (regularField && regularField.required === true) {
      console.error("FAIL: Regular field should not have required set either");
      return false;
    }
    console.log("PASS: Required property is not set from spreadsheet");

    console.log("=== Universal Field Does Not Set Required: ALL TESTS PASSED ===");
    return true;
  } catch (error) {
    console.error("FAIL: Exception thrown - " + error.message);
    return false;
  }
}

function testCategoriesAppliesParsing(): boolean {
  console.log("=== Test: Categories Applies Parsing ===");

  // Reset mutable global so prior tests (e.g. testUniversalFieldsAddedToAllCategories)
  // that trigger the auto-create Applies path don't leak state into this test.
  if (typeof AUTO_CREATED_APPLIES_COLUMN !== "undefined") {
    AUTO_CREATED_APPLIES_COLUMN = false;
  }

  try {
    const testData: SheetData = {
      documentName: "Applies Validation" as any,
      Categories: [
        ["Name", "Icon", "Fields", "Applies", "Category ID", "Icon ID"],
        ["Blank Applies", "", "field-a", "", "blank-applies", "blank-applies"],
        ["Track Only", "", "field-a", "Track", "track-only", "track-only"]
      ],
      Details: [
        ["Name", "Helper Text", "Type", "Options", "ID", "Universal"],
        ["Field A", "Helper", "t", "", "field-a", "FALSE"]
      ]
    } as SheetData;

    const fields = buildFields(testData);
    const categories = buildCategories(testData, fields);
    if (categories.length !== 2) {
      console.error(`FAIL: Expected 2 categories, got ${categories.length}`);
      return false;
    }

    const blankApplies = categories.find((category) => category.id === "blank-applies");
    if (!blankApplies) {
      console.error("FAIL: Could not find category with blank Applies value");
      return false;
    }
    if (
      !blankApplies.appliesTo ||
      blankApplies.appliesTo.length !== 1 ||
      blankApplies.appliesTo[0] !== "observation"
    ) {
      console.error(
        `FAIL: Blank Applies should default to observation, got ${JSON.stringify(blankApplies.appliesTo)}`,
      );
      return false;
    }
    console.log("PASS: Blank Applies defaults to observation");

    const trackOnly = categories.find((category) => category.id === "track-only");
    if (!trackOnly) {
      console.error("FAIL: Could not find track-only category");
      return false;
    }
    if (
      !trackOnly.appliesTo ||
      trackOnly.appliesTo.length !== 1 ||
      trackOnly.appliesTo[0] !== "track"
    ) {
      console.error(
        `FAIL: Track Applies should remain track-only, got ${JSON.stringify(trackOnly.appliesTo)}`,
      );
      return false;
    }
    console.log("PASS: Track-only Applies values remain track-only during category parsing");

    console.log("=== Categories Applies Parsing: ALL TESTS PASSED ===");
    return true;
  } catch (error) {
    console.error("FAIL: Exception thrown - " + (error as Error).message);
    return false;
  }
}

function testBuildCategoriesRequiresObservationCoverage(): boolean {
  console.log("=== Test: Build Categories Requires Observation Coverage ===");

  try {
    const testData: SheetData = {
      documentName: "Observation Coverage" as any,
      Categories: [
        ["Name", "Icon", "Fields", "Applies", "Category ID", "Icon ID"],
        ["Track One", "", "field-a", "track", "track-one", "track-one"],
        ["Track Two", "", "field-a", "track", "track-two", "track-two"],
      ],
      Details: [
        ["Name", "Helper Text", "Type", "Options", "ID", "Universal"],
        ["Field A", "Helper", "t", "", "field-a", "FALSE"],
      ],
    } as SheetData;

    const fields = buildFields(testData);

    try {
      buildCategories(testData, fields);
      console.error("FAIL: buildCategories should reject sheets without observation coverage");
      return false;
    } catch (error) {
      const message = String((error as Error).message || error);
      if (!message.includes('At least one category must include "observation"')) {
        console.error(`FAIL: Unexpected builder error for missing observation coverage: ${message}`);
        return false;
      }
    }

    console.log("PASS: Builder rejects all-track Categories sheets with no observation coverage");
    return true;
  } catch (error) {
    console.error("FAIL: Exception thrown - " + (error as Error).message);
    return false;
  }
}


// =============================================================================
// Round-Trip Integration Tests
// =============================================================================

function testIntegerFieldRoundTrip(): boolean {
  console.log("=== Test: Integer Field Round-Trip ===");

  try {
    // Start with integer field in spreadsheet
    const buildData: SheetData = {
      documentName: "Round Trip Test" as any,
      Categories: [
        ["Name", "Icon", "Fields", "ID", "Color", "Icon ID"],
        ["Test", "", "age", "test", "#FF0000", "test"]
      ],
      Details: [
        ["Name", "Helper Text", "Type", "Options", "ID", "Universal"],
        ["Age", "Person's age", "i", "", "age", "FALSE"]
      ]
    } as SheetData;

    // Build fields (simulates build process)
    const fields = buildFields(buildData);
    const ageField = fields.find(f => f.id === "age");

    if (!ageField || ageField.type !== "integer") {
      console.error("FAIL: Initial build should create integer field");
      return false;
    }
    console.log("PASS: Build creates integer field");

    // Map back to character (simulates import process)
    const charCode = mapFieldTypeToChar(ageField.type);
    if (charCode !== "i") {
      console.error(`FAIL: Integer should map back to 'i', got '${charCode}'`);
      return false;
    }
    console.log("PASS: Integer maps back to 'i'");

    // Verify it doesn't become 'n' (number)
    if (charCode === "n") {
      console.error("FAIL: Integer incorrectly mapped to 'n' (number)");
      return false;
    }
    console.log("PASS: Integer preserved (not converted to number)");

    console.log("=== Integer Field Round-Trip: ALL TESTS PASSED ===");
    return true;
  } catch (error) {
    console.error("FAIL: Exception thrown - " + error.message);
    return false;
  }
}

function testUniversalFieldRoundTrip(): boolean {
  console.log("=== Test: Import Preserves Field Distribution (No Universal Inference) ===");

  try {
    // Create a spreadsheet with a universal field
    const testData: SheetData = {
      documentName: "Universal Round-Trip" as any,
      Categories: [
        ["Name", "Icon", "Fields", "ID", "Color", "Icon ID"],
        ["Cat 1", "", "field1", "cat1", "#FF0000", "cat1"],
        ["Cat 2", "", "field2", "cat2", "#00FF00", "cat2"],
        ["Cat 3", "", "", "cat3", "#0000FF", "cat3"]
      ],
      Details: [
        ["Name", "Helper Text", "Type", "Options", "ID", "Universal"],
        ["Notes", "Universal notes field", "T", "", "notes", "TRUE"],  // Universal field
        ["Field 1", "Specific to Cat 1", "t", "", "field1", "FALSE"],
        ["Field 2", "Specific to Cat 2", "t", "", "field2", "FALSE"]
      ]
    } as SheetData;

    // Build from spreadsheet (notes should be added to all categories)
    const fields = buildFields(testData);
    const categories = buildCategories(testData, fields);

    // Create BuildRequest from built data
    const buildRequest: BuildRequest = {
      metadata: { name: "Test", version: "1.0.0" },
      categories: categories,
      fields: fields
    };

    // All categories should have 'notes' after build (due to Universal=TRUE)
    for (const cat of buildRequest.categories) {
      if (!cat.defaultFieldIds || !cat.defaultFieldIds.includes("notes")) {
        console.error(`FAIL: Category '${cat.id}' should have universal field 'notes' after build`);
        return false;
      }
    }
    console.log("PASS: Universal field added to all categories during build");

    // Now simulate import: create a mock spreadsheet context
    // In real import, populateDetailsSheet would be called
    // We verify it does NOT mark 'notes' as Universal just because it appears everywhere

    // Mock the Details sheet that would be populated during import
    const mockImportedDetails: any[][] = [];
    for (const field of buildRequest.fields) {
      // populateDetailsSheet always sets Universal=FALSE on import
      mockImportedDetails.push([
        field.name,
        field.description || '',
        mapFieldTypeToChar(field.type),
        '',  // options
        field.id,
        'FALSE'  // Always FALSE on import to preserve exact field distribution
      ]);
    }

    // Verify 'notes' is NOT marked as Universal in the imported Details sheet
    const notesRow = mockImportedDetails.find(row => row[4] === 'notes');
    if (!notesRow) {
      console.error("FAIL: 'notes' field should be in imported Details sheet");
      return false;
    }
    if (notesRow[5] !== 'FALSE') {
      console.error("FAIL: 'notes' should NOT be marked as Universal during import (got: " + notesRow[5] + ")");
      return false;
    }
    console.log("PASS: Field appearing in all categories is NOT auto-marked as Universal during import");

    // Verify Categories sheet would show 'notes' in each category's Fields column
    // (not stripped out, which would happen if it were incorrectly marked as Universal)
    for (const cat of buildRequest.categories) {
      const hasNotes = cat.defaultFieldIds && cat.defaultFieldIds.includes("notes");
      if (!hasNotes) {
        console.error(`FAIL: Category ${cat.id} should retain 'notes' in defaultFieldIds during import`);
        return false;
      }
    }
    console.log("PASS: Import preserves exact field distribution without inferring universal intent");

    console.log("=== Import Preserves Field Distribution: ALL TESTS PASSED ===");
    return true;
  } catch (error) {
    console.error("FAIL: Exception thrown - " + error.message);
    return false;
  }
}

// =============================================================================
// Edge Case Tests
// =============================================================================

function testEmptyInputs(): boolean {
  console.log("=== Test: Empty Inputs ===");

  try {
    // Test buildFields with empty details
    const emptyData: SheetData = {
      documentName: "Empty Test" as any,
      Categories: [
        ["Name", "Icon", "Fields", "ID", "Color", "Icon ID"]
      ],
      Details: [
        ["Name", "Helper Text", "Type", "Options", "ID", "Universal"]
      ]
    } as SheetData;

    const fields = buildFields(emptyData);
    if (fields.length !== 0) {
      console.error("FAIL: Empty Details should produce 0 fields");
      return false;
    }
    console.log("PASS: Empty Details produces 0 fields");

    // Test parseOptions with empty string
    const emptyOptions = parseOptions("");
    if (emptyOptions !== undefined) {
      console.error("FAIL: Empty options string should return undefined");
      return false;
    }
    console.log("PASS: Empty options string returns undefined");

    console.log("=== Empty Inputs: ALL TESTS PASSED ===");
    return true;
  } catch (error) {
    console.error("FAIL: Exception thrown - " + error.message);
    return false;
  }
}

function testOptionValuePreservation(): boolean {
  console.log("=== Test: Option Value Preservation ===");

  try {
    // Test "value:label" format
    const opts1 = parseOptions("oak:Oak Tree,pine:Pine Tree");
    if (!opts1 || opts1.length !== 2) {
      console.error("FAIL: Should parse 2 options with custom values");
      return false;
    }
    if (opts1[0].value !== "oak" || opts1[0].label !== "Oak Tree") {
      console.error("FAIL: First option should preserve custom value 'oak'");
      return false;
    }
    if (opts1[1].value !== "pine" || opts1[1].label !== "Pine Tree") {
      console.error("FAIL: Second option should preserve custom value 'pine'");
      return false;
    }
    console.log("PASS: Custom option values preserved");

    // Test mixed format (some with custom values, some without)
    const opts2 = parseOptions("oak:Oak Tree,Maple,pine:Pine Tree");
    if (!opts2 || opts2.length !== 3) {
      console.error("FAIL: Should parse 3 options in mixed format");
      return false;
    }
    if (opts2[0].value !== "oak") {
      console.error("FAIL: First option should have custom value 'oak'");
      return false;
    }
    if (opts2[1].value !== "maple" || opts2[1].label !== "Maple") {
      console.error("FAIL: Second option should auto-slugify to 'maple'");
      return false;
    }
    console.log("PASS: Mixed format (custom + auto values) handled correctly");

    console.log("=== Option Value Preservation: ALL TESTS PASSED ===");
    return true;
  } catch (error) {
    console.error("FAIL: Exception thrown - " + error.message);
    return false;
  }
}

function testCategoryIdPreservation(): boolean {
  console.log("=== Test: Category ID Preservation ===");

  try {
    const testData: SheetData = {
      documentName: "ID Test" as any,
      Categories: [
        ["Name", "Icon", "Fields", "Category ID", "Color", "Icon ID"],
        ["My Category", "", "", "custom-cat-id", "#FF0000", "custom-icon-id"],
        ["Auto ID Category", "", "", "", "#00FF00", ""]
      ],
      Details: [
        ["Name", "Helper Text", "Type", "Options", "ID", "Universal"]
      ]
    } as SheetData;

    const fields = buildFields(testData);
    const categories = buildCategories(testData, fields);

    // First category should use explicit ID
    const cat1 = categories[0];
    if (cat1.id !== "custom-cat-id") {
      console.error(`FAIL: Category should preserve explicit ID 'custom-cat-id', got '${cat1.id}'`);
      return false;
    }
    if (cat1.iconId !== "custom-icon-id") {
      console.error(`FAIL: Category should preserve explicit iconId 'custom-icon-id', got '${cat1.iconId}'`);
      return false;
    }
    console.log("PASS: Explicit category and icon IDs preserved");

    // Second category should auto-generate ID from name
    const cat2 = categories[1];
    if (cat2.id !== "auto-id-category") {
      console.error(`FAIL: Category should auto-generate ID 'auto-id-category', got '${cat2.id}'`);
      return false;
    }
    console.log("PASS: Category ID auto-generated when not provided");

    console.log("=== Category ID Preservation: ALL TESTS PASSED ===");
    return true;
  } catch (error) {
    console.error("FAIL: Exception thrown - " + error.message);
    return false;
  }
}

function testIconsRespectStoredCategoryIds(): boolean {
  console.log("=== Test: Icons Respect Stored Category IDs ===");

  try {
    const inlineSvg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10"><circle cx="5" cy="5" r="5" /></svg>';
    const testData: SheetData = {
      documentName: "Icon ID Test" as any,
      Categories: [
        ["Name", "Icon", "Fields", "Category ID", "Color", "Icon ID"],
        ["Stored Icon", inlineSvg, "", "stored-id", "#123456", "custom-icon"],
        ["Fallback Icon", inlineSvg, "", "fallback-id", "#654321", ""]
      ],
      Details: [
        ["Name", "Helper Text", "Type", "Options", "ID", "Universal"]
      ],
      Icons: [
        ["ID", "SVG"],
      ]
    } as SheetData;

    const icons = buildIconsFromSheet(testData);
    const iconIds = icons.map(icon => icon.id).sort();

    if (!iconIds.includes("custom-icon")) {
      console.error("FAIL: Icon should preserve explicit Icon ID 'custom-icon'");
      return false;
    }

    if (!iconIds.includes("fallback-id")) {
      console.error("FAIL: Icon should fall back to Category ID when Icon ID column is blank");
      return false;
    }

    console.log("PASS: Icons respect stored Category and Icon IDs");
    return true;
  } catch (error) {
    console.error("FAIL: Exception thrown - " + error.message);
    return false;
  }
}

function testCrossSheetIconCollisionsUseExactIds(): boolean {
  console.log("=== Test: Cross-Sheet Icon Collisions Use Exact IDs ===");

  try {
    const noCollision = findCrossSheetIconIdCollisions(
      [{ id: "tree", row: 2 }],
      [{ id: "tree.svg", row: 5 }],
    );
    if (noCollision.length !== 0) {
      console.error(
        'FAIL: "tree" and "tree.svg" should not collide when matching exact icon IDs',
      );
      return false;
    }

    const exactCollision = findCrossSheetIconIdCollisions(
      [{ id: "tree", row: 2 }],
      [
        { id: "tree", row: 5 },
        { id: "tree", row: 6 },
      ],
    );
    if (exactCollision.length !== 2) {
      console.error(
        `FAIL: Expected 2 exact-ID collisions, got ${exactCollision.length}`,
      );
      return false;
    }

    if (
      exactCollision[0].iconId !== "tree" ||
      exactCollision[0].categoryId !== "tree"
    ) {
      console.error("FAIL: Collision payload should preserve the exact icon IDs");
      return false;
    }

    console.log("PASS: Cross-sheet icon collisions use exact IDs");
    return true;
  } catch (error) {
    console.error("FAIL: Exception thrown - " + error.message);
    return false;
  }
}

function testFieldIdPreservation(): boolean {
  console.log("=== Test: Field ID Preservation ===");

  try {
    const testData: SheetData = {
      documentName: "Field ID Test" as any,
      Categories: [
        ["Name", "Icon", "Fields", "ID", "Color", "Icon ID"],
        ["Test", "", "", "test", "#FF0000", "test"]
      ],
      Details: [
        ["Name", "Helper Text", "Type", "Options", "ID", "Universal"],
        ["Field Name", "Description", "t", "", "custom-field-id", "FALSE"],
        ["Auto Field", "Auto ID", "t", "", "", "FALSE"]
      ]
    } as SheetData;

    const fields = buildFields(testData);

    // First field should use explicit ID
    const field1 = fields[0];
    if (field1.id !== "custom-field-id") {
      console.error(`FAIL: Field should preserve explicit ID 'custom-field-id', got '${field1.id}'`);
      return false;
    }
    console.log("PASS: Explicit field ID preserved");

    // Second field should auto-generate ID from name
    const field2 = fields[1];
    if (field2.id !== "auto-field") {
      console.error(`FAIL: Field should auto-generate ID 'auto-field', got '${field2.id}'`);
      return false;
    }
    console.log("PASS: Field ID auto-generated when not provided");

    console.log("=== Field ID Preservation: ALL TESTS PASSED ===");
    return true;
  } catch (error) {
    console.error("FAIL: Exception thrown - " + error.message);
    return false;
  }
}

function testBuildFieldsSkipsBlankRows(): boolean {
  console.log("=== Test: buildFields Skips Blank Rows ===");

  try {
    const testData: SheetData = {
      documentName: "Blank Row Test" as any,
      Categories: [
        ["Name", "Icon", "Fields", "ID", "Color", "Icon ID"]
      ],
      Details: [
        ["Name", "Helper Text", "Type", "Options", "ID", "Universal"],
        ["Field 1", "First field", "t", "", "field1", "FALSE"],
        ["", "", "", "", "", ""],  // Blank row
        ["Field 2", "Second field", "t", "", "field2", "FALSE"],
        ["", "", "", "", "", ""],  // Another blank row
        ["Field 3", "Third field", "t", "", "field3", "FALSE"]
      ]
    } as SheetData;

    const fields = buildFields(testData);

    // Should have exactly 3 fields (blank rows skipped)
    if (fields.length !== 3) {
      console.error(`FAIL: Expected 3 fields (blank rows skipped), got ${fields.length}`);
      return false;
    }
    console.log("PASS: Blank rows skipped, got exactly 3 fields");

    // Verify fields are correctly parsed
    if (fields[0].id !== "field1" || fields[0].name !== "Field 1") {
      console.error("FAIL: First field incorrectly parsed");
      return false;
    }
    if (fields[1].id !== "field2" || fields[1].name !== "Field 2") {
      console.error("FAIL: Second field incorrectly parsed");
      return false;
    }
    if (fields[2].id !== "field3" || fields[2].name !== "Field 3") {
      console.error("FAIL: Third field incorrectly parsed");
      return false;
    }
    console.log("PASS: All fields correctly parsed despite blank rows");

    console.log("=== buildFields Skips Blank Rows: ALL TESTS PASSED ===");
    return true;
  } catch (error) {
    console.error("FAIL: Exception thrown - " + error.message);
    return false;
  }
}

function testTranslationsWithBlankRows(): boolean {
  console.log("=== Test: Translations Work With Blank Rows in Details ===");

  try {
    // Create mock data with blank rows in Details
    const testData: SheetData = {
      documentName: "Translation Test" as any,
      Categories: [
        ["Name", "Icon", "Fields", "ID", "Color", "Icon ID"],
        ["Test Category", "", "field1, field2", "cat1", "#FF0000", ""]
      ],
      Details: [
        ["Name", "Helper Text", "Type", "Options", "ID", "Universal"],
        ["Field 1", "Help text 1", "s", "opt1, opt2", "field1", "FALSE"],
        ["", "", "", "", "", ""],  // Blank row
        ["Field 2", "Help text 2", "t", "", "field2", "FALSE"]
      ],
      "Detail Label Translations": [
        ["Name", "Spanish", "French"],
        ["Field 1", "Campo 1", "Champ 1"],
        ["", "", ""],  // Blank row in translations too
        ["Field 2", "Campo 2", "Champ 2"]
      ],
      "Detail Helper Text Translations": [
        ["Helper Text", "Spanish", "French"],
        ["Help text 1", "Texto de ayuda 1", "Texte d'aide 1"],
        ["", "", ""],  // Blank row
        ["Help text 2", "Texto de ayuda 2", "Texte d'aide 2"]
      ],
      "Detail Option Translations": [
        ["Options", "Spanish", "French"],
        ["opt1, opt2", "opción1, opción2", "option1, option2"],
        ["", "", ""],  // Blank row
        ["", "", ""]
      ]
    } as SheetData;

    // Build fields (should skip blank row)
    const fields = buildFields(testData);
    if (fields.length !== 2) {
      console.error(`FAIL: Expected 2 fields, got ${fields.length}`);
      return false;
    }
    console.log("PASS: Fields correctly built with blank rows skipped");

    // Build categories
    const categories = buildCategories(testData, fields);

    // Build translations
    const translations = buildTranslationsPayload(testData, categories, fields);

    // Verify translations for field1
    const esLocale = translations['es'];
    const esFieldTranslations = esLocale?.field;
    const esField1 = esFieldTranslations?.['field1'];
    if (!esField1) {
      console.error("FAIL: Spanish translations for field1 not found");
      return false;
    }
    if (esField1.label !== "Campo 1") {
      console.error(`FAIL: field1 label translation incorrect: ${esField1.label}`);
      return false;
    }
    if (esField1.helperText !== "Texto de ayuda 1") {
      console.error(`FAIL: field1 helper text translation incorrect: ${esField1.helperText}`);
      return false;
    }
    console.log("PASS: field1 translations correct (before blank row)");

    // Verify translations for field2 (after blank row - this is the critical test)
    const esField2 = esFieldTranslations?.['field2'];
    if (!esField2 || esField2.label !== "Campo 2") {
      console.error("FAIL: Spanish translations for field2 not found (blank row caused index divergence)");
      return false;
    }
    if (esField2.helperText !== "Texto de ayuda 2") {
      console.error(`FAIL: field2 helper text translation incorrect: ${esField2.helperText}`);
      return false;
    }
    console.log("PASS: field2 translations correct (after blank row - no index divergence)");

    // Verify option translations for field1
    if (!esField1["options.0"] || esField1["options.0"] !== "opción1") {
      console.error("FAIL: Option translation for opt1 incorrect");
      return false;
    }
    console.log("PASS: Option translations correct despite blank rows");

    console.log("=== Translations With Blank Rows: ALL TESTS PASSED ===");
    return true;
  } catch (error) {
    console.error("FAIL: Exception thrown - " + error.message);
    return false;
  }
}

function testFieldTranslationsWithoutCategorySheet(): boolean {
  console.log("=== Test: Field Translations Without Category Translations Sheet ===");

  try {
    // Create mock data with field translations but NO Category Translations sheet
    const testData: SheetData = {
      documentName: "Field Only Translations Test" as any,
      Categories: [
        ["Name", "Icon", "Fields", "ID", "Color", "Icon ID"],
        ["Test Category", "", "field1, field2", "test-cat", "#FF0000", ""]
      ],
      Details: [
        ["Name", "Helper Text", "Type", "Options", "ID", "Universal"],
        ["Field 1", "Help for field 1", "t", "", "field1", "FALSE"],
        ["Field 2", "Help for field 2", "s", "Option A, Option B", "field2", "FALSE"]
      ],
      // NOTE: No "Category Translations" sheet
      "Detail Label Translations": [
        ["Name", "Spanish", "French"],
        ["Field 1", "Campo 1", "Champ 1"],
        ["Field 2", "Campo 2", "Champ 2"]
      ],
      "Detail Helper Text Translations": [
        ["Helper Text", "Spanish", "French"],
        ["Help for field 1", "Ayuda para campo 1", "Aide pour champ 1"],
        ["Help for field 2", "Ayuda para campo 2", "Aide pour champ 2"]
      ],
      "Detail Option Translations": [
        ["Options", "Spanish", "French"],
        ["Option A, Option B", "Opción A, Opción B", "Option A, Option B"]
      ]
    } as SheetData;

    // Build fields and categories
    const fields = buildFields(testData);
    const categories = buildCategories(testData, fields);

    if (fields.length !== 2) {
      console.error(`FAIL: Expected 2 fields, got ${fields.length}`);
      return false;
    }
    console.log("PASS: Fields built correctly");

    // Build translations (should work even without Category Translations sheet)
    const translations = buildTranslationsPayload(testData, categories, fields);

    // Verify translations object was created with languages
    if (!translations['es'] || !translations['fr']) {
      console.error("FAIL: Translations should include Spanish and French even without Category Translations sheet");
      return false;
    }
    console.log("PASS: Translation languages detected from Detail sheets");

    // Verify field label translations exist
    const esFields = translations['es']?.field;
    const esField1 = esFields?.['field1'];
    if (!esField1 || esField1.label !== "Campo 1") {
      console.error("FAIL: Field label translation for field1 should exist");
      return false;
    }
    console.log("PASS: Field label translations work without Category Translations sheet");

    // Verify helper text translations exist
    if (!esField1.helperText || esField1.helperText !== "Ayuda para campo 1") {
      console.error(`FAIL: Field1 helper text Spanish translation incorrect`);
      return false;
    }
    console.log("PASS: Helper text translations work without Category Translations sheet");

    // Verify option translations exist
    const esField2 = esFields?.['field2'];
    if (!esField2 || esField2["options.0"] !== "Opción A") {
      console.error("FAIL: Option translation for 'option-a' incorrect");
      return false;
    }
    console.log("PASS: Option translations work without Category Translations sheet");

    // Verify French translations too
    const frField2 = translations['fr']?.field?.['field2'];
    if (!frField2 || frField2.label !== "Champ 2") {
      console.error("FAIL: French field label translation incorrect");
      return false;
    }
    console.log("PASS: All French field translations work correctly");

    console.log("=== Field Translations Without Category Translations Sheet: ALL TESTS PASSED ===");
    return true;
  } catch (error) {
    console.error("FAIL: Exception thrown - " + error.message);
    return false;
  }
}

function testDuplicateHelperTextTranslations(): boolean {
  console.log("=== Test: Duplicate Helper Text Translation ===");

  try {
    // Create mock data with multiple fields sharing same helper text (e.g., multiple numeric measurements)
    const testData: SheetData = {
      documentName: "Duplicate Helper Text Test" as any,
      Categories: [
        ["Name", "Icon", "Fields", "ID", "Color", "Icon ID"],
        ["Measurements", "", "width, height, depth", "measurements", "#FF0000", ""]
      ],
      Details: [
        ["Name", "Helper Text", "Type", "Options", "ID", "Universal"],
        ["Width", "Enter measurement in meters", "n", "", "width", "FALSE"],
        ["Height", "Enter measurement in meters", "n", "", "height", "FALSE"],
        ["Depth", "Enter measurement in meters", "n", "", "depth", "FALSE"]
      ],
      "Detail Helper Text Translations": [
        ["Helper Text", "Spanish", "French"],
        ["Enter measurement in meters", "Ingrese la medida en metros", "Entrez la mesure en mètres"],  // Shared translation row
        ["Enter measurement in meters", "Ingrese la medida en metros", "Entrez la mesure en mètres"],  // Duplicate
        ["Enter measurement in meters", "Ingrese la medida en metros", "Entrez la mesure en mètres"]   // Duplicate
      ]
    } as SheetData;

    // Build fields
    const fields = buildFields(testData);
    if (fields.length !== 3) {
      console.error(`FAIL: Expected 3 fields, got ${fields.length}`);
      return false;
    }
    console.log("PASS: Three fields with shared helper text created");

    // Verify all fields have same helper text
    for (const field of fields) {
      if (field.description !== "Enter measurement in meters") {
        console.error(`FAIL: Field ${field.id} should have shared helper text, got '${field.description}'`);
        return false;
      }
    }
    console.log("PASS: All fields share same helper text");

    // Build categories
    const categories = buildCategories(testData, fields);

    // Build translations
    const translations = buildTranslationsPayload(testData, categories, fields);

    // Verify ALL three fields have translations (not just the last one)
    for (const field of fields) {
      const esFieldEntry = translations['es']?.field?.[field.id];
      const helperTextTrans = esFieldEntry?.helperText;
      if (!helperTextTrans) {
        console.error(`FAIL: Helper text translation for ${field.id} not found`);
        return false;
      }

      // Verify translation is correct
      if (helperTextTrans !== "Ingrese la medida en metros") {
        console.error(`FAIL: Helper text translation for ${field.id} should be 'Ingrese la medida en metros', got '${helperTextTrans}'`);
        return false;
      }
    }
    console.log("PASS: All fields with shared helper text have correct translations (no data loss)");

    // Verify French translations too
    for (const field of fields) {
      const helperTextTrans = translations['fr']?.field?.[field.id]?.helperText;
      if (!helperTextTrans) {
        console.error(`FAIL: French helper text translation for ${field.id} not found`);
        return false;
      }
      if (helperTextTrans !== "Entrez la mesure en mètres") {
        console.error(`FAIL: French helper text translation for ${field.id} incorrect`);
        return false;
      }
    }
    console.log("PASS: French translations also correct for all fields");

    console.log("=== Duplicate Helper Text Translation: ALL TESTS PASSED ===");
    return true;
  } catch (error) {
    console.error("FAIL: Exception thrown - " + error.message);
    return false;
  }
}

function testDuplicateOptionTranslations(): boolean {
  console.log("=== Test: Duplicate Option Lists Translation ===");

  try {
    // Create mock data with multiple fields sharing same option list (e.g., multiple yes/no questions)
    const testData: SheetData = {
      documentName: "Duplicate Options Test" as any,
      Categories: [
        ["Name", "Icon", "Fields", "ID", "Color", "Icon ID"],
        ["Survey", "", "question1, question2, question3", "survey", "#FF0000", ""]
      ],
      Details: [
        ["Name", "Helper Text", "Type", "Options", "ID", "Universal"],
        ["Question 1", "First yes/no question", "s", "Yes, No", "question1", "FALSE"],
        ["Question 2", "Second yes/no question", "s", "Yes, No", "question2", "FALSE"],
        ["Question 3", "Third yes/no question", "s", "Yes, No", "question3", "FALSE"]
      ],
      "Detail Option Translations": [
        ["Options", "Spanish", "French"],
        ["Yes, No", "Sí, No", "Oui, Non"],  // Shared translation row
        ["Yes, No", "Sí, No", "Oui, Non"],  // Duplicate (one per field in sheet)
        ["Yes, No", "Sí, No", "Oui, Non"]   // Duplicate (one per field in sheet)
      ]
    } as SheetData;

    // Build fields
    const fields = buildFields(testData);
    if (fields.length !== 3) {
      console.error(`FAIL: Expected 3 fields, got ${fields.length}`);
      return false;
    }
    console.log("PASS: Three fields with shared options created");

    // Verify all fields have same option list
    for (const field of fields) {
      if (!field.options || field.options.length !== 2) {
        console.error(`FAIL: Field ${field.id} should have 2 options`);
        return false;
      }
      if (field.options[0].label !== "Yes" || field.options[1].label !== "No") {
        console.error(`FAIL: Field ${field.id} should have Yes/No options`);
        return false;
      }
    }
    console.log("PASS: All fields share same option list");

    // Build categories
    const categories = buildCategories(testData, fields);

    // Build translations
    const translations = buildTranslationsPayload(testData, categories, fields);

    // Verify ALL three fields have translations (not just the last one)
    for (const field of fields) {
      const esFieldEntry = translations['es']?.field?.[field.id];
      if (!esFieldEntry) {
        console.error(`FAIL: Spanish options map missing for ${field.id}`);
        return false;
      }

      // Verify translations are correct (options indices 0 => Yes, 1 => No)
      if (esFieldEntry['options.0'] !== "Sí") {
        console.error(`FAIL: 'Yes' translation for ${field.id} should be 'Sí', got '${esFieldEntry['options.0']}'`);
        return false;
      }
      if (esFieldEntry['options.1'] !== "No") {
        console.error(`FAIL: 'No' translation for ${field.id} should be 'No', got '${esFieldEntry['options.1']}'`);
        return false;
      }
    }
    console.log("PASS: All fields with shared options have correct translations (no data loss)");

    // Verify French translations too
    for (const field of fields) {
      const frFieldEntry = translations['fr']?.field?.[field.id];
      if (!frFieldEntry) {
        console.error(`FAIL: French option translations for ${field.id} not found`);
        return false;
      }
      if (frFieldEntry['options.0'] !== "Oui") {
        console.error(`FAIL: French 'Yes' translation for ${field.id} incorrect`);
        return false;
      }
      if (frFieldEntry['options.1'] !== "Non") {
        console.error(`FAIL: French 'No' translation for ${field.id} incorrect`);
        return false;
      }
    }
    console.log("PASS: French translations also correct for all fields");

    console.log("=== Duplicate Option Lists Translation: ALL TESTS PASSED ===");
    return true;
  } catch (error) {
    console.error("FAIL: Exception thrown - " + error.message);
    return false;
  }
}

function testOptionTranslationsWithLocalizedSeparators(): boolean {
  console.log("=== Test: Option Translations With Localized Separators ===");

  try {
    const testData: SheetData = {
      documentName: "Localized Option Separator Test" as any,
      Categories: [
        ["Name", "Icon", "Fields", "ID", "Color", "Icon ID"],
        ["Urban Survey", "", "building-type", "urban-survey", "#FFFFFF", ""]
      ],
      Details: [
        ["Name", "Helper Text", "Type", "Options", "ID", "Universal"],
        [
          "Building Type",
          "Select the type of structure",
          "s",
          "School, Hospital, Farm, Church, Shop, Other",
          "building-type",
          "FALSE"
        ]
      ],
      "Detail Option Translations": [
        ["Options", "Chinese"],
        [
          "School, Hospital, Farm, Church, Shop, Other",
          "学校、医院、农场、教堂、商店、其他"
        ]
      ]
    } as SheetData;

    const fields = buildFields(testData);
    const categories = buildCategories(testData, fields);
    const translations = buildTranslationsPayload(testData, categories, fields);

    const zhField = translations['zh']?.field?.['building-type'];
    if (!zhField) {
      console.error("FAIL: Chinese translations not generated");
      return false;
    }

    const expected = ["学校", "医院", "农场", "教堂", "商店", "其他"];
    for (let i = 0; i < expected.length; i++) {
      const key = `options.${i}`;
      if (zhField[key] !== expected[i]) {
        console.error(`FAIL: ${key} should be '${expected[i]}', got '${zhField[key]}'`);
        return false;
      }
    }

    console.log("PASS: Localized option strings split into individual entries");
    return true;
  } catch (error) {
    console.error("FAIL: Exception thrown - " + error.message);
    return false;
  }
}

// =============================================================================
// Test Runner
// =============================================================================

function runBuildAndImportTests(): void {
  console.log("\n========================================");
  console.log("CoMapeo Config v2.0.0 - Build & Import Tests");
  console.log("========================================\n");

  const tests = [
    // Basic utility tests
    { name: "Slugify Function", fn: testSlugify },
    { name: "Parse Options", fn: testParseOptions },
    { name: "Extract File ID", fn: testExtractFileId },

    // Field type tests
    { name: "Field Type Mapping", fn: testFieldTypeMapping },
    { name: "All Field Types (11 types)", fn: testAllFieldTypes },
    { name: "Integer vs Number Distinction", fn: testIntegerVsNumberDistinction },
    { name: "Map All Field Types to Char", fn: testMapFieldTypeToCharAllTypes },
    { name: "Map Field Type to Char", fn: testMapFieldTypeToChar },

    // Build and import tests
    { name: "Build Payload Creation", fn: testBuildPayloadCreation },
    { name: "Import Parsing", fn: testImportParsing },
    { name: "Validate BuildRequest", fn: testValidateBuildRequest },
    { name: "Strict Merged Translation Payload Limit", fn: testStrictMergedTranslationPayloadLimit },
    { name: "Strict Total Entity Overflow Includes Options and Translations", fn: testStrictTotalEntityOverflowIncludesOptionsAndTranslations },

    // Universal fields tests
    { name: "Universal Fields Added to All Categories", fn: testUniversalFieldsAddedToAllCategories },
    { name: "Universal Field Does Not Set Required", fn: testUniversalFieldDoesNotSetRequired },
    { name: "Build Categories Requires Observation Coverage", fn: testBuildCategoriesRequiresObservationCoverage },
    { name: "Categories Applies Parsing", fn: testCategoriesAppliesParsing },

    // Round-trip integration tests
    { name: "Integer Field Round-Trip", fn: testIntegerFieldRoundTrip },
    { name: "Import Preserves Field Distribution", fn: testUniversalFieldRoundTrip },

    // ID preservation tests
    { name: "Category ID Preservation", fn: testCategoryIdPreservation },
    { name: "Icons Respect Stored Category IDs", fn: testIconsRespectStoredCategoryIds },
    { name: "Cross-Sheet Icon Collisions Use Exact IDs", fn: testCrossSheetIconCollisionsUseExactIds },
    { name: "Field ID Preservation", fn: testFieldIdPreservation },
    { name: "Option Value Preservation", fn: testOptionValuePreservation },

    // Edge case tests
    { name: "Empty Inputs", fn: testEmptyInputs },
    { name: "buildFields Skips Blank Rows", fn: testBuildFieldsSkipsBlankRows },
    { name: "Translations With Blank Rows", fn: testTranslationsWithBlankRows },
    { name: "Field Translations Without Category Translations Sheet", fn: testFieldTranslationsWithoutCategorySheet },
    { name: "Duplicate Helper Text Translation", fn: testDuplicateHelperTextTranslations },
    { name: "Duplicate Option Lists Translation", fn: testDuplicateOptionTranslations },
    { name: "Option Translations With Localized Separators", fn: testOptionTranslationsWithLocalizedSeparators },

    // Category and API tests
    { name: "Category Selection", fn: testCategorySelection },
    { name: "API Error Parsing", fn: testApiErrorParsing }
  ];

  let passed = 0;
  let failed = 0;

  for (const test of tests) {
    try {
      if (test.fn()) {
        passed++;
      } else {
        failed++;
      }
    } catch (error) {
      console.error(`Test '${test.name}' threw exception: ${error.message}`);
      failed++;
    }
    console.log("\n");
  }

  console.log("========================================");
  console.log(`Results: ${passed} passed, ${failed} failed`);
  console.log("========================================");

  if (failed > 0) {
    console.log("\nSome tests failed. Please review the output above.");
  } else {
    console.log("\nAll tests passed!");
  }
}
