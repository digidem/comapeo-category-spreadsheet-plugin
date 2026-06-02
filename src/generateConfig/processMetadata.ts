/**
 * Processes metadata from spreadsheet data and generates metadata objects
 *
 * Creates or retrieves a Metadata sheet and generates both metadata
 * and package.json objects for CoMapeo configuration.
 *
 * @param data - Spreadsheet data object containing documentName
 * @returns Object containing metadata and packageJson
 *
 * @example
 * const data = getSpreadsheetData();
 * const { metadata, packageJson } = processMetadata(data);
 * // metadata: { dataset_id, name, version }
 * // packageJson: { name, version, description, dependencies, scripts }
 */
function processMetadata(data: { documentName: string }) {
  const { documentName } = data;
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let metadataSheet = spreadsheet.getSheetByName("Metadata");

  if (!metadataSheet) {
    metadataSheet = createMetadataSheet(spreadsheet);
  }

  const metadata = getOrCreateMetadata(metadataSheet, documentName);
  const packageJson = createPackageJson(metadata);

  return { metadata, packageJson };
}

/**
 * Creates a new Metadata sheet in the spreadsheet
 *
 * @param spreadsheet - The active spreadsheet
 * @returns The newly created Metadata sheet
 *
 * @example
 * const sheet = createMetadataSheet(SpreadsheetApp.getActiveSpreadsheet());
 * // Creates sheet with "Key" and "Value" headers
 */
function createMetadataSheet(
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
): GoogleAppsScript.Spreadsheet.Sheet {
  const sheet = spreadsheet.insertSheet("Metadata");
  const headerRange = sheet.getRange(1, 1, 1, 2);
  headerRange.setValues([["Key", "Value"]]);
  headerRange.setFontWeight("bold");
  return sheet;
}

/**
 * Retrieves or creates metadata from the Metadata sheet
 *
 * Reads existing metadata values from the sheet or generates defaults.
 * Version is always updated to current date in yy.MM.dd format.
 *
 * @param sheet - The Metadata sheet
 * @param documentName - Name of the spreadsheet document
 * @returns CoMapeoMetadata object with dataset_id, name, and version
 *
 * @example
 * const metadata = getOrCreateMetadata(sheet, "Wildlife Survey");
 * // Returns: { dataset_id: "comapeo-wildlife-survey", name: "config-wildlife-survey", version: "25.10.12" }
 */
function getOrCreateMetadata(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  documentName: string,
): CoMapeoMetadata {
  const metadata: CoMapeoMetadata = {
    dataset_id: getOrSetValue(
      sheet,
      "dataset_id",
      `comapeo-${slugify(documentName)}`,
    ),
    name: getOrSetValue(sheet, "name", `config-${slugify(documentName)}`),
    version: getOrSetValue(
      sheet,
      "version",
      Utilities.formatDate(new Date(), "UTC", "yy.MM.dd"),
    ),
  };
  return metadata;
}

/**
 * Retrieves or sets a value in the metadata sheet.
 *
 * @param sheet - The metadata sheet to operate on
 * @param key - The key to search for or set
 * @param defaultValue - The default value to set if the key is not found
 * @returns The existing value if found, or the default value if not found
 */
function getOrSetValue(sheet: GoogleAppsScript.Spreadsheet.Sheet, key: string, defaultValue: string): string {
  // Get all data from the sheet
  const data = sheet.getDataRange().getValues();

  // Search for the key in the existing data
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      // If key is found, return its corresponding value
      if (key === "version") {
        // Always update version with current date
        const newVersion = Utilities.formatDate(new Date(), "UTC", "yy.MM.dd");
        sheet.getRange(i + 1, 2).setValue(newVersion);
        return newVersion;
      }
      return data[i][1];
    }
  }

  // If key is not found, append a new row with the key and default value
  sheet.appendRow([key, defaultValue]);

  // Return the default value
  return defaultValue;
}

/**
 * Generates random hexadecimal bytes string
 *
 * @param length - Number of hexadecimal characters to generate
 * @returns Random hex string of specified length
 *
 * @example
 * generateRandomBytes(8) // "a3f5c2d1"
 */
function generateRandomBytes(length: number): string {
  const chars = "0123456789abcdef";
  let result = "";
  for (let i = 0; i < length; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
}

/**
 * Creates a package.json object for CoMapeo configuration
 *
 * Generates a valid package.json structure with dependencies and build scripts
 * for the CoMapeo configuration package.
 *
 * @param metadata - CoMapeo metadata object containing dataset_id, name, and version
 * @returns Package.json object with name, version, description, dependencies, and scripts
 *
 * @example
 * const metadata = { dataset_id: "comapeo-wildlife", name: "Wildlife Survey", version: "25.10.12" };
 * const pkg = createPackageJson(metadata);
 * // Returns: { name: "comapeo-wildlife", version: "25.10.12", ... }
 */
function createPackageJson(metadata: CoMapeoMetadata) {
  return {
    name: metadata.dataset_id,
    version: metadata.version,
    description: `CoMapeo configuration for ${metadata.name}`,
    dependencies: {
      "mapeo-settings-builder": "^6.0.0",
    },
    scripts: {
      build:
        "mkdir -p build && mapeo-settings build -l 'en' -o build/${npm_package_name}-v${npm_package_version}.comapeocat",
      lint: "mapeo-settings lint",
    },
  };
}
