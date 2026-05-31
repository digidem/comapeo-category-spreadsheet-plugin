interface SheetSnapshot {
  exists: boolean;
  data: any[][];
}

const IMPORT_REGRESSION_SHEETS = [
  "Categories",
  "Details",
  "Metadata",
  "Category Translations",
  "Detail Label Translations",
  "Detail Helper Text Translations",
  "Detail Option Translations",
];

function getOrCreateSheet(name: string): GoogleAppsScript.Spreadsheet.Sheet {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(name);
  return sheet || spreadsheet.insertSheet(name);
}

function snapshotSheets(
  sheetNames: string[],
): Record<string, SheetSnapshot> {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const snapshots: Record<string, SheetSnapshot> = {};

  sheetNames.forEach((name) => {
    const sheet = spreadsheet.getSheetByName(name);
    if (!sheet) {
      snapshots[name] = { exists: false, data: [] };
      return;
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow === 0 || lastCol === 0) {
      snapshots[name] = { exists: true, data: [] };
      return;
    }

    const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    snapshots[name] = { exists: true, data };
  });

  return snapshots;
}

function restoreSheets(
  snapshots: Record<string, SheetSnapshot>,
): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  Object.entries(snapshots).forEach(([name, snapshot]) => {
    const existingSheet = spreadsheet.getSheetByName(name);

    if (!snapshot.exists) {
      if (existingSheet && spreadsheet.getSheets().length > 1) {
        spreadsheet.deleteSheet(existingSheet);
      } else if (existingSheet) {
        existingSheet.clear();
      }
      return;
    }

    const sheet = existingSheet || spreadsheet.insertSheet(name);
    sheet.clear();

    if (snapshot.data.length > 0 && snapshot.data[0].length > 0) {
      sheet
        .getRange(1, 1, snapshot.data.length, snapshot.data[0].length)
        .setValues(snapshot.data);
    }
  });
}

function buildValidRegressionArchive(): GoogleAppsScript.Base.Blob {
  const metadataBlob = Utilities.newBlob(
    JSON.stringify({ name: "Regression Dataset", version: "1.0.0" }, null, 2),
    "application/json",
    "metadata.json",
  );

  const presetsBlob = Utilities.newBlob(
    JSON.stringify(
      {
        presets: {
          "category/regression": {
            name: "Regression Category",
            icon: "regression-icon",
            color: "#00FFAA",
            fields: ["field/regression"],
          },
        },
        fields: {
          "field/regression": {
            label: "Regression Field",
            type: "select_one",
            options: ["Yes", "No"],
            helperText: "Helper",
          },
        },
      },
      null,
      2,
    ),
    "application/json",
    "presets.json",
  );

  const translationsBlob = Utilities.newBlob(
    JSON.stringify(
      {
        en: {
          presets: {
            "category/regression": {
              name: "Regression Category",
            },
          },
        },
      },
      null,
      2,
    ),
    "application/json",
    "translations.json",
  );

  const iconSprite =
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">' +
    '<symbol id="regression-icon"><circle cx="12" cy="12" r="10" fill="#00FFAA"/></symbol>' +
    "</svg>";
  const iconsBlob = Utilities.newBlob(
    iconSprite,
    "image/svg+xml",
    "icons.svg",
  );

  const archive = Utilities.zip(
    [metadataBlob, presetsBlob, translationsBlob, iconsBlob],
    "regression.comapeocat",
  );
  archive.setContentType("application/zip");
  archive.setName("regression.comapeocat");
  return archive;
}

function buildInvalidRegressionArchive(): GoogleAppsScript.Base.Blob {
  const metadataBlob = Utilities.newBlob(
    JSON.stringify({ name: "Invalid Regression" }, null, 2),
    "application/json",
    "metadata.json",
  );
  const archive = Utilities.zip([metadataBlob], "invalid-regression.comapeocat");
  archive.setContentType("application/zip");
  archive.setName("invalid-regression.comapeocat");
  return archive;
}

function testImportRegressionSuccess() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const snapshots = snapshotSheets(IMPORT_REGRESSION_SHEETS);

  try {
    IMPORT_REGRESSION_SHEETS.forEach((name) => {
      const sheet = getOrCreateSheet(name);
      sheet.clear();
    });

    const archiveBlob = buildValidRegressionArchive();
    const base64 = Utilities.base64Encode(archiveBlob.getBytes());
    const result = processImportedCategoryFile(archiveBlob.getName(), base64);

    if (!result.success) {
      throw new Error(result.message || "Import failed unexpectedly");
    }

    const categoriesSheet = spreadsheet.getSheetByName("Categories");
    const detailsSheet = spreadsheet.getSheetByName("Details");

    if (!categoriesSheet || categoriesSheet.getLastRow() < 2) {
      throw new Error("Categories sheet missing imported rows");
    }

    const categoryRow = categoriesSheet.getRange(2, 1, 1, 3).getValues()[0];
    if (categoryRow[0] !== "Regression Category") {
      throw new Error("Category name mismatch after import");
    }

    if (!detailsSheet || detailsSheet.getLastRow() < 2) {
      throw new Error("Details sheet missing imported fields");
    }

    const detailRow = detailsSheet.getRange(2, 1, 1, 4).getValues()[0];
    if (detailRow[0] !== "Regression Field") {
      throw new Error("Detail label mismatch after import");
    }

    return {
      success: true,
      message: "Regression success import verified",
    };
  } catch (error) {
    return {
      success: false,
      message: `Regression success test failed: ${error instanceof Error ? error.message : error}`,
    };
  } finally {
    restoreSheets(snapshots);
  }
}

function testImportRegressionValidationFailure() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const snapshots = snapshotSheets(IMPORT_REGRESSION_SHEETS);

  try {
    const categoriesSheet = getOrCreateSheet("Categories");
    categoriesSheet.clear();
    categoriesSheet.getRange(1, 1, 1, 3).setValues([["Name", "Icon", "Fields"]]);
    categoriesSheet.getRange(2, 1, 1, 3).setValues([["Sentinel Category", "", "Sentinel Field"]]);

    const archiveBlob = buildInvalidRegressionArchive();
    const base64 = Utilities.base64Encode(archiveBlob.getBytes());
    const result = processImportedCategoryFile(archiveBlob.getName(), base64);

    if (result.success) {
      throw new Error("Invalid archive should not succeed");
    }

    const sentinelValue = categoriesSheet.getRange(2, 1).getValue();
    if (sentinelValue !== "Sentinel Category") {
      throw new Error("Categories sheet was modified on validation failure");
    }

    return {
      success: true,
      message: "Validation failure preserves spreadsheet data",
    };
  } catch (error) {
    return {
      success: false,
      message: `Regression validation test failed: ${error instanceof Error ? error.message : error}`,
    };
  } finally {
    restoreSheets(snapshots);
  }
}
