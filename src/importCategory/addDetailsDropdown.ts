/**
 * Functions for adding data validation dropdowns to the spreadsheet.
 * This file contains functions to add dropdowns to specific columns after import.
 */

/**
 * Adds a dropdown to the Details column in the Categories sheet.
 * The dropdown options are populated from the Label column in the Details sheet.
 * This function should be called after the import process is complete to avoid validation errors.
 *
 * Note: This implementation uses standard single-select dropdowns as a fallback
 * since multi-select chip-style dropdowns are not yet fully supported in Apps Script.
 * When Google adds official support for multi-select validations, this code should be updated.
 */
function addDetailsDropdown(): void {
  try {
    console.log("Adding dropdown to Details column in Categories sheet...");

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const categoriesSheet = spreadsheet.getSheetByName("Categories");
    const detailsSheet = spreadsheet.getSheetByName("Details");

    if (!categoriesSheet || !detailsSheet) {
      console.error("Categories or Details sheet not found");
      return;
    }

    // First, clear any existing data validations to prevent conflicts
    try {
      if (categoriesSheet.getLastRow() > 1) {
        // Clear validations from the Details column (column 3)
        categoriesSheet
          .getRange(2, 3, categoriesSheet.getLastRow() - 1, 1)
          .clearDataValidations();
      }
    } catch (clearError) {
      console.error("Error clearing existing validations:", clearError);
      // Continue anyway
    }

    // Check if Details sheet has data
    const detailsLastRow = Math.max(detailsSheet.getLastRow(), 2);
    if (detailsLastRow <= 1) {
      console.log("Details sheet is empty or contains only header");
      return;
    }

    // Get the source range for dropdown values (entire column A in Details sheet)
    const valueRange = detailsSheet.getRange("A:A");

    // Create standard validation rule (single-select as fallback)
    // Note: Multi-select chip display is not yet supported in Apps Script
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(valueRange, true) // true = show dropdown
      .setAllowInvalid(true)
      // .setAllowMultipleSelections(true)
      // .setDisplayStyle(SpreadsheetApp.DataValidationDisplayStyle.CHIP)
      .build();

    // Apply validation to the Details column in Categories sheet
    const categoriesLastRow = Math.max(categoriesSheet.getLastRow(), 2);
    if (categoriesLastRow <= 1) {
      console.log("Categories sheet is empty or contains only header");
      return;
    }

    // Apply validation to all cells in the Details column at once
    try {
      const detailsColumn = categoriesSheet.getRange(
        2,
        3,
        categoriesLastRow - 1,
        1,
      );
      detailsColumn.clearDataValidations();
      detailsColumn.setDataValidation(rule);

      // Add note to column header about comma-separation for multiple values
      categoriesSheet
        .getRange(1, 3)
        .setNote(
          "Multiple selections: Separate field labels with commas (e.g., 'Name, Type, Notes').\n\n" +
            "The dropdown validates against field labels from the Details sheet.\n\n" +
            "Note: Multi-select chip display will be enabled when Google Apps Script adds official support.",
        );

      // Set up an installable onEdit trigger for comma-separated values
      setupMultiSelectFallbackTrigger();
    } catch (validationError) {
      console.error(
        "Error setting validation for Details column:",
        validationError,
      );

      // Fallback: Apply validation one cell at a time if bulk operation fails
      for (let row = 2; row <= categoriesLastRow; row++) {
        try {
          const cell = categoriesSheet.getRange(row, 3);
          cell.clearDataValidations();
          cell.setDataValidation(rule);
        } catch (cellError) {
          console.error(
            `Error setting validation for cell C${row}:`,
            cellError,
          );
        }
      }
    }

    console.log(
      "Dropdown added successfully to Details column in Categories sheet (single-select fallback mode)",
    );
  } catch (error) {
    console.error("Error adding dropdown to Details column:", error);
  }
}

/**
 * Sets up an installable onEdit trigger to handle comma-separated values
 * as a fallback for multi-select dropdowns.
 * This is a temporary solution until Google Apps Script supports multi-select validations.
 */
function setupMultiSelectFallbackTrigger(): void {
  try {
    // Check if the trigger already exists
    const triggers = ScriptApp.getProjectTriggers();
    const triggerExists = triggers.some(
      (trigger) =>
        trigger.getHandlerFunction() === "handleMultiSelectEdit" &&
        trigger.getEventType() === ScriptApp.EventType.ON_EDIT,
    );

    if (!triggerExists) {
      ScriptApp.newTrigger("handleMultiSelectEdit")
        .forSpreadsheet(SpreadsheetApp.getActive())
        .onEdit()
        .create();
      console.log("Multi-select fallback trigger installed");
    }
  } catch (error) {
    console.error("Error setting up multi-select fallback trigger:", error);
  }
}

/**
 * Handles edits to the Details column to support comma-separated values
 * as a fallback for multi-select dropdowns.
 * @param e The onEdit event object
 */
function handleMultiSelectEdit(e: GoogleAppsScript.Events.SheetsOnEdit): void {
  try {
    const sheet = e.range.getSheet();
    if (sheet.getName() !== "Categories") return;

    // Only process edits to column 3 (Details column)
    if (e.range.getColumn() !== 3 || e.range.getRow() <= 1) return;

    const value = e.range.getValue();
    console.log(
      `Details column edit detected: row=${e.range.getRow()}, value="${value}"`,
    );

    if (!value || typeof value !== "string") return;

    // Get valid options from Details sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const detailsSheet = ss.getSheetByName("Details");
    if (!detailsSheet) {
      console.warn("Details sheet not found for validation");
      return;
    }

    const lastRow = detailsSheet.getLastRow();
    if (lastRow <= 1) {
      console.warn("Details sheet is empty");
      return;
    }

    const validOptions = detailsSheet
      .getRange("A2:A" + lastRow)
      .getValues()
      .flat()
      .filter((val) => val && typeof val === "string")
      .map((val) => val.toString().trim());

    console.log(`Valid options from Details sheet: ${validOptions.length} items`);

    if (validOptions.length === 0) return;

    // Process comma-separated values
    const selectedValues = value.split(",").map((v) => v.trim());
    console.log(`Processing ${selectedValues.length} selected values`);

    // Validate each value (case-insensitive)
    const validatedValues = selectedValues.filter((val) => {
      const isValid = validOptions.some(
        (option) => option.toLowerCase() === val.toLowerCase(),
      );
      if (!isValid) {
        console.warn(`Invalid value: "${val}" (not in Details sheet)`);
      }
      return isValid;
    });

    console.log(`Validated ${validatedValues.length} values`);

    // Update the cell with validated values if changed
    const newValue = validatedValues.join(", ");
    if (validatedValues.length > 0 && newValue !== value) {
      console.log(`Updating cell from "${value}" to "${newValue}"`);
      e.range.setValue(newValue);
    } else if (validatedValues.length === 0) {
      console.warn(`All values invalid, keeping original: "${value}"`);
    }
  } catch (error) {
    console.error("Error in multi-select edit handler:", error);
  }
}

/**
 * Adds all dropdowns to the spreadsheet.
 * This function should be called after the import process is complete.
 */
function addAllDropdowns(): void {
  try {
    console.log("Adding all dropdowns to the spreadsheet...");

    // Add dropdown to Details column in Categories sheet
    try {
      addDetailsDropdown();
    } catch (dropdownError) {
      console.error("Error adding Details dropdown:", dropdownError);
      // Continue anyway
    }

    console.log("All dropdowns added successfully");
  } catch (error) {
    console.error("Error adding dropdowns:", error);
  }
}

/**
 * Checks if multi-select validation is supported in the current Apps Script environment.
 * This function can be used to determine when to switch from the fallback to the official implementation.
 *
 * @returns {boolean} True if multi-select validation is supported, false otherwise
 */
function isMultiSelectValidationSupported(): boolean {
  try {
    // Attempt to create a validation rule with multi-select
    // This will throw an error if not supported
    const testRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(SpreadsheetApp.getActiveSpreadsheet().getRange("A1"))
      .setAllowInvalid(true);

    // Check if the method exists
    return typeof (testRule as unknown as Record<string, unknown>)["setAllowMultipleSelections"] === "function";
  } catch (error) {
    return false;
  }
}
