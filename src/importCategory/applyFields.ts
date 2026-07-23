/// <reference path="../utils.ts" />

/**
 * Fields application functions for the import category functionality.
 * This file contains functions related to applying fields to the spreadsheet.
 */

/**
 * Applies fields (details) to the Details sheet.
 * @param sheet - The details sheet
 * @param fields - Array of field objects
 */
function applyFields(sheet: GoogleAppsScript.Spreadsheet.Sheet, fields: any[]) {
  console.log(`Applying ${fields.length} fields to Details sheet`);
  console.log("Fields", fields);
  // Set headers
  sheet
    .getRange(1, 1, 1, 4)
    .setValues([["Label", "Helper Text", "Type", "Options"]]);
  sheet.getRange(1, 1, 1, 4).setFontWeight("bold");

  // Prepare field rows
  const fieldRows = fields.map((field) => {
    // Convert field type to spreadsheet format
    let typeStr = "text";
    if (field.type === "selectOne" || field.type === "single")
      typeStr = "single";
    if (field.type === "selectMultiple" || field.type === "multiple")
      typeStr = "multiple";
    if (field.type === "number") typeStr = "number";

    // Convert options to comma-separated string
    let optionsStr = "";
    if (field.options && field.options.length > 0) {
      optionsStr = field.options
        .map((opt: any) => {
          // Handle both string and object formats
          if (typeof opt === "string") return opt;
          if (typeof opt === "object" && opt !== null) {
            const label = opt.label || opt.name || "";
            const value = opt.value || "";
            if (!label) return value;
            // Preserve an explicit value that differs from the canonical label
            // form by writing it back as "value:label". Otherwise a legacy
            // config whose stored value predates canonicalizeOptionValue (e.g.
            // diacritic-folded "cafe" for label "Café") would be re-written as
            // the bare label and drift to "café" on the next export.
            return value && value !== canonicalizeOptionValue(label)
              ? `${value}:${label}`
              : label;
          }
          return "";
        })
        .filter(Boolean) // Remove empty strings
        .join(", ");
    }

    // Strip commas from label
    const sanitizedLabel = field.label ? field.label.replace(/,/g, "") : "";

    return [sanitizedLabel, field.helperText || "", typeStr, optionsStr];
  });

  try {
    // Clear the entire sheet to remove any data validations
    sheet.clear();

    // Set headers again after clearing
    sheet
      .getRange(1, 1, 1, 4)
      .setValues([["Label", "Helper Text", "Type", "Options"]]);
    sheet.getRange(1, 1, 1, 4).setFontWeight("bold");

    // Add field rows one by one to avoid validation errors
    if (fieldRows.length > 0) {
      for (let i = 0; i < fieldRows.length; i++) {
        try {
          // Set each cell individually to avoid validation errors
          for (let j = 0; j < 4; j++) {
            sheet.getRange(i + 2, j + 1).setValue(fieldRows[i][j]);
          }
        } catch (rowError) {
          console.error(`Error setting field row ${i + 2}:`, rowError);
        }
      }
    }
  } catch (error) {
    console.error("Error in applyFields:", error);
  }

  // Auto-resize columns for better readability
  sheet.autoResizeColumns(1, 4);
}
