/// <reference path="../loggingHelpers.ts" />

/**
 * Processes field definitions from spreadsheet Details sheet
 *
 * Validates all field definitions and converts them to CoMapeo field format.
 * Fields include tagKey, type, label, helperText, options, and universal flag.
 * Validates field types, required columns, and option formats before processing.
 *
 * @param data - Spreadsheet data object containing Details sheet
 * @returns Array of CoMapeo field objects
 * @throws Error if field validation fails with details of all validation errors
 *
 * @example
 * const data = getSpreadsheetData();
 * const fields = processFields(data);
 * // Returns: [{ tagKey: "name", type: "text", label: "Name", helperText: "Enter name", options: [], universal: false }, ...]
 */
function processFields(data) {
  const details = data.Details.slice(1);
  const errors: string[] = [];
  const warnings: string[] = [];

  // Validate all fields first
  details.forEach((detail, index) => {
    const validation = validateFieldDefinition(detail, index + 2); // +2 for header row and 0-index
    if (!validation.valid && validation.error) {
      errors.push(validation.error);
    }
    if (validation.warnings) {
      warnings.push(...validation.warnings);
    }
  });

  // If there are validation errors, throw with details
  if (errors.length > 0) {
    throw new Error(
      `Field validation failed with ${errors.length} error(s) in Details sheet:\n${errors.join("\n")}`,
    );
  }

  // Log warnings if any
  if (warnings.length > 0) {
    const logger = getScopedLogger("ProcessFields");
    logger.warn(
      `Field validation warnings (${warnings.length}):\n${warnings.join("\n")}`,
    );
  }

  return details.map((detail, index) => {
    const tagKey = createFieldTagKey(detail[0], index);

    return {
      tagKey,
      type: getFieldType(detail[2]),
      label: detail[0],
      helperText: detail[1],
      options: getFieldOptions(detail[2], detail[3], tagKey),
      universal: detail[5] === "TRUE",
    };
  });
}
