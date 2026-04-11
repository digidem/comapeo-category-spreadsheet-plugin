/// <reference path="../types.ts" />
/// <reference path="../loggingHelpers.ts" />

/**
 * Processes preset data from spreadsheet
 *
 * @param data - Spreadsheet data object
 * @param categoriesSheet - Categories sheet reference (passed to avoid uncached access)
 * @param fields - Processed CoMapeo fields (used for canonical key mapping)
 * @returns Array of CoMapeo preset objects
 */
function processPresets(
  data,
  categoriesSheet: GoogleAppsScript.Spreadsheet.Sheet,
  fields: CoMapeoField[],
) {
  const categories = data["Categories"].slice(1);
  const errors: string[] = [];
  const warnings: string[] = [];

  const fieldKeyByLabel = fields.reduce<Record<string, string>>((acc, field) => {
    if (field && field.label) {
      acc[field.label] = field.tagKey;
    }
    return acc;
  }, {});

  // Validate all categories first
  categories.forEach((category, index) => {
    const validation = validateCategoryDefinition(category, index + 2); // +2 for header row and 0-index
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
      `Category validation failed with ${errors.length} error(s) in Categories sheet:\n${errors.join("\n")}`,
    );
  }

  // Log warnings if any
  const logger = getScopedLogger("ProcessPresets");

  if (warnings.length > 0) {
    logger.warn(
      `Category validation warnings (${warnings.length}):\n${warnings.join("\n")}`,
    );
  }

  const backgroundColors = categoriesSheet
    .getRange(2, 1, categories.length, 1)
    .getBackgrounds();
  return categories.map((category, index) => {
    logger.info(`Processing category index ${index}: ${category[0]}`);
    const presetSlug = createPresetSlug(category[0], index);

    const referencedFields = category[2]
      ? category[2]
          .split(",")
          .map((field) => field.trim())
          .filter((field) => field !== "")
          .map((fieldName, fieldIndex) => {
            const canonicalKey = fieldKeyByLabel[fieldName];
            if (!canonicalKey) {
              logger.warn(
                `⚠️  Field "${fieldName}" referenced by category "${category[0]}" not found in processed fields. Falling back to slug.`,
              );
            }
            return canonicalKey || buildSlugWithFallback(fieldName, "field", fieldIndex);
          })
      : [];

    const terms = [
      category[0],
      ...referencedFields.map((fieldKey) => fieldKey.replace(/-/g, " ")),
    ];
    return {
      icon: presetSlug,
      color: backgroundColors[index][0] || "#0000FF",
      fields: referencedFields,
      geometry: ["point", "line", "area"],
      tags: { [presetSlug]: "yes" },
      name: category[0],
      sort: index + 1,
      terms,
    };
  });
}
