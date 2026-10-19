/**
 * Categories application functions for the import category functionality.
 * This file contains functions related to applying categories to the spreadsheet.
 */

function safeCategoryDebugLog(message: string, forceFlush = false): void {
  try {
    if (typeof createSafeDebugLogger === "function") {
      createSafeDebugLogger("Categories")(message, forceFlush);
      return;
    }
    if (typeof debugLog === "function") {
      debugLog(`[Categories] ${message}`, forceFlush);
      return;
    }
  } catch (_e) {
    // Fall through to console.
  }
  console.log(`[Categories] ${message}`);
}

/**
 * Applies categories (presets) to the Categories sheet.
 * @param sheet - The categories sheet
 * @param presets - Array of preset objects
 * @param fields - Array of field objects
 * @param icons - Array of icon objects
 */
function applyCategories(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  presets: any[],
  fields: any[],
  icons: any[],
) {
  safeCategoryDebugLog("=== APPLYING CATEGORIES TO SPREADSHEET ===");
  safeCategoryDebugLog(`Applying ${presets.length} categories with ${icons.length} icons`);

  // Set headers ("Name" is the canonical column header; primary language is
  // stored in the Metadata sheet under the "primaryLanguage" key.)
  // Canonical layout (issue #61 / #52): Name | Icon | Fields | Applies |
  // Category ID | Color | Icon ID, shared with importService and apiService.
  const headers = CATEGORY_HEADERS.slice();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  // Create a map of icon name to icon URL for quick lookup
  safeCategoryDebugLog("Building icon map for lookup...");
  safeCategoryDebugLog(`Processing ${icons.length} icon(s) from extraction...`);
  const iconMap: Record<string, string> = {};
  let overwriteCount = 0;
  for (const icon of icons) {
    if (icon.name && icon.svg) {
      if (iconMap[icon.name]) {
        overwriteCount++;
        safeCategoryDebugLog(`  ⚠️  OVERWRITE: iconMap["${icon.name}"] was "${iconMap[icon.name].substring(0, 40)}..."`);
        safeCategoryDebugLog(`              now replaced with "${icon.svg.substring(0, 40)}..."`);
      } else {
        safeCategoryDebugLog(`  ✓ Added: iconMap["${icon.name}"] = ${icon.svg.substring(0, 50)}...`);
      }
      iconMap[icon.name] = icon.svg;
    }
  }
  safeCategoryDebugLog(`\n=== ICON MAP SUMMARY ===`);
  safeCategoryDebugLog(`Processed ${icons.length} icons, created ${Object.keys(iconMap).length} map entries`);
  if (overwriteCount > 0) {
    safeCategoryDebugLog(`⚠️  WARNING: ${overwriteCount} duplicate icon name(s) detected - entries were overwritten!`);
    safeCategoryDebugLog(`This means ${icons.length - Object.keys(iconMap).length} icon(s) are hidden and won't appear in spreadsheet.`);
  } else {
    safeCategoryDebugLog(`✓ All icons have unique names - no overwrites`);
  }
  safeCategoryDebugLog(`=== END ICON MAP SUMMARY ===\n`);

  // Prepare category rows
  safeCategoryDebugLog("Matching icons to presets...");
  const categoryRows = presets.map((preset) => {
    // Find matching icon
    const iconUrl = preset.icon ? iconMap[preset.icon] || "" : "";

    if (preset.icon) {
      if (iconUrl) {
        safeCategoryDebugLog(`  ✓ Matched preset "${preset.name}" (icon: "${preset.icon}") → ${iconUrl.substring(0, 50)}...`);
      } else {
        safeCategoryDebugLog(`  ✗ No icon found for preset "${preset.name}" (looking for icon: "${preset.icon}")`);
      }
    } else {
      safeCategoryDebugLog(`  - Preset "${preset.name}" has no icon specified`);
    }

    // Get fields as comma-separated string with actual Label values
    let fieldsList = "";
    if (
      preset.fields &&
      Array.isArray(preset.fields) &&
      preset.fields.length > 0
    ) {
      // Create a map of field IDs to their labels for quick lookup
      const fieldMap: Record<string, string> = {};
      for (const field of fields) {
        if (field.id && field.label) {
          fieldMap[field.id] = field.label;
        }
      }

      // Map each field ID to its Label
      const fieldLabels = preset.fields.map((fieldId: string) => {
        // Use the Label from the fieldMap if available, otherwise use the original ID
        const label = fieldMap[fieldId] || fieldId;
        // Strip commas from the label to avoid issues in the Details column
        return label.replace(/,/g, "");
      });

      fieldsList = fieldLabels.join(", ");
    }

    // Canonical layout: write appliesTo, category id, color, and icon id
    // explicitly so a Generate run reads back what was imported (issue #61).
    // Default only when the legacy preset genuinely lacks appliesTo.
    const appliesValue =
      Array.isArray(preset.appliesTo) && preset.appliesTo.length > 0
        ? preset.appliesTo.join(", ")
        : "track, observation";
    const colorValue =
      typeof preset.color === "string" &&
      /^#([0-9A-F]{3}){1,2}$/i.test(preset.color)
        ? preset.color
        : "";

    // Only carry the Icon ID reference when the icon itself was actually
    // extracted from the archive: writing an Icon ID with an empty Icon cell
    // makes generate fail hard ("Could not load icons for categories ..."),
    // while leaving both blank degrades gracefully to an icon-less category.
    const iconIdValue = iconUrl ? preset.iconId || "" : "";

    return [
      preset.name,
      iconUrl,
      fieldsList,
      appliesValue,
      preset.id || "",
      colorValue,
      iconIdValue,
    ];
  });

  // Add category rows
  if (categoryRows.length > 0) {
    sheet
      .getRange(2, 1, categoryRows.length, CATEGORY_HEADERS.length)
      .setValues(categoryRows);

    // Apply background colors based on preset.color
    presets.forEach((preset, index) => {
      if (preset.color && typeof preset.color === "string") {
        // Check if it's a valid hex color (starts with # and has 3 or 6 hex digits)
        const isValidHexColor = /^#([0-9A-F]{3}){1,2}$/i.test(preset.color);
        if (isValidHexColor) {
          // Apply the color to the first cell (category name)
          sheet.getRange(index + 2, 1).setBackground(preset.color);
        }
      }
    });
  }

  // Data validation for the Details column is handled separately
  // to avoid import errors

  // Auto-resize columns for better readability
  sheet.autoResizeColumns(1, CATEGORY_HEADERS.length);
}
