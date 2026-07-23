/// <reference path="shared.ts" />

/**
 * Task 2: Checks for duplicate effective IDs in Details column E.
 * Mirrors checkDuplicateCategoryIds() for field IDs.
 */
function checkDuplicateDetailIds(): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("Details");
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  // Clear previous lint artifacts on column E (ID column, index 5)
  const idRange = sheet.getRange(2, 5, lastRow - 1, 1);
  clearLintArtifacts(idRange);

  // Read name column (A=1) and ID column (E=5)
  const names = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const ids = idRange.getValues();

  // Strict validation treats field IDs case-insensitively, so lint should too.
  // Mirror builder: explicit ID → slugify(name) → `field-${index+1}` fallback.
  // Skip blank-name rows — builder returns early on those.
  const effectiveIdMap = new Map<string, { rows: number[]; displayId: string }>();
  for (let i = 0; i < ids.length; i++) {
    const explicitId = String(ids[i][0] || "").trim();
    const name = String(names[i][0] || "").trim();
    if (!name) continue;
    const effectiveId = explicitId || slugify(name) || `field-${i + 1}`;
    const normalizedId = effectiveId.toLowerCase();

    if (!effectiveIdMap.has(normalizedId)) {
      effectiveIdMap.set(normalizedId, {
        rows: [i + 2],
        displayId: effectiveId,
      });
    } else {
      effectiveIdMap.get(normalizedId)?.rows.push(i + 2);
    }
  }

  // Annotate duplicates
  effectiveIdMap.forEach(({ rows, displayId }) => {
    if (rows.length > 1) {
      const logger = getScopedLogger("LintDuplicateDetailIds");
      logger.warn(
        `Duplicate field ID "${displayId}" in rows: ${rows.join(", ")}`,
      );
      for (const row of rows) {
        const cell = sheet.getRange(row, 5);
        setLintNote(
          cell,
          `Duplicate field ID "${displayId}" (case-insensitive match; also in rows ${rows.filter((r) => r !== row).join(", ")})`,
          "error",
        );
      }
    }
  });
}

/**
 * Check for unreferenced details (details that no category uses).
 * Uses normalizeFieldTokens() to match build parsing and resolves against
 * both slugified names and explicit IDs from Details column E.
 */
function checkUnreferencedDetails(): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const categoriesSheet = spreadsheet.getSheetByName("Categories");
  const detailsSheet = spreadsheet.getSheetByName("Details");

  if (!categoriesSheet || !detailsSheet) {
    return;
  }

  try {
    // Get all detail names and explicit IDs from Details sheet
    const detailsLastRow = detailsSheet.getLastRow();
    if (detailsLastRow <= 1) return; // No details to check

    const detailRange = detailsSheet.getRange(2, 1, detailsLastRow - 1, 1);
    clearRangeLintNotesWithPrefix(
      detailRange,
      `${LINT_NOTE_PREFIX}Detail `,
    );
    // Clear warning backgrounds on cells whose notes were fully removed
    // (i.e., cells that only had unreferenced-detail warnings, now resolved).
    // Cells still carrying other annotations (e.g., duplicate-name errors)
    // keep their higher-severity backgrounds.
    const notesAfter = detailRange.getNotes();
    const backgrounds = detailRange.getBackgrounds();
    let bgUpdated = false;
    for (let r = 0; r < backgrounds.length; r++) {
      if (
        !notesAfter[r][0] &&
        LINT_WARNING_BACKGROUND_COLORS.some(
          (c) => c.toUpperCase() === backgrounds[r][0].toUpperCase(),
        )
      ) {
        backgrounds[r][0] = null;
        bgUpdated = true;
      }
    }
    if (bgUpdated) {
      detailRange.setBackgrounds(backgrounds);
    }

    // Read both name (col A) and ID (col E) columns
    const detailData = detailsSheet
      .getRange(2, 1, detailsLastRow - 1, 5)
      .getValues();

    // Build list of detail entries: { slug, explicitId, row }
    const detailEntries: Array<{
      name: string;
      slug: string;
      explicitId: string;
      row: number;
    }> = [];
    for (let i = 0; i < detailData.length; i++) {
      const name = String(detailData[i][0] || "").trim();
      const explicitId = String(detailData[i][4] || "").trim();
      const slug = slugify(name);
      if (slug || explicitId) {
        detailEntries.push({ name, slug, explicitId, row: i + 2 });
      }
    }

    // Get all field references from Categories sheet — resolve "Fields" column
    // dynamically by header name to match builder behavior.
    const categoriesLastRow = categoriesSheet.getLastRow();
    if (categoriesLastRow <= 1) {
      // No categories exist, so all details are unreferenced
      getScopedLogger("LintUnreferencedDetails").info("No categories exist - all details are unreferenced");
      for (const entry of detailEntries) {
        appendLintNote(
          detailsSheet.getRange(entry.row, 1),
          `Detail "${entry.name}" is not referenced by any category's field list. No categories exist.`,
          "warning",
        );
      }
      return;
    }

    const categoriesLastCol = categoriesSheet.getLastColumn();
    const catHeaders = categoriesSheet
      .getRange(1, 1, 1, categoriesLastCol)
      .getValues()[0];
    const catHeaderMap = createLintHeaderMap(catHeaders);
    const fieldsColZeroBased =
      getLintColumnIndex(catHeaderMap, "fields", "details") ??
      CATEGORY_COL.FIELDS;
    const fieldsColOneBased = fieldsColZeroBased + 1;

    const fieldRange = categoriesSheet.getRange(
      2,
      fieldsColOneBased,
      categoriesLastRow - 1,
      1,
    );
    const displayFieldValues = fieldRange.getDisplayValues();
    const rawFieldValues = fieldRange.getValues();
    const categoryFields: string[] = [];
    for (let i = 0; i < displayFieldValues.length; i++) {
      const displayStr = String(displayFieldValues[i][0] || "");
      const rawStr = String(rawFieldValues[i][0] || "");
      // Mirror builder: prefer display value, fall back to raw
      const tokens = normalizeFieldTokens(displayStr);
      const value = tokens.length > 0 ? displayStr : rawStr;
      const strValue = String(value || "");
      if (strValue.trim() !== "") {
        categoryFields.push(strValue);
      }
    }

    // Build set of all referenced field identifiers using normalizeFieldTokens
    const referencedFields = new Set<string>();
    for (const fieldsStr of categoryFields) {
      const tokens = normalizeFieldTokens(fieldsStr);
      for (const token of tokens) {
        const slugified = slugify(token);
        if (slugified) referencedFields.add(slugified);
        if (token) {
          referencedFields.add(token);
          referencedFields.add(token.toLowerCase());
        }
      }
    }

    // Check each detail against the category field tokens. Mirror the builder's
    // resolution (payloadBuilder fieldNameToId): a category references a detail
    // by NAME (exact or lowercase), falling back to the slugified name. So match
    // on the detail's raw/lowercase name and slug, plus its explicit ID. The
    // name check is essential for non-Latin names — slugify() returns "" for
    // Thai etc., so without it every Thai detail referenced by name would be
    // falsely reported as unreferenced.
    for (const entry of detailEntries) {
      const isReferenced =
        (entry.slug && referencedFields.has(entry.slug)) ||
        referencedFields.has(entry.name) ||
        referencedFields.has(entry.name.toLowerCase()) ||
        (entry.explicitId &&
          (referencedFields.has(entry.explicitId) ||
            referencedFields.has(entry.explicitId.toLowerCase())));
      if (!isReferenced) {
        getScopedLogger("LintUnreferencedDetails").info(
          `Unreferenced detail: "${entry.slug}" at row ${entry.row}`,
        );
        const cell = detailsSheet.getRange(entry.row, 1);
        appendLintNote(
          cell,
          `Detail "${entry.name}" is not referenced by any category's field list. It will be excluded from the generated config.`,
          "warning",
        );
      }
    }
  } catch (error) {
    getScopedLogger("LintUnreferencedDetails").error("Error checking unreferenced details:", error);
  }
}

/**
 * Validate Universal flag column (should be TRUE, FALSE, or blank)
 */
function validateUniversalFlag(
  value: string,
  row: number,
  col: number,
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
): void {
  if (isEmptyOrWhitespace(value)) return; // Blank is allowed

  const upperValue = value.toString().trim().toUpperCase();
  if (upperValue !== "TRUE" && upperValue !== "FALSE") {
    getScopedLogger("LintDetails").info(
      `Invalid Universal flag value "${value}" at row ${row} - must be TRUE, FALSE, or blank`,
    );
    setInvalidCellBackground(sheet, row, col, LINT_ERROR_BG); // Light red for invalid
  }
}

function lintDetailsSheet(): void {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Details");
  if (!sheet) {
    getScopedLogger("LintDetails").info("Details sheet not found");
    return;
  }

  // Task 5: Check if Details sheet is empty before running lint
  if (checkEmptySheet(sheet, "Details")) {
    return;
  }

  // Clear stale empty-sheet note from A1 (left by a previous run when the
  // sheet was empty). The non-empty path only touches data rows, so A1
  // would otherwise retain the old error.
  clearLintArtifacts(sheet.getRange(1, 1));

  // Pre-cleanup: capture whitespace-only IDs before cleanWhitespaceOnlyCells() erases them
  const rawDetailIds = inspectRawIds("Details", 5);

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const lastColumn = sheet.getLastColumn();
    if (lastColumn >= 3) {
      clearRangeBackgroundIfMatches(sheet.getRange(2, 3, lastRow - 1, 1), [
        LINT_ERROR_BG,
      ]);
    }
    if (lastColumn >= 4) {
      clearRangeBackgroundIfMatches(sheet.getRange(2, 4, lastRow - 1, 1), [
        LINT_ERROR_BG,
      ]);
    }
    if (lastColumn >= 6) {
      clearRangeBackgroundIfMatches(sheet.getRange(2, 6, lastRow - 1, 1), [
        LINT_ERROR_BG,
      ]);
    }
  }

  const detailsValidations = [
    // Rule 1: Capitalize the first letter of the detail name
    (value: string, row: number, col: number) => {
      if (isEmptyOrWhitespace(value)) return;

      const capitalizedValue = capitalizeFirstLetter(value);
      if (capitalizedValue !== value) {
        try {
          SpreadsheetApp.getActiveSpreadsheet()
            .getSheetByName("Details")
            ?.getRange(row, col)
            .setValue(capitalizedValue);
        } catch (error) {
          getScopedLogger("LintDetails").error(
            "Error capitalizing detail name at row " +
              row +
              ", col " +
              col +
              ":",
            error,
          );
        }
      }
    },
    // Rule 2: Capitalize the first letter of the helper text
    (value: string, row: number, col: number) => {
      if (isEmptyOrWhitespace(value)) return;

      const capitalizedValue = capitalizeFirstLetter(value);
      if (capitalizedValue !== value) {
        try {
          SpreadsheetApp.getActiveSpreadsheet()
            .getSheetByName("Details")
            ?.getRange(row, col)
            .setValue(capitalizedValue);
        } catch (error) {
          getScopedLogger("LintDetails").error(
            "Error capitalizing helper text at row " +
              row +
              ", col " +
              col +
              ":",
            error,
          );
        }
      }
    },
    // Rule 3: Validate the type column (t, n, m, blank, s, or select* are valid)
    (value: string, row: number, col: number) => {
      // Type column validation logic:
      // - blank/empty → selectOne (mirrors payloadBuilder case "" default)
      // - "s*" (select, single, etc.) → selectOne (valid)
      // - "m*" (multi, multiple, etc.) → selectMultiple (valid)
      // - "n*" (number, numeric, etc.) → number (valid)
      // - "t*" (text, textual, etc.) → text (valid)
      // - Any other value → invalid

      const detailsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Details");

      // Empty/blank defaults to selectOne (mirrors getFieldType() fallback) - add advisory note
      if (isEmptyOrWhitespace(value)) {
        if (detailsSheet) {
          setLintNote(
            detailsSheet.getRange(row, col),
            "Blank type defaults to selectOne. Enter 'text', 'number', 'selectOne', or 'selectMultiple' to be explicit.",
            "advisory",
          );
        }
        return;
      }

      const firstChar = value.toLowerCase().charAt(0);
      const validTypes = ["t", "n", "m", "s"];

      if (!validTypes.includes(firstChar)) {
        try {
          getScopedLogger("LintDetails").info("Invalid type '" + value + "' at row " + row);
          if (detailsSheet) {
            setLintNote(
              detailsSheet.getRange(row, col),
              `Invalid type '${value}'. Expected: text (t), number (n), selectOne (s), or selectMultiple (m).`,
              "error",
            );
          }
        } catch (error) {
          getScopedLogger("LintDetails").error(
            "Error highlighting invalid type at row " +
              row +
              ", col " +
              col +
              ":",
            error,
          );
        }
      } else if (value.length > 1) {
        // Valid first char but longer than a single abbreviation - advisory for clarity
        if (detailsSheet) {
          setLintNote(
            detailsSheet.getRange(row, col),
            `Type '${value}' is valid (starts with '${firstChar}') but consider using the full type name for clarity.`,
            "advisory",
          );
        }
      }
    },
    // Rule 4: Validate options column (canonical parsing + ambiguous colon + ignored options)
    (value: string, row: number, col: number) => {
      try {
        // Get the type from column 3 (index 2) to determine if options are required
        const typeValue = sheet.getRange(row, 3).getValue();
        const typeStr = String(typeValue || "").trim();

        // Determine the field type category
        // Mirror payloadBuilder: empty type defaults to selectOne (case "" in switch)
        const firstChar = isEmptyOrWhitespace(typeStr)
          ? "s"
          : typeStr.toLowerCase().charAt(0);
        const isSelectField = firstChar === "s" || firstChar === "m";
        const isTextOrNumberField = firstChar === "t" || firstChar === "n";

        // Task 7: Warn when text/number fields contain ignored options
        if (isTextOrNumberField && !isEmptyOrWhitespace(value)) {
          const resolvedType = firstChar === "t" ? "text" : "number";
          const cell = sheet.getRange(row, col);
          setLintNote(
            cell,
            `Options on ${resolvedType} fields are ignored during config generation. Consider removing them.`,
            "warning",
          );
          // Still capitalize for presentation, but the warning is the main feedback
          const capitalizedList = validateAndCapitalizeCommaList(value);
          if (capitalizedList !== value) {
            sheet.getRange(row, col).setValue(capitalizedList);
          }
          return;
        }

        if (isSelectField) {
          // Select fields MUST have options
          if (isEmptyOrWhitespace(value)) {
            getScopedLogger("LintDetails").info(
              "Select field at row " + row + " is missing required options",
            );
            setLintNote(
              sheet.getRange(row, col),
              "Select fields (selectOne / selectMultiple) require at least one option.",
              "error",
            );
            return;
          }

          // Parse options using canonical parser (mirrors builder's parseOptions)
          const parsed = parseCanonicalOptions(value);

          if (parsed.length === 0) {
            getScopedLogger("LintDetails").info(
              "Select field at row " +
                row +
                " has empty options after trimming",
            );
            setLintNote(
              sheet.getRange(row, col),
              "Options column appears non-empty but contains no valid options after parsing.",
              "error",
            );
            return;
          }

          // Task 5: Check for ambiguous colon usage
          const ambiguityWarnings = detectAmbiguousColonUsage(parsed);
          if (ambiguityWarnings.length > 0) {
            const cell = sheet.getRange(row, col);
            appendLintNote(
              cell,
              ambiguityWarnings.join(" "),
              "warning",
            );
          }

          // Task 4: Check for duplicate canonical values
          const seenValues = new Map<string, string>(); // canonical value -> raw representation
          const uniqueEntries: string[] = [];
          const removedDuplicates: string[] = [];

          for (const opt of parsed) {
            // Preserve the original format: if the raw entry used "value:label",
            // keep that format; otherwise just use the label.
            const displayForm = opt.raw.includes(":") ? opt.raw : (opt.label || opt.value);
            if (seenValues.has(opt.value)) {
              removedDuplicates.push(opt.label || opt.value);
            } else {
              seenValues.set(opt.value, displayForm);
              uniqueEntries.push(displayForm);
            }
          }

          if (removedDuplicates.length > 0) {
            // Update cell with deduplicated options, preserving value:label format
            const deduplicatedValue = uniqueEntries.join(", ");
            sheet.getRange(row, col).setValue(deduplicatedValue);

            // Add warning note about removed duplicates (append to preserve
            // any colon-ambiguity warning already on the cell)
            const cell = sheet.getRange(row, col);
            appendLintNote(
              cell,
              `Removed ${removedDuplicates.length} duplicate option(s): "${removedDuplicates.join('", "')}". Each option must produce a unique canonical value.`,
              "warning",
            );

            getScopedLogger("LintDetails").info(
              `Row ${row}: Removed ${removedDuplicates.length} duplicate option(s): ${removedDuplicates.join(", ")}`,
            );
            return;
          }

          // Capitalize and format the options
          const capitalizedList = validateAndCapitalizeCommaList(value);
          if (capitalizedList !== value) {
            sheet.getRange(row, col).setValue(capitalizedList);
          }
        } else {
          // For other non-select, non-text/number fields, just capitalize if options are provided
          if (!isEmptyOrWhitespace(value)) {
            const capitalizedList = validateAndCapitalizeCommaList(value);
            if (capitalizedList !== value) {
              sheet.getRange(row, col).setValue(capitalizedList);
            }
          }
        }
      } catch (error) {
        getScopedLogger("LintDetails").error(
          "Error validating options at row " + row + ", col " + col + ":",
          error,
        );
      }
    },
    // Rule 5: Placeholder for column 5 (no validation needed)
    () => {
      // Column 5 - no validation
    },
    // Rule 6: Validate Universal flag column (TRUE, FALSE, or blank only)
    (value: string, row: number, col: number) => {
      validateUniversalFlag(value, row, col, sheet);
    },
  ];

  // Detail name and type are required fields
  lintSheet("Details", detailsValidations, [0, 2]);

  // Phase 2: Post-lintSheet() edge-case checks
  checkUnreferencedDetails(); // Task: unreferenced details (must run after lintSheet clears col A)
  checkDuplicateDetailIds(); // Task 2: duplicate effective IDs in col E
  checkSlugCollisions("Details", 1); // Task 3: slug collisions in col A
  checkManualIdHygiene("Details", 5, rawDetailIds); // Task 6: ID hygiene in col E
}
