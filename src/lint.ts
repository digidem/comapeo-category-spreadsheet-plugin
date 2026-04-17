// Import slugify function from utils
/// <reference path="./utils.ts" />

// Helper functions
function capitalizeFirstLetter(str: string): string {
  if (!str || typeof str !== "string") return "";
  return str.charAt(0).toUpperCase() + str.slice(1);
}

function validateAndCapitalizeCommaList(value: string): string {
  if (!value || typeof value !== "string") return "";
  return value
    .split(",")
    .map((item) => capitalizeFirstLetter(item.trim()))
    .filter((item) => item)
    .join(", ");
}

function setInvalidCellBackground(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  row: number,
  col: number,
  color: string,
): void {
  sheet.getRange(row, col).setBackground(color);
}

const LINT_WARNING_BACKGROUND_COLORS = [
  "#FFC7CE",
  "#FFEB9C",
  "#FFFFCC",
  "#FFF2CC",
  "#FFF3CD",
];
const LINT_WARNING_FONT_COLORS = ["red", "orange", "#FF0000"];
const LINT_NOTE_PREFIX = "[Lint] ";
const SLUG_COLLISION_LINT_NOTE_PREFIX = `${LINT_NOTE_PREFIX}Slug collision:`;

function clearRangeBackgroundIfMatches(
  range: GoogleAppsScript.Spreadsheet.Range,
  colorsToClear: string[],
): void {
  if (!range) return;
  if (range.getNumRows() === 0 || range.getNumColumns() === 0) return;

  const normalized = colorsToClear.map((color) => color.toUpperCase());
  const backgrounds = range.getBackgrounds();
  let updated = false;

  for (let row = 0; row < backgrounds.length; row++) {
    for (let col = 0; col < backgrounds[row].length; col++) {
      const background = backgrounds[row][col];
      if (background && normalized.includes(background.toUpperCase())) {
        backgrounds[row][col] = null;
        updated = true;
      }
    }
  }

  if (updated) {
    range.setBackgrounds(backgrounds);
  }
}

function clearRangeFontColorIfMatches(
  range: GoogleAppsScript.Spreadsheet.Range,
  colorsToClear: string[],
): void {
  if (!range) return;
  if (range.getNumRows() === 0 || range.getNumColumns() === 0) return;

  const normalized = colorsToClear.map((color) => color.toUpperCase());
  const fontColors = range.getFontColors();
  let updated = false;

  for (let row = 0; row < fontColors.length; row++) {
    for (let col = 0; col < fontColors[row].length; col++) {
      const fontColor = fontColors[row][col];
      if (fontColor && normalized.includes(fontColor.toUpperCase())) {
        fontColors[row][col] = null;
        updated = true;
      }
    }
  }

  if (updated) {
    range.setFontColors(fontColors);
  }
}

function clearRangeNotesWithPrefix(
  range: GoogleAppsScript.Spreadsheet.Range,
  prefix: string,
): void {
  if (!range) return;
  if (range.getNumRows() === 0 || range.getNumColumns() === 0) return;

  const notes = range.getNotes();
  let updated = false;

  for (let row = 0; row < notes.length; row++) {
    for (let col = 0; col < notes[row].length; col++) {
      const note = notes[row][col];
      if (
        note &&
        (note.startsWith(prefix) ||
          note.includes('Icon slug "') ||
          note.includes("No SVG icon found") ||
          note.includes("Unable to determine an icon name"))
      ) {
        notes[row][col] = "";
        updated = true;
      }
    }
  }

  if (updated) {
    range.setNotes(notes);
  }
}

function clearRangeLintNoteLinesWithPrefix(
  range: GoogleAppsScript.Spreadsheet.Range,
  prefix: string,
): void {
  if (!range) return;
  if (range.getNumRows() === 0 || range.getNumColumns() === 0) return;

  const normalizedWarningColors = LINT_WARNING_FONT_COLORS.map((color) =>
    color.toUpperCase(),
  );
  const notes = range.getNotes();
  const fontColors = range.getFontColors();
  let notesUpdated = false;
  let fontColorsUpdated = false;

  for (let row = 0; row < notes.length; row++) {
    for (let col = 0; col < notes[row].length; col++) {
      const note = notes[row][col];
      if (!note || !note.includes(prefix)) continue;

      const noteLines = note.split("\n");
      const remainingLines = noteLines.filter((line) => !line.startsWith(prefix));

      if (remainingLines.length === noteLines.length) continue;

      let start = 0;
      while (start < remainingLines.length && remainingLines[start] === "") {
        start++;
      }

      let end = remainingLines.length;
      while (end > start && remainingLines[end - 1] === "") {
        end--;
      }

      const remainingNote = remainingLines.slice(start, end).join("\n");

      notes[row][col] = remainingNote;
      notesUpdated = true;

      const fontColor = fontColors[row][col];
      if (
        remainingNote === "" &&
        fontColor &&
        normalizedWarningColors.includes(fontColor.toUpperCase())
      ) {
        fontColors[row][col] = null;
        fontColorsUpdated = true;
      }
    }
  }

  if (notesUpdated) {
    range.setNotes(notes);
  }

  if (fontColorsUpdated) {
    range.setFontColors(fontColors);
  }
}

/**
 * Standardized lint note writer. Sets a [Lint]-prefixed note on a cell and applies
 * severity-appropriate background and font colors so that cleanup and UI behavior
 * stay consistent across all lint checks.
 *
 * severity → background / font color mapping:
 *   error    → #FFC7CE / red
 *   warning  → #FFF2CC / orange
 *   advisory → #FFFFCC / (default)
 */
function setLintNote(
  cell: GoogleAppsScript.Spreadsheet.Range,
  message: string,
  severity: "error" | "warning" | "advisory",
): void {
  cell.setNote(`${LINT_NOTE_PREFIX}${message}`);

  switch (severity) {
    case "error":
      cell.setBackground("#FFC7CE");
      cell.setFontColor("red");
      break;
    case "warning":
      cell.setBackground("#FFF2CC");
      cell.setFontColor("orange");
      break;
    case "advisory":
      cell.setBackground("#FFFFCC");
      break;
  }
}

/**
 * Appends a lint note to a cell, preserving any existing [Lint]-prefixed note.
 * Uses the same severity-based styling as setLintNote but concatenates messages
 * when a note already exists, preventing overwrites from sequential lint passes.
 *
 * severity escalation: if the existing note is a lower severity than the new one,
 * the visual styling is upgraded to match the highest severity present.
 */
function appendLintNote(
  cell: GoogleAppsScript.Spreadsheet.Range,
  message: string,
  severity: "error" | "warning" | "advisory",
): void {
  const existingNote = cell.getNote() || "";
  const newMessage = `${LINT_NOTE_PREFIX}${message}`;

  if (
    existingNote &&
    existingNote.startsWith(LINT_NOTE_PREFIX)
  ) {
    cell.setNote(`${existingNote}\n${newMessage}`);
  } else {
    cell.setNote(newMessage);
  }

  // Apply severity styling (upgrade if higher severity than current)
  const currentBg = cell.getBackground();
  const isAlreadyError =
    currentBg.toUpperCase() === "#FFC7CE";
  const isAlreadyWarning =
    currentBg.toUpperCase() === "#FFF2CC";

  switch (severity) {
    case "error":
      cell.setBackground("#FFC7CE");
      cell.setFontColor("red");
      break;
    case "warning":
      // Only set warning styling if not already at error level
      if (!isAlreadyError) {
        cell.setBackground("#FFF2CC");
        cell.setFontColor("orange");
      }
      break;
    case "advisory":
      // Only set advisory styling if no higher severity is present
      if (!isAlreadyError && !isAlreadyWarning) {
        cell.setBackground("#FFFFCC");
      }
      break;
  }
}

/**
 * Combined helper that clears all lint-managed visual artifacts from a range:
 * background colors, font colors, and [Lint]-prefixed notes.
 * Range-scoped so category color backgrounds in column A are preserved.
 */
function clearLintArtifacts(
  range: GoogleAppsScript.Spreadsheet.Range,
): void {
  if (!range) return;
  if (range.getNumRows() === 0 || range.getNumColumns() === 0) return;

  clearRangeBackgroundIfMatches(range, LINT_WARNING_BACKGROUND_COLORS);
  clearRangeFontColorIfMatches(range, LINT_WARNING_FONT_COLORS);
  clearRangeNotesWithPrefix(range, LINT_NOTE_PREFIX);
}

/**
 * Like setLintNote but does NOT change the cell background.
 * Use this for columns where user-set backgrounds must be preserved
 * (e.g., Categories column A category colors).
 */
function setLintNotePreserveBackground(
  cell: GoogleAppsScript.Spreadsheet.Range,
  message: string,
  severity: "error" | "warning" | "advisory",
): void {
  cell.setNote(`${LINT_NOTE_PREFIX}${message}`);

  switch (severity) {
    case "error":
      cell.setFontColor("red");
      break;
    case "warning":
      cell.setFontColor("orange");
      break;
    case "advisory":
      break;
  }
}

/**
 * Like appendLintNote but does NOT change the cell background.
 * Use this for columns where user-set backgrounds must be preserved
 * (e.g., Categories column A category colors).
 */
function appendLintNotePreserveBackground(
  cell: GoogleAppsScript.Spreadsheet.Range,
  message: string,
  severity: "error" | "warning" | "advisory",
): void {
  const existingNote = cell.getNote() || "";
  const newMessage = `${LINT_NOTE_PREFIX}${message}`;

  if (
    existingNote &&
    existingNote.startsWith(LINT_NOTE_PREFIX)
  ) {
    cell.setNote(`${existingNote}\n${newMessage}`);
  } else {
    cell.setNote(newMessage);
  }

  // Apply font color only (no background)
  const currentBg = cell.getBackground();
  const isAlreadyError =
    currentBg.toUpperCase() === "#FFC7CE";

  switch (severity) {
    case "error":
      cell.setFontColor("red");
      break;
    case "warning":
      if (!isAlreadyError) {
        cell.setFontColor("orange");
      }
      break;
    case "advisory":
      break;
  }
}

/**
 * Pre-cleanup inspection pass for ID columns (column E, index 4).
 * Reads raw values BEFORE cleanWhitespaceOnlyCells() runs so that
 * whitespace-only IDs (which would be erased by cleanup) are captured.
 *
 * Returns a Map of row number → raw ID value for cells that contain
 * only whitespace characters (not empty, but whitespace-only).
 */
function inspectRawIds(
  sheetName: string,
  columnEIndex: number,
): Map<number, string> {
  const result = new Map<number, string>();
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return result;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return result;

  const range = sheet.getRange(2, columnEIndex, lastRow - 1, 1);
  const values = range.getValues();

  for (let i = 0; i < values.length; i++) {
    const value = values[i][0];
    // Capture cells that are strings with content but only whitespace
    if (
      typeof value === "string" &&
      value.length > 0 &&
      value.trim() === ""
    ) {
      result.set(i + 2, value); // +2 because data starts at row 2
    }
  }

  return result;
}

/**
 * Task 1: Checks for duplicate effective IDs in Categories column E.
 * The "effective ID" is the explicit ID if present, otherwise the slugified name.
 * Annotates duplicate cells with error-level lint notes.
 */
function checkDuplicateCategoryIds(): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("Categories");
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  // Clear previous lint artifacts on column E (ID column, index 5)
  const idRange = sheet.getRange(2, 5, lastRow - 1, 1);
  clearLintArtifacts(idRange);

  // Read name column (A=1) and ID column (E=5)
  const names = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const ids = idRange.getValues();

  // Build map of effective ID → rows
  const effectiveIdMap = new Map<string, number[]>();
  for (let i = 0; i < ids.length; i++) {
    const explicitId = String(ids[i][0] || "").trim();
    const name = String(names[i][0] || "").trim();
    const effectiveId = explicitId || slugify(name) || "";
    if (!effectiveId) continue;

    if (!effectiveIdMap.has(effectiveId)) {
      effectiveIdMap.set(effectiveId, [i + 2]);
    } else {
      effectiveIdMap.get(effectiveId)?.push(i + 2);
    }
  }

  // Annotate duplicates
  effectiveIdMap.forEach((rows, effectiveId) => {
    if (rows.length > 1) {
      const logger = getScopedLogger("LintDuplicateCategoryIds");
      logger.warn(
        `Duplicate category ID "${effectiveId}" in rows: ${rows.join(", ")}`,
      );
      for (const row of rows) {
        const cell = sheet.getRange(row, 5);
        setLintNote(
          cell,
          `Duplicate category ID "${effectiveId}" (also in rows ${rows.filter((r) => r !== row).join(", ")})`,
          "error",
        );
      }
    }
  });
}

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
  const effectiveIdMap = new Map<string, { rows: number[]; displayId: string }>();
  for (let i = 0; i < ids.length; i++) {
    const explicitId = String(ids[i][0] || "").trim();
    const name = String(names[i][0] || "").trim();
    const effectiveId = explicitId || slugify(name) || "";
    if (!effectiveId) continue;
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
 * Task 3: Warns when two names in column A produce the same fallback slug,
 * even if their explicit IDs may later differ. This is the earliest visible
 * signal that auto-generated IDs may collide.
 * SEPARATE from checkForDuplicates() which checks exact string matches.
 */
function checkSlugCollisions(
  sheetName: string,
  nameColumnIndex: number,
): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const nameRange = sheet.getRange(2, nameColumnIndex, lastRow - 1, 1);
  // Remove only prior slug-collision warnings so duplicate-name errors from
  // checkForDuplicates() remain intact while stale slug warnings are refreshed.
  clearRangeLintNoteLinesWithPrefix(nameRange, SLUG_COLLISION_LINT_NOTE_PREFIX);

  const names = nameRange.getValues();

  // Build map of slug → rows (original names for display)
  const slugMap = new Map<string, { rows: number[]; originals: string[] }>();
  for (let i = 0; i < names.length; i++) {
    const name = String(names[i][0] || "").trim();
    if (!name) continue;
    const slug = slugify(name);
    if (!slug) continue;

    if (!slugMap.has(slug)) {
      slugMap.set(slug, { rows: [i + 2], originals: [name] });
    } else {
      const entry = slugMap.get(slug)!;
      entry.rows.push(i + 2);
      entry.originals.push(name);
    }
  }

  // Annotate slug collisions with warnings
  slugMap.forEach(({ rows, originals }, slug) => {
    if (rows.length > 1) {
      const logger = getScopedLogger("LintSlugCollisions");
      logger.warn(
        `Slug collision "${slug}" in ${sheetName}: rows ${rows.join(", ")} (names: ${originals.join(", ")})`,
      );
      for (let idx = 0; idx < rows.length; idx++) {
        const row = rows[idx];
        const otherRows = rows.filter((r) => r !== row);
        const otherNames = originals.filter((_, i) => rows[i] !== row);
        const cell = sheet.getRange(row, nameColumnIndex);
        appendLintNotePreserveBackground(
          cell,
          `Slug collision: "${originals[idx]}" → "${slug}" (also produced by "${otherNames.join('", "')}" in row${otherRows.length > 1 ? "s" : ""} ${otherRows.join(", ")})`,
          "warning",
        );
      }
    }
  });
}

/**
 * Task 4: Validates the Applies column in the Categories sheet.
 * - Resolves the Applies column by header label, matching the builder
 * - Warns when tokens would be ignored by the builder
 * - Warns if no category includes "track" (non-blocking per strict build validation)
 * - Warns if the column is missing and the builder would need to auto-create it
 */
function validateAppliesColumn(): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const categoriesSheet = spreadsheet.getSheetByName("Categories");
  if (!categoriesSheet) return;

  const lastRow = categoriesSheet.getLastRow();
  const lastCol = categoriesSheet.getLastColumn();
  if (lastCol > 0) {
    const headerRowRange = categoriesSheet.getRange(1, 1, 1, lastCol);
    clearRangeLintNoteLinesWithPrefix(
      headerRowRange,
      `${LINT_NOTE_PREFIX}No "Applies" header found.`,
    );
    clearRangeLintNoteLinesWithPrefix(
      headerRowRange,
      `${LINT_NOTE_PREFIX}No category includes "observation" in Applies.`,
    );
    clearRangeLintNoteLinesWithPrefix(
      headerRowRange,
      `${LINT_NOTE_PREFIX}No category includes "track" in Applies.`,
    );

    if (lastRow > 1) {
      const bodyRange = categoriesSheet.getRange(2, 1, lastRow - 1, lastCol);
      clearRangeLintNoteLinesWithPrefix(
        bodyRange,
        `${LINT_NOTE_PREFIX}Unrecognized Applies token(s):`,
      );
    }
  }
  const headerValues =
    lastCol > 0 ? categoriesSheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  const normalizedHeaders = headerValues.map((header) =>
    String(header || "")
      .trim()
      .toLowerCase(),
  );
  const appliesColZeroBased = normalizedHeaders.findIndex(
    (header) =>
      header === "applies" ||
      header === "tracks" ||
      header === "applies to" ||
      header === "appliesto",
  );

  if (appliesColZeroBased === -1) {
    // The Applies header was removed or renamed. We've already cleared
    // Applies-specific notes above. We intentionally do NOT clear backgrounds
    // across the entire sheet here because other checks (validateCategoryIcons,
    // field validation, etc.) may have set severity backgrounds that should be
    // preserved. Stale Applies backgrounds will be naturally cleaned up on the
    // next full lint cycle via the targeted clearLintArtifacts calls.
    const headerCell = categoriesSheet.getRange(1, 1);
    appendLintNote(
      headerCell,
      'No "Applies" header found. The builder resolves this column by header name and may auto-create it, seeding the first category with "track, observation".',
      "warning",
    );
    return;
  }

  const appliesColIndex = appliesColZeroBased + 1; // 1-based for Sheet ranges

  // Clear previous lint artifacts on column D (data rows + header)
  if (lastRow > 1) {
    const appliesRange = categoriesSheet.getRange(
      2,
      appliesColIndex,
      lastRow - 1,
      1,
    );
    clearLintArtifacts(appliesRange);
  }
  // Also clear header cell artifacts
  clearLintArtifacts(categoriesSheet.getRange(1, appliesColIndex));

  if (lastRow <= 1) return;

  // Read both name (col A) and Applies so we can skip blank/spacer rows,
  // matching buildCategories() which returns early for rows without a name.
  const appliesValues = categoriesSheet.getRange(
    2,
    appliesColIndex,
    lastRow - 1,
    1,
  ).getValues();
  const nameValues = categoriesSheet.getRange(
    2,
    1,
    lastRow - 1,
    1,
  ).getValues();
  let hasObservation = false;
  let hasTrack = false;

  for (let i = 0; i < appliesValues.length; i++) {
    const rawValue = String(appliesValues[i][0] || "").trim();
    const categoryName = String(nameValues[i][0] || "").trim();
    const row = i + 2;

    // Mirror buildCategories(): skip rows without a category name entirely.
    // Blank Applies on a real category row still falls back to observation.
    if (!categoryName) continue;

    if (!rawValue) {
      // Blank Applies cells fall back to observation in the builder.
      hasObservation = true;
      continue;
    }

    // Parse tokens: split by comma, trim, lowercase
    const tokens = rawValue
      .split(",")
      .map((t) => t.trim().toLowerCase())
      .filter(Boolean);

    const normalizedTokens = tokens
      .map((token) => {
        if (token.startsWith("o")) return "observation";
        if (token.startsWith("t")) return "track";
        return "";
      })
      .filter(Boolean);
    const invalidTokens = tokens.filter((token) => {
      return !token.startsWith("o") && !token.startsWith("t");
    });

    if (invalidTokens.length > 0) {
      const cell = categoriesSheet.getRange(row, appliesColIndex);
      if (normalizedTokens.length === 0 && tokens.length > 0) {
        // All tokens are unrecognized — builder silently defaults to observation
        appendLintNote(
          cell,
          `All Applies tokens are unrecognized ("${invalidTokens.join('", "')}"). The builder will silently default this row to "observation". Use "observation" or "track" explicitly.`,
          "warning",
        );
      } else {
        appendLintNote(
          cell,
          `Unrecognized Applies token(s): "${invalidTokens.join('", "')}". The builder only keeps observation/track prefixes and ignores the rest.`,
          "warning",
        );
      }
    }

    // If nothing matched, the builder falls back to observation for this row.
    if (normalizedTokens.length === 0 || normalizedTokens.includes("observation")) {
      hasObservation = true;
    }
    if (normalizedTokens.includes("track")) hasTrack = true;
  }

  // The payload builder still requires observation coverage somewhere in the sheet.
  if (!hasObservation) {
    const cell = categoriesSheet.getRange(1, appliesColIndex);
    appendLintNote(
      cell,
      'No category includes "observation" in Applies. Config generation currently fails unless at least one category resolves to observation.',
      "error",
    );
  }

  // When the Applies column exists, missing track is a hard error — the builder
  // throws 'At least one category must include "track"'. Only a warning when the
  // column is absent entirely (builder auto-creates + seeds it).
  if (!hasTrack) {
    const cell = categoriesSheet.getRange(1, appliesColIndex);
    appendLintNote(
      cell,
      'No category includes "track" in Applies. Config generation will fail — at least one category must include "track".',
      "error",
    );
  }
}

/**
 * Task 5: Checks if a sheet is empty (only header row or no rows).
 * Returns true if empty (caller should return early), false otherwise.
 */
function checkEmptySheet(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  sheetDisplayName: string,
): boolean {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    const cell = sheet.getRange(1, 1);
    clearLintArtifacts(cell);
    setLintNote(
      cell,
      `${sheetDisplayName} sheet is empty. At least one ${sheetDisplayName.toLowerCase()} entry is required for config generation.`,
      "error",
    );
    return true;
  }
  return false;
}

/**
 * Task 6: Checks manual ID hygiene in column E.
 * - Warns on whitespace-only IDs (from pre-cleanup capture)
 * - Warns on manually entered IDs that are not slug-safe
 * - Skips blank/empty IDs (builder auto-generates these)
 */
function checkManualIdHygiene(
  sheetName: string,
  columnEIndex: number,
  rawIds: Map<number, string>,
): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  // NOTE: Do NOT blanket-clear column E here. Duplicate-ID checks (Tasks 1/2)
  // have already set their annotations on column E. We only clear notes/fonts
  // for rows we are about to annotate, to avoid wiping prior findings.
  const logger = getScopedLogger("LintManualIdHygiene");

  // 1. Warn on whitespace-only IDs captured before cleanup
  rawIds.forEach((rawValue, row) => {
    const cell = sheet.getRange(row, columnEIndex);
    appendLintNote(
      cell,
      `Whitespace-only ID "${rawValue}" will be treated as empty and auto-generated.`,
      "warning",
    );
    logger.warn(
      `${sheetName} row ${row}: whitespace-only ID will be treated as empty`,
    );
  });

  // 2. Check all non-empty IDs for slug-safety
  const idRange = sheet.getRange(2, columnEIndex, lastRow - 1, 1);
  const idValues = idRange.getValues();
  const slugSafePattern = /^[a-z0-9]+(-[a-z0-9]+)*$/;

  for (let i = 0; i < idValues.length; i++) {
    const row = i + 2;
    const idValue = String(idValues[i][0] || "").trim();

    // Skip blank/empty (builder auto-generates)
    if (!idValue) continue;

    // Skip if this row was already flagged as whitespace-only
    if (rawIds.has(row)) continue;

    // Check slug-safety
    if (!slugSafePattern.test(idValue)) {
      const cell = sheet.getRange(row, columnEIndex);
      appendLintNote(
        cell,
        `Manual ID "${idValue}" is used as entered by the builder. Recommended format: lowercase letters, numbers, and hyphens (e.g., "my-category-id").`,
        "warning",
      );
      logger.warn(
        `${sheetName} row ${row}: non-slug-safe ID "${idValue}"`,
      );
    }
  }
}

/**
 * Canonical option parser that mirrors the builder's parseOptions() logic.
 * Splits by comma, trims, and for each option uses first-colon value:label format.
 * No-colon entries use slugified label as value.
 */
function parseCanonicalOptions(optionsStr: string): Array<{
  value: string;
  label: string;
  raw: string;
}> {
  if (!optionsStr) return [];

  const opts = optionsStr
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);
  if (opts.length === 0) return [];

  return opts.map((opt) => {
    const colonIndex = opt.indexOf(":");
    if (colonIndex > 0) {
      const value = opt.substring(0, colonIndex);
      const label = opt.substring(colonIndex + 1);
      return { value, label, raw: opt };
    }
    return {
      value: slugify(opt),
      label: opt,
      raw: opt,
    };
  });
}

/**
 * Detects ambiguous colon usage in select options that may indicate
 * accidental label splits rather than intentional value:label pairs.
 * Returns an array of warning messages for ambiguous patterns.
 */
function detectAmbiguousColonUsage(
  parsed: Array<{ value: string; label: string; raw: string }>,
): string[] {
  const warnings: string[] = [];

  for (const opt of parsed) {
    // Use the raw option text for colon analysis
    const rawText = opt.raw;
    const colonCount = (rawText.match(/:/g) || []).length;

    // Leading colon with empty value (e.g., ":label")
    if (colonCount > 0 && rawText.startsWith(":")) {
      warnings.push(
        `Option "${rawText}" has an empty value before the colon.`,
      );
      continue;
    }

    // Multiple colons (e.g., "value:label:extra")
    if (colonCount > 1) {
      warnings.push(
        `Option "${rawText}" contains multiple colons. Only the first colon is used to split value:label.`,
      );
      continue;
    }

    // Value part looks like natural language with spaces (e.g., "not applicable:No")
    if (
      opt.value.includes(" ") &&
      opt.label &&
      opt.value.length > 1
    ) {
      warnings.push(
        `Option "${rawText}" has spaces in the value portion ("${opt.value}"), which may be unintentional.`,
      );
    }
  }

  return warnings;
}

function isEmptyOrWhitespace(value: any): boolean {
  return (
    value === undefined ||
    value === null ||
    (typeof value === "string" && value.trim() === "")
  );
}

function cleanWhitespaceOnlyCells(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  startRow: number,
  startCol: number,
  numRows: number,
  numCols: number,
): void {
  const range = sheet.getRange(startRow, startCol, numRows, numCols);
  const values = range.getValues();
  let changesMade = false;

  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
      const value = values[i][j];
      if (typeof value === "string" && value.trim() === "" && value !== "") {
        values[i][j] = "";
        changesMade = true;
      }
    }
  }

  if (changesMade) {
    range.setValues(values);
    console.log(`Cleaned whitespace-only cells in ${sheet.getName()}`);
  }
}

function extractDriveFileId(url: string): string | null {
  if (!url) return null;
  const match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

function isSupportedSvgDataUri(dataUri: string): boolean {
  if (!dataUri.toLowerCase().startsWith("data:image/svg+xml")) return false;

  try {
    let svgText: string;
    if (dataUri.includes(";base64,")) {
      const base64 = dataUri.split(";base64,")[1];
      if (!base64) return false;
      svgText = Utilities.newBlob(
        Utilities.base64Decode(base64),
      ).getDataAsString();
    } else {
      const commaIndex = dataUri.indexOf(",");
      if (commaIndex === -1) return false;
      svgText = decodeURIComponent(dataUri.substring(commaIndex + 1));
    }
    return svgText.trim().startsWith("<svg");
  } catch (_error) {
    return false;
  }
}

function columnNumberToLetter(columnNumber: number): string {
  let dividend = columnNumber;
  let columnName = "";
  while (dividend > 0) {
    const modulo = (dividend - 1) % 26;
    columnName = String.fromCharCode(65 + modulo) + columnName;
    dividend = Math.floor((dividend - modulo) / 26);
  }
  return columnName;
}

// normalizeIconSlug is now defined in utils.ts and imported via reference path

function checkForDuplicates(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  columnIndex: number,
  startRow: number = 2,
): void {
  const lastRow = sheet.getLastRow();
  if (lastRow <= startRow) return;

  const range = sheet.getRange(
    startRow,
    columnIndex,
    lastRow - startRow + 1,
    1,
  );
  clearLintArtifacts(range);
  const values = range
    .getValues()
    .map((row) => row[0].toString().trim().toLowerCase());
  const duplicates = new Map<string, number[]>();

  values.forEach((value, index) => {
    if (value === "") return;

    if (!duplicates.has(value)) {
      duplicates.set(value, [index + startRow]);
    } else {
      duplicates.get(value)?.push(index + startRow);
    }
  });

  // Highlight duplicates
  duplicates.forEach((rows, value) => {
    if (rows.length > 1) {
      console.log(
        'Found duplicate value "' + value + '" in rows: ' + rows.join(", "),
      );
      const otherRowsStr = rows.join(", ");
      for (const row of rows) {
        setLintNote(
          sheet.getRange(row, columnIndex),
          `Duplicate value "${value}" found in rows: ${otherRowsStr}`,
          "error",
        );
      }
    }
  });
}

// Additional validation functions

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
    clearRangeLintNoteLinesWithPrefix(
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
      console.log("No categories exist - all details are unreferenced");
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

    const categoryFields = categoriesSheet
      .getRange(2, fieldsColOneBased, categoriesLastRow - 1, 1)
      .getValues()
      .map((row) => String(row[0] || ""))
      .filter((fields) => fields.trim() !== "");

    // Build set of all referenced field identifiers using normalizeFieldTokens
    const referencedFields = new Set<string>();
    for (const fieldsStr of categoryFields) {
      const tokens = normalizeFieldTokens(fieldsStr);
      for (const token of tokens) {
        const slugified = slugify(token);
        if (slugified) referencedFields.add(slugified);
        if (token) referencedFields.add(token);
      }
    }

    // Check each detail to see if it's referenced by slugified name or explicit ID
    for (const entry of detailEntries) {
      const isReferenced =
        (entry.slug && referencedFields.has(entry.slug)) ||
        (entry.explicitId && referencedFields.has(entry.explicitId));
      if (!isReferenced) {
        console.log(
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
    console.error("Error checking unreferenced details:", error);
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
    console.log(
      `Invalid Universal flag value "${value}" at row ${row} - must be TRUE, FALSE, or blank`,
    );
    setInvalidCellBackground(sheet, row, col, "#FFC7CE"); // Light red for invalid
  }
}

/**
 * Optional check for duplicate slugs in translation sheets.
 * Not invoked by default because duplicate translation values are allowed.
 */
function checkDuplicateTranslationSlugs(): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const translationSheets = sheets(true);

  for (const sheetName of translationSheets) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) continue;

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) continue;

    try {
      // Get all translated values and check for duplicate slugs
      const lastCol = sheet.getLastColumn();
      if (lastCol < 4) continue; // Need at least Name, ISO, Source, and one translation

      clearLintArtifacts(
        sheet.getRange(2, 4, lastRow - 1, lastCol - 3),
      );

      // Check each translation column (starting from column 4)
      for (let col = 4; col <= lastCol; col++) {
        const values = sheet
          .getRange(2, col, lastRow - 1, 1)
          .getValues()
          .map((row) => String(row[0] || "").trim())
          .filter((v) => v !== "");

        // Build slug frequency map
        const slugCounts = new Map<string, number[]>();
        for (let i = 0; i < values.length; i++) {
          const value = values[i];
          const slug = slugify(value);
          if (!slug) continue;

          if (!slugCounts.has(slug)) {
            slugCounts.set(slug, [i + 2]); // +2 for header and 0-index
          } else {
            slugCounts.get(slug)?.push(i + 2);
          }
        }

        // Highlight cells with duplicate slugs
        for (const [slug, rows] of slugCounts.entries()) {
          if (rows.length > 1) {
            console.log(
              `Duplicate slug "${slug}" in ${sheetName} column ${col} at rows: ${rows.join(", ")}`,
            );
            const otherRowsStr = rows.join(", ");
            for (const row of rows) {
              setLintNote(
                sheet.getRange(row, col),
                `Duplicate translation slug "${slug}" in rows: ${otherRowsStr}`,
                "warning",
              );
            }
          }
        }
      }
    } catch (error) {
      console.error(`Error checking duplicate slugs in ${sheetName}:`, error);
    }
  }
}

/**
 * Validate translation headers are valid language names, ISO codes, or "Name - ISO" format
 */
function validateTranslationHeaders(): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const translationSheets = sheets(true);
  const allLanguages = getAllLanguages();
  const resolveHeaderCode = createTranslationHeaderResolver(allLanguages);

  for (const sheetName of translationSheets) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) continue;

    try {
      const lastCol = sheet.getLastColumn();
      if (lastCol < 2) continue; // Need at least a source column and one language column

      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      clearLintArtifacts(sheet.getRange(1, 1, 1, lastCol));

      const headerB = String(headers[1] || "")
        .trim()
        .toLowerCase();
      const headerC = String(headers[2] || "")
        .trim()
        .toLowerCase();
      const hasMetaColumns =
        headerB.includes("iso") && headerC.includes("source");
      const languageStartIndex = hasMetaColumns ? 3 : 1;

      // Check language columns (skip meta columns if present)
      // Track resolved locale codes to detect duplicates across columns
      const localeToColumns = new Map<string, number[]>();

      for (let i = languageStartIndex; i < headers.length; i++) {
        const header = String(headers[i] || "").trim();
        if (!header) continue;

        const parsedCode = resolveHeaderCode(header);

        if (!parsedCode) {
          console.log(
            `Invalid translation header "${header}" in ${sheetName} column ${i + 1} - should be a language name, ISO code, or "Name - ISO" format`,
          );
          setLintNote(
            sheet.getRange(1, i + 1),
            `Invalid translation header "${header}" — should be a language name, ISO code, or "Name - ISO" format`,
            "error",
          );
          continue;
        }

        // Track locale → column indices for duplicate detection
        if (!localeToColumns.has(parsedCode)) {
          localeToColumns.set(parsedCode, [i]);
        } else {
          localeToColumns.get(parsedCode)?.push(i);
        }
      }

      // Flag duplicate locale headers
      localeToColumns.forEach((columns, code) => {
        if (columns.length > 1) {
          const firstCol = columns[0] + 1; // 1-based for display
          for (let idx = 1; idx < columns.length; idx++) {
            const colIndex = columns[idx];
            const header = String(headers[colIndex] || "").trim();
            const cell = sheet.getRange(1, colIndex + 1);
            setLintNote(
              cell,
              `Duplicate locale: header "${header}" resolves to "${code}" which is already used by column ${firstCol}`,
              "error",
            );
          }
        }
      });
    } catch (error) {
      console.error(`Error validating headers in ${sheetName}:`, error);
    }
  }
}

// Generic linting function
function lintSheet(
  sheetName: string,
  columnValidations: ((value: string, row: number, col: number) => void)[],
  requiredColumns: number[] = [],
  preserveBackgroundColumns: number[] = [],
): void {
  console.time(`Linting ${sheetName}`);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    console.log(`${sheetName} sheet not found`);
    console.timeEnd(`Linting ${sheetName}`);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    console.log(`${sheetName} sheet is empty or contains only header`);
    console.timeEnd(`Linting ${sheetName}`);
    return;
  }

  try {
    if (lastRow > 1 && columnValidations.length > 0) {
      // Clear backgrounds, font colors, and stale lint notes for all columns
      // except those that should be preserved
      for (let col = 0; col < columnValidations.length; col++) {
        // Skip columns that should preserve their backgrounds (e.g., Categories icon column with user colors)
        const shouldClearBackground = !preserveBackgroundColumns.includes(col);
        const colRange = sheet.getRange(2, col + 1, lastRow - 1, 1);

        if (shouldClearBackground) {
          clearRangeBackgroundIfMatches(
            colRange,
            LINT_WARNING_BACKGROUND_COLORS,
          );
        }
        clearRangeFontColorIfMatches(colRange, LINT_WARNING_FONT_COLORS);
        // Clear stale [Lint] notes so re-linting produces a clean result
        clearRangeLintNoteLinesWithPrefix(colRange, LINT_NOTE_PREFIX);
      }
    }

    // First clean any whitespace-only cells
    console.time(`Cleaning whitespace cells for ${sheetName}`);
    cleanWhitespaceOnlyCells(
      sheet,
      2,
      1,
      lastRow - 1,
      columnValidations.length,
    );
    console.timeEnd(`Cleaning whitespace cells for ${sheetName}`);

    // Check for duplicates in the first column (usually the name/identifier column)
    console.time(`Checking for duplicates in ${sheetName}`);
    checkForDuplicates(sheet, 1);
    console.timeEnd(`Checking for duplicates in ${sheetName}`);

    console.time(`Getting data for ${sheetName}`);
    // Get all data from the sheet, excluding the header row
    const dataRange = sheet.getRange(
      2,
      1,
      lastRow - 1,
      columnValidations.length,
    );
    const data = dataRange.getValues();
    console.timeEnd(`Getting data for ${sheetName}`);

    console.time(`Validating cells for ${sheetName}`);

    // Highlight required fields in batches before running column validations
    if (requiredColumns.length > 0) {
      const rangesToReset: string[] = [];
      requiredColumns.forEach((colIndex) => {
        // Skip columns that should preserve their backgrounds
        if (!preserveBackgroundColumns.includes(colIndex)) {
          const columnLetter = columnNumberToLetter(colIndex + 1);
          rangesToReset.push(`${columnLetter}2:${columnLetter}${lastRow}`);
        }
      });

      if (rangesToReset.length > 0) {
        sheet.getRangeList(rangesToReset).setBackground(null);
      }

      const requiredHighlights = new Map<number, number[]>();
      requiredColumns.forEach((colIndex) => {
        // Skip columns that should preserve their backgrounds
        if (!preserveBackgroundColumns.includes(colIndex)) {
          requiredHighlights.set(colIndex, []);
        }
      });

      data.forEach((row, rowIndex) => {
        requiredColumns.forEach((colIndex) => {
          // Skip columns that should preserve their backgrounds
          if (
            !preserveBackgroundColumns.includes(colIndex) &&
            isEmptyOrWhitespace(row[colIndex])
          ) {
            const rows = requiredHighlights.get(colIndex);
            if (rows) {
              rows.push(rowIndex + 2); // +2 accounts for header row
            }
          }
        });
      });

      requiredHighlights.forEach((rows, colIndex) => {
        if (!rows || rows.length === 0) {
          return;
        }
        const columnLetter = columnNumberToLetter(colIndex + 1);
        const rangeAddresses = rows.map(
          (rowNumber) => `${columnLetter}${rowNumber}`,
        );
        sheet.getRangeList(rangeAddresses).setBackground("#FFF2CC"); // Light yellow for required fields
      });
    }

    // Iterate through each cell and apply the corresponding validation function
    data.forEach((row, rowIndex) => {
      row.forEach((cellValue, colIndex) => {
        if (columnValidations[colIndex]) {
          columnValidations[colIndex](
            String(cellValue || ""),
            rowIndex + 2,
            colIndex + 1,
          );
        }
      });
    });
    console.timeEnd(`Validating cells for ${sheetName}`);

    console.log(`${sheetName} sheet linting completed`);
  } catch (error) {
    console.error(`Error linting ${sheetName} sheet:`, error);
  } finally {
    console.timeEnd(`Linting ${sheetName}`);
  }
}

function getDriveIconInfo(fileId: string): {
  slug: string | null;
  isSvg: boolean;
  errorMessage?: string;
} {
  try {
    const file = DriveApp.getFileById(fileId);
    const fileName = file.getName();
    const mimeType = file.getMimeType();
    const mimeTypeLower = mimeType?.toLowerCase() ?? "";
    const nameWithoutExt = fileName.replace(/\.[^/.]+$/, "");
    const slug = normalizeIconSlug(slugify(nameWithoutExt));

    // Fast path: MIME type already tells us it's SVG
    if (mimeTypeLower.includes("svg")) {
      return { slug: slug || null, isSvg: true };
    }

    // Fallback: try reading as text to detect SVG content
    // Binary files (PNG/JPEG) will fail here, so treat as non-SVG
    let isSvg = false;
    try {
      isSvg = file.getBlob().getDataAsString().trim().startsWith("<svg");
    } catch {
      // Binary file — can't read as text, treat as non-SVG
      isSvg = false;
    }

    return { slug: slug || null, isSvg };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return {
      slug: null,
      isSvg: false,
      errorMessage: `Unable to access icon file (Drive ID ${fileId}): ${message}`,
    };
  }
}

function validateCategoryIcons(): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const categoriesSheet = spreadsheet.getSheetByName("Categories");

  if (!categoriesSheet) {
    console.log("Categories sheet not found during icon validation");
    return;
  }

  const lastRow = categoriesSheet.getLastRow();
  if (lastRow <= 1) {
    console.log("No category rows available for icon validation");
    return;
  }

  const iconRange = categoriesSheet.getRange(2, 2, lastRow - 1, 1);
  // Clear icon column lint artifacts: backgrounds, font colors, and notes
  clearLintArtifacts(iconRange);

  const iconValues = iconRange.getValues();
  const driveFileCache = new Map<
    string,
    { slug: string | null; isSvg: boolean; errorMessage?: string }
  >();
  const rowIssues = new Map<number, string[]>();

  const addIssue = (row: number, message: string): void => {
    if (!rowIssues.has(row)) {
      rowIssues.set(row, []);
    }
    const issues = rowIssues.get(row)!;
    if (!issues.includes(message)) {
      issues.push(message);
    }
  };

  iconValues.forEach((row, index) => {
    const rowNumber = index + 2;
    const iconCellValue = row[0];

    if (
      iconCellValue === null ||
      iconCellValue === undefined ||
      iconCellValue === ""
    ) {
      // Missing icon — warn (builder creates category without icon, not a hard error)
      const cell = categoriesSheet.getRange(rowNumber, 2);
      setLintNote(
        cell,
        "Icon is empty — a default icon will be used during config generation",
        "warning",
      );
      return;
    }

    if (typeof iconCellValue === "string") {
      const iconValue = iconCellValue.trim();
      if (!iconValue) {
        // Whitespace-only — treat as missing
        const cell = categoriesSheet.getRange(rowNumber, 2);
        setLintNote(
          cell,
          "Icon is empty — a default icon will be used during config generation",
          "warning",
        );
        return;
      }

      if (iconValue.startsWith("<svg")) {
        // Inline SVG markup - passed through
        return;
      } else if (isSupportedSvgDataUri(iconValue)) {
        // Valid SVG data URI - passed through
        return;
      } else if (iconValue.startsWith("data:")) {
        addIssue(
          rowNumber,
          "Only SVG data URIs (data:image/svg+xml) are supported for direct config export; other data URIs are silently dropped during generation.",
        );
      } else if (iconValue.startsWith("https://drive.google.com/file/d/")) {
        // Drive URL in builder-supported form - validate access
        const fileId = extractDriveFileId(iconValue);
        if (fileId) {
          let info = driveFileCache.get(fileId);
          if (!info) {
            info = getDriveIconInfo(fileId);
            driveFileCache.set(fileId, info);
          }

          if (info.slug && info.isSvg) {
            // Access is valid; no lint issue.
            return;
          }

          if (info.errorMessage) {
            addIssue(rowNumber, info.errorMessage);
          } else if (info.slug) {
            addIssue(
              rowNumber,
              "Google Drive icon files must be SVG for direct config export; non-SVG Drive files are silently dropped during generation.",
            );
          }
        } else {
          addIssue(
            rowNumber,
            "Icon URL must contain a valid Google Drive file ID.",
          );
        }
      } else if (iconValue.startsWith("https://drive.google.com/")) {
        addIssue(
          rowNumber,
          "Google Drive icon URLs must use the /file/d/ form (for example, https://drive.google.com/file/d/<FILE_ID>/view) so config generation can package them.",
        );
      } else if (/^http:\/\//i.test(iconValue)) {
        if (iconValue.toLowerCase().includes(".svg")) {
          // HTTP SVG URL - warn about security (should use HTTPS)
          const cell = categoriesSheet.getRange(rowNumber, 2);
          setLintNote(
            cell,
            "Icon URL should use HTTPS instead of HTTP for security",
            "warning",
          );
          return;
        }
        addIssue(
          rowNumber,
          'HTTP(S) icon URLs must point directly to an SVG file (URL must contain ".svg"); non-SVG URLs are silently dropped during config generation.',
        );
      } else if (/^https:\/\//i.test(iconValue)) {
        if (iconValue.toLowerCase().includes(".svg")) {
          // HTTPS SVG URL - passed through
          return;
        }
        addIssue(
          rowNumber,
          'HTTP(S) icon URLs must point directly to an SVG file (URL must contain ".svg"); non-SVG URLs are silently dropped during config generation.',
        );
      } else {
        // Plain text - works with 'Generate Category Icons' search but not direct export
        setLintNote(
          categoriesSheet.getRange(rowNumber, 2),
          `Plain text icon '${iconValue}' works with 'Generate Category Icons' search, but direct config export only packages supported icon sources (inline SVG, SVG data URI, Drive SVG, or direct SVG URL).`,
          "warning",
        );
        return;
      }
    } else if (
      iconCellValue &&
      typeof iconCellValue === "object" &&
      iconCellValue.toString() === "CellImage"
    ) {
      // Cell images ARE supported and will be processed via icon API
      return;
    } else {
      addIssue(
        rowNumber,
        "Unrecognized icon cell value. Expected text, Drive URL, or cell image.",
      );
    }
  });

  rowIssues.forEach((messages, rowNumber) => {
    const cell = categoriesSheet.getRange(rowNumber, 2);
    setLintNote(cell, messages.join("\n"), "error");
    console.warn(
      `Icon issue in Categories row ${rowNumber}: ${messages.join(" | ")}`,
    );
  });

  if (rowIssues.size === 0) {
    console.log("Category icon validation completed with no issues found.");
  }
}

function validatePrimaryLanguageInA1(): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const categoriesSheet = spreadsheet.getSheetByName("Categories");
  if (!categoriesSheet) return;

  const cell = categoriesSheet.getRange(1, 1);
  clearLintArtifacts(cell);

  // Mirror getPrimaryLanguageName(): if Metadata sheet has a primaryLanguage
  // entry, the builder uses that and ignores A1 entirely. Only validate A1
  // when it is the effective primary-language source.
  const metadataSheet = spreadsheet.getSheetByName("Metadata");
  if (metadataSheet) {
    const metadataValues = metadataSheet.getDataRange().getValues();
    for (let i = 1; i < metadataValues.length; i++) {
      if (String(metadataValues[i][0]).trim() === "primaryLanguage") {
        const lang = String(metadataValues[i][1] || "").trim();
        if (lang) return; // Metadata has primaryLanguage — A1 is not used
      }
    }
  }

  const a1Value = String(categoriesSheet.getRange(1, 1).getValue() || "").trim();
  if (!a1Value) return; // Empty A1 – builder falls back to Metadata

  // Skip standard headers that are clearly not language names
  const lowerA1 = a1Value.toLowerCase();
  const HEADER_WORDS = ["category", "categories", "name", "label", "type"];
  if (HEADER_WORDS.includes(lowerA1)) return;

  // Validate as a language name
  const validation = validateLanguageName(a1Value);
  if (!validation.valid) {
    setLintNote(
      cell,
      `Invalid primary language in A1: ${validation.error || "Unknown error"}`,
      "error",
    );
  }
}

// Specific sheet linting functions
function lintCategoriesSheet(): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const categoriesSheetRef = spreadsheet.getSheetByName("Categories");
  const detailsSheetRef = spreadsheet.getSheetByName("Details");

  // Task 5: Check if Categories sheet is empty before running lint.
  // Must run BEFORE validatePrimaryLanguageInA1() to avoid the empty-sheet
  // check wiping the A1 language error note (both target cell A1).
  if (categoriesSheetRef && checkEmptySheet(categoriesSheetRef, "Categories")) {
    return;
  }

  // Phase 4: Validate primary language in A1 (only when sheet has data rows)
  validatePrimaryLanguageInA1();

  // Pre-cleanup: capture whitespace-only IDs before cleanWhitespaceOnlyCells() erases them
  const rawCategoryIds = inspectRawIds("Categories", 5);

  if (categoriesSheetRef) {
    const lastRow = categoriesSheetRef.getLastRow();
    if (lastRow > 1) {
      // Clear font colors in icon column (column 2) for lint warnings
      clearRangeFontColorIfMatches(
        categoriesSheetRef.getRange(2, 2, lastRow - 1, 1),
        LINT_WARNING_FONT_COLORS,
      );
      // Clear background colors in fields column (column 3)
      clearRangeBackgroundIfMatches(
        categoriesSheetRef.getRange(2, 3, lastRow - 1, 1),
        ["#FFC7CE"],
      );
    }
  }

  // Build a Set of all valid field identifiers: slugified names AND explicit IDs
  let cachedValidFieldIds: Set<string>;
  if (detailsSheetRef) {
    const lastRow = detailsSheetRef.getLastRow();
    if (lastRow > 1) {
      const detailData = detailsSheetRef
        .getRange(2, 1, lastRow - 1, 5)
        .getValues();
      cachedValidFieldIds = new Set<string>();
      for (const row of detailData) {
        const name = String(row[0] || "").trim();
        const explicitId = String(row[4] || "").trim();
        if (name) {
          cachedValidFieldIds.add(slugify(name));
        }
        if (explicitId) {
          cachedValidFieldIds.add(explicitId);
        }
      }
    }
  }
  cachedValidFieldIds = cachedValidFieldIds || new Set<string>();

  const categoriesValidations = [
    // Rule 1: Capitalize the first letter of the category name
    (value, row, col) => {
      if (isEmptyOrWhitespace(value)) return;

      const capitalizedValue = capitalizeFirstLetter(value);
      if (capitalizedValue !== value) {
        try {
          categoriesSheetRef?.getRange(row, col).setValue(capitalizedValue);
        } catch (error) {
          console.error(
            "Error capitalizing value in Categories sheet at row " +
              row +
              ", col " +
              col +
              ":",
            error,
          );
        }
      }
    },
    // Rule 2: Icon validation is handled entirely by validateCategoryIcons()
    // which runs after lintSheet() to avoid clear/write race conditions.
    () => {
      // Icon column — no inline validation; see validateCategoryIcons()
    },
    // Rule 3: Validate field references using normalizeFieldTokens (matches build parsing)
    (value, row, col) => {
      if (isEmptyOrWhitespace(value)) return;

      try {
        // Use normalizeFieldTokens to match build parsing across commas, semicolons, newlines, bullets, fullwidth
        if (cachedValidFieldIds.size > 0) {
          const tokens = normalizeFieldTokens(value);
          const invalidFields: string[] = [];
          for (const token of tokens) {
            const slugified = slugify(token);
            if (slugified && !cachedValidFieldIds.has(slugified) && !cachedValidFieldIds.has(token)) {
              invalidFields.push(token);
            }
          }

          if (invalidFields.length > 0) {
            console.log(
              "Invalid fields in row " + row + ": " + invalidFields.join(", "),
            );
            const cell = categoriesSheetRef?.getRange(row, col);
            if (cell) {
              setLintNote(
                cell,
                `Invalid fields: ${invalidFields.join(", ")}. These fields do not exist in the Details sheet`,
                "error",
              );
            }
          }
        }
      } catch (error) {
        console.error(
          "Error validating fields in Categories sheet at row " +
            row +
            ", col " +
            col +
            ":",
          error,
        );
      }
    },
  ];

  // Category name and icon are required
  // Preserve backgrounds in category name column (index 0) because they are user-set category colors
  lintSheet("Categories", categoriesValidations, [0, 1], [0]);
  validateCategoryIcons();

  // Phase 2: Post-lintSheet() edge-case checks
  checkDuplicateCategoryIds(); // Task 1: duplicate effective IDs in col E
  checkSlugCollisions("Categories", 1); // Task 3: slug collisions in col A
  validateAppliesColumn(); // Task 4: Applies column validation
  checkManualIdHygiene("Categories", 5, rawCategoryIds); // Task 6: ID hygiene in col E
}

function lintDetailsSheet(): void {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Details");
  if (!sheet) {
    console.log("Details sheet not found");
    return;
  }

  // Task 5: Check if Details sheet is empty before running lint
  if (checkEmptySheet(sheet, "Details")) {
    return;
  }

  // Pre-cleanup: capture whitespace-only IDs before cleanWhitespaceOnlyCells() erases them
  const rawDetailIds = inspectRawIds("Details", 5);

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const lastColumn = sheet.getLastColumn();
    if (lastColumn >= 3) {
      clearRangeBackgroundIfMatches(sheet.getRange(2, 3, lastRow - 1, 1), [
        "#FFC7CE",
      ]);
    }
    if (lastColumn >= 4) {
      clearRangeBackgroundIfMatches(sheet.getRange(2, 4, lastRow - 1, 1), [
        "#FFC7CE",
      ]);
    }
    if (lastColumn >= 6) {
      clearRangeBackgroundIfMatches(sheet.getRange(2, 6, lastRow - 1, 1), [
        "#FFC7CE",
      ]);
    }
  }

  const detailsValidations = [
    // Rule 1: Capitalize the first letter of the detail name
    (value, row, col) => {
      if (isEmptyOrWhitespace(value)) return;

      const capitalizedValue = capitalizeFirstLetter(value);
      if (capitalizedValue !== value) {
        try {
          SpreadsheetApp.getActiveSpreadsheet()
            .getSheetByName("Details")
            ?.getRange(row, col)
            .setValue(capitalizedValue);
        } catch (error) {
          console.error(
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
    (value, row, col) => {
      if (isEmptyOrWhitespace(value)) return;

      const capitalizedValue = capitalizeFirstLetter(value);
      if (capitalizedValue !== value) {
        try {
          SpreadsheetApp.getActiveSpreadsheet()
            .getSheetByName("Details")
            ?.getRange(row, col)
            .setValue(capitalizedValue);
        } catch (error) {
          console.error(
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
    (value, row, col) => {
      // Type column validation logic:
      // - blank/empty → text (valid, mirrors payloadBuilder default)
      // - "s*" (select, single, etc.) → selectOne (valid)
      // - "m*" (multi, multiple, etc.) → selectMultiple (valid)
      // - "n*" (number, numeric, etc.) → number (valid)
      // - "t*" (text, textual, etc.) → text (valid)
      // - Any other value → invalid

      const detailsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Details");

      // Empty/blank is valid (defaults to text per payloadBuilder) - add advisory note
      if (isEmptyOrWhitespace(value)) {
        if (detailsSheet) {
          setLintNote(
            detailsSheet.getRange(row, col),
            "Blank type defaults to text. Enter 'text', 'number', 'selectOne', or 'selectMultiple' to be explicit.",
            "advisory",
          );
        }
        return;
      }

      const firstChar = value.toLowerCase().charAt(0);
      const validTypes = ["t", "n", "m", "s"];

      if (!validTypes.includes(firstChar)) {
        try {
          console.log("Invalid type '" + value + "' at row " + row);
          setInvalidCellBackground(
            SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Details")!,
            row,
            col,
            "#FFC7CE",
          ); // Light red for invalid type
        } catch (error) {
          console.error(
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
    (value, row, col) => {
      try {
        // Get the type from column 3 (index 2) to determine if options are required
        const typeValue = sheet.getRange(row, 3).getValue();
        const typeStr = String(typeValue || "").trim();

        // Determine the field type category
        // Mirror payloadBuilder: empty type defaults to "text" (not selectOne)
        const firstChar = isEmptyOrWhitespace(typeStr)
          ? "t"
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
            console.log(
              "Select field at row " + row + " is missing required options",
            );
            setInvalidCellBackground(sheet, row, col, "#FFC7CE"); // Light red for missing options
            return;
          }

          // Parse options using canonical parser (mirrors builder's parseOptions)
          const parsed = parseCanonicalOptions(value);

          if (parsed.length === 0) {
            console.log(
              "Select field at row " +
                row +
                " has empty options after trimming",
            );
            setInvalidCellBackground(sheet, row, col, "#FFC7CE"); // Light red for empty options
            return;
          }

          // Task 5: Check for ambiguous colon usage
          const ambiguityWarnings = detectAmbiguousColonUsage(parsed);
          if (ambiguityWarnings.length > 0) {
            const cell = sheet.getRange(row, col);
            setLintNote(
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

            console.log(
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
        console.error(
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
    (value, row, col) => {
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

/**
 * Phase 5 Task 1: Validates the Icons sheet.
 * Checks for missing icon IDs, duplicate icon IDs, missing icon sources,
 * and unsupported icon source formats.
 */
function lintIconsSheet(): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const iconsSheet = spreadsheet.getSheetByName("Icons");
  if (!iconsSheet) return; // Icons sheet is optional

  const lastRow = iconsSheet.getLastRow();
  if (lastRow <= 1) return; // Header-only or empty

  const logger = getScopedLogger("LintIconsSheet");

  // Clear previous lint artifacts on columns A and B
  const dataRange = iconsSheet.getRange(2, 1, lastRow - 1, 2);
  clearLintArtifacts(dataRange);

  const data = dataRange.getValues();
  const seenIds = new Map<string, number[]>();
  const driveFileCache = new Map<
    string,
    { slug: string | null; isSvg: boolean; errorMessage?: string }
  >();

  for (let i = 0; i < data.length; i++) {
    const row = i + 2;
    const iconId = String(data[i][0] || "").trim();
    const iconSource = String(data[i][1] || "").trim();

    // Check for missing icon ID (col A)
    if (!iconId) {
      setLintNote(
        iconsSheet.getRange(row, 1),
        "Icon ID is required.",
        "error",
      );
      logger.warn(`Row ${row}: missing icon ID`);
      continue;
    }

    // Track for duplicate detection
    if (!seenIds.has(iconId)) {
      seenIds.set(iconId, [row]);
    } else {
      seenIds.get(iconId)?.push(row);
    }

    // Check for missing icon source (col B)
    if (!iconSource) {
      setLintNote(
        iconsSheet.getRange(row, 2),
        "Icon source is required.",
        "error",
      );
      logger.warn(`Row ${row}: missing icon source for ID "${iconId}"`);
      continue;
    }

    // Check for unsupported icon source format.
    // Mirrors parseIconSource() in payloadBuilder: only inline SVG, data:image/svg+xml,
    // Google Drive links, and HTTP(S) URLs ending in .svg are accepted.
    const isSvg = iconSource.startsWith("<svg");
    const isDataUri = iconSource.toLowerCase().startsWith("data:image/svg+xml");
    const isDriveUrl = iconSource.startsWith("https://drive.google.com/file/d/");
    const isHttpUrl = /^https?:\/\//i.test(iconSource);
    // Builder's isSvgUrl() requires the URL to contain ".svg"
    const isSvgHttpUrl = isHttpUrl && iconSource.toLowerCase().includes(".svg");

    if (isSvg) {
      // Inline SVG — check for basic structural validity
      if (!iconSource.includes("</svg>")) {
        setLintNote(
          iconsSheet.getRange(row, 2),
          'Inline SVG markup appears incomplete (missing closing </svg> tag). This icon will be dropped during config generation.',
          "error",
        );
      }
    } else if (isDataUri) {
      // Data URI — verify it actually decodes to valid SVG content
      if (!isSupportedSvgDataUri(iconSource)) {
        setLintNote(
          iconsSheet.getRange(row, 2),
          "SVG data URI is malformed or does not decode to valid SVG content. This icon will be silently dropped during config generation.",
          "error",
        );
      }
    } else if (isDriveUrl) {
      // Drive URL — verify the file is accessible and is SVG
      const fileId = extractDriveFileId(iconSource);
      if (fileId) {
        let info = driveFileCache.get(fileId);
        if (!info) {
          info = getDriveIconInfo(fileId);
          driveFileCache.set(fileId, info);
        }
        if (info.errorMessage) {
          setLintNote(
            iconsSheet.getRange(row, 2),
            info.errorMessage,
            "error",
          );
        } else if (!info.isSvg) {
          setLintNote(
            iconsSheet.getRange(row, 2),
            "Google Drive icon file is not SVG. Non-SVG Drive files are silently dropped during config generation.",
            "error",
          );
        }
      } else {
        setLintNote(
          iconsSheet.getRange(row, 2),
          "Google Drive URL does not contain a valid file ID.",
          "error",
        );
      }
    } else if (!isSvgHttpUrl) {
      if (isHttpUrl) {
        // HTTP URL but not an SVG — will be silently dropped by parseIconSource()
        setLintNote(
          iconsSheet.getRange(row, 2),
          "HTTP(S) icon URLs must point directly to an SVG file (URL must contain \".svg\"). Non-SVG URLs are silently dropped during config generation.",
          "error",
        );
      } else {
        setLintNote(
          iconsSheet.getRange(row, 2),
          "Unsupported icon source format. Expected inline SVG (<svg…>), data:image/svg+xml URI, Google Drive URL (must use /file/d/ format), or HTTP(S) URL ending in .svg.",
          "warning",
        );
      }
      logger.warn(
        `Row ${row}: unsupported icon source format for ID "${iconId}"`,
      );
    }
  }

  // Flag duplicate IDs
  seenIds.forEach((rows, id) => {
    if (rows.length > 1) {
      logger.warn(
        `Duplicate icon ID "${id}" in rows: ${rows.join(", ")}`,
      );
      for (const row of rows) {
        const otherRows = rows.filter((r) => r !== row);
        appendLintNote(
          iconsSheet.getRange(row, 1),
          `Duplicate icon ID "${id}" (also in row${otherRows.length > 1 ? "s" : ""} ${otherRows.join(", ")}).`,
          "error",
        );
      }
    }
  });
}

/**
 * Returns exact icon ID collisions between Icons sheet rows and
 * category-derived icon rows. This must mirror buildIconsFromSheet(), which
 * deduplicates by exact icon.id values rather than sanitized variants.
 */
function findCrossSheetIconIdCollisions(
  iconsEntries: Array<{ id: string; row: number }>,
  categoryEntries: Array<{ id: string; row: number }>,
): Array<{
  iconId: string;
  iconRow: number;
  categoryId: string;
  categoryRow: number;
}> {
  const categoryEntriesById = new Map<string, Array<{ id: string; row: number }>>();
  for (const entry of categoryEntries) {
    if (!entry.id) continue;
    if (!categoryEntriesById.has(entry.id)) {
      categoryEntriesById.set(entry.id, []);
    }
    categoryEntriesById.get(entry.id)?.push(entry);
  }

  const collisions: Array<{
    iconId: string;
    iconRow: number;
    categoryId: string;
    categoryRow: number;
  }> = [];

  for (const iconEntry of iconsEntries) {
    const matchingCategoryEntries = categoryEntriesById.get(iconEntry.id);
    if (!matchingCategoryEntries) continue;
    for (const categoryEntry of matchingCategoryEntries) {
      collisions.push({
        iconId: iconEntry.id,
        iconRow: iconEntry.row,
        categoryId: categoryEntry.id,
        categoryRow: categoryEntry.row,
      });
    }
  }

  return collisions;
}

/**
 * Phase 5 Task 2: Checks for exact icon ID collisions between the Icons sheet
 * and Categories sheet. The builder merges icons from both sources,
 * deduplicating by exact ID, so lint must do the same.
 */
function checkCrossSheetIconCollisions(): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const iconsSheet = spreadsheet.getSheetByName("Icons");
  const categoriesSheet = spreadsheet.getSheetByName("Categories");
  if (!iconsSheet || !categoriesSheet) return;

  const iconsLastRow = iconsSheet.getLastRow();
  const categoriesLastRow = categoriesSheet.getLastRow();
  if (iconsLastRow <= 1 || categoriesLastRow <= 1) return;

  const logger = getScopedLogger("LintCrossSheetIconCollisions");

  // Read Icons sheet IDs (col A)
  const iconsData = iconsSheet
    .getRange(2, 1, iconsLastRow - 1, 1)
    .getValues();
  const iconsEntries: Array<{ id: string; row: number }> = [];
  for (let i = 0; i < iconsData.length; i++) {
    const iconId = String(iconsData[i][0] || "").trim();
    if (iconId) {
      iconsEntries.push({ id: iconId, row: i + 2 });
    }
  }

  // Read Categories: resolve columns by header name (mirrors buildIconsFromSheet).
  const catLastCol = Math.max(categoriesSheet.getLastColumn(), 6);
  const catHeaders = categoriesSheet.getRange(1, 1, 1, catLastCol).getValues()[0];
  const catHeaderMap = createLintHeaderMap(catHeaders);
  const catNameCol = getLintColumnIndex(catHeaderMap, "name") ?? 0;
  const catIconCol = getLintColumnIndex(catHeaderMap, "icon", "icons") ?? 1;
  const catCategoryIdCol = getLintColumnIndex(catHeaderMap, "category id", "id");
  const catIconIdCol = getLintColumnIndex(catHeaderMap, "icon id", "iconid");

  const catData = categoriesSheet.getRange(2, 1, categoriesLastRow - 1, catLastCol).getValues();
  const categoryEntries: Array<{ id: string; row: number }> = [];
  for (let i = 0; i < catData.length; i++) {
    const name = String(catData[i][catNameCol] || "").trim();
    const iconRaw = catData[i][catIconCol];
    const iconStr = typeof iconRaw === "string" ? iconRaw.trim() : "";
    // Mirror buildIconsFromSheet(): only category rows with a name and a
    // supported icon source contribute packaged icon assets.
    if (!name || !iconStr || !parseIconSource(iconStr)) continue;
    // Mirror buildIconsFromSheet(): slugify(name) || `category-${index+1}`
    // The builder falls back to `category-N` when slugify produces an empty
    // string (e.g. non-Latin or punctuation-only names).
    const explicitCategoryId = catCategoryIdCol !== undefined
      ? String(catData[i][catCategoryIdCol] || "").trim()
      : "";
    const categoryId = explicitCategoryId || slugify(name) || `category-${i + 1}`;
    // Builder uses iconIdCol ("icon id") when present, otherwise falls back to categoryId
    const explicitIconId = catIconIdCol !== undefined
      ? String(catData[i][catIconIdCol] || "").trim()
      : "";
    const iconId = explicitIconId || categoryId;
    if (iconId) {
      categoryEntries.push({ id: iconId, row: i + 2 });
    }
  }

  const collisions = findCrossSheetIconIdCollisions(iconsEntries, categoryEntries);
  for (const collision of collisions) {
    logger.warn(
      `Icon ID collision: "${collision.iconId}" (Icons row ${collision.iconRow}) vs "${collision.categoryId}" (Categories row ${collision.categoryRow})`,
    );
    appendLintNote(
      iconsSheet.getRange(collision.iconRow, 1),
      `Icon ID "${collision.iconId}" collides with a category-derived ID in Categories row ${collision.categoryRow}.`,
      "error",
    );
    appendLintNote(
      categoriesSheet.getRange(collision.categoryRow, 2),
      `Category icon ID "${collision.categoryId}" collides with an Icons sheet entry in row ${collision.iconRow}.`,
      "error",
    );
  }
}

/**
 * Warns when multiple translation rows in the same sheet have source values
 * that slugify to the same key. Later rows would silently overwrite earlier ones
 * during config generation.
 */
function checkTranslationSourceOverwrites(): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const translationSheets = sheets(true);

  for (const sheetName of translationSheets) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) continue;

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) continue;

    try {
      // Source column is always column A (index 0) for all translation sheets
      const sourceRange = sheet.getRange(2, 1, lastRow - 1, 1);
      clearLintArtifacts(sourceRange);
      const sourceValues = sourceRange.getValues();

      // Track slug → array of { row, originalValue }
      const slugMap = new Map<string, Array<{ row: number; value: string }>>();

      for (let i = 0; i < sourceValues.length; i++) {
        const value = String(sourceValues[i][0] || "").trim();
        if (!value) continue;

        const slug = slugify(value);
        if (!slug) continue;

        const entry = { row: i + 2, value };
        if (!slugMap.has(slug)) {
          slugMap.set(slug, [entry]);
        } else {
          slugMap.get(slug)?.push(entry);
        }
      }

      // Warn on duplicate slugs
      slugMap.forEach((entries, slug) => {
        if (entries.length > 1) {
          for (const entry of entries) {
            const otherRows = entries
              .filter((e) => e.row !== entry.row)
              .map((e) => e.row);
            const cell = sheet.getRange(entry.row, 1);
            appendLintNote(
              cell,
              `Source value "${entry.value}" produces the same key as row ${otherRows.join(", ")}. Later values may overwrite earlier ones.`,
              "warning",
            );
          }
        }
      });
    } catch (error) {
      console.error(
        `Error checking source overwrites in ${sheetName}:`,
        error,
      );
    }
  }
}

function lintTranslationSheets(): void {
  // First validate translation headers
  console.log("Validating translation headers...");
  validateTranslationHeaders();

  const translationSheets = sheets(true);
  translationSheets.forEach((sheetName) => {
    const sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      console.error(`Sheet "${sheetName}" not found`);
      return;
    }

    try {
      console.log("Linting translation sheet: " + sheetName);

      // First clean any whitespace-only cells
      cleanWhitespaceOnlyCells(
        sheet,
        1,
        1,
        sheet.getLastRow(),
        sheet.getLastColumn(),
      );

      // Get all data from the sheet
      const data = sheet.getDataRange().getValues();
      // Capitalize the first letter of each cell if it's a string and not empty
      const updatedData = data.map((row) =>
        row.map((cell) =>
          typeof cell === "string" && cell.trim() !== ""
            ? capitalizeFirstLetter(cell)
            : cell,
        ),
      );

      // Update the sheet with the capitalized data
      sheet.getDataRange().setValues(updatedData);
      console.log("Finished linting translation sheet: " + sheetName);
    } catch (error) {
      console.error(
        "Error linting translation sheet " + sheetName + ":",
        error,
      );
    }
  });

  // After basic linting, validate translation sheet consistency
  validateTranslationSheetConsistency();

  // Phase 4: Warn on source-value slug collisions that cause silent overwrites
  checkTranslationSourceOverwrites();
}

/**
 * Validates that translation sheets have consistent headers and row counts with their source sheets.
 */
function validateTranslationSheetConsistency(): void {
  console.log("Validating translation sheet consistency...");

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // Validate Category Translations
    const categoriesSheet = spreadsheet.getSheetByName("Categories");
    const categoryTranslationsSheet = spreadsheet.getSheetByName(
      "Category Translations",
    );

    if (categoriesSheet && categoryTranslationsSheet) {
      validateSheetConsistency(
        categoriesSheet,
        categoryTranslationsSheet,
        "Category Translations",
        false,
      );
    }

    // Validate Detail translations
    const detailsSheet = spreadsheet.getSheetByName("Details");

    if (detailsSheet) {
      const detailLabelTranslations = spreadsheet.getSheetByName(
        "Detail Label Translations",
      );
      const detailHelperTranslations = spreadsheet.getSheetByName(
        "Detail Helper Text Translations",
      );
      const detailOptionTranslations = spreadsheet.getSheetByName(
        "Detail Option Translations",
      );

      if (detailLabelTranslations) {
        validateSheetConsistency(
          detailsSheet,
          detailLabelTranslations,
          "Detail Label Translations",
          false,
        );
      }

      if (detailHelperTranslations) {
        validateSheetConsistency(
          detailsSheet,
          detailHelperTranslations,
          "Detail Helper Text Translations",
          false,
        );
      }

      if (detailOptionTranslations) {
        validateSheetConsistency(
          detailsSheet,
          detailOptionTranslations,
          "Detail Option Translations",
          true, // Special handling for option count validation
        );
      }
    }

    console.log("Translation sheet consistency validation complete");
  } catch (error) {
    console.error("Error validating translation sheet consistency:", error);
  }
}

/**
 * Validates consistency between a source sheet and its translation sheet.
 *
 * @param sourceSheet - The source sheet (Categories or Details)
 * @param translationSheet - The translation sheet to validate
 * @param translationSheetName - Name for logging
 * @param validateOptionCounts - Whether to validate option counts (for Detail Option Translations)
 */
function validateSheetConsistency(
  sourceSheet: GoogleAppsScript.Spreadsheet.Sheet,
  translationSheet: GoogleAppsScript.Spreadsheet.Sheet,
  translationSheetName: string,
  validateOptionCounts: boolean,
): void {
  console.log(
    `Validating consistency for ${translationSheetName} against ${sourceSheet.getName()}`,
  );

  try {
    // OPTIMIZATION: Read row counts once
    const sourceRowCount = sourceSheet.getLastRow();
    const translationRowCount = translationSheet.getLastRow();

    if (translationRowCount > 0 && translationSheet.getLastColumn() > 0) {
      const fullRange = translationSheet.getRange(
        1,
        1,
        translationRowCount,
        translationSheet.getLastColumn(),
      );
      clearRangeBackgroundIfMatches(
        fullRange,
        ["#FFC7CE", "#FFF2CC", "#FF0000"], // Include bright red for primary column mismatches
      );
      clearRangeFontColorIfMatches(
        fullRange,
        ["#FFFFFF"], // White text paired with red backgrounds for primary column mismatches
      );
    }

    // Check row count consistency (excluding header)
    if (sourceRowCount !== translationRowCount) {
      console.warn(
        `Row count mismatch in ${translationSheetName}: ` +
          `Source has ${sourceRowCount} rows, translation has ${translationRowCount} rows`,
      );

      // Highlight the discrepancy in the translation sheet
      // Use cached translationRowCount instead of calling getLastRow() again
      if (translationRowCount > 0) {
        setLintNote(
          translationSheet.getRange(1, 1),
          `Row count mismatch: source has ${sourceRowCount} rows, translation has ${translationRowCount} rows. Re-sync translation sheets before generating config.`,
          "warning",
        );
      }
    }

    // CRITICAL: Validate primary language column values match between source and translation sheets
    // This prevents translation lookup failures during config generation (e.g., "Animal" vs "Animal Terrs")
    const minRowCheck = Math.min(sourceRowCount, translationRowCount) - 1; // Exclude header
    if (minRowCheck > 0) {
      // Determine source column based on translation sheet type
      let sourceColumn = 1; // Default to column A
      if (translationSheetName === "Detail Helper Text Translations") {
        sourceColumn = 2; // Column B for helper text
      } else if (translationSheetName === "Detail Option Translations") {
        sourceColumn = 4; // Column D for options
      }

      // Read source and translation primary columns
      const sourceData = sourceSheet
        .getRange(2, sourceColumn, minRowCheck, 1)
        .getValues();
      const translationData = translationSheet
        .getRange(2, 1, minRowCheck, 1)
        .getValues();

      // Collect mismatched cells
      const mismatchedCells: Array<{
        row: number;
        sourceValue: string;
        translationValue: string;
      }> = [];

      for (let i = 0; i < minRowCheck; i++) {
        const sourceValue = String(sourceData[i][0] || "").trim();
        const translationValue = String(translationData[i][0] || "").trim();

        // Normalize comparison (case-insensitive, whitespace-normalized)
        const normalizedSource = sourceValue.toLowerCase().replace(/\s+/g, " ");
        const normalizedTranslation = translationValue
          .toLowerCase()
          .replace(/\s+/g, " ");

        if (normalizedSource !== normalizedTranslation) {
          mismatchedCells.push({
            row: i + 2, // +2 for header row and 0-index
            sourceValue,
            translationValue,
          });

          console.warn(
            `Primary column mismatch in ${translationSheetName} row ${i + 2}: ` +
              `Source="${sourceValue}", Translation="${translationValue}"`,
          );
        }
      }

      // Highlight mismatched cells in bright red to indicate critical error
      if (mismatchedCells.length > 0) {
        const rangeStrings = mismatchedCells.map(({ row }) =>
          translationSheet.getRange(row, 1).getA1Notation(),
        );
        const rangeList = translationSheet.getRangeList(rangeStrings);
        rangeList.setBackground("#FF0000"); // Bright red for critical mismatch
        rangeList.setFontColor("#FFFFFF"); // White text for visibility
      }
    }

    // Validate option counts for Detail Option Translations
    if (validateOptionCounts && sourceRowCount > 1) {
      const detailsSheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Details");
      if (!detailsSheet) return;

      // OPTIMIZATION: Read all data once in a single batch operation
      const translationLastCol = translationSheet.getLastColumn();

      // Read options column from Details sheet (column 4)
      const detailsData = detailsSheet
        .getRange(2, 4, sourceRowCount - 1, 1)
        .getValues();

      // Read all translation data in one operation
      const translationData = translationSheet
        .getRange(2, 1, translationRowCount - 1, translationLastCol)
        .getValues();

      // OPTIMIZATION: Collect all cells that need highlighting instead of setting individually
      const cellsToHighlight: Array<{ row: number; col: number }> = [];

      // Validate each row in a single pass
      const minRows = Math.min(detailsData.length, translationData.length);
      for (let i = 0; i < minRows; i++) {
        const sourceOptions = String(detailsData[i][0] || "").trim();
        if (!sourceOptions) continue; // Skip if no options in source

        // Source options use comma-only splitting to match parseOptions() in the builder.
        const sourceOptionCount = sourceOptions
          .split(",")
          .map((opt) => opt.trim())
          .filter((opt) => opt !== "").length;

        // Check each translation column (starting from column 4, after Name, ISO, Source columns)
        for (let col = 3; col < translationData[i].length; col++) {
          const translatedOptions = String(
            translationData[i][col] || "",
          ).trim();
          if (!translatedOptions) continue;

          const translatedOptionCount = translatedOptions
            .split(/[;,，、]/)
            .map((opt) => opt.trim())
            .filter((opt) => opt !== "").length;

          if (sourceOptionCount !== translatedOptionCount) {
            console.warn(
              `Option count mismatch in ${translationSheetName} at row ${i + 2}, column ${col + 1}: ` +
                `Expected ${sourceOptionCount} options, found ${translatedOptionCount}`,
            );

            // Collect cell for batch highlighting
            cellsToHighlight.push({ row: i + 2, col: col + 1 });
          }
        }
      }

      // OPTIMIZATION: Apply all highlights in a single batch operation using RangeList
      if (cellsToHighlight.length > 0) {
        const rangeStrings = cellsToHighlight.map(({ row, col }) =>
          translationSheet.getRange(row, col).getA1Notation(),
        );
        const rangeList = translationSheet.getRangeList(rangeStrings);
        rangeList.setBackground("#FFC7CE"); // Light red for mismatch
      }
    }
  } catch (error) {
    console.error(
      `Error validating ${translationSheetName} consistency:`,
      error,
    );
  }
}

/**
 * Validates HTML content for common issues that could cause "Malformed HTML content" errors
 * @param html - The HTML string to validate
 * @returns Object with isValid flag and errors array
 */
function validateHtmlContent(html: string): {
  isValid: boolean;
  errors: string[];
} {
  const errors: string[] = [];

  // Check for basic HTML structure issues
  if (!html || typeof html !== "string") {
    errors.push("HTML content is empty or not a string");
    return { isValid: false, errors };
  }

  // Check for unclosed tags (basic validation)
  const tagStack: string[] = [];
  const selfClosingTags = new Set([
    "img",
    "br",
    "hr",
    "input",
    "meta",
    "link",
    "area",
    "base",
    "col",
    "embed",
    "param",
    "source",
    "track",
    "wbr",
  ]);

  // Match opening and closing tags
  const tagRegex = /<\/?([a-zA-Z][a-zA-Z0-9]*)\b[^>]*>/g;
  let match;

  while ((match = tagRegex.exec(html)) !== null) {
    const fullTag = match[0];
    const tagName = match[1].toLowerCase();

    // Skip self-closing tags
    if (selfClosingTags.has(tagName) || fullTag.endsWith("/>")) {
      continue;
    }

    // Check if it's a closing tag
    if (fullTag.startsWith("</")) {
      if (tagStack.length === 0) {
        errors.push(
          `Closing tag </${tagName}> found without matching opening tag`,
        );
      } else {
        const lastTag = tagStack.pop();
        if (lastTag !== tagName) {
          errors.push(
            `Mismatched tags: Expected </${lastTag}>, found </${tagName}>`,
          );
        }
      }
    } else {
      // Opening tag
      tagStack.push(tagName);
    }
  }

  // Check for unclosed tags
  if (tagStack.length > 0) {
    errors.push(
      `Unclosed tags: ${tagStack.map((tag) => `<${tag}>`).join(", ")}`,
    );
  }

  // Check for common HTML errors
  if (html.includes("<script>") && !html.includes("</script>")) {
    errors.push("Unclosed <script> tag detected");
  }

  if (html.includes("<style>") && !html.includes("</style>")) {
    errors.push("Unclosed <style> tag detected");
  }

  // Check for unescaped special characters in attribute values
  const attrRegex = /(\w+)="([^"]*)"/g;
  let attrMatch;
  while ((attrMatch = attrRegex.exec(html)) !== null) {
    const attrValue = attrMatch[2];
    if (attrValue.includes("<") && !attrValue.startsWith("data:")) {
      errors.push(
        `Unescaped '<' in attribute ${attrMatch[1]}="${attrValue}" - should use &lt;`,
      );
    }
  }

  // Check for malformed attribute syntax (e.g., style="value";> instead of style="value">)
  const malformedAttrRegex = /(\w+)="([^"]*)";>/g;
  let malformedMatch;
  while ((malformedMatch = malformedAttrRegex.exec(html)) !== null) {
    errors.push(
      `Malformed attribute syntax: ${malformedMatch[0]} - semicolon should be inside quotes or removed`,
    );
  }

  // NOTE: Unclosed quote validation removed due to false positives
  // The regex pattern was incorrectly capturing attributes without their closing quotes,
  // causing all valid attributes to be flagged as errors.
  // Other validations (tag matching, script/style tags, etc.) are sufficient.

  // Check for multiple DOCTYPE declarations
  const doctypeCount = (html.match(/<!DOCTYPE/gi) || []).length;
  if (doctypeCount > 1) {
    errors.push("Multiple DOCTYPE declarations found");
  }

  return {
    isValid: errors.length === 0,
    errors,
  };
}

/**
 * Validates HTML before showing a dialog to prevent "Malformed HTML content" errors
 * @param html - The HTML string to validate
 * @param context - Context description for error messages (e.g., "Language Selection Dialog")
 * @throws Error if HTML is malformed
 */
function validateDialogHtml(html: string, context: string = "Dialog"): void {
  const validation = validateHtmlContent(html);

  if (!validation.isValid) {
    const errorMessage = `HTML validation failed for ${context}:\n${validation.errors.join("\n")}`;
    console.error(errorMessage);
    throw new Error(
      `Malformed HTML detected in ${context}. Please check the console for details.`,
    );
  }

  // Additional checks specific to Google Apps Script dialogs
  if (html.length > 500000) {
    console.warn(
      `HTML content for ${context} is very large (${html.length} characters). This may cause performance issues.`,
    );
  }
}

/**
 * Test function to validate HTML dialog generation
 * Run this from the Apps Script editor to test HTML validation
 */
function testHtmlValidation(): void {
  console.log("=== Testing HTML Validation ===");

  // Test cases
  const testCases = [
    {
      name: "Valid HTML",
      html: "<p>Hello <strong>world</strong></p>",
      shouldPass: true,
    },
    {
      name: "Unclosed tag",
      html: "<p>Hello <strong>world</p>",
      shouldPass: false,
    },
    {
      name: "Mismatched tags",
      html: "<div><p>Content</div></p>",
      shouldPass: false,
    },
    {
      name: "Unclosed script tag",
      html: "<script>console.log('test')",
      shouldPass: false,
    },
    {
      name: "Valid self-closing tags",
      html: "<p>Line 1<br/>Line 2<img src='test.png' /></p>",
      shouldPass: true,
    },
    {
      name: "Unescaped < in attribute",
      html: '<p data-value="test<value">Content</p>',
      shouldPass: false,
    },
    {
      name: "Valid data URI",
      html: '<img src="data:image/svg+xml,%3Csvg%3E" />',
      shouldPass: true,
    },
    {
      name: "Multiple DOCTYPE declarations",
      html: "<!DOCTYPE html><!DOCTYPE html><html></html>",
      shouldPass: false,
    },
    {
      name: "Malformed attribute with semicolon",
      html: '<ol style="text-align: left";><li>Item</li></ol>',
      shouldPass: false,
    },
    // NOTE: Unclosed quote validation removed due to false positives
    // The test case for unclosed quotes has been removed
    {
      name: "Valid complex HTML",
      html: "<!DOCTYPE html><html><head><style>body { color: red; }</style></head><body><p>Test</p></body></html>",
      shouldPass: true,
    },
  ];

  let passed = 0;
  let failed = 0;

  testCases.forEach((testCase) => {
    console.log(`\nTesting: ${testCase.name}`);
    const validation = validateHtmlContent(testCase.html);

    if (validation.isValid === testCase.shouldPass) {
      console.log(`✅ PASS: ${testCase.name}`);
      passed++;
    } else {
      console.log(`❌ FAIL: ${testCase.name}`);
      console.log(`  Expected: ${testCase.shouldPass ? "valid" : "invalid"}`);
      console.log(`  Got: ${validation.isValid ? "valid" : "invalid"}`);
      if (validation.errors.length > 0) {
        console.log(`  Errors: ${validation.errors.join(", ")}`);
      }
      failed++;
    }
  });

  console.log("\n=== Test Results ===");
  console.log(`Passed: ${passed}/${testCases.length}`);
  console.log(`Failed: ${failed}/${testCases.length}`);

  if (failed === 0) {
    console.log("✅ All tests passed!");
  } else {
    console.log(`❌ ${failed} test(s) failed`);
  }
}

/**
 * INTEGRATION TEST: Test actual dialog HTML generation
 * This validates the real HTML that would be shown to users
 */
function testDialogHtmlGeneration(): void {
  console.log("\n=== Testing Dialog HTML Generation ===");

  let totalTests = 0;
  let passedTests = 0;

  // Test 1: Simple dialog
  try {
    totalTests++;
    const html = generateDialog("Test Title", "<p>Test message</p>");
    validateDialogHtml(html, "Test Dialog");
    console.log("✅ Simple dialog HTML is valid");
    passedTests++;
  } catch (error) {
    console.log("❌ Simple dialog HTML is INVALID:", error);
  }

  // Test 2: Dialog with button
  try {
    totalTests++;
    const html = generateDialog(
      "Test",
      "<p>Message</p>",
      "Click",
      "https://example.com",
    );
    validateDialogHtml(html, "Dialog with Button");
    console.log("✅ Dialog with button HTML is valid");
    passedTests++;
  } catch (error) {
    console.log("❌ Dialog with button HTML is INVALID:", error);
  }

  // Test 3: Dialog with function button
  try {
    totalTests++;
    const html = generateDialog(
      "Test",
      "<p>Message</p>",
      "Submit",
      null,
      "submitForm",
    );
    validateDialogHtml(html, "Dialog with Function");
    console.log("✅ Dialog with function button HTML is valid");
    passedTests++;
  } catch (error) {
    console.log("❌ Dialog with function button HTML is INVALID:", error);
  }

  // Test 4: Dialog with special characters (should be escaped)
  try {
    totalTests++;
    const title = 'Test <> & "Title"';
    const message = "<p>" + escapeHtml('Message with <> & "quotes"') + "</p>";
    const html = generateDialog(title, message);
    validateDialogHtml(html, "Dialog with Special Chars");
    console.log("✅ Dialog with special characters HTML is valid");
    passedTests++;
  } catch (error) {
    console.log("❌ Dialog with special characters HTML is INVALID:", error);
  }

  console.log(`\n=== Dialog Generation Test Results ===`);
  console.log(`Passed: ${passedTests}/${totalTests}`);

  if (passedTests === totalTests) {
    console.log("✅ All dialog generation tests passed!");
  } else {
    console.log(
      `❌ ${totalTests - passedTests} dialog generation test(s) failed`,
    );
  }
}

/**
 * Run all HTML validation tests
 */
function runAllHtmlValidationTests(): void {
  testHtmlValidation();
  testDialogHtmlGeneration();
  console.log("\n=== All HTML Validation Tests Complete ===");
}

/**
 * Detects translation sheet mismatches without modifying anything.
 * Returns details about mismatches found.
 *
 * @returns Object with mismatch details or null if no mismatches found
 */
function detectTranslationMismatches(): {
  hasMismatches: boolean;
  details: Array<{
    sheetName: string;
    sourceSheet: string;
    sourceColumn: string;
    mismatches: Array<{
      row: number;
      sourceValue: string;
      translationValue: string;
    }>;
  }>;
} | null {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const allDetails: Array<{
    sheetName: string;
    sourceSheet: string;
    sourceColumn: string;
    mismatches: Array<{
      row: number;
      sourceValue: string;
      translationValue: string;
    }>;
  }> = [];

  // Check Category Translations
  const categoriesSheet = spreadsheet.getSheetByName("Categories");
  const categoryTranslationsSheet = spreadsheet.getSheetByName(
    "Category Translations",
  );

  if (categoriesSheet && categoryTranslationsSheet) {
    const catMismatches = checkSheetMismatches(
      categoriesSheet,
      categoryTranslationsSheet,
      1, // Column A
    );
    if (catMismatches.length > 0) {
      allDetails.push({
        sheetName: "Category Translations",
        sourceSheet: "Categories",
        sourceColumn: "A",
        mismatches: catMismatches,
      });
    }
  }

  // Check Detail translations
  const detailsSheet = spreadsheet.getSheetByName("Details");
  if (detailsSheet) {
    const detailLabelTranslations = spreadsheet.getSheetByName(
      "Detail Label Translations",
    );
    const detailHelperTranslations = spreadsheet.getSheetByName(
      "Detail Helper Text Translations",
    );
    const detailOptionTranslations = spreadsheet.getSheetByName(
      "Detail Option Translations",
    );

    if (detailLabelTranslations) {
      const labelMismatches = checkSheetMismatches(
        detailsSheet,
        detailLabelTranslations,
        1,
      );
      if (labelMismatches.length > 0) {
        allDetails.push({
          sheetName: "Detail Label Translations",
          sourceSheet: "Details",
          sourceColumn: "A",
          mismatches: labelMismatches,
        });
      }
    }

    if (detailHelperTranslations) {
      const helperMismatches = checkSheetMismatches(
        detailsSheet,
        detailHelperTranslations,
        2,
      );
      if (helperMismatches.length > 0) {
        allDetails.push({
          sheetName: "Detail Helper Text Translations",
          sourceSheet: "Details",
          sourceColumn: "B",
          mismatches: helperMismatches,
        });
      }
    }

    if (detailOptionTranslations) {
      const optionMismatches = checkSheetMismatches(
        detailsSheet,
        detailOptionTranslations,
        4,
      );
      if (optionMismatches.length > 0) {
        allDetails.push({
          sheetName: "Detail Option Translations",
          sourceSheet: "Details",
          sourceColumn: "D",
          mismatches: optionMismatches,
        });
      }
    }
  }

  if (allDetails.length === 0) {
    return null;
  }

  return {
    hasMismatches: true,
    details: allDetails,
  };
}

/**
 * Helper function to check mismatches between source and translation sheet
 */
function checkSheetMismatches(
  sourceSheet: GoogleAppsScript.Spreadsheet.Sheet,
  translationSheet: GoogleAppsScript.Spreadsheet.Sheet,
  sourceColumn: number,
): Array<{ row: number; sourceValue: string; translationValue: string }> {
  const sourceRowCount = sourceSheet.getLastRow();
  const translationRowCount = translationSheet.getLastRow();
  const minRowCheck = Math.min(sourceRowCount, translationRowCount) - 1;

  if (minRowCheck <= 0) {
    return [];
  }

  const sourceData = sourceSheet
    .getRange(2, sourceColumn, minRowCheck, 1)
    .getValues();
  const translationData = translationSheet
    .getRange(2, 1, minRowCheck, 1)
    .getValues();

  const mismatches: Array<{
    row: number;
    sourceValue: string;
    translationValue: string;
  }> = [];

  for (let i = 0; i < minRowCheck; i++) {
    const sourceValue = String(sourceData[i][0] || "").trim();
    const translationValue = String(translationData[i][0] || "").trim();

    const normalizedSource = sourceValue.toLowerCase().replace(/\s+/g, " ");
    const normalizedTranslation = translationValue
      .toLowerCase()
      .replace(/\s+/g, " ");

    if (normalizedSource !== normalizedTranslation) {
      mismatches.push({
        row: i + 2,
        sourceValue,
        translationValue,
      });
    }
  }

  return mismatches;
}

/**
 * Fixes translation sheet mismatches by re-syncing formulas from source sheets
 * and optionally re-running translation for configured languages.
 *
 * @param reTranslate - Whether to re-run translation after fixing formulas
 * @param mismatchData - Optional pre-detected mismatch data (avoids re-detection after formula sync)
 */
function fixTranslationMismatches(
  reTranslate: boolean = true,
  mismatchData?: {
    hasMismatches: boolean;
    details: Array<{
      sheetName: string;
      sourceSheet: string;
      sourceColumn: string;
      mismatches: Array<{
        row: number;
        sourceValue: string;
        translationValue: string;
      }>;
    }>;
  } | null,
): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Re-sync Category Translations formulas
  const categoriesSheet = spreadsheet.getSheetByName("Categories");
  const categoryTranslationsSheet = spreadsheet.getSheetByName(
    "Category Translations",
  );

  if (categoriesSheet && categoryTranslationsSheet) {
    const lastRow = categoriesSheet.getLastRow();
    if (lastRow > 1) {
      const formula = `=Categories!A2:A${lastRow}`;
      categoryTranslationsSheet
        .getRange(2, 1, lastRow - 1, 1)
        .setFormula(formula);
      console.log(
        `Re-synced Category Translations formulas (rows 2-${lastRow})`,
      );
    }
  }

  // Re-sync Detail translations formulas
  const detailsSheet = spreadsheet.getSheetByName("Details");
  if (detailsSheet) {
    const lastRow = detailsSheet.getLastRow();

    if (lastRow > 1) {
      // Detail Label Translations (column A)
      const detailLabelSheet = spreadsheet.getSheetByName(
        "Detail Label Translations",
      );
      if (detailLabelSheet) {
        const formula = `=Details!A2:A${lastRow}`;
        detailLabelSheet.getRange(2, 1, lastRow - 1, 1).setFormula(formula);
        console.log(
          `Re-synced Detail Label Translations formulas (rows 2-${lastRow})`,
        );
      }

      // Detail Helper Text Translations (column B)
      const detailHelperSheet = spreadsheet.getSheetByName(
        "Detail Helper Text Translations",
      );
      if (detailHelperSheet) {
        const formula = `=Details!B2:B${lastRow}`;
        detailHelperSheet.getRange(2, 1, lastRow - 1, 1).setFormula(formula);
        console.log(
          `Re-synced Detail Helper Text Translations formulas (rows 2-${lastRow})`,
        );
      }

      // Detail Option Translations (column D)
      const detailOptionSheet = spreadsheet.getSheetByName(
        "Detail Option Translations",
      );
      if (detailOptionSheet) {
        const formula = `=Details!D2:D${lastRow}`;
        detailOptionSheet.getRange(2, 1, lastRow - 1, 1).setFormula(formula);
        console.log(
          `Re-synced Detail Option Translations formulas (rows 2-${lastRow})`,
        );
      }
    }
  }

  // Re-translate if requested
  if (reTranslate) {
    console.log("Re-running translation for configured languages...");

    // Get configured target languages from Category Translations headers
    if (categoryTranslationsSheet) {
      const lastColumn = categoryTranslationsSheet.getLastColumn();
      if (lastColumn > 1) {
        const headers = categoryTranslationsSheet
          .getRange(1, 1, 1, lastColumn)
          .getValues()[0];
        const targetLanguages: TranslationLanguage[] = [];
        const allLanguages = getAllLanguages();
        const resolveHeaderCode = createTranslationHeaderResolver(allLanguages);
        const seenLanguages = new Set<string>();
        const headerB = String(headers[1] || "")
          .trim()
          .toLowerCase();
        const headerC = String(headers[2] || "")
          .trim()
          .toLowerCase();
        const hasMetaColumns =
          headerB.includes("iso") && headerC.includes("source");
        const languageStartIndex = hasMetaColumns ? 3 : 1;

        // Extract language codes from headers (skip first column which is primary language)
        for (let i = languageStartIndex; i < headers.length; i++) {
          const header = String(headers[i] || "").trim();
          if (!header) continue;

          const langCode = resolveHeaderCode(header);
          if (!langCode) continue;

          const normalizedCode = langCode.toLowerCase();
          if (seenLanguages.has(normalizedCode)) continue;
          seenLanguages.add(normalizedCode);
          targetLanguages.push(langCode as TranslationLanguage);
        }

        if (targetLanguages.length > 0) {
          console.log(
            `Re-translating to languages: ${targetLanguages.join(", ")}`,
          );

          // OPTIMIZATION: Only clear and re-translate rows that have mismatches
          // Use pre-detected mismatch data if provided, otherwise detect now
          const mismatchResult = mismatchData || detectTranslationMismatches();

          if (mismatchResult && mismatchResult.hasMismatches) {
            console.log(
              `Found ${mismatchResult.details.length} translation sheet(s) with mismatches`,
            );

            // Clear only the mismatched rows in each sheet
            mismatchResult.details.forEach((detail) => {
              const sheet = spreadsheet.getSheetByName(detail.sheetName);
              if (!sheet || sheet.getLastColumn() <= 1) return;

              const colCount = sheet.getLastColumn() - 1; // Exclude primary language column
              const mismatchedRows = detail.mismatches.map((m) => m.row);

              // Clear translation columns for only the mismatched rows
              mismatchedRows.forEach((row) => {
                sheet.getRange(row, 2, 1, colCount).clearContent();
              });

              console.log(
                `Cleared translations for ${mismatchedRows.length} mismatched row(s) in ${detail.sheetName}: ` +
                  `rows ${mismatchedRows.join(", ")}`,
              );
            });
          } else {
            console.log("No mismatches detected, skipping clearing step");
          }

          autoTranslateSheetsBidirectional(targetLanguages);
        } else {
          console.log(
            "No target languages found in headers, skipping re-translation",
          );
        }
      }
    }
  }

  // Clear any bright red highlighting and white font from the primary column after fixing
  // This removes the visual indicators since the issues have been resolved
  const translationSheetNames = [
    "Category Translations",
    "Detail Label Translations",
    "Detail Helper Text Translations",
    "Detail Option Translations",
  ];

  translationSheetNames.forEach((sheetName) => {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet && sheet.getLastRow() > 1) {
      const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1);
      clearRangeBackgroundIfMatches(range, ["#FF0000"]);
      clearRangeFontColorIfMatches(range, ["#FFFFFF"]);
    }
  });

  console.log("Translation sheet fix complete");
}

/**
 * Validates explicit Metadata sheet values for unsafe characters.
 * Only checks rows that are present – does NOT flag missing rows.
 * Keys validated: name, version, primaryLanguage.
 */
function lintMetadataSheet(): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const metadataSheet = spreadsheet.getSheetByName("Metadata");
  if (!metadataSheet) return;

  const lastRow = metadataSheet.getLastRow();
  if (lastRow <= 1) return; // Header-only or empty

  // Clear previous lint artifacts on value column (B)
  clearLintArtifacts(metadataSheet.getRange(2, 2, lastRow - 1, 1));

  // Matches containsUnsafeNameCharacters() in importService: slashes, backslashes, ellipsis
  // These are the only characters that strict build validation actually rejects (errors).
  const STRICT_UNSAFE_PATTERN = /[\\/]|\.\.\./;

  const data = metadataSheet.getRange(2, 1, lastRow - 1, 2).getValues();

  // Track seen keys to detect duplicates. The builder uses exact key
  // matching (sheetData[i][0] === key), so lint must do the same to
  // avoid false-positive duplicate warnings on differently-cased keys
  // (e.g. "PrimaryLanguage" vs "primaryLanguage" are distinct to the builder).
  const seenKeys = new Set<string>();

  for (let i = 0; i < data.length; i++) {
    const key = String(data[i][0] || "").trim();
    const value = String(data[i][1] || "").trim();
    if (!value) continue; // Skip empty values

    const row = i + 2;

    // Flag duplicate keys as a warning — the builder uses only the first row.
    // Use exact key matching to match builder semantics.
    const isDuplicate = key && seenKeys.has(key);
    if (isDuplicate) {
      const cell = metadataSheet.getRange(row, 2);
      appendLintNote(
        cell,
        `Duplicate metadata key "${key}". The builder only reads the first occurrence — this row is ignored.`,
        "warning",
      );
    }
    if (key) seenKeys.add(key);

    // Skip further validation for duplicate rows — the builder ignores them,
    // so unsafe-char and language errors would be false positives.
    if (isDuplicate) continue;

    // Only validate "name" for unsafe characters.
    // The export pipeline overwrites "version" with the current date
    // (yy.MM.dd) before strict validation runs, so the sheet value is
    // never used as-is and linting it would produce false errors.
    if (key === "name") {
      if (STRICT_UNSAFE_PATTERN.test(value)) {
        const cell = metadataSheet.getRange(row, 2);
        appendLintNote(
          cell,
          `Metadata "${key}" contains characters that will fail config generation: slashes, backslashes, and ellipses (…) are not allowed.`,
          "error",
        );
      }
    }

    if (key === "primaryLanguage") {
      // Validate language using the same normalization path as the builder.
      // The builder's normalizeLocaleInput() first checks for ISO codes
      // (e.g. "en", "pt-BR") via regex, then falls back to validateLanguageName()
      // for display names (e.g. "English", "Português"). Lint must accept both.
      const ISO_LOCALE_PATTERN = /^[a-z]{2,3}(-[a-z]{2,3})?$/i;
      const isValid = ISO_LOCALE_PATTERN.test(value) ? true : (() => {
        const validation = validateLanguageName(value);
        return validation.valid;
      })();
      if (!isValid) {
        const cell = metadataSheet.getRange(row, 2);
        appendLintNote(
          cell,
          `Metadata primaryLanguage: "${value}" is not a recognized language. Use an ISO code (e.g. "en", "pt-BR") or a language name (e.g. "English", "Português").`,
          "error",
        );
      }
    }
  }
}

/**
 * Decodes a data:image/svg+xml URI (plain or base64) to extract the raw SVG
 * content. Returns null if the URI cannot be decoded. Mirrors the decode logic
 * from payloadBuilder's decodeDataSvg() for use in lint checks.
 */
function decodeDataSvgForLint(dataUri: string): string | null {
  try {
    if (dataUri.includes(";base64,")) {
      const base64 = dataUri.split(";base64,")[1];
      const decoded = Utilities.newBlob(
        Utilities.base64Decode(base64),
      ).getDataAsString();
      return decoded;
    }

    const commaIndex = dataUri.indexOf(",");
    if (commaIndex === -1) return null;
    const payload = dataUri.substring(commaIndex + 1);
    return decodeURIComponent(payload);
  } catch (_err) {
    return null;
  }
}

/**
 * Phase 6 Task 1: Checks inline SVG sizes in Categories column B and Icons column B.
 * - Warning if SVG > 300 KB (307200 bytes)
 * - Error if SVG > 2 MB (2097152 bytes)
 */
function checkInlineSvgSizes(): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const logger = getScopedLogger("LintInlineSvgSizes");

  const SVG_WARN_BYTES = 307200; // 300 KB
  const SVG_ERROR_BYTES = 2097152; // 2 MB

  const sheetsToCheck = [
    { name: "Categories", col: 2 },
    { name: "Icons", col: 2 },
  ];

  for (const { name, col } of sheetsToCheck) {
    const sheet = spreadsheet.getSheetByName(name);
    if (!sheet) continue;

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) continue;

    const values = sheet.getRange(2, col, lastRow - 1, 1).getValues();

    for (let i = 0; i < values.length; i++) {
      const value = String(values[i][0] || "").trim();
      // Normalize icon source to extract inline SVG content for size checking.
      // parseIconSource() handles inline <svg>, data:image/svg+xml URIs, and
      // Drive URLs. Only measure sources that resolve to inline svgData.
      let svgContent: string | null = null;
      if (value.startsWith("<svg")) {
        svgContent = value;
      } else if (value.toLowerCase().startsWith("data:image/svg+xml")) {
        // Decode data URI to get the actual SVG content for size measurement
        svgContent = decodeDataSvgForLint(value);
      }
      // Drive URLs and remote URLs are not measured here — they require
      // network access and the builder validates them during generation.
      if (!svgContent) continue;

      const sizeBytes = Utilities.newBlob(svgContent).getBytes().length;
      const sizeKB = Math.round(sizeBytes / 1024);

      if (sizeBytes > SVG_ERROR_BYTES) {
        const row = i + 2;
        appendLintNote(
          sheet.getRange(row, col),
          `Inline SVG is ${sizeKB}KB (limit: 300KB warning, 2MB error)`,
          "error",
        );
        logger.warn(
          `${name} row ${row}: inline SVG is ${sizeKB}KB (exceeds 2MB error limit)`,
        );
      } else if (sizeBytes > SVG_WARN_BYTES) {
        const row = i + 2;
        appendLintNote(
          sheet.getRange(row, col),
          `Inline SVG is ${sizeKB}KB (limit: 300KB warning, 2MB error)`,
          "warning",
        );
        logger.warn(
          `${name} row ${row}: inline SVG is ${sizeKB}KB (exceeds 300KB warning limit)`,
        );
      }
    }
  }
}

function createLintHeaderMap(headers: any[]): Record<string, number> {
  const headerMap: Record<string, number> = {};
  headers.forEach((header, index) => {
    const key = String(header || "")
      .trim()
      .toLowerCase();
    if (key) {
      headerMap[key] = index;
    }
  });
  return headerMap;
}

function getLintColumnIndex(
  headerMap: Record<string, number>,
  ...names: string[]
): number | undefined {
  for (const name of names) {
    const key = name.toLowerCase();
    if (headerMap[key] !== undefined) {
      return headerMap[key];
    }
  }
  return undefined;
}

function buildLintCategorySummaries(
  data: SheetData,
): Array<{ id: string; name: string }> {
  const rows = data.Categories?.slice(1) || [];
  const headers = data.Categories?.[0] || [];
  const headerMap = createLintHeaderMap(headers);
  const nameCol = getLintColumnIndex(headerMap, "name") ?? CATEGORY_COL.NAME;
  const idCol = getLintColumnIndex(headerMap, "category id", "id");

  return rows
    .map((row, index) => {
      const name = String(row[nameCol] || "").trim();
      if (!name) {
        return null;
      }

      const explicitId =
        idCol !== undefined ? String(row[idCol] || "").trim() : "";
      return {
        id: explicitId || slugify(name) || `category-${index + 1}`,
        name,
      };
    })
    .filter(
      (
        category,
      ): category is {
        id: string;
        name: string;
      } => Boolean(category),
    );
}

function buildLintFieldSummaries(
  data: SheetData,
): Array<{
  id: string;
  name: string;
  helperText?: string;
  options?: SelectOption[];
}> {
  const rows = data.Details?.slice(1) || [];
  const headers = data.Details?.[0] || [];
  const headerMap = createLintHeaderMap(headers);
  const nameCol = getLintColumnIndex(headerMap, "name", "label") ?? DETAILS_COL.NAME;
  const helperCol =
    getLintColumnIndex(headerMap, "helper text", "helper", "help") ??
    DETAILS_COL.HELPER_TEXT;
  const typeCol = getLintColumnIndex(headerMap, "type") ?? DETAILS_COL.TYPE;
  const optionsCol = getLintColumnIndex(headerMap, "options") ?? DETAILS_COL.OPTIONS;
  const idCol = getLintColumnIndex(headerMap, "id") ?? DETAILS_COL.ID;
  type LintFieldSummary = {
    id: string;
    name: string;
    helperText?: string;
    options?: SelectOption[];
  };

  return rows
    .map((row, index) => {
      const name = String(row[nameCol] || "").trim();
      if (!name) {
        return null;
      }

      const helperText = String(row[helperCol] || "");
      // Mirror payloadBuilder: empty type cells default to "text" (not selectOne)
      const typeRaw = String(row[typeCol] || "text")
        .trim()
        .toLowerCase();
      const optionsStr = String(row[optionsCol] || "");
      const explicitId = String(row[idCol] || "").trim();

      let options: SelectOption[] | undefined;
      const typeKey = typeRaw.charAt(0);
      switch (typeKey) {
        case "m":
          options = parseOptions(optionsStr);
          break;
        case "n":
        case "t":
          options = undefined;
          break;
        case "s":
        default:
          options = parseOptions(optionsStr);
          break;
      }

      const isSelectType =
        typeKey === "m" || typeKey === "s";
      if (isSelectType && (!options || options.length === 0)) {
        return null;
      }

      if (options && options.length > 0) {
        const seenValues = new Set<string>();
        options = options.filter((option) => {
          if (seenValues.has(option.value)) {
            return false;
          }
          seenValues.add(option.value);
          return true;
        });
      }

      const fieldSummary: LintFieldSummary = {
        id: explicitId || slugify(name) || `field-${index + 1}`,
        name,
      };

      if (helperText) {
        fieldSummary.helperText = helperText;
      }

      if (options && options.length > 0) {
        fieldSummary.options = options;
      }

      return fieldSummary;
    })
    .filter(
      (field): field is LintFieldSummary => field !== null,
    );
}

function buildLintTranslationsByLocale(
  data: SheetData,
  categories: Array<{ id: string; name: string }>,
  fields: Array<{
    id: string;
    name: string;
    helperText?: string;
    options?: SelectOption[];
  }>,
): {
  translationsByLocale: Record<string, any>;
  localeHeaderRefs: Record<
    string,
    Array<{ sheetName: string; column: number; header: string }>
  >;
} {
  const translationsByLocale: Record<string, any> = {};
  const localeHeaderRefs: Record<
    string,
    Array<{ sheetName: string; column: number; header: string }>
  > = {};

  const ensureLocaleEntry = (locale: string): Record<string, any> => {
    if (!translationsByLocale[locale]) {
      translationsByLocale[locale] = {};
    }
    return translationsByLocale[locale];
  };

  const ensureCategoryEntry = (
    locale: string,
    categoryId: string,
  ): Record<string, string> => {
    const localeEntry = ensureLocaleEntry(locale);
    localeEntry.category = localeEntry.category || {};
    localeEntry.category[categoryId] = localeEntry.category[categoryId] || {};
    return localeEntry.category[categoryId];
  };

  const ensureFieldEntry = (
    locale: string,
    fieldId: string,
  ): Record<string, string> => {
    const localeEntry = ensureLocaleEntry(locale);
    localeEntry.field = localeEntry.field || {};
    localeEntry.field[fieldId] = localeEntry.field[fieldId] || {};
    return localeEntry.field[fieldId];
  };

  const categoryIdByName = new Map<string, string>();
  categories.forEach((category) => {
    categoryIdByName.set(category.name, category.id);
  });

  const fieldIdByName = new Map<string, string>();
  fields.forEach((field) => {
    fieldIdByName.set(field.name, field.id);
  });

  const fieldsByHelperText = new Map<string, string[]>();
  fields.forEach((field) => {
    if (!field.helperText) {
      return;
    }
    const fieldIds = fieldsByHelperText.get(field.helperText) || [];
    fieldIds.push(field.id);
    fieldsByHelperText.set(field.helperText, fieldIds);
  });

  const fieldsByOptionsString = new Map<
    string,
    Array<{ id: string; optionCount: number }>
  >();
  fields.forEach((field) => {
    if (!field.options || field.options.length === 0) {
      return;
    }
    const optionsString = field.options
      .map((option) => {
        const value = option?.value || "";
        const label = option?.label || "";
        if (!label) {
          return "";
        }
        return value === slugify(label) ? label : `${value}:${label}`;
      })
      .filter(Boolean)
      .join(", ");
    if (!optionsString) {
      return;
    }
    const matches = fieldsByOptionsString.get(optionsString) || [];
    matches.push({ id: field.id, optionCount: field.options.length });
    fieldsByOptionsString.set(optionsString, matches);
  });

  const processSheet = (
    sheetName: string,
    handler: (row: any[], locale: string, value: string) => void,
  ) => {
    const sheetData = data[sheetName];
    if (!sheetData || sheetData.length <= 1) {
      return;
    }

    const headers = sheetData[0];
    const langColumns = extractLanguagesFromHeaders(headers);
    langColumns.forEach((entry) => {
      const refs = localeHeaderRefs[entry.lang] || [];
      refs.push({
        sheetName,
        column: entry.index + 1,
        header: String(headers[entry.index] || "").trim(),
      });
      localeHeaderRefs[entry.lang] = refs;
    });

    sheetData.slice(1).forEach((row) => {
      langColumns.forEach((entry) => {
        const value = String(row[entry.index] || "").trim();
        if (!value) {
          return;
        }
        handler(row, entry.lang, value);
      });
    });
  };

  processSheet("Category Translations", (row, locale, value) => {
    const sourceName = String(row[TRANSLATION_COL.SOURCE_TEXT] || "").trim();
    const categoryId = categoryIdByName.get(sourceName);
    if (!categoryId) {
      return;
    }
    ensureCategoryEntry(locale, categoryId).name = value;
  });

  processSheet("Detail Label Translations", (row, locale, value) => {
    const sourceName = String(row[TRANSLATION_COL.SOURCE_TEXT] || "").trim();
    const fieldId = fieldIdByName.get(sourceName);
    if (!fieldId) {
      return;
    }
    ensureFieldEntry(locale, fieldId).label = value;
  });

  processSheet("Detail Helper Text Translations", (row, locale, value) => {
    const sourceHelperText = String(
      row[TRANSLATION_COL.SOURCE_TEXT] || "",
    ).trim();
    const matchingFieldIds = fieldsByHelperText.get(sourceHelperText) || [];
    matchingFieldIds.forEach((fieldId) => {
      ensureFieldEntry(locale, fieldId).helperText = value;
    });
  });

  processSheet("Detail Option Translations", (row, locale, value) => {
    const sourceOptions = String(row[TRANSLATION_COL.SOURCE_TEXT] || "").trim();
    const matchingFields = fieldsByOptionsString.get(sourceOptions) || [];
    const translatedOptions = splitTranslatedOptions(value);
    matchingFields.forEach((field) => {
      for (
        let optionIndex = 0;
        optionIndex < translatedOptions.length &&
        optionIndex < field.optionCount;
        optionIndex++
      ) {
        ensureFieldEntry(locale, field.id)[`options.${optionIndex}`] =
          translatedOptions[optionIndex];
      }
    });
  });

  Object.keys(translationsByLocale).forEach((locale) => {
    const entry = translationsByLocale[locale];
    const hasCategories = Boolean(
      entry.category && Object.keys(entry.category).length > 0,
    );
    const hasFields = Boolean(entry.field && Object.keys(entry.field).length > 0);
    if (!hasCategories && !hasFields) {
      delete translationsByLocale[locale];
    }
  });

  return { translationsByLocale, localeHeaderRefs };
}

function collectStrictLintMetrics(): {
  categoryCount: number;
  fieldCount: number;
  iconCount: number;
  optionCount: number;
  translationEntryCount: number;
  translationsByLocale: Record<string, any>;
  localeHeaderRefs: Record<
    string,
    Array<{ sheetName: string; column: number; header: string }>
  >;
} {
  const data = getSpreadsheetData();
  const categories = buildLintCategorySummaries(data);
  const fields = buildLintFieldSummaries(data);
  const { translationsByLocale, localeHeaderRefs } = buildLintTranslationsByLocale(
    data,
    categories,
    fields,
  );

  let iconCount = 0;
  try {
    iconCount = buildIconsFromSheet(data).length;
  } catch (error) {
    getScopedLogger("LintStrictMetrics").warn(
      "Falling back to Icons sheet row count while calculating entity totals:",
      error,
    );
    const iconsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Icons");
    iconCount = iconsSheet ? Math.max(0, iconsSheet.getLastRow() - 1) : 0;
  }

  const optionCount = fields.reduce(
    (total, field) => total + (field.options?.length || 0),
    0,
  );
  const translationEntryCount =
    typeof countTranslationEntries === "function"
      ? countTranslationEntries(translationsByLocale)
      : 0;

  return {
    categoryCount: categories.length,
    fieldCount: fields.length,
    iconCount,
    optionCount,
    translationEntryCount,
    translationsByLocale,
    localeHeaderRefs,
  };
}

/**
 * Phase 6 Task 2: Checks per-locale translation payload sizes.
 * Uses the merged per-locale translation object so lint matches strict validation.
 */
function checkTranslationPayloadSizes(metrics: {
  translationsByLocale: Record<string, any>;
  localeHeaderRefs: Record<
    string,
    Array<{ sheetName: string; column: number; header: string }>
  >;
}): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const logger = getScopedLogger("LintTranslationPayloadSizes");

  const PAYLOAD_LIMIT_BYTES = 1048576; // 1 MB

  Object.keys(metrics.translationsByLocale).forEach((locale) => {
    const localePayload = metrics.translationsByLocale[locale];
    const localeBytes = getByteLength(JSON.stringify(localePayload || {}));
    if (localeBytes <= PAYLOAD_LIMIT_BYTES) {
      return;
    }

    const sizeKB = Math.round(localeBytes / 1024);
    const refs = metrics.localeHeaderRefs[locale] || [];
    refs.forEach((ref) => {
      const sheet = spreadsheet.getSheetByName(ref.sheetName);
      if (!sheet) {
        return;
      }
      appendLintNote(
        sheet.getRange(1, ref.column),
        `Combined translations for locale '${locale}' exceed 1MB limit (${sizeKB}KB) after merging all translation sheets. This will fail strict validation.`,
        "error",
      );
    });
    logger.warn(
      `Merged translation payload for locale ${locale} is ${sizeKB}KB across ${refs.length} sheet column(s)`,
    );
  });
}

/**
 * Phase 6 Task 3: Checks strict total entity count parity.
 * Includes categories, fields, icons, option entries, and translation entries.
 */
function checkTotalEntityCounts(metrics: {
  categoryCount: number;
  fieldCount: number;
  iconCount: number;
  optionCount: number;
  translationEntryCount: number;
}): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const logger = getScopedLogger("LintTotalEntityCounts");

  const ENTITY_LIMIT = 10000;
  const totalCount =
    metrics.categoryCount +
    metrics.fieldCount +
    metrics.iconCount +
    metrics.optionCount +
    metrics.translationEntryCount;

  if (totalCount > ENTITY_LIMIT) {
    const categoriesSheet = spreadsheet.getSheetByName("Categories");
    const cell = categoriesSheet
      ? categoriesSheet.getRange(1, 1)
      : spreadsheet.getActiveSheet().getRange(1, 1);
    appendLintNote(
      cell,
      `Total entity count (${totalCount}) exceeds 10,000 limit. Categories: ${metrics.categoryCount}, Details: ${metrics.fieldCount}, Icons: ${metrics.iconCount}, Options: ${metrics.optionCount}, Translations: ${metrics.translationEntryCount}. This will fail strict validation.`,
      "error",
    );
    logger.warn(
      `Total entity count ${totalCount} exceeds limit (Categories: ${metrics.categoryCount}, Details: ${metrics.fieldCount}, Icons: ${metrics.iconCount}, Options: ${metrics.optionCount}, Translations: ${metrics.translationEntryCount})`,
    );
  }
}

/**
 * Main linting function that validates all sheets in the spreadsheet.
 *
 * @param showAlerts - Whether to show UI alerts (default: true). Set to false when called from other functions.
 */
function lintAllSheets(showAlerts: boolean = true): void {
  try {
    console.log("Starting linting process...");

    console.log("Linting Categories sheet...");
    lintCategoriesSheet();

    console.log("Linting Details sheet...");
    lintDetailsSheet();

    console.log("Linting Icons sheet...");
    lintIconsSheet();

    // Phase 5: Cross-sheet icon checks
    checkCrossSheetIconCollisions();

    // Phase 6 Task 1: Inline SVG size warnings/errors
    checkInlineSvgSizes();

    console.log("Linting Translation sheets...");
    lintTranslationSheets();

    const strictLintMetrics = collectStrictLintMetrics();

    // Phase 6 Task 2: Per-locale translation payload size checks
    checkTranslationPayloadSizes(strictLintMetrics);

    console.log("Linting Metadata sheet...");
    lintMetadataSheet();

    // Phase 6 Task 3: Total entity count check
    checkTotalEntityCounts(strictLintMetrics);

    console.log("Finished linting all sheets.");

    // Add a summary of issues found, but only if showAlerts is true
    if (showAlerts) {
      const ui = SpreadsheetApp.getUi();
      ui.alert(
        "Linting Complete",
        "All sheets have been linted. Please check for:\n" +
          "- BRIGHT RED cells with white text: CRITICAL translation mismatch (primary language values don't match source sheet)\n" +
          "- Light red cells (#FFC7CE): Invalid values, missing options, duplicate IDs, missing observation/track coverage, invalid Applies tokens, unsafe Metadata values, duplicate resolved locales, oversized SVGs/locale payloads, entity overflow\n" +
          "- Yellow highlighted cells (#FFF2CC): Required fields missing, unreferenced details, slug collisions, manual ID hygiene, ambiguous option colons, ignored options on text/number fields, translation source overwrites\n" +
          "- Light yellow cells (#FFFFCC): Advisory guidance (blank type defaults, type clarity, plain-text icon workflow warnings, icon ID normalization)\n" +
          "- Red/orange text in icon columns: Missing icons, Drive access issues, HTTP URLs, plain-text workflow warnings\n" +
          "- Icons sheet issues: Missing/duplicate IDs, unsupported formats, cross-sheet collisions\n" +
          "- Metadata validation: Unsafe characters in name/version/primaryLanguage\n\n" +
          "IMPORTANT: Bright red cells will cause translation failures. Re-sync translation sheets before generating config.\n" +
          "TIP: For icons, paste inline SVG from https://icons.earthdefenderstoolkit.com for best results. Plain text still works (auto lookup), but is less accurate.",
        ui.ButtonSet.OK,
      );
    }
  } catch (error) {
    console.error("Error during linting process:", error);

    // Only show error alert if showAlerts is true
    if (showAlerts) {
      const ui = SpreadsheetApp.getUi();
      ui.alert(
        "Linting Error",
        "An error occurred during the linting process: " +
          error.message +
          "\n\nSome sheets may not have been fully processed.",
        ui.ButtonSet.OK,
      );
    }
  }
}
